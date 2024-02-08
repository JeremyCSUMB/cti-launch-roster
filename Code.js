function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Code Executions')
    .addItem('Update Attendance', 'processAttendance')
    .addItem('Update Canvas Gradebook Sheet', 'getCanvasGradebook')
    .addItem('Update Current Milestone Column', 'updateCurrentMilestones')
    .addItem('Update Deep Work Session Location/Day', 'updateRosterFromExternalSheet')
    .addItem('Generate Not Signed Contract List', 'checkSignedContracts')
    .addItem('Generate Deep Work Session Assignments', 'createDeepWorkSessionAssignments')
    .addItem('Generate YAMM Email', 'generateYAMMEmail')
    .addToUi();
}
function checkSignedContracts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('Spring 2024');
  const data = sourceSheet.getDataRange().getValues();

  const headers = data[0];
  const firstNameIndex = headers.indexOf('First Name');
  const lastNameIndex = headers.indexOf('Last Name');
  const emailAddressIndex = headers.indexOf('Email Address');
  const signedContractIndex = headers.indexOf('Signed Contract');

  const unsignedStudents = data.filter((row, index) => {
    // Skip header row
    if (index === 0) return false;
    return !row[signedContractIndex];
  });

  const unsignedData = unsignedStudents.map(student => [
    student[firstNameIndex],
    student[lastNameIndex],
    student[emailAddressIndex]
  ]);

  const resultSheetName = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd") + " NOT SIGNED CONTRACT";
  let resultSheet = ss.getSheetByName(resultSheetName);
  if (!resultSheet) {
    resultSheet = ss.insertSheet(resultSheetName);
    resultSheet.appendRow(['First Name', 'Last Name', 'Email Address']);
  }

  if (unsignedData.length > 0) {
    resultSheet.getRange(resultSheet.getLastRow() + 1, 1, unsignedData.length, 3).setValues(unsignedData);
  }
}

function boilerPlate() {
  Logger.log("Hello, I am boilerplate!");
  return;
}

function getCanvasGradebook() {
  var courseId = 148;
  var courseTitle = 'Canvas Gradebook';
  var ctiCanvasGradebook = CanvasGradebook.getCanvasGradebook(courseId, courseTitle);
  CanvasGradebook.updateSpreadSheetView(courseTitle, ctiCanvasGradebook);
}
function updateCanvasIDs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var springSheet = ss.getSheetByName('Spring 2024');
  var canvasSheet = ss.getSheetByName('Canvas Gradebook');

  // Find the column indexes
  var springHeaders = springSheet.getRange(1, 1, 1, springSheet.getLastColumn()).getValues()[0];
  var canvasHeaders = canvasSheet.getRange(1, 1, 1, canvasSheet.getLastColumn()).getValues()[0];
  var sisLoginIdIndexSpring = springHeaders.indexOf('SIS Login ID') + 1;
  var canvasIdIndexSpring = springHeaders.indexOf('Canvas ID') + 1;
  var sisLoginIdIndexCanvas = canvasHeaders.indexOf('SIS Login ID') + 1;
  var idIndexCanvas = canvasHeaders.indexOf('ID') + 1;  // Adjusted to correctly identify the 'ID' column

  if (sisLoginIdIndexSpring < 1 || canvasIdIndexSpring < 1 || sisLoginIdIndexCanvas < 1 || idIndexCanvas < 1) {
    throw new Error('One or more required columns not found.');
  }

  // Extract data from both sheets
  var springData = springSheet.getRange(2, sisLoginIdIndexSpring, springSheet.getLastRow() - 1).getValues();
  var canvasData = canvasSheet.getRange(2, 1, canvasSheet.getLastRow() - 1, canvasSheet.getLastColumn()).getValues();

  // Create a map for faster lookup
  var canvasIdMap = new Map();
  for (var i = 0; i < canvasData.length; i++) {
    canvasIdMap.set(canvasData[i][sisLoginIdIndexCanvas - 1], canvasData[i][idIndexCanvas - 1]);
  }

  // Update 'Canvas ID' in 'Spring 2024' sheet
  for (var j = 0; j < springData.length; j++) {
    var sisLoginId = springData[j][0];
    if (canvasIdMap.has(sisLoginId)) {
      springSheet.getRange(j + 2, canvasIdIndexSpring).setValue(canvasIdMap.get(sisLoginId));
    }
  }
}
function processAttendance() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var attendanceSheet = ss.getSheetByName('Attendance');
  var springSheet = ss.getSheetByName('Spring 2024');

  var attendanceData = attendanceSheet.getRange(1, 1, attendanceSheet.getLastRow(), attendanceSheet.getLastColumn()).getDisplayValues();
  var springData = springSheet.getRange(1, 1, springSheet.getLastRow(), springSheet.getLastColumn()).getValues();

  var springHeader = springSheet.getRange(1, 1, 1, springSheet.getLastColumn()).getValues()[0];
  var emailCol = springHeader.indexOf('Email Address');
  var sisLoginIdCol = springHeader.indexOf('SIS Login ID');
  var alternateEmailCol = springHeader.indexOf('Alternate Email');
  var session1Col = springHeader.indexOf('Deep Work Session #1');
  var session2Col = springHeader.indexOf('Deep Work Session #2');
  var springInstitutionCol = springHeader.indexOf('Deep Work Session Location');

  var dateCol = attendanceData[0].indexOf('Date');
  var sessionTypeCol = attendanceData[0].indexOf('Session Type');
  var institutionCol = attendanceData[0].indexOf('Institution');
  var peardeckLinkCol = attendanceData[0].indexOf('Peardeck Link');
  var processedCol = attendanceData[0].indexOf('Processed');

  if (emailCol < 0 || sisLoginIdCol < 0 || alternateEmailCol < 0) {
    throw new Error('Required columns not found in Spring 2024 sheet.');
  }

  for (var i = 1; i < attendanceData.length; i++) {
    if (!attendanceData[i][processedCol]) {
      var attendanceInstitution = attendanceData[i][institutionCol];
      var sessionType = attendanceData[i][sessionTypeCol];
      var peardeckLink = attendanceData[i][peardeckLinkCol];
      var date = attendanceData[i][dateCol];

      var pearSheet = SpreadsheetApp.openByUrl(peardeckLink).getSheets()[0];
      var pearData = pearSheet.getDataRange().getValues();

      for (var j = 1; j < springData.length; j++) {
        var springInstitution = springData[j][springInstitutionCol];
        var session1 = springData[j][session1Col];
        var session2 = springData[j][session2Col];

        if (attendanceInstitution === springInstitution && (sessionType === session1 || sessionType === session2)) {
          var studentPresent = checkAttendance(pearData, springData[j], emailCol, sisLoginIdCol, alternateEmailCol);
          updateAttendanceRecords(springSheet, j, date, studentPresent);
        }
      }

      attendanceSheet.getRange(i + 1, processedCol + 1).setValue('Yes');
    }
  }
}

function checkAttendance(pearData, springRow, emailCol, sisLoginIdCol, alternateEmailCol) {
  var studentEmail = springRow[emailCol].toLowerCase();
  var studentSisLoginId = springRow[sisLoginIdCol].toLowerCase();
  var studentAlternateEmail = springRow[alternateEmailCol].toLowerCase();

  for (var i = 0; i < pearData.length; i++) {
    var pearEmail = (pearData[i][2] || "").toLowerCase(); // Assuming email is in the third column of Peardeck data
    if (pearEmail === studentEmail || pearEmail === studentSisLoginId || pearEmail === studentAlternateEmail) {
      return true;
    }
  }
  return false;
}

function updateAttendanceRecords(sheet, rowIndex, date, didAttend) {
  var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var lastAttendedCol = headerRow.indexOf('Last Attended Deep Work Session') + 1;
  var attendedHistoryCol = headerRow.indexOf('Attended Deep Work Session History') + 1;
  var missedHistoryCol = headerRow.indexOf('Missed Deep Work Session History') + 1;
  var totalAttendedCol = headerRow.indexOf('Total # of Attended DW Sessions') + 1;
  var totalMissedCol = headerRow.indexOf('Total # of Missed DW Sessions') + 1;

  if (lastAttendedCol <= 0 || attendedHistoryCol <= 0 || missedHistoryCol <= 0 || totalAttendedCol <= 0 || totalMissedCol <= 0) {
    throw new Error('Required columns not found.');
  }

  var historyCol = didAttend ? attendedHistoryCol : missedHistoryCol;
  var totalSessionsCol = didAttend ? totalAttendedCol : totalMissedCol;
  var historyUpdate = date + ' (' + (didAttend ? 'Attended' : 'Missed') + ')';
  var cellToUpdate = sheet.getRange(rowIndex + 1, historyCol);
  var currentHistory = cellToUpdate.getValue();
  var updatedHistory = currentHistory ? currentHistory + ', ' + historyUpdate : historyUpdate;
  cellToUpdate.setValue(updatedHistory);

  var totalSessionsCell = sheet.getRange(rowIndex + 1, totalSessionsCol);
  var currentTotal = totalSessionsCell.getValue() || 0;
  totalSessionsCell.setValue(currentTotal + 1);

  if (didAttend) {
    var lastAttendedCell = sheet.getRange(rowIndex + 1, lastAttendedCol);
    lastAttendedCell.setValue(date);
  }
}

function generateYAMMEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var springSheet = ss.getSheetByName('Spring 2024');

  // Retrieve headers and find column indices
  var headers = springSheet.getRange(1, 1, 1, springSheet.getLastColumn()).getValues()[0];
  var firstNameCol = headers.indexOf('First Name') + 1;
  var lastNameCol = headers.indexOf('Last Name') + 1;
  var sisLoginIDCol = headers.indexOf('Email Address') + 1;

  // Verify that columns exist
  if (firstNameCol < 1 || lastNameCol < 1 || sisLoginIDCol < 1) {
    throw new Error('One or more required columns not found in Spring 2024 sheet.');
  }

  // Extract data from the Spring 2024 sheet
  var data = springSheet.getRange(2, 1, springSheet.getLastRow() - 1, springSheet.getLastColumn()).getValues();

  // Create a new sheet for YAMM within the same spreadsheet
  var yammSheetName = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + ' - YAMM';
  var yammSheet = ss.getSheetByName(yammSheetName) || ss.insertSheet(yammSheetName);

  // Set headers for the new sheet
  yammSheet.appendRow(['First Name', 'Last Name', 'Email Address']);

  // Process and transfer data to the new sheet
  data.forEach(function (row) {
    var firstName = row[firstNameCol - 1];
    var lastName = row[lastNameCol - 1];
    var sisLoginID = row[sisLoginIDCol - 1];
    if (firstName && lastName && sisLoginID) { // Ensure all fields are present
      yammSheet.appendRow([firstName, lastName, sisLoginID]);
    }
  });

  // Log a message to indicate completion
  Logger.log('YAMM data sheet created: ' + yammSheetName);
}
function updateCurrentMilestones() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Access the necessary sheets
  var springSheet = ss.getSheetByName('Spring 2024');
  var canvasGradebookSheet = ss.getSheetByName('Canvas Gradebook');
  var milestoneListSheet = ss.getSheetByName('Milestone List');

  // Get data from sheets
  var springData = springSheet.getDataRange().getValues();
  var canvasData = canvasGradebookSheet.getDataRange().getValues();
  var milestoneData = milestoneListSheet.getDataRange().getValues();

  // Find column indexes in Spring 2024 and Canvas Gradebook sheets
  var canvasIdIndexSpring = springData[0].indexOf('Canvas ID');
  var currentMilestoneIndexSpring = springData[0].indexOf('Current Milestone');
  var idIndexCanvas = canvasData[0].indexOf('ID');

  // Create a map for Canvas IDs to row numbers in Spring 2024 sheet
  var canvasIdToRowMap = {};
  for (var i = 1; i < springData.length; i++) {
    var canvasId = springData[i][canvasIdIndexSpring];
    canvasIdToRowMap[canvasId] = i;
  }

  // Standardize assignment names in Canvas Gradebook
  var standardizedCanvasAssignments = canvasData[0].map(function (assignment) {
    return assignment.trim().replace(/\s+/g, ' ');
  });

  // Iterate over each milestone
  for (var j = 1; j < milestoneData.length; j++) {
    var milestone = milestoneData[j][0]; // Milestone number
    var assignment = milestoneData[j][1].trim().replace(/\s+/g, ' '); // Standardize milestone assignment name

    // Find column index for this assignment in Canvas Gradebook
    var assignmentIndexCanvas = standardizedCanvasAssignments.indexOf(assignment);
    if (assignmentIndexCanvas == -1) continue; // Skip if assignment not found in Canvas Gradebook

    // Check each student's completion status for this assignment
    for (var k = 1; k < canvasData.length; k++) {
      var studentId = canvasData[k][idIndexCanvas];
      var assignmentCompletion = canvasData[k][assignmentIndexCanvas];

      // Update the current milestone in Spring 2024 sheet if the student completed the assignment
      if (assignmentCompletion && canvasIdToRowMap.hasOwnProperty(studentId)) {
        var springRow = canvasIdToRowMap[studentId];
        springSheet.getRange(springRow + 1, currentMilestoneIndexSpring + 1).setValue(milestone);
      }
    }
  }
}

function generateMilestoneReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var springSheet = ss.getSheetByName('Spring 2024');

  // Get data from the Spring 2024 sheet
  var springData = springSheet.getDataRange().getValues();

  // Find column indexes for First Name, Last Name, Email Address, and Current Milestone
  var firstNameIndex = springData[0].indexOf('First Name');
  var lastNameIndex = springData[0].indexOf('Last Name');
  var emailIndex = springData[0].indexOf('Email Address');
  var currentMilestoneIndex = springData[0].indexOf('Current Milestone');

  // Generate a new sheet titled with the current date and 'Milestone Report'
  var reportSheetName = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd') + ' Milestone Report';
  var reportSheet = ss.getSheetByName(reportSheetName) || ss.insertSheet(reportSheetName);

  // Set headers for the report sheet
  reportSheet.appendRow(['First Name', 'Last Name', 'Email Address', 'Current Milestone']);

  // Process each student and populate the report
  for (var i = 1; i < springData.length; i++) {
    var firstName = springData[i][firstNameIndex];
    var lastName = springData[i][lastNameIndex];
    var email = springData[i][emailIndex];
    var currentMilestone = springData[i][currentMilestoneIndex];

    // Check for null, undefined, or empty string (after trimming) in the current milestone cell
    var milestoneStatus = (currentMilestone === null || currentMilestone === undefined || String(currentMilestone).trim() === '')
      ? 'MILESTONE 0 NOT COMPLETED'
      : currentMilestone;

    // Append row to the report sheet
    reportSheet.appendRow([firstName, lastName, email, milestoneStatus]);
  }
}

function createUserList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var list = ss.getSheetByName('Spring 2024');

  var data = list.getRange(1, 1, list.getLastRow(), list.getLastColumn()).getValues();
  var first_col = data[0].indexOf("First Name");
  var last_col = data[0].indexOf("Last Name");
  var email_col = data[0].indexOf("Email Address");
  var canvas_id_col = data[0].indexOf("Canvas ID"); // Add this line to get the Canvas ID column index

  var output = ss.getSheetByName('Canvas Import CSV New');
  output.clearContents();
  output.appendRow(['user_id', 'integration_id', 'login_id', 'password', 'first_name', 'last_name', 'full_name', 'sortable_name', 'short_name', 'email', 'status']);

  for (var i = 0; i < data.length; i++) {
    var first = data[i][first_col];
    var last = data[i][last_col];
    var email = data[i][email_col];
    var canvas_id = data[i][canvas_id_col]; // Add this line to get the Canvas ID value

    // Check if Canvas ID is blank before appending the row
    if (canvas_id === '') {
      output.appendRow([email, '', email, 'LAUNCH123456', first, last, '', '', '', email, 'active']);
    }
  }
}

function checkSignedContracts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('Spring 2024');
  const data = sourceSheet.getDataRange().getValues();

  const headers = data[0];
  const firstNameIndex = headers.indexOf('First Name');
  const lastNameIndex = headers.indexOf('Last Name');
  const emailAddressIndex = headers.indexOf('Email Address');
  const signedContractIndex = headers.indexOf('Signed Contract');

  const unsignedStudents = data.filter((row, index) => {
    // Skip header row
    if (index === 0) return false;
    return !row[signedContractIndex];
  });

  const unsignedData = unsignedStudents.map(student => [
    student[firstNameIndex],
    student[lastNameIndex],
    student[emailAddressIndex]
  ]);

  const resultSheetName = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd") + " NOT SIGNED CONTRACT";
  let resultSheet = ss.getSheetByName(resultSheetName);
  if (!resultSheet) {
    resultSheet = ss.insertSheet(resultSheetName);
    resultSheet.appendRow(['First Name', 'Last Name', 'Email Address']);
  }

  if (unsignedData.length > 0) {
    resultSheet.getRange(resultSheet.getLastRow() + 1, 1, unsignedData.length, 3).setValues(unsignedData);
  }
}
function updateRosterFromExternalSheet() {
  var ui = SpreadsheetApp.getUi();
  var url = 'https://docs.google.com/spreadsheets/d/17ZFuo96ydKhFZz74Iz-etPj6eWY_NJiyZeQ1J2slyTM/edit?usp=sharing';
  Logger.log("Opening external sheet: " + url);
  var externalSheet = SpreadsheetApp.openByUrl(url);
  var externalData = externalSheet.getSheets()[0].getDataRange().getValues();

  var rosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Spring 2024');
  Logger.log("Accessing roster sheet: 'Spring 2024 Roster'");
  var rosterData = rosterSheet.getDataRange().getValues();

  var emailIndexExt = externalData[0].indexOf('Email Address');
  var fullNameIndexExt = externalData[0].indexOf('Please enter your full name:');
  var session1IndexExt = externalData[0].indexOf('Deep Work Session #1');
  var session2IndexExt = externalData[0].indexOf('Deep Work Session #2');
  var deepWorkLocationIndexExt = externalData[0].indexOf('Please select the institution where you will be joining deep work sessions.');

  var emailIndexRoster = rosterData[0].indexOf('Email Address');
  var firstNameIndexRoster = rosterData[0].indexOf('First Name');
  var lastNameIndexRoster = rosterData[0].indexOf('Last Name');
  var session1IndexRoster = rosterData[0].indexOf('Deep Work Session #1');
  var session2IndexRoster = rosterData[0].indexOf('Deep Work Session #2');
  var deepWorkLocationIndexRoster = rosterData[0].indexOf('Deep Work Session Location');
  var alternateEmailIndexRoster = rosterData[0].indexOf('Alternate Email');

  Logger.log("Column indices (External): Email - " + emailIndexExt + ", Name - " + fullNameIndexExt + ", Session 1 - " + session1IndexExt + ", Session 2 - " + session2IndexExt + ", Location - " + deepWorkLocationIndexExt);
  Logger.log("Column indices (Roster): Email - " + emailIndexRoster + ", First Name - " + firstNameIndexRoster + ", Last Name - " + lastNameIndexRoster + ", Session 1 - " + session1IndexRoster + ", Session 2 - " + session2IndexRoster + ", Location - " + deepWorkLocationIndexRoster);

  for (var i = 1; i < externalData.length; i++) {
    var emailExt = externalData[i][emailIndexExt] || '';
    var fullNameExt = externalData[i][fullNameIndexExt] || '';
    var session1Response = externalData[i][session1IndexExt] || '';
    var session2Response = externalData[i][session2IndexExt] || '';
    var deepWorkLocation = externalData[i][deepWorkLocationIndexExt] || '';

    var session1Day = extractDayOfWeek(session1Response);
    var session2Day = extractDayOfWeek(session2Response);

    Logger.log("Processing: Email - " + emailExt + ", Name - " + fullNameExt + ", Session 1 - " + session1Day + ", Session 2 - " + session2Day + ", Location - " + deepWorkLocation);

    var foundMatch = false;
    for (var j = 1; j < rosterData.length; j++) {
      var rosterEmail = rosterData[j][emailIndexRoster] || '';
      if (rosterEmail.trim().toLowerCase() === emailExt.trim().toLowerCase()) {
        rosterSheet.getRange(j + 1, session1IndexRoster + 1).setValue(session1Day);
        rosterSheet.getRange(j + 1, session2IndexRoster + 1).setValue(session2Day);
        rosterSheet.getRange(j + 1, deepWorkLocationIndexRoster + 1).setValue(deepWorkLocation);
        rosterSheet.getRange(j + 1, alternateEmailIndexRoster + 1).setValue(emailExt);
        Logger.log("Updated Session for Email: " + emailExt + " - Session 1: " + session1Day + ", Session 2: " + session2Day + ", Location: " + deepWorkLocation + ", Alternate Email: " + emailExt);
        foundMatch = true;
        break;
      }
    }

    if (!foundMatch) {
      Logger.log("No direct match found for " + fullNameExt + " (" + emailExt + "). Trying alternate email match.");
      for (var k = 1; k < rosterData.length; k++) {
        var alternateEmail = rosterData[k][alternateEmailIndexRoster] || '';
        if (alternateEmail.trim().toLowerCase() === emailExt.trim().toLowerCase()) {
          rosterSheet.getRange(k + 1, session1IndexRoster + 1).setValue(session1Day);
          rosterSheet.getRange(k + 1, session2IndexRoster + 1).setValue(session2Day);
          rosterSheet.getRange(k + 1, deepWorkLocationIndexRoster + 1).setValue(deepWorkLocation);
          Logger.log("Updated Session for Alternate Email: " + emailExt + " - Session 1: " + session1Day + ", Session 2: " + session2Day + ", Location: " + deepWorkLocation);
          foundMatch = true;
          break;
        }
      }

      if (!foundMatch) {
        Logger.log("No match found in Alternate Email for " + fullNameExt + " (" + emailExt + ")");
        var response = ui.prompt('Enter the email of the student to find the row for manual update:');
        var inputEmail = response.getResponseText().trim().toLowerCase();

        for (var l = 1; l < rosterData.length; l++) {
          if (rosterData[l][emailIndexRoster].trim().toLowerCase() === inputEmail) {
            rosterSheet.getRange(l + 1, session1IndexRoster + 1).setValue(session1Day);
            rosterSheet.getRange(l + 1, session2IndexRoster + 1).setValue(session2Day);
            rosterSheet.getRange(l + 1, deepWorkLocationIndexRoster + 1).setValue(deepWorkLocation);
            rosterSheet.getRange(l + 1, alternateEmailIndexRoster + 1).setValue(emailExt);
            Logger.log("Manually updated row " + l + " for Email: " + inputEmail + " - Session 1: " + session1Day + ", Session 2: " + session2Day + ", Location: " + deepWorkLocation + ", Alternate Email: " + emailExt);
            break;
          }
        }
      }
    }
  }

  SpreadsheetApp.flush(); // Force the spreadsheet to update
  Logger.log("Function execution completed.");
}

function extractDayOfWeek(response) {
  var daysOfWeek = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
  for (var i = 0; i < daysOfWeek.length; i++) {
    if (response.includes(daysOfWeek[i])) {
      Logger.log("Extracted Day: " + daysOfWeek[i] + " from Response: " + response);
      return daysOfWeek[i];
    }
  }
  Logger.log("No Day of the Week Found in Response: " + response);
  return '';
}

function createDeepWorkSessionAssignments() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getSheetByName('Spring 2024');
  var sourceData = sourceSheet.getDataRange().getValues();

  var assignmentSheetName = 'Deep Work Session Assignments';
  var assignmentSheet = spreadsheet.getSheetByName(assignmentSheetName);
  if (!assignmentSheet) {
    assignmentSheet = spreadsheet.insertSheet(assignmentSheetName);
  } else {
    assignmentSheet.clear(); // Clear existing data
  }

  // Define column indices from the source sheet
  var firstNameIndex = sourceData[0].indexOf('First Name');
  var lastNameIndex = sourceData[0].indexOf('Last Name');
  var emailIndex = sourceData[0].indexOf('Email Address');
  var locationIndex = sourceData[0].indexOf('Deep Work Session Location');
  var session1Index = sourceData[0].indexOf('Deep Work Session #1');
  var session2Index = sourceData[0].indexOf('Deep Work Session #2');

  // Headers for the assignment sheet
  var headers = ['First Name', 'Last Name', 'Email Address', 'Deep Work Session Location', 'Deep Work Session #1', 'Deep Work Session #1 Room Number', 'Deep Work Session #2', 'Deep Work Session #2 Room Number'];
  assignmentSheet.appendRow(headers);

  for (var i = 1; i < sourceData.length; i++) {
    var row = sourceData[i];
    var firstName = row[firstNameIndex];
    var lastName = row[lastNameIndex];
    var email = row[emailIndex];
    var location = row[locationIndex];
    var session1 = row[session1Index] + ' ' + getSessionTime(row[session1Index]);
    var session2 = row[session2Index] + ' ' + getSessionTime(row[session2Index]);
    var session1Room = determineRoomNumber(location, row[session1Index]);
    var session2Room = determineRoomNumber(location, row[session2Index]);

    assignmentSheet.appendRow([firstName, lastName, email, location, session1, session1Room, session2, session2Room]);
  }
}

function determineRoomNumber(location, session) {
  // Define the room number logic based on the location and session
  if (location === 'CSUMB') {
    return mapCSUMBSessionsToRooms(session);
  } else if (location === 'Hartnell Alisal') {
    return mapHartnellAlisalSessionsToRooms(session);
  } else if (location === 'CSUDH') {
    return 'NSM A-143';
  } else if (location === 'ECC') {
    return (session.includes('Monday') || session.includes('Wednesday') || session.includes('Thursday')) ? 'MBA 103' : 'MBA 111';
  } else {
    return 'Room Not Found';
  }
}

function mapCSUMBSessionsToRooms(session) {
  // Map sessions to room numbers for CSUMB
  var sessionToRoomMap = {
    'Monday': '506-111',
    'Tuesday': '504-1301',
    'Wednesday': '506-223',
    'Thursday': '506-108',
    'Friday': '506-108'
  };
  return sessionToRoomMap[session.split(' ')[0]] || 'Room Not Found';
}

function mapHartnellAlisalSessionsToRooms(session) {
  // Map sessions to room numbers for Hartnell Alisal
  var sessionToRoomMap = {
    'Monday': 'AC C108',
    'Tuesday': 'AC C106',
    'Wednesday': 'AC A114',
    'Thursday': 'AC C110',
    'Friday': 'AC C106'
  };
  return sessionToRoomMap[session.split(' ')[0]] || 'Room Not Found';
}

function getSessionTime(session) {
  // Extract the day from the session string
  var day = session.split(' ')[0];

  // Define session times for all days, common across all institutions
  var sessionTimeMap = {
    'Monday': '12-2pm',
    'Tuesday': '2-4pm',
    'Wednesday': '4-6pm',
    'Thursday': '6-8pm',
    'Friday': '2-4pm'
  };

  return sessionTimeMap[day] || '';
}

function logCurrentDayOfWeek() {
  var daysOfWeek = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  var currentDate = new Date(); // Gets the current date and time
  var dayOfWeekIndex = currentDate.getDay(); // Gets the day of the week as a number (0-6)
  var dayOfWeekName = daysOfWeek[dayOfWeekIndex]; // Maps the number to the corresponding day of the week
  
  Logger.log(dayOfWeekName); // Logs the day of the week
}


function assignStudentAssistants() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var rosterSheet = spreadsheet.getSheetByName("Spring 2024");
  var dataRange = rosterSheet.getDataRange();
  var data = dataRange.getValues();
  var headers = data[0];
  
  // Dynamic column identification
  var session1Index = headers.indexOf('Deep Work Session #1') + 1;
  var session2Index = headers.indexOf('Deep Work Session #2') + 1;
  var studentAssistantIndex = headers.indexOf('Student Assistant') + 1;
  var agGroupIndex = headers.indexOf('AG Group') + 1;
  var sessionLocationIndex = headers.indexOf('Deep Work Session Location') + 1;
  
  // Updated SA availability based on days of the week, including CSUMB SAs
  var saAvailability = {
    "Alexis Guzman": ["Tuesday", "Thursday", "Friday"],
    "Haider Syed": ["Monday", "Wednesday", "Friday"],
    "Rodrigo Hernandez": ["Tuesday", "Friday"],
    "Elizabeth Barco Lopez": ["Wednesday", "Friday"],
    "Nishat Nawshin": ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"],
    "Geovanny Martinez": ["Monday", "Tuesday", "Thursday", "Friday"],
    "Aileen Dong": ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"],
    "Zachery Rouzaud": ["Wednesday", "Friday"],
    "Nicolas Garcia": ["Monday", "Thursday", "Friday"],
    "Sebastian Santoyo": ["Tuesday", "Wednesday"],
    "Mariana Duran": ["Tuesday","Friday"],
    "Jesus Garcia": ["Monday", "Thursday"]
  };

  // SA institutions mapping, including CSUMB
  var saInstitutions = {
    "Alexis Guzman": "Hartnell Alisal",
    "Haider Syed": "Hartnell Alisal",
    "Rodrigo Hernandez": "Hartnell Alisal",
    "Elizabeth Barco Lopez": "CSUDH",
    "Nishat Nawshin": "CSUDH",
    "Geovanny Martinez": "CSUDH",
    "Aileen Dong": "ECC",
    "Zachery Rouzaud": "ECC",
    "Nicolas Garcia": "CSUMB",
    "Sebastian Santoyo": "CSUMB",
    "Mariana Duran": "CSUMB",
    "Jesus Garcia": "CSUMB"
  };

  // Assign fixed AG Group numbers to each SA
  var saAgGroupNumbers = {
    "Alexis Guzman": 1,
    "Haider Syed": 2,
    "Rodrigo Hernandez": 3,
    "Elizabeth Barco Lopez": 4,
    "Nishat Nawshin": 5,
    "Geovanny Martinez": 6,
    "Aileen Dong": 7,
    "Zachery Rouzaud": 8,
    "Nicolas Garcia": 9,
    "Sebastian Santoyo": 10,
    "Mariana Duran": 11,
    "Jesus Garcia": 12
  };

  // Initialize counts for equitable distribution (if needed for other logic)
  var saCounts = {};
  Object.keys(saAvailability).forEach(function(sa) {
    saCounts[sa] = 0;
  });

  // Assign SAs to students based on session preferences and institution match
  for (var i = 1; i < data.length; i++) { // Skip header row
    var session1 = data[i][session1Index - 1];
    var session2 = data[i][session2Index - 1];
    var sessionLocation = data[i][sessionLocationIndex - 1];
    
    var assignedSA = "";
    var minCount = Number.MAX_SAFE_INTEGER;
    
    Object.keys(saAvailability).forEach(function(sa) {
      var saDays = saAvailability[sa];
      var saInstitution = saInstitutions[sa];
      if ((saDays.includes(session1) || saDays.includes(session2)) && saInstitution === sessionLocation && saCounts[sa] < minCount) {
        assignedSA = sa;
        minCount = saCounts[sa];
      }
    });
    
    // Update the spreadsheet if an SA is assigned
    if (assignedSA !== "") {
      rosterSheet.getRange(i + 1, studentAssistantIndex).setValue(assignedSA); // +1 to adjust for header and zero-based index
      saCounts[assignedSA] += 1; // Increment SA's count if needed for other logic
      
      // Assign the fixed AG Group number based on the SA
      var agGroup = saAgGroupNumbers[assignedSA];
      rosterSheet.getRange(i + 1, agGroupIndex).setValue(agGroup);
    }
  }
  
  // Apply changes to the spreadsheet
  SpreadsheetApp.flush();
}

function processDeepWorkSessions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName('Spring 2024');
  var dataRange = sourceSheet.getDataRange();
  var values = dataRange.getValues();

  // Find column indices dynamically
  var headers = values[0];
  var lastNameIndex = headers.indexOf('Last Name');
  var firstNameIndex = headers.indexOf('First Name');
  var sisLoginIdIndex = headers.indexOf('SIS Login ID');
  var locationIndex = headers.indexOf('Deep Work Session Location');
  var session1Index = headers.indexOf('Deep Work Session #1');
  var session2Index = headers.indexOf('Deep Work Session #2');

  // Headers for the new sheets
  var newSheetHeaders = ['Last Name', 'First Name', 'SIS Login ID', 'Deep Work Session Location', 'Deep Work Session Time', 'Deep Work Session Room Number'];

  // Create the sheets for each day of the week
  var daysOfWeek = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
  daysOfWeek.forEach(function(day) {
    if (!ss.getSheetByName(day)) {
      var sheet = ss.insertSheet(day);
      sheet.appendRow(newSheetHeaders);
    }
  });

  // Process each row (student)
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var lastName = row[lastNameIndex];
    var firstName = row[firstNameIndex];
    var sisLoginId = row[sisLoginIdIndex];
    var location = row[locationIndex];
    var session1 = row[session1Index];
    var session2 = row[session2Index];

    // Process each session
    [session1, session2].forEach(function(session) {
      if (session && daysOfWeek.includes(session.split(' ')[0])) {
        var daySheet = ss.getSheetByName(session.split(' ')[0]);
        var roomNumber = determineRoomNumber(location, session);
        var sessionTime = getSessionTime(session);
        daySheet.appendRow([lastName, firstName, sisLoginId, location, sessionTime, roomNumber]);
      }
    });
  }
}

