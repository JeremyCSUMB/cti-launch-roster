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
    daysOfWeek.forEach(function (day) {
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
        [session1, session2].forEach(function (session) {
            if (session && daysOfWeek.includes(session.split(' ')[0])) {
                var daySheet = ss.getSheetByName(session.split(' ')[0]);
                var roomNumber = determineRoomNumber(location, session);
                var sessionTime = getSessionTime(session);
                daySheet.appendRow([lastName, firstName, sisLoginId, location, sessionTime, roomNumber]);
            }
        });
    }
}

