
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

function updateForIncompleteAssignment() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi(); // For user interaction

  // Prompt for assignment name
  var response = ui.prompt('Enter the assignment name from the Canvas Gradebook:');
  var assignmentName = response.getResponseText().trim().replace(/\s+/g, ' ');
  if (!assignmentName) {
    ui.alert('No assignment name entered. Operation cancelled.');
    return;
  }

  // Access sheets
  var canvasGradebookSheet = ss.getSheetByName('Canvas Gradebook');
  var canvasData = canvasGradebookSheet.getDataRange().getValues();
  var idIndexCanvas = canvasData[0].indexOf('ID'); // Get the index of the 'ID' column

  // Standardize assignment names
  var standardizedCanvasAssignments = canvasData[0].map(function (assignment) {
    return assignment.trim().replace(/\s+/g, ' ');
  });

  var assignmentIndexCanvas = standardizedCanvasAssignments.indexOf(assignmentName);
  if (assignmentIndexCanvas === -1) {
    ui.alert('Assignment not found in Canvas Gradebook. Check name and try again.');
    return;
  }

  // Access Spring 2024 sheet and prepare new sheet
  var springSheet = ss.getSheetByName('Spring 2024');
  var springData = springSheet.getDataRange().getValues();
  var incompleteSheetName = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + ' - Incomplete ' + assignmentName;
  var incompleteSheet = ss.getSheetByName(incompleteSheetName) || ss.insertSheet(incompleteSheetName);
  incompleteSheet.clear(); // Clear existing data
  incompleteSheet.appendRow(['Canvas ID', 'First Name', 'Last Name', 'Email Address']); // Example headers

  // Iterate over Canvas Gradebook data to find and add students who have not completed the assignment
  for (var k = 1; k < canvasData.length; k++) {
    var studentId = canvasData[k][idIndexCanvas];
    var assignmentCompletion = canvasData[k][assignmentIndexCanvas];
    if (!assignmentCompletion) { // Check if student has not completed the assignment
      // Find student in Spring 2024 sheet using ID
      var studentRow = springData.find(row => row.includes(studentId));
      if (studentRow) {
        // Extract desired columns from Spring 2024 sheet for the student
        var canvasId = studentRow[springData[0].indexOf('Canvas ID')];
        var firstName = studentRow[springData[0].indexOf('First Name')];
        var lastName = studentRow[springData[0].indexOf('Last Name')];
        var email = studentRow[springData[0].indexOf('Email Address')];
        incompleteSheet.appendRow([canvasId, firstName, lastName, email]); // Add to new sheet
      }
    }
  }

  Logger.log('Incomplete assignment sheet created: ' + incompleteSheetName);
}

function generateProgressReport(expectedMilestone) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var springSheet = ss.getSheetByName('Spring 2024');
  var data = springSheet.getDataRange().getValues();

  var lastNameIndex = data[0].indexOf('Last Name');
  var firstNameIndex = data[0].indexOf('First Name');
  var sisLoginIdIndex = data[0].indexOf('SIS Login ID');
  var attendanceAverageIndex = data[0].indexOf('Attendance Average');
  var currentMilestoneIndex = data[0].indexOf('Current Milestone');

  var reportSheetName = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + ' - Weekly Progress Report';
  var reportSheet = ss.getSheetByName(reportSheetName) || ss.insertSheet(reportSheetName);
  reportSheet.clear();

  reportSheet.appendRow(['Last Name', 'First Name', 'SIS Login ID', 'Attendance Average', 'Current Milestone', 'At Expected Milestone', 'AI Feedback']);

  for (var i = 1; i < data.length; i++) {
    var lastName = data[i][lastNameIndex];
    var firstName = data[i][firstNameIndex];
    var sisLoginId = data[i][sisLoginIdIndex];
    var attendanceAverage = data[i][attendanceAverageIndex];
    var currentMilestone = data[i][currentMilestoneIndex];

    var atExpectedMilestone = currentMilestone >= expectedMilestone ? 'Yes' : 'No';
    var aiFeedback = ''; // Initialize AI feedback as empty
    var inputText = `As an educational mentor, provide a concise, encouraging feedback summary for a student in a course called 'CTI Launch'. The student has an attendance rate of ${attendanceAverage * 100}%, is currently at milestone ${currentMilestone}, and is expected to be at milestone ${expectedMilestone}. Offer guidance on how they can improve and reassure them of the support available to help them meet their goals, such as meeting with their SA's (Student Assistants). Keep the feedback to two sentences.`;
    aiFeedback = callOpenAI(inputText, 'gpt-4-0125-preview'); 

    // Log the AI feedback for demonstration
    Logger.log(`AI Feedback for ${firstName} ${lastName}: ${aiFeedback}`);

    reportSheet.appendRow([lastName, firstName, sisLoginId, attendanceAverage * 100, currentMilestone, atExpectedMilestone, aiFeedback]);
  }

  Logger.log('Weekly Progress Report generated: ' + reportSheetName);
}

function processProgressReport() {
  generateProgressReport(3);
}

function withCategorizeStudents() {
  categorizeStudents(3);
}
function categorizeStudents(atExpectedMilestone) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Spring 2024');
  const data = sheet.getDataRange().getValues();

  // Find column indexes
  const institutionIndex = data[0].indexOf('Institution');
  const attendanceIndex = data[0].indexOf('Attendance Average');
  const milestoneIndex = data[0].indexOf('Current Milestone');
  const categoryIndex = data[0].indexOf('Student Category') || data[0].length;
  const statusIndex = data[0].indexOf('Program Status') || data[0].length;

  // Add column headers if needed
  if (categoryIndex === data[0].length) sheet.getRange(1, categoryIndex + 1).setValue('Student Category');
  if (statusIndex === data[0].length) sheet.getRange(1, statusIndex + 1).setValue('Program Status');

  // Color definitions
  const highRiskColor = '#FFC0CB'; // Light pink
  const moderateRiskColor = '#FFE4B5'; // Light orange
  const lowRiskColor = '#90EE90'; // Light green

  // Categorize students 
  for (let i = 1; i < data.length; i++) {
    const institution = data[i][institutionIndex];
    const attendance = data[i][attendanceIndex];
    const milestone = data[i][milestoneIndex];

    let category, status;
    let expectedMilestone = atExpectedMilestone;

    // Adjust expectedMilestone for ECC institution
    if (institution === 'ECC') {
      expectedMilestone -= 1;
    }

    // Logging for clarity
    Logger.log("Student " + i + ": Institution: " + institution + ", Attendance: " + attendance + ", Milestone: " + milestone);

    if (milestone < expectedMilestone && attendance < 0.8) { 
      category = 'High Risk';
      status = 'Escalate';

    } else if ((milestone < expectedMilestone && attendance < 0.9) || 
               (milestone >= expectedMilestone && attendance < 0.8)) {
      category = 'Moderate Risk';
      status = 'Active'; 

    } else {
      category = 'Low Risk';
      status = 'Active';
    }

    Logger.log("Student " + i + " classified as: " + category);

    sheet.getRange(i + 1, categoryIndex + 1).setValue(category);
    sheet.getRange(i + 1, statusIndex + 1).setValue(status);

    // Apply colors
    const categoryCell = sheet.getRange(i + 1, categoryIndex + 1);
    const statusCell = sheet.getRange(i + 1, statusIndex + 1);

    switch (category) {
      case 'High Risk':
        categoryCell.setBackground(highRiskColor);
        statusCell.setBackground(highRiskColor);
        break;
      case 'Moderate Risk':
        categoryCell.setBackground(moderateRiskColor);
        statusCell.setBackground(moderateRiskColor);
        break;
      case 'Low Risk':
        categoryCell.setBackground(lowRiskColor);
        statusCell.setBackground(lowRiskColor);
        break;
    }
  }
}
