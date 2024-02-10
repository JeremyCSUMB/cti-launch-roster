
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
    var standardizedCanvasAssignments = canvasData[0].map(function(assignment) {
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

