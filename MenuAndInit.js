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
        .addItem('Generate Not Completed Assignments', 'updateForIncompleteAssignment')
        .addToUi();
}