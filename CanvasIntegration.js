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