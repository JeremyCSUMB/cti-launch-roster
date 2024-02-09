
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