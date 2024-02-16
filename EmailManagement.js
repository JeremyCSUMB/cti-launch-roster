
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
function prepareGuidedSessionData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName('Spring 2024');
  var targetSheetName = 'Guided Session';
  var targetSheet = ss.getSheetByName(targetSheetName);
  
  // If the target sheet does not exist, create it
  if (!targetSheet) {
    targetSheet = ss.insertSheet(targetSheetName);
  } else {
    // Clear the existing data
    targetSheet.clear();
  }
  
  // Get the header row from the source sheet
  var headerRow = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  
  // Find the indexes for each column
  var lastNameIndex = headerRow.indexOf('Last Name') + 1;
  var firstNameIndex = headerRow.indexOf('First Name') + 1;
  var emailAddressIndex = headerRow.indexOf('Email Address') + 1;
  var sessionLocationIndex = headerRow.indexOf('Deep Work Session Location') + 1;
  
  // Validate if all columns are found
  if (lastNameIndex === 0 || firstNameIndex === 0 || emailAddressIndex === 0 || sessionLocationIndex === 0) {
    Logger.log("Error: One or more column names not found.");
    return; // Exit the function if any column name is not found
  }
  
  // Get the data from the source sheet
  var dataRange = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn());
  var data = dataRange.getValues();
  
  // Prepare the data for the target sheet
  var updatedData = data.map(function(row) {
    var location = row[sessionLocationIndex - 1]; // Adjust for 0-based index
    var roomNumber = '';
    switch (location) {
      case 'CSUMB':
        roomNumber = 'BIT RM 111';
        break;
      case 'Hartnell Alisal':
        roomNumber = 'AC 106';
        break;
      case 'CSUDH':
        roomNumber = 'NSM A 143';
        break;
      case 'ECC':
        roomNumber = 'MBA 103';
        break;
      default:
        roomNumber = 'NA'; // Default case set to 'NA'
    }
    // Return only the required columns and append the room number
    return [row[lastNameIndex - 1], row[firstNameIndex - 1], row[emailAddressIndex - 1], row[sessionLocationIndex - 1], roomNumber];
  });
  
  // Prepare the header for the target sheet and append it
  var header = ['Last Name', 'First Name', 'Email Address', 'Deep Work Session Location', 'Room Number'];
  targetSheet.appendRow(header);
  
  // Write the updated data to the target sheet
  if (updatedData.length > 0) {
    targetSheet.getRange(2, 1, updatedData.length, updatedData[0].length).setValues(updatedData);
  }
}