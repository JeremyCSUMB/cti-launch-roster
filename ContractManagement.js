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