function sendDeepWorkReminders() {
  // Get the day of the week
  const today = new Date();
  const dayOfWeek = today.toLocaleDateString('en-US', { weekday: 'long' }); // e.g., 'Friday'

  // Access the correct sheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(dayOfWeek);

  // Get the email data
  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); // Remove header row

  // Iterate through each row and send emails
  data.forEach(row => {
    const firstName = row[headers.indexOf('First Name')];
    const sisLoginId = row[headers.indexOf('SIS Login ID')];
    const location = row[headers.indexOf('Deep Work Session Location')];
    const time = row[headers.indexOf('Deep Work Session Time')];
    const roomNumber = row[headers.indexOf('Deep Work Session Room Number')];

    // Construct the email message (Customize as needed)
    const subject = "Deep Work Session Reminder";
    const body = `Hello ${firstName},\n\nJust a friendly reminder about your Deep Work Session today:\n
               * Location: ${location}
               * Time: ${time}
               * Room Number: ${roomNumber}\n\nBest of luck with your focused work!`;

    // Send the email (Assume 'SIS Login ID' column contains email address)
    MailApp.sendEmail(sisLoginId, subject, body);
  });
}
