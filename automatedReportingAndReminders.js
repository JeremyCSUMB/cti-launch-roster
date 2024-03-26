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

function sendMissedDeepWorkReminders() {
  // Get the 'Missed Last Deep Work Session' sheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('Missed Last Deep Work Session');

  // Get the data from the sheet
  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); // Remove header row

  // Iterate through each row and send emails
  data.forEach(row => {
    const firstName = row[headers.indexOf('First Name')];
    const lastName = row[headers.indexOf('Last Name')];
    const sisLoginId = row[headers.indexOf('SIS Login ID')];
    const dateString = row[headers.indexOf('Date')];
    const date = new Date(dateString);
    const day = row[headers.indexOf('Day')];
    const location = row[headers.indexOf('Deep Work Session Location')];

    // Format the date as "3/13"
    const dateFormat = `${date.getMonth() + 1}/${date.getDate()}`;

    // Construct the email message
    const subject = `${firstName}, Quick Follow-Up on missed DW session ${day} ${dateFormat}`;
    const body = `Hi ${firstName},\n\nI hope everything is okay on your end. I was looking at the attendance record and it looks like you missed the recent Deep Work Session on ${day}, ${dateFormat}. These sessions are super beneficial – I want to make sure you're not falling behind!\n\nHere are the details that I saw:\nMissed Deep Work Session: ${day} ${dateFormat}\nDeep Work Session Location: ${location}\n\nIf this was a mistake, please let me know right away. As always, let me know if you have any questions.\n\nKindly,`;

    // Send the email
    MailApp.sendEmail(sisLoginId, subject, body);
  });
}