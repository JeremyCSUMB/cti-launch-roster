function processAttendance() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var attendanceSheet = ss.getSheetByName('Attendance');
    var springSheet = ss.getSheetByName('Spring 2024');

    var attendanceData = attendanceSheet.getRange(1, 1, attendanceSheet.getLastRow(), attendanceSheet.getLastColumn()).getDisplayValues();
    var springData = springSheet.getRange(1, 1, springSheet.getLastRow(), springSheet.getLastColumn()).getValues();

    var springHeader = springSheet.getRange(1, 1, 1, springSheet.getLastColumn()).getValues()[0];
    var emailCol = springHeader.indexOf('Email Address');
    var sisLoginIdCol = springHeader.indexOf('SIS Login ID');
    var alternateEmailCol = springHeader.indexOf('Alternate Email');
    var session1Col = springHeader.indexOf('Deep Work Session #1');
    var session2Col = springHeader.indexOf('Deep Work Session #2');
    var springInstitutionCol = springHeader.indexOf('Deep Work Session Location');

    var dateCol = attendanceData[0].indexOf('Date');
    var sessionTypeCol = attendanceData[0].indexOf('Session Type');
    var institutionCol = attendanceData[0].indexOf('Institution');
    var peardeckLinkCol = attendanceData[0].indexOf('Peardeck Link');
    var processedCol = attendanceData[0].indexOf('Processed');

    if (emailCol < 0 || sisLoginIdCol < 0 || alternateEmailCol < 0) {
        throw new Error('Required columns not found in Spring 2024 sheet.');
    }

    for (var i = 1; i < attendanceData.length; i++) {
        if (!attendanceData[i][processedCol]) {
            var attendanceInstitution = attendanceData[i][institutionCol];
            var sessionType = attendanceData[i][sessionTypeCol];
            var peardeckLink = attendanceData[i][peardeckLinkCol];
            var date = attendanceData[i][dateCol];

            var pearSheet = SpreadsheetApp.openByUrl(peardeckLink).getSheets()[0];
            var pearData = pearSheet.getDataRange().getValues();

            for (var j = 1; j < springData.length; j++) {
                var springInstitution = springData[j][springInstitutionCol];
                var session1 = springData[j][session1Col];
                var session2 = springData[j][session2Col];

                if (attendanceInstitution === springInstitution && (sessionType === session1 || sessionType === session2)) {
                    var studentPresent = checkAttendance(pearData, springData[j], emailCol, sisLoginIdCol, alternateEmailCol);
                    updateAttendanceRecords(springSheet, j, date, studentPresent);
                }
            }

            attendanceSheet.getRange(i + 1, processedCol + 1).setValue('Yes');
        }
    }
}

function checkAttendance(pearData, springRow, emailCol, sisLoginIdCol, alternateEmailCol) {
    var studentEmail = springRow[emailCol].toLowerCase();
    var studentSisLoginId = springRow[sisLoginIdCol].toLowerCase();
    var studentAlternateEmail = springRow[alternateEmailCol].toLowerCase();

    for (var i = 0; i < pearData.length; i++) {
        var pearEmail = (pearData[i][2] || "").toLowerCase(); // Assuming email is in the third column of Peardeck data
        if (pearEmail === studentEmail || pearEmail === studentSisLoginId || pearEmail === studentAlternateEmail) {
            return true;
        }
    }
    return false;
}
function updateAttendanceRecords(sheet, rowIndex, date, didAttend) {
    var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var lastAttendedCol = headerRow.indexOf('Last Attended Deep Work Session') + 1;
    var attendedHistoryCol = headerRow.indexOf('Attended Deep Work Session History') + 1;
    var missedHistoryCol = headerRow.indexOf('Missed Deep Work Session History') + 1;
    var totalAttendedCol = headerRow.indexOf('Total # of Attended DW Sessions') + 1;
    var totalMissedCol = headerRow.indexOf('Total # of Missed DW Sessions') + 1;
    var firstNameCol = headerRow.indexOf('First Name') + 1;
    var lastNameCol = headerRow.indexOf('Last Name') + 1;
    var sisLoginIdCol = headerRow.indexOf('SIS Login ID') + 1;
    var sessionLocationCol = headerRow.indexOf('Deep Work Session Location') + 1;

    if (lastAttendedCol <= 0 || attendedHistoryCol <= 0 || missedHistoryCol <= 0 || totalAttendedCol <= 0 || totalMissedCol <= 0 || firstNameCol <= 0 || lastNameCol <= 0 || sisLoginIdCol <= 0 || sessionLocationCol <= 0) {
        throw new Error('Required columns not found.');
    }

    var missedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Missed Last Deep Work Session');
    if (!missedSheet) {
        missedSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Missed Last Deep Work Session');
        missedSheet.appendRow(['First Name', 'Last Name', 'SIS Login ID', 'Date', 'Day', 'Deep Work Session Location']);
    } else {
        var lastRowDate;
        if (missedSheet.getLastRow() > 1) {
            lastRowDate = missedSheet.getRange(2, 4, missedSheet.getLastRow() - 1, 1).getValues().flat().pop();
        }
        var lastRowDateObj = lastRowDate ? new Date(lastRowDate) : null;
        var currentDateObj = new Date(date);
        if (lastRowDateObj && lastRowDateObj.getDate() !== currentDateObj.getDate()) {
            missedSheet.getDataRange().clearContent();
        }
    }

    var historyCol = didAttend ? attendedHistoryCol : missedHistoryCol;
    var totalSessionsCol = didAttend ? totalAttendedCol : totalMissedCol;
    var historyUpdate = date + ' (' + (didAttend ? 'Attended' : 'Missed') + ')';
    var cellToUpdate = sheet.getRange(rowIndex + 1, historyCol);
    var currentHistory = cellToUpdate.getValue();
    var updatedHistory = currentHistory ? currentHistory + ', ' + historyUpdate : historyUpdate;
    cellToUpdate.setValue(updatedHistory);

    var totalSessionsCell = sheet.getRange(rowIndex + 1, totalSessionsCol);
    var currentTotal = totalSessionsCell.getValue() || 0;
    totalSessionsCell.setValue(currentTotal + 1);

    if (didAttend) {
        var lastAttendedCell = sheet.getRange(rowIndex + 1, lastAttendedCol);
        lastAttendedCell.setValue(date);
    } else {
        var firstName = sheet.getRange(rowIndex + 1, firstNameCol).getValue();
        var lastName = sheet.getRange(rowIndex + 1, lastNameCol).getValue();
        var sisLoginId = sheet.getRange(rowIndex + 1, sisLoginIdCol).getValue();
        var sessionLocation = sheet.getRange(rowIndex + 1, sessionLocationCol).getValue();
        var dateFormatted = Utilities.formatDate(new Date(date), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'MM/dd');
        var dayFormatted = Utilities.formatDate(new Date(date), SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'EEE');
        missedSheet.appendRow([firstName, lastName, sisLoginId, dateFormatted, dayFormatted, sessionLocation]);
    }
}
function updateAttendanceAverage() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Spring 2024');
  var data = sheet.getDataRange().getValues(); 
  
  // Define column headers (modify as needed)
  var attendedHeader = 'Total # of Attended DW Sessions';
  var missedHeader = 'Total # of Missed DW Sessions';
  var averageHeader = 'Attendance Average';

  // Get column indices directly
  var attendedIndex = data[0].indexOf(attendedHeader);
  var missedIndex = data[0].indexOf(missedHeader);
  var averageIndex = data[0].indexOf(averageHeader);

  // Process the data 
  for (var i = 1; i < data.length; i++) {
    var attended = data[i][attendedIndex]; 
    var missed = data[i][missedIndex]; 

    var average;
    if (attended == 0 && missed == 0) {
      average = '';  // Both are 0
    } else {
      // Calculate percentage: (attended / (attended + missed)) * 100
      average = (attended / (attended + missed)); 
    } 

    // Update with percentage formatting
    sheet.getRange(i + 1, averageIndex + 1).setValue(average).setNumberFormat('0.00%'); 
  }
}
function processGuidedSessionAttendance() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var attendanceSheet = ss.getSheetByName('Attendance');
    var springSheet = ss.getSheetByName('Spring 2024');

    var attendanceData = attendanceSheet.getRange(1, 1, attendanceSheet.getLastRow(), attendanceSheet.getLastColumn()).getDisplayValues();

    var dateCol = attendanceData[0].indexOf('Date');
    var sessionTypeCol = attendanceData[0].indexOf('Session Type');
    var peardeckLinkCol = attendanceData[0].indexOf('Peardeck Link');
    var processedCol = attendanceData[0].indexOf('Processed');

    var springData = springSheet.getDataRange().getValues();
    var emailCol = springData[0].indexOf('Email Address');
    var sisLoginIdCol = springData[0].indexOf('SIS Login ID');
    var alternateEmailCol = springData[0].indexOf('Alternate Email');
    var lastAttendedGuidedSessionCol = springData[0].indexOf('Last Attended Guided Session') + 1;
    var totalAttendedGuidedSessionsCol = springData[0].indexOf('Guided Sessions Attendance Average') + 1;

    if (emailCol < 0 || sisLoginIdCol < 0 || alternateEmailCol < 0 || lastAttendedGuidedSessionCol <= 0 || totalAttendedGuidedSessionsCol <= 0) {
        throw new Error('Required columns not found in sheets.');
    }

    for (var i = 1; i < attendanceData.length; i++) {
        if (attendanceData[i][processedCol] !== 'Yes' && attendanceData[i][sessionTypeCol] === 'Guided Session') {
            var peardeckLink = attendanceData[i][peardeckLinkCol];
            var pearSheet = SpreadsheetApp.openByUrl(peardeckLink).getSheets()[0];
            var pearData = pearSheet.getDataRange().getValues();
            var date = attendanceData[i][dateCol];

            for (var j = 1; j < springData.length; j++) {
                var studentPresent = checkAttendance(pearData, springData[j], emailCol, sisLoginIdCol, alternateEmailCol);
                if (studentPresent) {
                    var currentTotal = springSheet.getRange(j + 1, totalAttendedGuidedSessionsCol).getValue() || 0;
                    springSheet.getRange(j + 1, totalAttendedGuidedSessionsCol).setValue(currentTotal + 1);
                    springSheet.getRange(j + 1, lastAttendedGuidedSessionCol).setValue(date);
                }
            }
            attendanceSheet.getRange(i + 1, processedCol + 1).setValue('Yes');
        }
    }
}
