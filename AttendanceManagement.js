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

    if (lastAttendedCol <= 0 || attendedHistoryCol <= 0 || missedHistoryCol <= 0 || totalAttendedCol <= 0 || totalMissedCol <= 0) {
        throw new Error('Required columns not found.');
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
    }
}