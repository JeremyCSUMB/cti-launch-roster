
function createUserList() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var list = ss.getSheetByName('Spring 2024');

    var data = list.getRange(1, 1, list.getLastRow(), list.getLastColumn()).getValues();
    var first_col = data[0].indexOf("First Name");
    var last_col = data[0].indexOf("Last Name");
    var email_col = data[0].indexOf("Email Address");
    var canvas_id_col = data[0].indexOf("Canvas ID"); // Add this line to get the Canvas ID column index

    var output = ss.getSheetByName('Canvas Import CSV New');
    output.clearContents();
    output.appendRow(['user_id', 'integration_id', 'login_id', 'password', 'first_name', 'last_name', 'full_name', 'sortable_name', 'short_name', 'email', 'status']);

    for (var i = 0; i < data.length; i++) {
        var first = data[i][first_col];
        var last = data[i][last_col];
        var email = data[i][email_col];
        var canvas_id = data[i][canvas_id_col]; // Add this line to get the Canvas ID value

        // Check if Canvas ID is blank before appending the row
        if (canvas_id === '') {
            output.appendRow([email, '', email, 'LAUNCH123456', first, last, '', '', '', email, 'active']);
        }
    }
}


function assignStudentAssistants() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var rosterSheet = spreadsheet.getSheetByName("Spring 2024");
    var dataRange = rosterSheet.getDataRange();
    var data = dataRange.getValues();
    var headers = data[0];

    // Dynamic column identification
    var session1Index = headers.indexOf('Deep Work Session #1') + 1;
    var session2Index = headers.indexOf('Deep Work Session #2') + 1;
    var studentAssistantIndex = headers.indexOf('Student Assistant') + 1;
    var agGroupIndex = headers.indexOf('AG Group') + 1;
    var sessionLocationIndex = headers.indexOf('Deep Work Session Location') + 1;

    // Updated SA availability based on days of the week, including CSUMB SAs
    var saAvailability = {
        "Alexis Guzman": ["Tuesday", "Thursday", "Friday"],
        "Haider Syed": ["Monday", "Wednesday", "Friday"],
        "Rodrigo Hernandez": ["Tuesday", "Friday"],
        "Elizabeth Barco Lopez": ["Wednesday", "Friday"],
        "Nishat Nawshin": ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"],
        "Geovanny Martinez": ["Monday", "Tuesday", "Thursday", "Friday"],
        "Aileen Dong": ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"],
        "Zachery Rouzaud": ["Wednesday", "Friday"],
        "Nicolas Garcia": ["Monday", "Thursday", "Friday"],
        "Sebastian Santoyo": ["Tuesday", "Wednesday"],
        "Mariana Duran": ["Tuesday", "Friday"],
        "Jesus Garcia": ["Monday", "Thursday"]
    };

    // SA institutions mapping, including CSUMB
    var saInstitutions = {
        "Alexis Guzman": "Hartnell Alisal",
        "Haider Syed": "Hartnell Alisal",
        "Rodrigo Hernandez": "Hartnell Alisal",
        "Elizabeth Barco Lopez": "CSUDH",
        "Nishat Nawshin": "CSUDH",
        "Geovanny Martinez": "CSUDH",
        "Aileen Dong": "ECC",
        "Zachery Rouzaud": "ECC",
        "Nicolas Garcia": "CSUMB",
        "Sebastian Santoyo": "CSUMB",
        "Mariana Duran": "CSUMB",
        "Jesus Garcia": "CSUMB"
    };

    // Assign fixed AG Group numbers to each SA
    var saAgGroupNumbers = {
        "Alexis Guzman": 1,
        "Haider Syed": 2,
        "Rodrigo Hernandez": 3,
        "Elizabeth Barco Lopez": 4,
        "Nishat Nawshin": 5,
        "Geovanny Martinez": 6,
        "Aileen Dong": 7,
        "Zachery Rouzaud": 8,
        "Nicolas Garcia": 9,
        "Sebastian Santoyo": 10,
        "Mariana Duran": 11,
        "Jesus Garcia": 12
    };

    // Initialize counts for equitable distribution (if needed for other logic)
    var saCounts = {};
    Object.keys(saAvailability).forEach(function (sa) {
        saCounts[sa] = 0;
    });

    // Assign SAs to students based on session preferences and institution match
    for (var i = 1; i < data.length; i++) { // Skip header row
        var session1 = data[i][session1Index - 1];
        var session2 = data[i][session2Index - 1];
        var sessionLocation = data[i][sessionLocationIndex - 1];

        var assignedSA = "";
        var minCount = Number.MAX_SAFE_INTEGER;

        Object.keys(saAvailability).forEach(function (sa) {
            var saDays = saAvailability[sa];
            var saInstitution = saInstitutions[sa];
            if ((saDays.includes(session1) || saDays.includes(session2)) && saInstitution === sessionLocation && saCounts[sa] < minCount) {
                assignedSA = sa;
                minCount = saCounts[sa];
            }
        });

        // Update the spreadsheet if an SA is assigned
        if (assignedSA !== "") {
            rosterSheet.getRange(i + 1, studentAssistantIndex).setValue(assignedSA); // +1 to adjust for header and zero-based index
            saCounts[assignedSA] += 1; // Increment SA's count if needed for other logic

            // Assign the fixed AG Group number based on the SA
            var agGroup = saAgGroupNumbers[assignedSA];
            rosterSheet.getRange(i + 1, agGroupIndex).setValue(agGroup);
        }
    }

    // Apply changes to the spreadsheet
    SpreadsheetApp.flush();
}