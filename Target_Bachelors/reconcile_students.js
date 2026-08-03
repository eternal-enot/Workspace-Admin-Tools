/************** RECONCILE CONFIG **************/
const RECONCILE_CONFIG = {
    PROCESS_ALL_SHEETS: false, // Flag: set to true to process all sheets
    USE_ACTIVE_SHEET: true,    // Default: verify active sheet only

    // Columns (1-based)
    COL_EMAIL: 6,       // F (Email Check)
    COL_FIRST_NAME: 7,  // G (Sheet First Name)
    COL_LAST_NAME: 8,   // H (Sheet Last Name)

    COL_LOG_STATUS: 2,  // B (Write 'false' here if not found)

    COL_UA_SURNAME: 18, // R (Output UA Surname)
    COL_UA_NAME: 19     // S (Output UA Name)
};

/**
 * Main function to reconcile students.
 * Can be run for active sheet or all sheets based on config.
 */
function reconcileStudents() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = getReconcileSheetsToProcess_(ss);

    // STUDENTS_DATA is a global constant from 'students_data.js'
    if (typeof STUDENTS_DATA === 'undefined') {
        throw new Error("STUDENTS_DATA not found. Make sure students_data.js is present.");
    }

    for (const sheet of sheets) {
        const sheetName = sheet.getName();
        // Assuming sheet name is the group name, e.g., "БМ-21"
        const groupName = sheetName.trim();

        // Filter JSON for this group
        const groupStudents = STUDENTS_DATA.filter(s => s.group === groupName);

        // Create a mutable copy to track found students
        let studentsToFind = [...groupStudents];

        const lastRow = sheet.getLastRow();
        if (lastRow < 2) continue; // No data

        const dataRange = sheet.getRange(2, 1, lastRow - 1, 20); // Read up to col T (20) just in case
        const values = dataRange.getValues();

        const outStatus = []; // for Col B
        const outUa = [];     // for Col R, S

        // Read existing Col B to preserve other values if we don't write 'false'
        const existingStatus = values.map(r => r[RECONCILE_CONFIG.COL_LOG_STATUS - 1]);

        for (let i = 0; i < values.length; i++) {
            const row = values[i];
            const email = row[RECONCILE_CONFIG.COL_EMAIL - 1]; // F
            const fName = String(row[RECONCILE_CONFIG.COL_FIRST_NAME - 1]).trim(); // G
            const lName = String(row[RECONCILE_CONFIG.COL_LAST_NAME - 1]).trim(); // H

            let statusVal = existingStatus[i];
            let surnameUa = "";
            let nameUa = "";

            // 1. Check Email
            if (!email) {
                // Just a check, user didn't specify action for missing email output, 
                // but implied checking it. We proceed to name matching.
            }

            // 2. Match with JSON
            // Try exact text match or fuzzy? User said "search by surname name"
            // We will look for a student in our group list
            const index = studentsToFind.findIndex(s =>
                isNameMatch_(fName, lName, s.firstName, s.lastName) ||
                isNameMatch_(fName, lName, s.firstNameEn, s.lastNameEn) // Fallback to EN match?
                // Usually sheet has Ukrainian names in G/H, but let's stick to simple logic first
            );

            if (index !== -1) {
                // MATCH FOUND
                const found = studentsToFind[index];
                surnameUa = found.lastName;
                nameUa = found.firstName;

                // Remove from list so we know who is left
                studentsToFind.splice(index, 1);

                // If status was 'dropout', maybe clear it? Or leave it? 
                // Implicitly if found, we don't mark dropout.
                if (statusVal === "dropout") statusVal = "active"; // Restore active if found again

            } else {
                // NO MATCH -> Potential Dropout
                // Only mark if the row actually has a name
                if (fName || lName) {
                    statusVal = "dropout";
                }
            }

            outStatus.push([statusVal]);
            outUa.push([surnameUa, nameUa]);
        }

        // Write Updates back to Sheet

        // Col B (Status)
        sheet.getRange(2, RECONCILE_CONFIG.COL_LOG_STATUS, outStatus.length, 1).setValues(outStatus);

        // Col R, S (UA Names)
        sheet.getRange(2, RECONCILE_CONFIG.COL_UA_SURNAME, outUa.length, 2).setValues(outUa);

        // 3. Handle "Lost" students (Remaining in studentsToFind)
        if (studentsToFind.length > 0) {
            const newRows = studentsToFind.map(s => {
                // Prepare a row. We only write to R and S.
                // But we need to append.
                // We can just set specific cells for the new rows.
                return [s.lastName, s.firstName];
            });

            const startRow = lastRow + 1;
            // We want to write into R and S.
            // We can't use appendRow easily for specific columns in the middle without messing up others if not careful.
            // But if we just assume empty rows below, we can write to R and S.

            // However, to make them visible "rows", usually we might want to put something in A or F?
            // User said: "insert surname and name into R and S (potential 'lost')"
            // I will write them starting at the first new row.

            if (newRows.length > 0) {
                sheet.getRange(startRow, RECONCILE_CONFIG.COL_UA_SURNAME, newRows.length, 2).setValues(newRows);
                // Maybe mark B as "MISSING_IN_SHEET"? User didn't ask, but helpful.
                // User asked: "insert surname and name... that's it"
            }
        }
    }
}

function isNameMatch_(sheetFirst, sheetLast, jsonFirst, jsonLast) {
    if (!sheetFirst || !sheetLast) return false;

    const normalize = (str) => {
        return String(str || "")
            .toLowerCase()
            .trim()
            // Replace various apostrophe-like characters with a standard single quote
            .replace(/[\u2019\u02BC\u0060\u2018]/g, "'");
    };

    return normalize(sheetFirst) === normalize(jsonFirst) &&
        normalize(sheetLast) === normalize(jsonLast);
}

/************** HELPERS **************/
function getReconcileSheetsToProcess_(ss) {
    if (RECONCILE_CONFIG.PROCESS_ALL_SHEETS) return ss.getSheets();
    if (RECONCILE_CONFIG.USE_ACTIVE_SHEET) return [ss.getActiveSheet()];
    return [];
}
