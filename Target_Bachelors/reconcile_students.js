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

            // Match by first/last name against STUDENTS_DATA (UA, then EN)
            const index = studentsToFind.findIndex(s =>
                isNameMatch_(fName, lName, s.firstName, s.lastName) ||
                isNameMatch_(fName, lName, s.firstNameEn, s.lastNameEn)
            );

            if (index !== -1) {
                const found = studentsToFind[index];
                surnameUa = found.lastName;
                nameUa = found.firstName;

                studentsToFind.splice(index, 1);

                // Restore active if previously marked dropout
                if (statusVal === "dropout") statusVal = "active";

            } else if (fName || lName) {
                statusVal = "dropout";
            }

            outStatus.push([statusVal]);
            outUa.push([surnameUa, nameUa]);
        }

        // Write Updates back to Sheet

        // Col B (Status)
        sheet.getRange(2, RECONCILE_CONFIG.COL_LOG_STATUS, outStatus.length, 1).setValues(outStatus);

        // Col R, S (UA Names)
        sheet.getRange(2, RECONCILE_CONFIG.COL_UA_SURNAME, outUa.length, 2).setValues(outUa);

        // Append students missing from the sheet into columns R and S
        if (studentsToFind.length > 0) {
            const newRows = studentsToFind.map(s => [s.lastName, s.firstName]);
            const startRow = lastRow + 1;
            sheet.getRange(startRow, RECONCILE_CONFIG.COL_UA_SURNAME, newRows.length, 2).setValues(newRows);
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
