/**
 * TEMPORARY SCRIPT TO UPDATE THE SHEET (Target_PhD)
 * You can safely delete this file once the run is complete.
 */

function runOneTimeMigration() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets().filter(s => 
        s.getName() !== "Архів" && 
        s.getName() !== "Відповіді форми (1)"
    );

    const headers = [
        "Status",                // A (1)
        "record_type",           // B (2)
        "actions",               // C (3)
        "",                      // D (4) - empty
        "Group",                 // E (5) - student group
        "User key / email",      // F (6)
        "First name",            // G (7)
        "Last name",             // H (8)
        "Recovery/extra email",  // I (9)
        "Last login (local)",    // J (10)
        "Comment",               // K (11)
        "Org unit path",         // L (12)
        "New OU",                // M (13)
        "Move status",           // N (14)
        "Move note",             // O (15)
        "",                      // P (16)
        "",                      // Q (17)
        "UA Surname",            // R (18)
        "UA Name",               // S (19)
        "Надіслано"              // T (20)
    ];

    let updatedCount = 0;

    for (const sheet of sheets) {
        // 1. Setup column headers
        const maxCols = sheet.getMaxColumns();
        if (maxCols < headers.length) {
            sheet.insertColumnsAfter(maxCols, headers.length - maxCols);
        }
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.setFrozenRows(1);

        const lastRow = sheet.getLastRow();
        if (lastRow >= 2) {
            // 2. Clear old Data Validation rules to avoid conflicts
            const dataRangeB = sheet.getRange(2, 2, lastRow - 1, 1);
            const dataRangeC = sheet.getRange(2, 3, lastRow - 1, 1);
            
            dataRangeB.clearDataValidations();
            dataRangeC.clearDataValidations();
            
            // Update legacy 'false' values to 'dropout' in column B
            // Fill empty cells: B becomes 'active', C becomes 'IDLE'
            const valsB = dataRangeB.getValues();
            const valsC = dataRangeC.getValues();
            
            const validActions = ["IDLE", "Send to Archive", "Restore", "Move to Custom OU"];

            for (let i = 0; i < valsB.length; i++) {
                const bStr = String(valsB[i][0] || "").trim().toLowerCase();
                if (bStr === "false") {
                    valsB[i][0] = "dropout";
                } else if (bStr === "" || bStr === "undefined") {
                    valsB[i][0] = "active";
                }
                
                let cStr = String(valsC[i][0] || "").trim();
                // Default to IDLE if the cell is empty or contains an invalid value
                if (!validActions.includes(cStr)) {
                    valsC[i][0] = "IDLE";
                }
            }
            
            dataRangeB.setValues(valsB);
            dataRangeC.setValues(valsC);
        }

        // 3. Apply Data Validation rules and conditional formatting
        if (typeof setupGroupActionsFormatting_ === 'function') {
            setupGroupActionsFormatting_(sheet);
        }

        updatedCount++;
    }

    SpreadsheetApp.getUi().alert(`✅ Migration successfully completed!\nSheets updated: ${updatedCount}.\n\nYou can now delete the "setup_migration.js" script and its menu button.`);
}
