/**
 * Syncs sheet headers, data validation, and formatting with the current script.
 */

function updateSheetLayout_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets().filter(s =>
        s.getName() !== "Архів" &&
        s.getName() !== "До АСПІРАНТУРИ" &&
        s.getName() !== "Відповіді форми (1)"
    );

    if (typeof setupGroupActionsFormatting_ !== "function") {
        SpreadsheetApp.getUi().alert(
            "❌ Функція setupGroupActionsFormatting_ не знайдена.\n\n" +
            "Переконайтесь, що execute_group_actions.js задеплоєно (clasp push) і спробуйте ще раз."
        );
        return;
    }

    const validActions = (typeof GROUP_ACTIONS_CFG !== "undefined" && GROUP_ACTIONS_CFG.ACTIONS)
        ? GROUP_ACTIONS_CFG.ACTIONS
        : ["IDLE", "Send to Archive", "Move to PhD", "Move to Custom OU", "Notify Deletion", "Notify Alumni", "Delete Account", "Restore", "Change Main Email", "Change Name"];

    const headers = [
        "Status",                // A (1)
        "record_type",           // B (2)
        "actions",               // C (3)
        "",                      // D (4)
        "",                      // E (5)
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
        const maxCols = sheet.getMaxColumns();
        if (maxCols < headers.length) {
            sheet.insertColumnsAfter(maxCols, headers.length - maxCols);
        }
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.setFrozenRows(1);

        const lastRow = sheet.getLastRow();
        const maxRows = sheet.getMaxRows();
        const dataRowCount = Math.max(lastRow - 1, 0);

        if (dataRowCount > 0) {
            const dataRangeA = sheet.getRange(2, 1, dataRowCount, 1);
            const dataRangeB = sheet.getRange(2, 2, dataRowCount, 1);
            const dataRangeC = sheet.getRange(2, 3, dataRowCount, 1);

            dataRangeA.clearDataValidations();
            dataRangeB.clearDataValidations();
            dataRangeC.clearDataValidations();

            const valsB = dataRangeB.getValues();
            const valsC = dataRangeC.getValues();

            for (let i = 0; i < valsB.length; i++) {
                const bStr = String(valsB[i][0] || "").trim().toLowerCase();
                if (bStr === "false") {
                    valsB[i][0] = "dropout";
                } else if (bStr === "true" || bStr === "" || bStr === "undefined") {
                    valsB[i][0] = "active";
                }

                const cStr = String(valsC[i][0] || "").trim();
                if (!validActions.includes(cStr)) {
                    valsC[i][0] = "IDLE";
                }
            }

            dataRangeB.setValues(valsB);
            dataRangeC.setValues(valsC);
        }

        // Clear stale validation on the rest of columns A/C (below lastRow)
        if (maxRows > 2) {
            const tailRows = maxRows - 1;
            sheet.getRange(2, 1, tailRows, 1).clearDataValidations();
            sheet.getRange(2, 3, tailRows, 1).clearDataValidations();
        }

        setupGroupActionsFormatting_(sheet);
        updatedCount++;
    }

    SpreadsheetApp.getUi().alert(
        `✅ Вигляд таблиці оновлено.\n` +
        `Оновлено аркушів: ${updatedCount}.\n` +
        `Список actions синхронізовано зі скриптом (${validActions.length} пунктів).`
    );
}

/** @deprecated Use updateSheetLayout_ */
function runOneTimeMigration() {
    updateSheetLayout_();
}
