/**
 * ТИМЧАСОВИЙ СКРИПТ ДЛЯ ОНОВЛЕННЯ ТАБЛИЦІ
 * Після успішного запуску цей файл можна видалити.
 */

function runOneTimeMigration() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets().filter(s => 
        s.getName() !== "Архів" && 
        s.getName() !== "До МАГІСТРІВ" &&
        s.getName() !== "Відповіді форми (1)"
    );

    const headers = [
        "Status",                // A (1)
        "record_type",           // B (2)
        "actions",               // C (3)
        "",                      // D (4) - порожня
        "",                      // E (5) - порожня
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
        // 1. Встановлюємо заголовки
        const maxCols = sheet.getMaxColumns();
        if (maxCols < headers.length) {
            sheet.insertColumnsAfter(maxCols, headers.length - maxCols);
        }
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.setFrozenRows(1);

        const lastRow = sheet.getLastRow();
        if (lastRow >= 2) {
            // 2. Очищаємо старі правила перевірки даних (Data Validation), щоб уникнути помилок
            const dataRangeB = sheet.getRange(2, 2, lastRow - 1, 1);
            const dataRangeC = sheet.getRange(2, 3, lastRow - 1, 1);
            
            dataRangeB.clearDataValidations();
            dataRangeC.clearDataValidations();
            
            // Оновлюємо старі значення 'false' на 'dropout' в колонці B
            // Також заповнюємо порожні комірки B -> 'active', C -> 'IDLE'
            const valsB = dataRangeB.getValues();
            const valsC = dataRangeC.getValues();
            
            const validActions = ["IDLE", "Send to Archive", "Move to Masters", "Move to Custom OU", "Notify Deletion", "Notify Alumni", "Delete Account", "Restore"];

            for (let i = 0; i < valsB.length; i++) {
                const bStr = String(valsB[i][0] || "").trim().toLowerCase();
                if (bStr === "false") {
                    valsB[i][0] = "dropout";
                } else if (bStr === "" || bStr === "undefined") {
                    valsB[i][0] = "active";
                }
                
                let cStr = String(valsC[i][0] || "").trim();
                // Якщо порожньо, undefined або будь-яке інше старе значення, ставимо IDLE
                if (!validActions.includes(cStr)) {
                    valsC[i][0] = "IDLE";
                }
            }
            
            dataRangeB.setValues(valsB);
            dataRangeC.setValues(valsC);
        }

        // 3. Застосовуємо Data Validation та кольорове форматування
        if (typeof setupGroupActionsFormatting_ === 'function') {
            setupGroupActionsFormatting_(sheet);
        }

        updatedCount++;
    }

    SpreadsheetApp.getUi().alert(`✅ Міграцію успішно завершено!\nОновлено аркушів: ${updatedCount}.\n\nТепер ви можете видалити скрипт "setup_migration.js" та відповідну кнопку з меню.`);
}
