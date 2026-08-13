function executeSupportDeletions() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SUPPORT_CFG.SHEET_NAME);
    if (!sheet) {
        sheet = ss.getActiveSheet();
        if (sheet.getName() !== SUPPORT_CFG.SHEET_NAME) {
            ui.alert(`❌ Цей скрипт призначено для аркуша "${SUPPORT_CFG.SHEET_NAME}".`);
            return;
        }
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < SUPPORT_CFG.START_ROW) return;

    const dataRange = sheet.getRange(SUPPORT_CFG.START_ROW, 1, lastRow - SUPPORT_CFG.START_ROW + 1, SUPPORT_CFG.COL_OU);
    const data = dataRange.getValues();

    // Check how many to delete
    const toDeleteIndices = [];
    for (let i = 0; i < data.length; i++) {
        const currentStatus = String(data[i][SUPPORT_CFG.COL_STATUS - 1] || "").trim().toUpperCase();
        if (currentStatus === "DELETE") {
            toDeleteIndices.push(i);
        }
    }

    if (toDeleteIndices.length === 0) {
        ui.alert("🤷‍♂️ Не знайдено жодного рядка зі статусом 'delete'.");
        return;
    }

    const response = ui.alert(
        "Увага! Незворотна дія",
        `Знайдено акаунтів для видалення: ${toDeleteIndices.length}.\nВи впевнені, що хочете безповоротно видалити ці акаунти з Google Workspace?`,
        ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
        ui.alert("Видалення скасовано.");
        return;
    }

    let deletedCount = 0;
    let errors = [];

    for (const i of toDeleteIndices) {
        const row = data[i];
        const email = String(row[SUPPORT_CFG.COL_EMAIL - 1] || "").trim();
        
        if (!email) continue;

        try {
            AdminDirectory.Users.remove(email);
            
            row[SUPPORT_CFG.COL_STATUS - 1] = "DELETED";
            // Clear other fields
            row[SUPPORT_CFG.COL_FIRST_NAME - 1] = "";
            row[SUPPORT_CFG.COL_LAST_NAME - 1] = "";
            row[SUPPORT_CFG.COL_EXTRA_EMAILS - 1] = "";
            row[SUPPORT_CFG.COL_PHONES - 1] = "";
            row[SUPPORT_CFG.COL_LAST_LOGIN - 1] = "";
            row[SUPPORT_CFG.COL_OU - 1] = "";
            
            deletedCount++;
        } catch (e) {
            errors.push(`${email}: ${e.message}`);
            row[SUPPORT_CFG.COL_STATUS - 1] = "ERROR: " + e.message;
        }
    }

    // Persist updated statuses
    dataRange.setValues(data);

    if (errors.length > 0) {
        ui.alert(`⚠️ Видалено акаунтів: ${deletedCount}.\n\nПомилки:\n${errors.join("\\n")}`);
    } else {
        ui.alert(`✅ Успішно видалено акаунтів: ${deletedCount}.`);
    }
}
