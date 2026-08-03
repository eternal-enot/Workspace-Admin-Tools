const GROUP_ACTIONS_CFG = {
    ACTION_COL: 3, // C
    RECORD_TYPE_COL: 2, // B
    STATUS_COL: 1, // A
    START_ROW: 2,
    TIMESTAMP_COL: 20, // T
    ACTIONS: ["IDLE", "Send to Archive", "Restore", "Move to Custom OU"],
    COLORS: {
        "IDLE": "#fff9c4",
        "Send to Archive": "#ffcc80",
        "Restore": "#c8e6c9",
        "Move to Custom OU": "#c8e6c9"
    }
};

function executeGroupActionsActiveSheet() {
    executeGroupActions_(false);
}

function executeGroupActionsAllSheets() {
    executeGroupActions_(true);
}

function getRestoreOuPath_(sheetName) {
    const dept = String(sheetName || '').trim();
    return `/2. Факультети/ФБМІ/${dept}/0. Співробітники`;
}

function executeGroupActions_(processAll) {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheets = [];

    if (processAll) {
        sheets = ss.getSheets().filter(s => 
            s.getName() !== "Архів" && 
            s.getName() !== "Відповіді форми (1)"
        );
    } else {
        sheets = [ss.getActiveSheet()];
        const name = sheets[0].getName();
        if (name === "Архів" || name === "Відповіді форми (1)") {
            ui.alert(`❌ Cannot execute group actions on "${name}". Please select a department sheet.`);
            return;
        }
    }

    let stats = { archive: 0, customOu: 0, restored: 0, errors: [] };

    for (const sheet of sheets) {
        const lastRow = sheet.getLastRow();
        if (lastRow < GROUP_ACTIONS_CFG.START_ROW) {
            setupGroupActionsFormatting_(sheet);
            continue;
        }
        
        const lastCol = sheet.getLastColumn();
        const maxCol = Math.max(lastCol, GROUP_ACTIONS_CFG.TIMESTAMP_COL);
        const dataRange = sheet.getRange(GROUP_ACTIONS_CFG.START_ROW, 1, lastRow - GROUP_ACTIONS_CFG.START_ROW + 1, maxCol);
        const data = dataRange.getValues();

        for (let i = 0; i < data.length; i++) {
            const rowIndex = GROUP_ACTIONS_CFG.START_ROW + i;
            const action = String(data[i][GROUP_ACTIONS_CFG.ACTION_COL - 1] || "").trim();
            const rowData = data[i];
            
            if (action === "IDLE" || !action) continue;

            const mainEmail = String(rowData[6 - 1] || "").split(/[,;\s]+/)[0]; // Col F (6) User key / email

            try {
                if (action === "Send to Archive") {
                    if (!mainEmail || !mainEmail.includes("@")) throw new Error("Missing or invalid gen email (Col F)");

                    const userObj = typeof callWithRetry_ !== "undefined" ? callWithRetry_(() => AdminDirectory.Users.get(mainEmail), 5) : AdminDirectory.Users.get(mainEmail);
                    const currentOu = userObj.orgUnitPath;
                    const newOu = currentOu.replace(/\/(1\. Бакалаврат|2\. Магістратура|3\. Аспірантура|0\. Співробітники)/i, '/Архів');
                    
                    if (newOu !== currentOu && !currentOu.includes("/Архів")) {
                        if (typeof archiveEnsureOrgUnit_ !== "undefined") archiveEnsureOrgUnit_(newOu);
                        if (typeof callWithRetry_ !== "undefined") {
                            callWithRetry_(() => AdminDirectory.Users.update({ orgUnitPath: newOu }, mainEmail), 5);
                        } else {
                            AdminDirectory.Users.update({ orgUnitPath: newOu }, mainEmail);
                        }
                    }
                    
                    sheet.getRange(rowIndex, 12).setValue(newOu); // Col L (12) Org unit path
                    sheet.getRange(rowIndex, GROUP_ACTIONS_CFG.RECORD_TYPE_COL).setValue(`to_archiv (In OU)`);
                    sheet.getRange(rowIndex, GROUP_ACTIONS_CFG.ACTION_COL).setValue("IDLE");
                    stats.archive++;

                } else if (action === "Restore") {
                    if (!mainEmail || !mainEmail.includes("@")) throw new Error("Missing or invalid gen email (Col F)");

                    const sheetName = sheet.getName();
                    const newOu = getRestoreOuPath_(sheetName);

                    const userObj = typeof callWithRetry_ !== "undefined" ? callWithRetry_(() => AdminDirectory.Users.get(mainEmail), 5) : AdminDirectory.Users.get(mainEmail);
                    const currentOu = userObj.orgUnitPath;

                    if (newOu !== currentOu) {
                        if (typeof archiveEnsureOrgUnit_ !== "undefined") archiveEnsureOrgUnit_(newOu);
                        if (typeof callWithRetry_ !== "undefined") {
                            callWithRetry_(() => AdminDirectory.Users.update({ orgUnitPath: newOu }, mainEmail), 5);
                        } else {
                            AdminDirectory.Users.update({ orgUnitPath: newOu }, mainEmail);
                        }
                    }

                    sheet.getRange(rowIndex, 12).setValue(newOu); // Update Org unit path (Col L)
                    sheet.getRange(rowIndex, GROUP_ACTIONS_CFG.RECORD_TYPE_COL).setValue(`restored`);
                    sheet.getRange(rowIndex, GROUP_ACTIONS_CFG.ACTION_COL).setValue("IDLE");
                    stats.restored++;

                } else if (action === "Move to Custom OU") {
                    const newOU = String(rowData[13 - 1] || "").trim(); // Col M (13)
                    const oldOU = String(rowData[12 - 1] || "").trim(); // Col L (12)
                    
                    if (!mainEmail || !newOU) throw new Error("Missing user key (Col F) or new OU (Col M)");
                    
                    const newOUNorm = typeof normalizeOuPath_ !== "undefined" ? normalizeOuPath_(newOU) : newOU;
                    const oldOUNorm = typeof normalizeOuPath_ !== "undefined" ? normalizeOuPath_(oldOU) : oldOU;
                    if (newOUNorm && newOUNorm !== oldOUNorm) {
                        if (typeof callWithRetry_ !== "undefined") {
                            callWithRetry_(() => AdminDirectory.Users.update({ orgUnitPath: newOUNorm }, mainEmail), 5);
                        } else {
                            AdminDirectory.Users.update({ orgUnitPath: newOUNorm }, mainEmail);
                        }
                    }
                    
                    sheet.getRange(rowIndex, 12).setValue(newOUNorm);
                    sheet.getRange(rowIndex, GROUP_ACTIONS_CFG.ACTION_COL).setValue("IDLE");
                    sheet.getRange(rowIndex, 13).setValue("");
                    stats.customOu++;
                }

            } catch (e) {
                stats.errors.push(`Sheet ${sheet.getName()}, row ${rowIndex}: action "${action}" - ${e.message}`);
            }
        }
        
        // Setup format for remaining rows
        setupGroupActionsFormatting_(sheet);
    }
    
    let msg = `✅ Execution finished.\n\n` +
              `Sent to Archive: ${stats.archive}\n` +
              `Restored: ${stats.restored}\n` +
              `Moved to Custom OU: ${stats.customOu}`;
    
    if (stats.errors.length > 0) {
        msg += `\n\n⚠️ Errors:\n` + stats.errors.join("\n");
    }
    
    if (stats.archive === 0 && stats.customOu === 0 && stats.restored === 0 && stats.errors.length === 0) {
        ui.alert("🤷‍♂️ No actionable rows found (all IDLE).");
    } else {
        ui.alert(msg);
    }
}

/**
 * Setup columns B and C on department sheets.
 */
function setupGroupActionsFormatting_(sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    
    const maxRows = sheet.getMaxRows();
    
    // Column A: Status (Data Validation & Colors)
    const colA = sheet.getRange(GROUP_ACTIONS_CFG.START_ROW, GROUP_ACTIONS_CFG.STATUS_COL, maxRows - 1, 1);
    const ruleA = SpreadsheetApp.newDataValidation()
        .requireValueInList(["FOUND", "PENDING", "NOT FOUND", "ERROR"], true)
        .setAllowInvalid(false)
        .build();
    colA.setDataValidation(ruleA);
    
    const rules = [];
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("FOUND").setBackground("#c8e6c9").setRanges([colA]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("PENDING").setBackground("#fff9c4").setRanges([colA]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("NOT FOUND").setBackground("#ffcc80").setRanges([colA]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("ERROR").setBackground("#ef9a9a").setRanges([colA]).build());
    
    // Column C: Actions (Data Validation & Colors)
    const colC = sheet.getRange(GROUP_ACTIONS_CFG.START_ROW, GROUP_ACTIONS_CFG.ACTION_COL, maxRows - 1, 1);
    const ruleC = SpreadsheetApp.newDataValidation()
        .requireValueInList(GROUP_ACTIONS_CFG.ACTIONS, true)
        .setAllowInvalid(false)
        .build();
    colC.setDataValidation(ruleC);
    
    for (const act of GROUP_ACTIONS_CFG.ACTIONS) {
        rules.push(SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(act)
            .setBackground(GROUP_ACTIONS_CFG.COLORS[act])
            .setRanges([colC])
            .build());
    }
    
    // Row Highlight (Light Orange) for specific record types (except cols A and C)
    const rangeB = sheet.getRange(GROUP_ACTIONS_CFG.START_ROW, 2, maxRows - 1, 1);
    const rangeDT = sheet.getRange(GROUP_ACTIONS_CFG.START_ROW, 4, maxRows - 1, Math.max(1, sheet.getMaxColumns() - 3));
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=($B2="to_archiv (In OU)") + ($B2="to_archiv (Notified)")`)
        .setBackground("#ffe0b2") // Light Orange
        .setRanges([rangeB, rangeDT])
        .build());
    
    sheet.setConditionalFormatRules(rules);
}
