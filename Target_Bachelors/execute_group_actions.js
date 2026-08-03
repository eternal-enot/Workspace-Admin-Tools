const GROUP_ACTIONS_CFG = {
    ACTION_COL: 3, // C
    RECORD_TYPE_COL: 2, // B
    STATUS_COL: 1, // A
    START_ROW: 2,
    TIMESTAMP_COL: 20, // T
    ACTIONS: ["IDLE", "Send to Archive", "Move to Masters", "Move to Custom OU", "Notify Deletion", "Notify Alumni", "Delete Account", "Restore"],
    COLORS: {
        "IDLE": "#fff9c4",
        "Send to Archive": "#ffcc80",
        "Move to Masters": "#90caf9",
        "Move to Custom OU": "#c8e6c9",
        "Notify Deletion": "#ce93d8",
        "Notify Alumni": "#90caf9",
        "Delete Account": "#ef9a9a",
        "Restore": "#c8e6c9"
    }
};

function executeGroupActionsActiveSheet() {
    executeGroupActions_(false);
}

function executeGroupActionsAllSheets() {
    executeGroupActions_(true);
}

function getRestoreOuPath_(groupRaw) {
    const g = String(groupRaw || '').trim();
    const gUpper = g.toUpperCase();
    
    const isBMI = (gUpper.includes('БМ') || gUpper.includes('ЗМ'));
    const isTMBI = (gUpper.includes('БФ') || gUpper.includes('ЗФ'));
    const isBMK = (gUpper.includes('БС') || gUpper.includes('ЗК'));
    const isBBZL = (gUpper.includes('БР') || gUpper.includes('ЗР'));

    let dept = 'Інше';
    if (isBMI) dept = 'БМІ';
    else if (isTMBI) dept = 'ТМБІ';
    else if (isBMK) dept = 'БМК';
    else if (isBBZL) dept = 'ББЗЛ';

    return `/2. Факультети/ФБМІ/${dept}/1. Бакалаврат/${gUpper}`;
}

function executeGroupActions_(processAll) {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheets = [];

    if (processAll) {
        sheets = ss.getSheets().filter(s => 
            s.getName() !== "Архів" && 
            s.getName() !== "До МАГІСТРІВ" &&
            s.getName() !== "Відповіді форми (1)"
        );
    } else {
        sheets = [ss.getActiveSheet()];
        const name = sheets[0].getName();
        if (name === "Архів" || name === "До МАГІСТРІВ" || name === "Відповіді форми (1)") {
            ui.alert(`❌ Cannot execute group actions on "${name}". Please select a group sheet.`);
            return;
        }
    }

    let stats = { archive: 0, masters: 0, customOu: 0, notified: 0, deleted: 0, restored: 0, errors: [] };

    // Calculate deletion date for emails (e.g. 6 months from now)
    const deletionDate = new Date();
    deletionDate.setMonth(deletionDate.getMonth() + 6);
    let deletionDateStr = deletionDate.toLocaleDateString();
    if (typeof formatUkrainianDate_ === "function") {
        deletionDateStr = formatUkrainianDate_(deletionDate);
    }

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

            const recordType = String(rowData[GROUP_ACTIONS_CFG.RECORD_TYPE_COL - 1] || "").trim();
            const isMaster = recordType.includes("master");
            const baseType = isMaster ? "master" : "to_archiv";
            const mainEmail = String(rowData[6 - 1] || "").split(/[,;\s]+/)[0]; // Col F (6) User key / email
            const recoveryEmail = String(rowData[9 - 1] || "").trim(); // Col I (9) Recovery email

            try {
                if (action === "Send to Archive") {
                    if (!mainEmail || !mainEmail.includes("@")) throw new Error("Missing or invalid gen email (Col I)");

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
                    sheet.getRange(rowIndex, GROUP_ACTIONS_CFG.RECORD_TYPE_COL).setValue(`${baseType} (In OU)`);
                    sheet.getRange(rowIndex, GROUP_ACTIONS_CFG.ACTION_COL).setValue("IDLE");
                    stats.archive++;

                } else if (action === "Move to Masters") {
                    if (!mainEmail || !mainEmail.includes("@")) throw new Error("Missing or invalid gen email (Col I)");

                    const userObj = typeof callWithRetry_ !== "undefined" ? callWithRetry_(() => AdminDirectory.Users.get(mainEmail), 5) : AdminDirectory.Users.get(mainEmail);
                    const currentOu = userObj.orgUnitPath;
                    const newOu = currentOu.replace('/1. Бакалаврат', '/2. Магістратура');
                    
                    if (newOu !== currentOu) {
                        if (typeof archiveEnsureOrgUnit_ !== "undefined") archiveEnsureOrgUnit_(newOu);
                        if (typeof callWithRetry_ !== "undefined") {
                            callWithRetry_(() => AdminDirectory.Users.update({ orgUnitPath: newOu }, mainEmail), 5);
                        } else {
                            AdminDirectory.Users.update({ orgUnitPath: newOu }, mainEmail);
                        }
                    }

                    sheet.getRange(rowIndex, 12).setValue(newOu); // Update Org unit path (Col L)
                    sheet.getRange(rowIndex, GROUP_ACTIONS_CFG.RECORD_TYPE_COL).setValue(`master (In OU)`);
                    sheet.getRange(rowIndex, GROUP_ACTIONS_CFG.ACTION_COL).setValue("IDLE");
                    stats.masters++;

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

                } else if (action === "Notify Deletion" || action === "Notify Alumni") {
                    const rawMain = mainEmail;
                    const rawRec = recoveryEmail;
                    
                    const splitEmails = (str) => str.split(/[,;\s]+/).map(e => e.trim()).filter(e => e.includes("@"));
                    const toEmails = [];
                    toEmails.push(...splitEmails(rawMain));
                    toEmails.push(...splitEmails(rawRec));
                    
                    const uniqueEmails = [...new Set(toEmails)];
                    const targetEmail = toEmails.length > 0 ? toEmails[0] : rawMain;
                    
                    if (uniqueEmails.length === 0) throw new Error("Не знайдено жодного валідного email.");

                    if (action === "Notify Deletion") {
                        if (typeof sendDeletionEmail_ !== "undefined") {
                            sendDeletionEmail_(targetEmail, uniqueEmails.join(","), deletionDateStr);
                        } else {
                            throw new Error("Функція sendDeletionEmail_ не знайдена.");
                        }
                    } else {
                        if (typeof sendAlumniEmail_ !== "undefined") {
                            sendAlumniEmail_(targetEmail, uniqueEmails.join(","), deletionDateStr);
                        } else {
                            throw new Error("Функція sendAlumniEmail_ не знайдена.");
                        }
                    }

                    sheet.getRange(rowIndex, GROUP_ACTIONS_CFG.TIMESTAMP_COL).setValue(new Date());
                    sheet.getRange(rowIndex, GROUP_ACTIONS_CFG.RECORD_TYPE_COL).setValue(`${baseType} (Notified)`);
                    sheet.getRange(rowIndex, GROUP_ACTIONS_CFG.ACTION_COL).setValue("IDLE");
                    stats.notified++;

                } else if (action === "Delete Account") {
                    sheet.getRange(rowIndex, GROUP_ACTIONS_CFG.RECORD_TYPE_COL).setValue(`${baseType} (Deleted)`);
                    sheet.getRange(rowIndex, GROUP_ACTIONS_CFG.ACTION_COL).setValue("IDLE");
                    stats.deleted++;

                } else if (action === "Restore") {
                    if (!mainEmail || !mainEmail.includes("@")) throw new Error("Missing or invalid gen email (Col I)");

                    const groupName = sheet.getName();
                    const newOu = getRestoreOuPath_(groupName);

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
              `Moved to Masters: ${stats.masters}\n` +
              `Moved to Custom OU: ${stats.customOu}\n` +
              `Notified: ${stats.notified}\n` +
              `Deleted: ${stats.deleted}\n` +
              `Restored: ${stats.restored}`;
    
    if (stats.errors.length > 0) {
        msg += `\n\n⚠️ Errors:\n` + stats.errors.join("\n");
    }
    
    if (stats.archive === 0 && stats.masters === 0 && stats.customOu === 0 && stats.notified === 0 && stats.deleted === 0 && stats.restored === 0 && stats.errors.length === 0) {
        ui.alert("🤷‍♂️ No actionable rows found (all IDLE).");
    } else {
        ui.alert(msg);
    }
}

/**
 * Setup columns B and C on group sheets.
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
