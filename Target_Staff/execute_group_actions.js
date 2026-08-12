const GROUP_ACTIONS_CFG = {
    ACTION_COL: 3, // C
    RECORD_TYPE_COL: 2, // B
    STATUS_COL: 1, // A
    START_ROW: 2,
    TIMESTAMP_COL: 20, // T
    ACTIONS: ["IDLE", "Send to Archive", "Restore", "Move to Custom OU", "Change Main Email"],
    COLORS: {
        "IDLE": "#fff9c4",
        "Send to Archive": "#ffcc80",
        "Restore": "#c8e6c9",
        "Move to Custom OU": "#c8e6c9",
        "Change Main Email": "#e1bee7"
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

    let stats = { archive: 0, customOu: 0, restored: 0, emailChanges: 0, errors: [] };

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

                } else if (action === "Change Main Email") {
                    const newEmail = String(rowData[4 - 1] || "").trim(); // Col D (4)
                    const backupEmail = String(rowData[9 - 1] || "").trim(); // Col I (9)
                    
                    if (!mainEmail || !mainEmail.includes("@")) throw new Error("Missing or invalid current main email (Col F)");
                    if (!newEmail || !newEmail.includes("@")) throw new Error("Missing or invalid new email (Col D)");
                    
                    if (newEmail !== mainEmail) {
                        // Get immutable user ID to avoid eventual consistency issues after rename
                        const userObj = typeof callWithRetry_ !== "undefined" ? callWithRetry_(() => AdminDirectory.Users.get(mainEmail), 5) : AdminDirectory.Users.get(mainEmail);
                        const userId = userObj.id;

                        // 1. Update primary email
                        if (typeof callWithRetry_ !== "undefined") {
                            callWithRetry_(() => AdminDirectory.Users.update({ primaryEmail: newEmail }, mainEmail), 5);
                        } else {
                            AdminDirectory.Users.update({ primaryEmail: newEmail }, mainEmail);
                        }
                        
                        // 2. Add old email as alias (Google usually does this automatically, explicitly adding just in case)
                        try {
                            // Wait a short moment to let the rename propagate before interacting with aliases
                            Utilities.sleep(1000);
                            
                            if (typeof callWithRetry_ !== "undefined") {
                                callWithRetry_(() => AdminDirectory.Users.Aliases.insert({ alias: mainEmail }, userId), 3);
                            } else {
                                AdminDirectory.Users.Aliases.insert({ alias: mainEmail }, userId);
                            }
                        } catch(aliasErr) {
                            // Ignore if alias already exists (e.g. automatically created by Google)
                            if (!String(aliasErr).includes("already exists") && !String(aliasErr).includes("Duplicate") && !String(aliasErr).includes("already an alias")) {
                                throw new Error("Error adding alias: " + aliasErr.message);
                            }
                        }

                        // 3. Send email to new email and backup email
                        const SENDER_NAME = 'Адміністрація Google Workspace for Education';
                        const subject = "Оновлення електронної адреси з lll.kpi.ua на edu.kpi.ua";
                        const bodyPlain = getChangeEmailBody_(mainEmail, newEmail, false);
                        const htmlBody = getChangeEmailBody_(mainEmail, newEmail, true);
                        
                        MailApp.sendEmail({
                            to: newEmail,
                            subject: subject,
                            name: SENDER_NAME,
                            body: bodyPlain,
                            htmlBody: htmlBody
                        });
                        
                        if (backupEmail && backupEmail.includes("@")) {
                            const backupList = backupEmail.split(/[,;\s]+/).map(e => e.trim()).filter(e => e.includes("@"));
                            const validBackups = backupList.filter(e => !e.endsWith("@edu.kpi.ua") && !e.endsWith("@lll.kpi.ua"));
                            
                            if (validBackups.length > 0) {
                                MailApp.sendEmail({
                                    to: validBackups.join(","),
                                    subject: subject,
                                    name: SENDER_NAME,
                                    body: bodyPlain,
                                    htmlBody: htmlBody
                                });
                            }
                        }
                    }
                    
                    // Update sheet
                    sheet.getRange(rowIndex, 6).setValue(newEmail); // Set new main email in Col F
                    sheet.getRange(rowIndex, GROUP_ACTIONS_CFG.RECORD_TYPE_COL).setValue(`Email changed`);
                    sheet.getRange(rowIndex, GROUP_ACTIONS_CFG.ACTION_COL).setValue("IDLE");
                    stats.emailChanges++;
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
              `Moved to Custom OU: ${stats.customOu}\n` +
              `Emails Changed: ${stats.emailChanges}`;
    
    if (stats.errors.length > 0) {
        msg += `\n\n⚠️ Errors:\n` + stats.errors.join("\n");
    }
    
    if (stats.archive === 0 && stats.customOu === 0 && stats.restored === 0 && stats.emailChanges === 0 && stats.errors.length === 0) {
        ui.alert("🤷‍♂️ No actionable rows found (all IDLE).");
    } else {
        ui.alert(msg);
    }
}

function sendTestChangeEmail() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
        "Тестовий лист", 
        "Введіть вашу email-адресу для отримання тестового листа:", 
        ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.OK) {
        const adminEmail = response.getResponseText().trim();
        if (adminEmail && adminEmail.includes("@")) {
            const SENDER_NAME = 'Адміністрація Google Workspace for Education';
            const subject = "Оновлення електронної адреси з lll.kpi.ua на edu.kpi.ua (ТЕСТ)";
            const demoOld = "old.email@lll.kpi.ua";
            const demoNew = "new.email@lll.kpi.ua";
            const bodyPlain = getChangeEmailBody_(demoOld, demoNew, false);
            const htmlBody = getChangeEmailBody_(demoOld, demoNew, true);
            
            try {
                MailApp.sendEmail({
                    to: adminEmail,
                    subject: subject,
                    name: SENDER_NAME,
                    body: bodyPlain,
                    htmlBody: htmlBody
                });
                ui.alert(`✅ Тестовий лист успішно надіслано на ${adminEmail}`);
            } catch (e) {
                ui.alert(`❌ Помилка надсилання листа: ${e.message}`);
            }
        } else {
            ui.alert("❌ Некоректна електронна адреса.");
        }
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

/**
 * Генерує HTML та Plain-текст листа про зміну адреси
 */
function getChangeEmailBody_(mainEmail, newEmail, isHtml) {
    if (!isHtml) {
        return `Шановний користувачу,\n\nПовідомляємо, що Вашу основну корпоративну електронну адресу було змінено.\n\nНова адреса: ${newEmail}\nПопередня адреса: ${mainEmail}\n\nСтара електронна адреса продовжуватиме функціонувати як псевдонім (аліас), й листи відправлені на неї будуть автоматично надходити на нову. Однак для входу в обліковий запис необхідно використовувати виключно нову адресу.\n\nУ разі виникнення додаткових запитань Ви можете звернутися до адміністратора, відповівши на цей лист.\n\nЗ повагою,\nАдміністрація Google Workspace for Education\nКПІ ім. Ігоря Сікорського\nФакультет біомедичної інженерії`;
    }

    const logoHtml = '<div style="text-align:center;margin-top:24px;"><img src="https://fbmi.kpi.ua/wp-content/uploads/2025/09/fbmi_blue_logo.png" alt="Лого університету" style="max-width:200px;height:auto;"></div>';
    
    return '<!doctype html>' +
        '<html><body style="margin:0;padding:0;background:#f6f7f9;">' +
        '  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#f6f7f9;padding:24px 0;">' +
        '    <tr><td align="center">' +
        '      <table role="presentation" width="640" cellspacing="0" cellpadding="0" style="background:#ffffff;border-radius:12px;overflow:hidden;font-family:Arial,system-ui,-apple-system,\'Segoe UI\',Roboto,\'Noto Sans\',sans-serif;color:#111">' +
        '        <tr><td style="padding:24px 32px 0;font-size:16px;line-height:1.6;">' +
        '          <p style="margin:0 0 12px 0;">Шановний користувачу,</p>' +
        '          <p style="margin:0 0 12px 0;">Повідомляємо, що Вашу основну корпоративну електронну адресу було змінено.</p>' +
        '        </td></tr>' +
        '        <tr><td style="padding:8px 32px 0">' +
        '          <table width="100%" cellspacing="0" cellpadding="0" style="border:1px solid #e6e6e6;border-radius:8px;">' +
        '            <tr>' +
        '              <td style="padding:12px 16px;font-size:14px;background:#fafafa;border-bottom:1px solid #e6e6e6;color:#555;">Попередня адреса:</td>' +
        '              <td style="padding:12px 16px;font-size:16px;background:#fafafa;border-bottom:1px solid #e6e6e6;text-align:right;color:#666;">' + mainEmail + '</td>' +
        '            </tr>' +
        '            <tr>' +
        '              <td style="padding:12px 16px;font-size:14px;background:#fff;font-weight:bold;color:#333;">Нова основна адреса:</td>' +
        '              <td style="padding:12px 16px;font-size:16px;background:#fff;text-align:right;font-weight:bold;color:#1a73e8;">' + newEmail + '</td>' +
        '            </tr>' +
        '          </table>' +
        '        </td></tr>' +
        '        <tr><td style="padding:24px 32px 0;font-size:16px;line-height:1.6;color:#333">' +
        '          Стара електронна адреса продовжуватиме функціонувати як <strong>псевдонім (аліас)</strong>, й листи відправлені на неї будуть автоматично надходити на нову. Однак для входу в обліковий запис необхідно використовувати виключно нову адресу.' +
        '        </td></tr>' +
        '        <tr><td style="padding:24px 32px 24px;font-size:14px;color:#333;line-height:1.6;">' +
        '          У разі виникнення додаткових запитань Ви можете звернутися до адміністратора, відповівши на цей лист.' +
        '        </td></tr>' +
        '        <tr><td style="padding:24px 32px 24px;font-size:14px;color:#333;line-height:1.5;border-top:1px solid #eee;">' +
        '          З повагою,<br>Адміністрація Google Workspace for Education<br>КПІ ім. Ігоря Сікорського<br>Факультет біомедичної інженерії' +
        '        </td></tr>' +
        '        <tr><td style="padding:0 32px 24px;">' + logoHtml + '</td></tr>' +
        '      </table>' +
        '    </td></tr>' +
        '  </table>' +
        '</body></html>';
}
