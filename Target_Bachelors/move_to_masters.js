const MASTERS_TRANSFER_CFG = {
    DEST_EMAIL_COL: 6,     // F
    DEST_START_ROW: 2,
    DEST_STATUS_COL: 1,    // A
    DEST_COMMENT_COL: 11,  // K
    STATUS_VALUE: "PENDING",
};

function getMastersSpreadsheetId_() {
    if (typeof PRIVATE_CONFIG !== "undefined" && PRIVATE_CONFIG.MASTERS_SPREADSHEET_ID) {
        const id = String(PRIVATE_CONFIG.MASTERS_SPREADSHEET_ID).trim();
        if (id && !id.includes("REPLACE")) return id;
    }
    return "";
}

/**
 * Inserts email into the Masters target spreadsheet on the sheet matching groupName.
 * Skips if the email already exists on that sheet.
 */
function transferEmailToMastersSpreadsheet_(email, groupName) {
    const spreadsheetId = getMastersSpreadsheetId_();
    if (!spreadsheetId) {
        throw new Error("MASTERS_SPREADSHEET_ID is not configured in Secrets.js (and was not deployed — run clasp push after adding .claspignore)");
    }

    const groupRaw = String(groupName || "").trim();
    if (!groupRaw) throw new Error("Missing target master group (Col E)");

    const emailLower = String(email || "").trim().toLowerCase();
    if (!emailLower || !emailLower.includes("@")) {
        throw new Error("Missing or invalid email (Col F)");
    }

    let tgtSS;
    try {
        tgtSS = SpreadsheetApp.openById(spreadsheetId);
    } catch (e) {
        throw new Error(`Cannot open Masters spreadsheet: ${e.message}`);
    }

    const targetSheet = findMastersGroupSheet_(tgtSS, groupRaw);
    if (!targetSheet) {
        throw new Error(`Sheet not found in Masters spreadsheet for group "${groupRaw}"`);
    }

    insertEmailIntoMastersSheet_(targetSheet, emailLower, groupRaw);
}

function buildMasterOuPath_(groupRaw) {
    const g = normalizeGroupTypos_(groupRaw);
    let groupForOu = g.toUpperCase().replace(/МП/g, "мп").replace(/МН/g, "мн");
    const dept = resolveDeptFromGroup_(g);
    return `/2. Факультети/ФБМІ/${dept}/2. Магістратура/${groupForOu}`;
}

function findMastersGroupSheet_(ss, groupRaw) {
    const candidates = candidateMastersGroupSheetNames_(groupRaw);
    for (const name of candidates) {
        const sh = ss.getSheetByName(name);
        if (sh) return sh;
    }
    return null;
}

function candidateMastersGroupSheetNames_(groupRaw) {
    const g = String(groupRaw || "").trim();
    if (!g) return [];
    const u = g.toUpperCase();

    const candidates = [g, u];
    const mpMnVariant = u.replace(/МП/g, "мп").replace(/МН/g, "мн");
    candidates.push(mpMnVariant);

    const seen = new Set();
    const out = [];
    for (const x of candidates) {
        if (x && !seen.has(x)) {
            seen.add(x);
            out.push(x);
        }
    }
    return out;
}

function insertEmailIntoMastersSheet_(sheet, emailLower, groupName) {
    const start = MASTERS_TRANSFER_CFG.DEST_START_ROW;
    const colKey = MASTERS_TRANSFER_CFG.DEST_EMAIL_COL;
    const lastRow = sheet.getLastRow();

    let existing = [];
    if (lastRow >= start) {
        existing = sheet
            .getRange(start, colKey, lastRow - start + 1, 1)
            .getValues()
            .map(r => String(r[0] || "").trim().toLowerCase());
    }

    if (existing.some(e => e === emailLower)) return;

    const now = new Date();
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const stamp = Utilities.formatDate(now, tz, "yyyy-MM-dd HH:mm:ss");
    const commentText = `Transferred from Bachelors: ${stamp} (group ${groupName})`;

    let rowIndex = null;
    for (let i = 0; i < existing.length; i++) {
        if (!existing[i]) {
            rowIndex = start + i;
            break;
        }
    }
    if (!rowIndex) {
        rowIndex = Math.max(sheet.getLastRow() + 1, start);
    }

    sheet.getRange(rowIndex, colKey).setValue(emailLower);
    sheet.getRange(rowIndex, MASTERS_TRANSFER_CFG.DEST_STATUS_COL).setValue(MASTERS_TRANSFER_CFG.STATUS_VALUE);
    sheet.getRange(rowIndex, 2).setValue("active");
    sheet.getRange(rowIndex, 3).setValue("IDLE");
    sheet.getRange(rowIndex, MASTERS_TRANSFER_CFG.DEST_COMMENT_COL).setValue(commentText);

    applyMastersRowDropdowns_(sheet, rowIndex);
}

function applyMastersRowDropdowns_(sheet, rowIndex) {
    const ruleA = sheet.getRange(2, 1).getDataValidation();
    const ruleC = sheet.getRange(2, 3).getDataValidation();

    const ruleA_fallback = SpreadsheetApp.newDataValidation()
        .requireValueInList(["FOUND", "NOT FOUND", "PENDING", "ERROR", "DELETED"], true)
        .build();
    const ruleC_fallback = SpreadsheetApp.newDataValidation()
        .requireValueInList(["IDLE"], true)
        .build();

    sheet.getRange(rowIndex, 1).setDataValidation(ruleA || ruleA_fallback);
    sheet.getRange(rowIndex, 3).setDataValidation(ruleC || ruleC_fallback);
}
