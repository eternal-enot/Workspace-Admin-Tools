/**
 * ============================================================================
 *  FILE: Export.gc
 *  CONTAINS:
 *   - Logic for exporting generated emails to external Group sheets
 * ============================================================================
 */

/**
 * EMAIL EXPORT CONFIG
 */
const PrivateConfigExport = (typeof getPrivateConfig === 'function') ? getPrivateConfig() : {};

const EMAIL_EXPORT_CFG = {
  // Source configuration
  // We use the active sheet, assuming it's STUDENTS or STAFF
  SOURCE_SHEET_NAME: "",

  ONLY_IF_SOURCE_STATUS_IS: "CREATED", // Only export created accounts

  // Target Spreadsheet IDs - FILL THESE IN!
  TARGET_SPREADSHEETS: (PrivateConfigExport && PrivateConfigExport.EXPORT_TARGET_SPREADSHEETS)
    ? PrivateConfigExport.EXPORT_TARGET_SPREADSHEETS
    : {
      BACHELORS: "REPLACE_WITH_BACHELORS_SHEET_ID", // Bachelors spreadsheet ID
      MASTERS: "REPLACE_WITH_MASTERS_SHEET_ID",   // Masters spreadsheet ID
      PHD: "REPLACE_WITH_PHD_SHEET_ID",       // PhD spreadsheet ID
      STAFF: "REPLACE_WITH_STAFF_SHEET_ID",     // Staff spreadsheet ID
    },

  // Target Sheet Layout (after reorder in target)
  DEST_EMAIL_COL: 6,     // F (User key/email input)
  DEST_START_ROW: 2,     // Start from row 2
  DEST_STATUS_COL: 1,    // A (Status)
  DEST_COMMENT_COL: 11,  // K (Comment)

  // PhD: single sheet instead of one sheet per group
  PHD_SINGLE_SHEET_NAME: "PhD",
  PHD_GROUP_COL: 5,      // E — write group name here

  // Values to set in target
  STATUS_VALUE: "PENDING",
};

/**
 * Main function called from MenuItem
 * Exports GEN_EMAIL to the matching target spreadsheet (Bachelors/Masters/Staff):
 *  - picks the spreadsheet by sheet type and group name
 *  - into column F (user key/email)
 *  - without duplicates
 *  - updates status (A) + comment (K) only for newly added rows
 */
function exportGenEmailsToGroupSheets_() {
  const ui = SpreadsheetApp.getUi();
  const srcSS = SpreadsheetApp.getActiveSpreadsheet();

  const srcSheet = EMAIL_EXPORT_CFG.SOURCE_SHEET_NAME
    ? srcSS.getSheetByName(EMAIL_EXPORT_CFG.SOURCE_SHEET_NAME)
    : srcSS.getActiveSheet();

  if (!srcSheet) throw new Error("Source sheet not found.");

  const srcSheetName = srcSheet.getName();

  if (srcSheetName !== APP_CONFIG.STUDENTS_SHEET_NAME &&
    srcSheetName !== APP_CONFIG.STAFF_SHEET_NAME &&
    srcSheetName !== APP_CONFIG.PHD_SHEET_NAME) {
    ui.alert(`❌ Будь ласка, запустіть це з аркуша "${APP_CONFIG.STUDENTS_SHEET_NAME}", "${APP_CONFIG.PHD_SHEET_NAME}" або "${APP_CONFIG.STAFF_SHEET_NAME}".`);
    return;
  }

  const res = exportGenEmailsFromSheetSilent_(srcSheetName);
  if (res.error) {
    ui.alert(`❌ ${res.error}`);
    return;
  }

  if (res.added === 0 && !res.missingID.length && !res.missingSheets.length) {
    ui.alert("Немає GEN_EMAIL для переносу (або не проходить фільтр по статусу).");
    return;
  }

  let msg = `✅ Готово. Додано адрес: ${res.added}`;

  if (res.missingID.length) {
    msg += `\n\n⚠️ Не знайдено ID таблиці (або помилка доступу) для:\n- ${res.missingID.join("\n- ")}`;
  }
  if (res.missingSheets.length) {
    msg += `\n\n⚠️ Не знайдено аркуші (в цільовій таблиці) для:\n- ${res.missingSheets.join("\n- ")}`;
  }

  ui.alert(msg);
}

/**
 * Silent export: pushes CREATED accounts from a source sheet to target spreadsheets.
 * @param {string} srcSheetName - STUDENTS, PHD, or STAFF sheet name
 * @param {Object} [options]
 * @param {string[]} [options.emailsLower] - if set, export only these emails (e.g. just deployed)
 * @returns {{ added: number, missingSheets: string[], missingID: string[], error?: string }}
 */
function exportGenEmailsFromSheetSilent_(srcSheetName, options) {
  options = options || {};
  const emailFilter = options.emailsLower
    ? new Set(options.emailsLower.map(e => String(e || "").trim().toLowerCase()).filter(Boolean))
    : null;

  const srcSS = SpreadsheetApp.getActiveSpreadsheet();
  const srcSheet = srcSS.getSheetByName(srcSheetName);
  if (!srcSheet) {
    return { added: 0, missingSheets: [], missingID: [], error: `Sheet not found: ${srcSheetName}` };
  }

  if (srcSheetName !== APP_CONFIG.STUDENTS_SHEET_NAME &&
    srcSheetName !== APP_CONFIG.STAFF_SHEET_NAME &&
    srcSheetName !== APP_CONFIG.PHD_SHEET_NAME) {
    return { added: 0, missingSheets: [], missingID: [], error: `Invalid source sheet: ${srcSheetName}` };
  }

  const lastRow = srcSheet.getLastRow();
  if (lastRow < 2) {
    return { added: 0, missingSheets: [], missingID: [] };
  }

  const data = srcSheet.getRange(2, 1, lastRow - 1, 15).getValues();

  const colGroup = SHEET_COLS.groupOrDept - 1;
  const colEmail = SHEET_COLS.genEmail - 1;
  const colStatus = SHEET_COLS.status - 1;

  const grouped = new Map();

  for (let i = 0; i < data.length; i++) {
    const r = data[i];

    const groupRaw = String(r[colGroup] || "").trim();
    const email = String(r[colEmail] || "").trim();

    if (!groupRaw || !email) continue;

    const emailLower = email.toLowerCase();
    if (emailFilter && !emailFilter.has(emailLower)) continue;

    if (EMAIL_EXPORT_CFG.ONLY_IF_SOURCE_STATUS_IS) {
      const st = String(r[colStatus] || "").trim().toUpperCase();
      if (st !== EMAIL_EXPORT_CFG.ONLY_IF_SOURCE_STATUS_IS.toUpperCase()) continue;
    }

    const key = groupRaw;
    if (!grouped.has(key)) grouped.set(key, new Set());
    grouped.get(key).add(emailLower);
  }

  if (grouped.size === 0) {
    return { added: 0, missingSheets: [], missingID: [] };
  }

  const openSpreadsheets = {};
  function getSpreadsheetById(id) {
    if (!id || id.includes("REPLACE")) return null;
    if (!openSpreadsheets[id]) {
      try {
        openSpreadsheets[id] = SpreadsheetApp.openById(id);
      } catch (e) {
        console.warn(`Could not open spreadsheet ${id}: ${e}`);
        return null;
      }
    }
    return openSpreadsheets[id];
  }

  const missingSheets = [];
  const missingID = [];
  let totalAdded = 0;

  const isPhdExport = (srcSheetName === APP_CONFIG.PHD_SHEET_NAME);

  for (const [groupName, emailSetLower] of grouped.entries()) {
    const targetId = determineTargetSpreadsheetId_(srcSheetName, groupName);

    if (!targetId || targetId.includes("REPLACE")) {
      missingID.push(groupName);
      continue;
    }

    const tgtSS = getSpreadsheetById(targetId);
    if (!tgtSS) {
      missingID.push(`${groupName} (Bad ID: ${targetId})`);
      continue;
    }

    if (isPhdExport) {
      const phdSheet = tgtSS.getSheetByName(EMAIL_EXPORT_CFG.PHD_SINGLE_SHEET_NAME);
      if (!phdSheet) {
        missingSheets.push(`${groupName} (sheet "${EMAIL_EXPORT_CFG.PHD_SINGLE_SHEET_NAME}" not found)`);
        continue;
      }
      totalAdded += insertEmailsIntoPhdSheet_(phdSheet, Array.from(emailSetLower), groupName);
    } else {
      const targetSheet = findGroupSheet_(tgtSS, groupName);
      if (!targetSheet) {
        missingSheets.push(groupName);
        continue;
      }
      totalAdded += insertEmailsIntoSheetWithStatus_(targetSheet, Array.from(emailSetLower));
    }
  }

  return { added: totalAdded, missingSheets, missingID };
}

/**
 * Logic to choose which spreadsheet to use
 */
function determineTargetSpreadsheetId_(sourceSheetName, groupName) {
  // If we are processing STAFF sheet -> always Staff Spreadsheet
  if (sourceSheetName === APP_CONFIG.STAFF_SHEET_NAME) {
    return EMAIL_EXPORT_CFG.TARGET_SPREADSHEETS.STAFF;
  }

  // If we are processing PHD sheet -> always PHD Spreadsheet
  if (sourceSheetName === APP_CONFIG.PHD_SHEET_NAME) {
    return EMAIL_EXPORT_CFG.TARGET_SPREADSHEETS.PHD;
  }

  // If we are processing STUDENTS sheet -> Check if Master or Bachelor
  // Masters groups contain 'мп' or 'мн'
  const g = String(groupName).toLowerCase();
  const isMaster = (g.includes('мп') || g.includes('мн'));

  if (isMaster) {
    return EMAIL_EXPORT_CFG.TARGET_SPREADSHEETS.MASTERS;
  } else {
    return EMAIL_EXPORT_CFG.TARGET_SPREADSHEETS.BACHELORS;
  }
}

/** ---- Core: insert emails into DEST_EMAIL_COL, avoid duplicates, update status/comment for inserted rows ---- */
function insertEmailsIntoSheetWithStatus_(sheet, emailsLower) {
  const lastRow = sheet.getLastRow();
  const start = EMAIL_EXPORT_CFG.DEST_START_ROW;

  const colKey = EMAIL_EXPORT_CFG.DEST_EMAIL_COL; // F

  let existing = [];
  if (lastRow >= start) {
    existing = sheet
      .getRange(start, colKey, lastRow - start + 1, 1)
      .getValues()
      .map(r => String(r[0] || "").trim());
  }

  const existingSet = new Set(existing.filter(Boolean).map(x => x.toLowerCase()));

  const toAddLower = [];
  for (const eLower of emailsLower) {
    if (!existingSet.has(eLower)) toAddLower.push(eLower);
  }
  if (toAddLower.length === 0) return 0;

  const now = new Date();
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const stamp = Utilities.formatDate(now, tz, "yyyy-MM-dd HH:mm:ss");
  const commentText = `Added automatically: ${stamp}`;

  // Fill gaps first, then append
  const emptyRowIndexes = [];
  for (let i = 0; i < existing.length; i++) {
    if (!existing[i]) emptyRowIndexes.push(start + i);
  }

  const rowsInserted = [];
  let cursor = 0;

  // 1) Fill empty rows
  for (; cursor < toAddLower.length && cursor < emptyRowIndexes.length; cursor++) {
    const rowIndex = emptyRowIndexes[cursor];
    const emailLower = toAddLower[cursor];
    sheet.getRange(rowIndex, colKey).setValue(emailLower);
    rowsInserted.push(rowIndex);
  }

  // 2) Append remaining
  if (cursor < toAddLower.length) {
    const remaining = toAddLower.slice(cursor);
    const appendStartRow = Math.max(sheet.getLastRow() + 1, start);

    const values = remaining.map(e => [e]);
    sheet.getRange(appendStartRow, colKey, values.length, 1).setValues(values);

    for (let k = 0; k < remaining.length; k++) {
      rowsInserted.push(appendStartRow + k);
    }
  }

  // Update status/comment ONLY for inserted rows
  writePerRowBatched_(sheet, rowsInserted, EMAIL_EXPORT_CFG.DEST_STATUS_COL, EMAIL_EXPORT_CFG.STATUS_VALUE); // A
  writePerRowBatched_(sheet, rowsInserted, EMAIL_EXPORT_CFG.DEST_COMMENT_COL, commentText); // K
  
  // Default values for new schema: B = record_type, C = actions
  writePerRowBatched_(sheet, rowsInserted, 2, "active"); // B (record_type)
  writePerRowBatched_(sheet, rowsInserted, 3, "IDLE"); // C (actions)
  
  // Add Data Validation dropdowns for new rows
  applyDropdownsToInsertedRows_(sheet, rowsInserted);

  return rowsInserted.length;
}

/** Write same value into given column for list of row indices (batched into contiguous blocks). */
function writePerRowBatched_(sheet, rowIdxs, col, value) {
  if (!rowIdxs.length) return;
  const sorted = rowIdxs.slice().sort((a, b) => a - b);

  let blockStart = sorted[0];
  let blockEnd = sorted[0];

  const flush = () => {
    const n = blockEnd - blockStart + 1;
    const vals = Array.from({ length: n }, () => [value]);
    sheet.getRange(blockStart, col, n, 1).setValues(vals);
  };

  for (let i = 1; i < sorted.length; i++) {
    const r = sorted[i];
    if (r === blockEnd + 1) {
      blockEnd = r;
    } else {
      flush();
      blockStart = r;
      blockEnd = r;
    }
  }
  flush();
}

/** ---- PhD: insert emails + group name into a single sheet ---- */
function insertEmailsIntoPhdSheet_(sheet, emailsLower, groupName) {
  const lastRow = sheet.getLastRow();
  const start = EMAIL_EXPORT_CFG.DEST_START_ROW;
  const colKey = EMAIL_EXPORT_CFG.DEST_EMAIL_COL; // F
  const colGroup = EMAIL_EXPORT_CFG.PHD_GROUP_COL; // E

  let existing = [];
  if (lastRow >= start) {
    existing = sheet
      .getRange(start, colKey, lastRow - start + 1, 1)
      .getValues()
      .map(r => String(r[0] || "").trim());
  }

  const existingSet = new Set(existing.filter(Boolean).map(x => x.toLowerCase()));

  const toAddLower = [];
  for (const eLower of emailsLower) {
    if (!existingSet.has(eLower)) toAddLower.push(eLower);
  }
  if (toAddLower.length === 0) return 0;

  const now = new Date();
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const stamp = Utilities.formatDate(now, tz, "yyyy-MM-dd HH:mm:ss");
  const commentText = `Added automatically: ${stamp}`;

  // Fill gaps first, then append
  const emptyRowIndexes = [];
  for (let i = 0; i < existing.length; i++) {
    if (!existing[i]) emptyRowIndexes.push(start + i);
  }

  const rowsInserted = [];
  let cursor = 0;

  // 1) Fill empty rows
  for (; cursor < toAddLower.length && cursor < emptyRowIndexes.length; cursor++) {
    const rowIndex = emptyRowIndexes[cursor];
    const emailLower = toAddLower[cursor];
    sheet.getRange(rowIndex, colKey).setValue(emailLower);
    sheet.getRange(rowIndex, colGroup).setValue(groupName);
    rowsInserted.push(rowIndex);
  }

  // 2) Append remaining
  if (cursor < toAddLower.length) {
    const remaining = toAddLower.slice(cursor);
    const appendStartRow = Math.max(sheet.getLastRow() + 1, start);

    const emailValues = remaining.map(e => [e]);
    sheet.getRange(appendStartRow, colKey, emailValues.length, 1).setValues(emailValues);

    const groupValues = remaining.map(() => [groupName]);
    sheet.getRange(appendStartRow, colGroup, groupValues.length, 1).setValues(groupValues);

    for (let k = 0; k < remaining.length; k++) {
      rowsInserted.push(appendStartRow + k);
    }
  }

  // Update status/comment ONLY for inserted rows
  writePerRowBatched_(sheet, rowsInserted, EMAIL_EXPORT_CFG.DEST_STATUS_COL, EMAIL_EXPORT_CFG.STATUS_VALUE); // A
  writePerRowBatched_(sheet, rowsInserted, EMAIL_EXPORT_CFG.DEST_COMMENT_COL, commentText); // K
  
  // Default values for new schema: B = record_type, C = actions
  writePerRowBatched_(sheet, rowsInserted, 2, "active"); // B (record_type)
  writePerRowBatched_(sheet, rowsInserted, 3, "IDLE"); // C (actions)

  // Add Data Validation dropdowns
  applyDropdownsToInsertedRows_(sheet, rowsInserted);

  return rowsInserted.length;
}

/** Try to find target sheet by group name with a couple of variants */
function findGroupSheet_(ss, groupRaw) {
  const candidates = candidateGroupSheetNames_(groupRaw);
  for (const name of candidates) {
    const sh = ss.getSheetByName(name);
    if (sh) return sh;
  }
  return null;
}

function candidateGroupSheetNames_(groupRaw) {
  const g = String(groupRaw || "").trim();
  if (!g) return [];
  const u = g.toUpperCase();

  const candidates = [g, u];

  // variant: uppercase but mp/mn lowercase (common pattern for master groups)
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

/**
 * Copies Data Validation from row 2 onto inserted rows.
 * Falls back to sensible defaults when row 2 has no validation yet.
 */
function applyDropdownsToInsertedRows_(sheet, rowsInserted) {
    if (!rowsInserted || rowsInserted.length === 0) return;
    
    const firstRow = Math.min(...rowsInserted);
    const lastRow = Math.max(...rowsInserted);
    const numRows = lastRow - firstRow + 1;
    
    const ruleA_fallback = SpreadsheetApp.newDataValidation()
        .requireValueInList(["FOUND", "NOT FOUND", "PENDING", "ERROR", "DELETED"], true)
        .build();
    const ruleC_fallback = SpreadsheetApp.newDataValidation()
        .requireValueInList(["IDLE"], true)
        .build();

    const ruleA = sheet.getRange(2, 1).getDataValidation();
    sheet.getRange(firstRow, 1, numRows, 1).setDataValidation(ruleA || ruleA_fallback);
    
    const ruleC = sheet.getRange(2, 3).getDataValidation();
    sheet.getRange(firstRow, 3, numRows, 1).setDataValidation(ruleC || ruleC_fallback);
}
