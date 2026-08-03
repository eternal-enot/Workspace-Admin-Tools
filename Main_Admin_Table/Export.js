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
      BACHELORS: "REPLACE_WITH_BACHELORS_SHEET_ID", // ID таблиці бакалаврів
      MASTERS: "REPLACE_WITH_MASTERS_SHEET_ID",   // ID таблиці магістрів
      PHD: "REPLACE_WITH_PHD_SHEET_ID",       // ID таблиці аспірантів
      STAFF: "REPLACE_WITH_STAFF_SHEET_ID",     // ID таблиці співробітників
    },

  // Target Sheet Layout (after reorder in target)
  DEST_EMAIL_COL: 6,     // F (User key/email input)
  DEST_START_ROW: 2,     // Start from row 2
  DEST_STATUS_COL: 1,    // A (Status)
  DEST_COMMENT_COL: 11,  // K (Comment)

  // PhD: єдиний аркуш замість окремих аркушів на групу
  PHD_SINGLE_SHEET_NAME: "PhD",
  PHD_GROUP_COL: 5,      // E — куди писати назву групи

  // Values to set in target
  STATUS_VALUE: "PENDING",
};

/**
 * Main function called from MenuItem
 * Переносить GEN_EMAIL у відповідну зовнішню таблицю (Бакалаври/Магістри/Співробітники):
 *  - визначає таблицю за типом аркуша та назвою групи
 *  - в колонку F (user key/email)
 *  - без дублікатів
 *  - оновлює статус (A) + коментар (K) тільки для доданих рядків
 */
function exportGenEmailsToGroupSheets_() {
  const ui = SpreadsheetApp.getUi();
  const srcSS = SpreadsheetApp.getActiveSpreadsheet();

  // Use configured sheet name if provided, else active sheet
  const srcSheet = EMAIL_EXPORT_CFG.SOURCE_SHEET_NAME
    ? srcSS.getSheetByName(EMAIL_EXPORT_CFG.SOURCE_SHEET_NAME)
    : srcSS.getActiveSheet();

  if (!srcSheet) throw new Error("Source sheet not found.");

  const srcSheetName = srcSheet.getName();

  // Validate we are on a valid source sheet
  if (srcSheetName !== APP_CONFIG.STUDENTS_SHEET_NAME &&
    srcSheetName !== APP_CONFIG.STAFF_SHEET_NAME &&
    srcSheetName !== APP_CONFIG.PHD_SHEET_NAME) {
    ui.alert(`❌ Будь ласка, запустіть це з аркуша "${APP_CONFIG.STUDENTS_SHEET_NAME}", "${APP_CONFIG.PHD_SHEET_NAME}" або "${APP_CONFIG.STAFF_SHEET_NAME}".`);
    return;
  }

  const lastRow = srcSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("Немає даних (є тільки заголовок).");
    return;
  }

  // Read the data range (assuming standard 15 columns as per Config.gc)
  const data = srcSheet.getRange(2, 1, lastRow - 1, 15).getValues();

  // Columns from Config.gc
  const colGroup = SHEET_COLS.groupOrDept - 1;   // 6 (0-based)
  const colEmail = SHEET_COLS.genEmail - 1;      // 8 (0-based)
  const colStatus = SHEET_COLS.status - 1;       // 12 (0-based)

  // group -> Set(emails lower)
  const grouped = new Map();

  for (let i = 0; i < data.length; i++) {
    const r = data[i];

    const groupRaw = String(r[colGroup] || "").trim();
    const email = String(r[colEmail] || "").trim();

    if (!groupRaw || !email) continue;

    // Filter by status if requested
    if (EMAIL_EXPORT_CFG.ONLY_IF_SOURCE_STATUS_IS) {
      const st = String(r[colStatus] || "").trim().toUpperCase();
      if (st !== EMAIL_EXPORT_CFG.ONLY_IF_SOURCE_STATUS_IS.toUpperCase()) continue;
    }

    const key = groupRaw;
    if (!grouped.has(key)) grouped.set(key, new Set());
    grouped.get(key).add(email.toLowerCase());
  }

  if (grouped.size === 0) {
    ui.alert("Немає GEN_EMAIL для переносу (або не проходить фільтр по статусу).");
    return;
  }

  // Cache open spreadsheets to avoid re-opening
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

  // Визначаємо, чи це PhD (одна таблиця — один аркуш)
  const isPhdExport = (srcSheetName === APP_CONFIG.PHD_SHEET_NAME);

  for (const [groupName, emailSetLower] of grouped.entries()) {
    // 1. Determine Target Spreadsheet ID
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
      // PhD: завжди пишемо на єдиний аркуш, групу — в колонку E
      const phdSheet = tgtSS.getSheetByName(EMAIL_EXPORT_CFG.PHD_SINGLE_SHEET_NAME);
      if (!phdSheet) {
        missingSheets.push(`${groupName} (sheet "${EMAIL_EXPORT_CFG.PHD_SINGLE_SHEET_NAME}" not found)`);
        continue;
      }
      const addedCount = insertEmailsIntoPhdSheet_(phdSheet, Array.from(emailSetLower), groupName);
      totalAdded += addedCount;
    } else {
      // Bachelors / Masters / Staff: аркуш = назва групи
      const targetSheet = findGroupSheet_(tgtSS, groupName);

      if (!targetSheet) {
        missingSheets.push(groupName);
        continue;
      }

      const isStaffExport = (targetId === EMAIL_EXPORT_CFG.TARGET_SPREADSHEETS.STAFF);
      const addedCount = insertEmailsIntoSheetWithStatus_(targetSheet, Array.from(emailSetLower), isStaffExport);
      totalAdded += addedCount;
    }
  }

  let msg = `✅ Готово. Додано адрес: ${totalAdded}`;

  if (missingID.length) {
    msg += `\n\n⚠️ Не знайдено ID таблиці (або помилка доступу) для:\n- ${missingID.join("\n- ")}`;
  }
  if (missingSheets.length) {
    msg += `\n\n⚠️ Не знайдено аркуші (в цільовій таблиці) для:\n- ${missingSheets.join("\n- ")}`;
  }

  ui.alert(msg);
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
  // We reuse the logic: if group contains 'мп' or 'мн' -> Master
  const g = String(groupName).toLowerCase();
  const isMaster = (g.includes('мп') || g.includes('мн'));

  if (isMaster) {
    return EMAIL_EXPORT_CFG.TARGET_SPREADSHEETS.MASTERS;
  } else {
    return EMAIL_EXPORT_CFG.TARGET_SPREADSHEETS.BACHELORS;
  }
}

/** ---- Core: insert emails into DEST_EMAIL_COL, avoid duplicates, update status/comment for inserted rows ---- */
function insertEmailsIntoSheetWithStatus_(sheet, emailsLower, isStaffExport = false) {
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

  // Update статус/коментар ТІЛЬКИ для вставлених рядків
  writePerRowBatched_(sheet, rowsInserted, EMAIL_EXPORT_CFG.DEST_STATUS_COL, EMAIL_EXPORT_CFG.STATUS_VALUE); // A
  writePerRowBatched_(sheet, rowsInserted, EMAIL_EXPORT_CFG.DEST_COMMENT_COL, commentText); // K
  
  // Дефолтні значення для випадаючих списків (Archive та trans_master)
  writePerRowBatched_(sheet, rowsInserted, 2, true); // B (Archive = TRUE)
  if (!isStaffExport) {
      writePerRowBatched_(sheet, rowsInserted, 3, "undefined"); // C (trans_master = undefined)
  }
  
  // Додаємо власне випадаючі списки (Data Validation) для нових рядків
  applyDropdownsToInsertedRows_(sheet, rowsInserted, isStaffExport);

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

  // Update статус/коментар ТІЛЬКИ для вставлених рядків
  writePerRowBatched_(sheet, rowsInserted, EMAIL_EXPORT_CFG.DEST_STATUS_COL, EMAIL_EXPORT_CFG.STATUS_VALUE); // A
  writePerRowBatched_(sheet, rowsInserted, EMAIL_EXPORT_CFG.DEST_COMMENT_COL, commentText); // K
  
  // Дефолтні значення для випадаючих списків
  writePerRowBatched_(sheet, rowsInserted, 2, true); // B (Archive = TRUE)
  writePerRowBatched_(sheet, rowsInserted, 3, "undefined"); // C (trans_master = undefined)

  // Додаємо власне випадаючі списки
  applyDropdownsToInsertedRows_(sheet, rowsInserted, false);

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
 * Перетворює текстові клітинки в колонках A, B, C у Data Validation списки 
 * для визначених рядків. В ідеалі пробує скопіювати форматування з рядка 2,
 * якщо його там немає — створює жорстко задані списки (Hardcoded Fallback).
 */
function applyDropdownsToInsertedRows_(sheet, rowsInserted, isStaffExpert) {
    if (!rowsInserted || rowsInserted.length === 0) return;
    
    // Більшість часу ці рядки йдуть підряд, тому знаходимо загальний діапазон.
    const firstRow = Math.min(...rowsInserted);
    const lastRow = Math.max(...rowsInserted);
    const numRows = lastRow - firstRow + 1;
    
    // Базові правила, якщо в рядку 2 порожньо:
    const ruleA_fallback = SpreadsheetApp.newDataValidation().requireValueInList(["FOUND", "NOT FOUND", "PENDING", "ERROR", "DELETED"], true).build();
    const ruleB_fallback = SpreadsheetApp.newDataValidation().requireValueInList(["TRUE", "FALSE"], true).build();
    const ruleC_fallback = SpreadsheetApp.newDataValidation().requireValueInList(["TRUE", "FALSE", "undefined"], true).build();

    // Колонка A
    let ruleA = sheet.getRange(2, 1).getDataValidation();
    sheet.getRange(firstRow, 1, numRows, 1).setDataValidation(ruleA || ruleA_fallback);
    
    // Колонка B
    let ruleB = sheet.getRange(2, 2).getDataValidation();
    sheet.getRange(firstRow, 2, numRows, 1).setDataValidation(ruleB || ruleB_fallback);
    
    // Колонка C
    if (!isStaffExpert) {
        let ruleC = sheet.getRange(2, 3).getDataValidation();
        sheet.getRange(firstRow, 3, numRows, 1).setDataValidation(ruleC || ruleC_fallback);
    }
}
