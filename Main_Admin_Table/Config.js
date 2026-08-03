/**
 * ============================================================================
 *  FILE: Config.gc
 *  CONTAINS:
 *   - Global Configuration (APP_CONFIG, CONSTANTS)
 *   - Menu & UI
 *   - Shared Utilities (Email, Password, Transliteration, Sheet Helpers, API)
 *   - Common Logic (Transfer from Form)
 * ============================================================================
 */

/** =========================
 *  GLOBAL CONFIG
 *  ========================= */
const PrivateConfig = (typeof getPrivateConfig === 'function') ? getPrivateConfig() : {};

const APP_CONFIG = {
  DOMAIN: PrivateConfig.DOMAIN || 'example.com',
  ADMIN_TEST_EMAIL: PrivateConfig.ADMIN_TEST_EMAIL || 'admin@example.com',

  FORM_SHEET_NAME: 'Відповіді форми (1)',
  STUDENTS_SHEET_NAME: 'STUDENTS',
  PHD_SHEET_NAME: 'PHD',
  STAFF_SHEET_NAME: 'STAFF',

  // Form columns (1-based)
  FORM_COL_ROLE: 2,          // B: "Студент / аспірант" | "Викладач / співробітник"
  FORM_COL_RECOVERY_HINT: 3, // C: "Якщо ви вже маєте..." (Recovery Email Hint)
  FORM_COL_NAME_UA: 4,       // D
  FORM_COL_SURNAME_UA: 5,    // E
  FORM_COL_PATR_UA: 6,       // F
  FORM_COL_NAME_EN: 7,       // G
  FORM_COL_SURNAME_EN: 8,    // H
  FORM_COL_PERSONAL_EMAIL: 9,// I
  FORM_COL_GROUP: 10,        // J: group (students)
  FORM_COL_DEPT: 11,         // K: кафедра (staff)

  // Students base OU
  STUDENTS_BASE_OU: '/2. Факультети/ФБМІ',

  // Rows with these statuses are treated as "finalized"
  FINAL_STATUSES: ['CREATED', 'EXISTS'],

  // Behavior
  SLEEP_MS_BETWEEN_CALLS: 150,
  SEND_CREDENTIALS_EMAIL: true,
  AUTO_DEPLOY_ENABLED: false,
};

const SHEET_COLS = {
  // Fixed layout A..O (1..15) for BOTH sheets
  mode: 1,            // A: ACCOUNT_MODE
  nameUa: 2,          // B
  surnameUa: 3,       // C
  patrUa: 4,          // D
  nameEn: 5,          // E
  surnameEn: 6,       // F
  groupOrDept: 7,     // G (group for students, кафедра for staff)
  personalEmail: 8,   // H
  genEmail: 9,        // I
  genOu: 10,          // J
  genRecovery: 11,    // K
  genPassword: 12,    // L
  status: 13,         // M
  createdAt: 14,      // N
  error: 15,          // O
};

const STUDENTS_HEADERS = [
  'ACCOUNT_MODE',  // A
  "Імʼя укр",      // B
  "прізвище укр",  // C
  "по батькові укр", // D
  "імʼя англ",     // E
  "прізвище англ", // F
  "група",         // G
  "особиста пошта",// H
  'GEN_EMAIL',     // I
  'GEN_OU',        // J
  'GEN_RECOVERY_EMAIL', // K
  'GEN_PASSWORD',  // L
  'ACCOUNT_STATUS',// M
  'CREATED_AT',    // N
  'ERROR',         // O
];

const STAFF_HEADERS = [
  'ACCOUNT_MODE',  // A
  "Імʼя укр",      // B
  "прізвище укр",  // C
  "по батькові укр", // D
  "імʼя англ",     // E
  "прізвище англ", // F
  "кафедра",       // G
  "особиста пошта",// H
  'GEN_EMAIL',     // I
  'GEN_OU',        // J
  'GEN_RECOVERY_EMAIL', // K
  'GEN_PASSWORD',  // L
  'ACCOUNT_STATUS',// M
  'CREATED_AT',    // N
  'ERROR',         // O
];

const PHD_HEADERS = [
  'ACCOUNT_MODE',  // A
  "Імʼя укр",      // B
  "прізвище укр",  // C
  "по батькові укр", // D
  "імʼя англ",     // E
  "прізвище англ", // F
  "група",         // G
  "особиста пошта",// H
  'GEN_EMAIL',     // I
  'GEN_OU',        // J
  'GEN_RECOVERY_EMAIL', // K
  'GEN_PASSWORD',  // L
  'ACCOUNT_STATUS',// M
  'CREATED_AT',    // N
  'ERROR',         // O
];

/** =========================
 *  MENU
 *  ========================= */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Accounts')
    .addItem('Setup/ensure ALL sheets (STUDENTS/PHD/STAFF)', 'setupTargetSheets_')
    .addSeparator()
    .addSeparator()
    .addItem('TRANSFER from Form -> TARGETS', 'transferFromFormResponses_')
    .addSeparator()
    .addItem('⚙️ Setup: Enable email notifications (Trigger)', 'setupFormSubmitTrigger')
    .addSeparator()
    .addItem('🐞 DEBUG: Test Search Strategies', 'debugTestSearch_')
    .addItem('🐞 DEBUG: Inspect User Info', 'debugShowUserInfo_')
    .addItem('🐞 DEBUG: Test Email in Cache (Why not rejected?)', 'debugTestEmailCache_')
    .addSeparator()

    .addItem('STUDENTS: Run PREVIEW (fill GEN_*)', 'runPreviewStudents_')
    .addItem('STUDENTS: Process rows (DEPLOY only)', 'processStudentsByMode_')
    .addSeparator()
    .addItem('PHD: Run PREVIEW (fill GEN_*)', 'runPreviewPhd_')
    .addItem('PHD: Process rows (DEPLOY only)', 'processPhdByMode_')
    .addSeparator()
    .addItem('STAFF: Run PREVIEW (fill GEN_*)', 'runPreviewStaff_')
    .addItem('STAFF: Process rows (DEPLOY only)', 'processStaffByMode_')
    .addSeparator()
    .addItem('Test: send credentials email to ADMIN (active sheet row)', 'testSendCredentialsEmailToAdmin_')
    .addSeparator()
    .addItem('Export GEN_EMAIL to group sheets', 'exportGenEmailsToGroupSheets_')
    .addToUi();
}

/** =========================
 *  SHEET SETUP & HELPERS
 *  ========================= */
function setupTargetSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s1 = ensureSheetWithHeaders_(ss, APP_CONFIG.STUDENTS_SHEET_NAME, STUDENTS_HEADERS);
  applyGlobalFormattingRules_(s1);

  const s2 = ensureSheetWithHeaders_(ss, APP_CONFIG.STAFF_SHEET_NAME, STAFF_HEADERS);
  applyGlobalFormattingRules_(s2);

  const s3 = ensureSheetWithHeaders_(ss, APP_CONFIG.PHD_SHEET_NAME, PHD_HEADERS);
  applyGlobalFormattingRules_(s3);

  SpreadsheetApp.getUi().alert('✅ ALL sheets ensured (A..O headers set) + Formatting Updated.');
}

function ensureSheetWithHeaders_(ss, sheetName, headers) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);

  // Ensure at least 15 columns
  if (sh.getMaxColumns() < headers.length) sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());

  // Set headers exactly in A..O
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Optional: freeze header row
  sh.setFrozenRows(1);

  return sh;
}

function appendRows15_(sheet, rows15) {
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows15.length, 15).setValues(rows15);
}

/** =========================
 *  TRANSFER: Form -> STUDENTS / STAFF
 *  ========================= */
function transferFromFormResponses_() {
  const ui = SpreadsheetApp.getUi();
  const res = transferFromFormResponsesSilent_();

  if (!res) {
    // Error or no sheet handled inside
    return;
  }

  if (res.empty) {
    ui.alert('Немає рядків для переносу (форма порожня).');
    return;
  }

  ui.alert(
    `✅ TRANSFER завершено.\n\n` +
    `STUDENTS додано: ${res.movedStudents}\n` +
    `PHD додано: ${res.movedPhd}\n` +
    `STAFF додано: ${res.movedStaff}\n` +
    `Помилок: ${res.errors}\n\n` +
    `Далі: запусти PREVIEW на потрібному аркуші.`
  );
}

function transferFromFormResponsesSilent_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName(APP_CONFIG.FORM_SHEET_NAME);
  if (!formSheet) {
    console.error(`❌ Не знайдено аркуш форми: "${APP_CONFIG.FORM_SHEET_NAME}"`);
    return null;
  }

  // Ensure target sheets exist + headers
  // REMOVED upfront calls: we will ensure them lazily below

  const lastRow = formSheet.getLastRow();
  const lastCol = formSheet.getLastColumn();
  if (lastRow < 2) {
    return { empty: true };
  }

  // Ensure transfer tracking columns in form sheet
  const transferCols = ensureFormTransferColumns_(formSheet);
  const n = lastRow - 1;

  // Read form data rows
  const data = formSheet.getRange(2, 1, n, lastCol).getValues();

  // Read transfer block (status/at/target/error) for all rows
  const trRange = formSheet.getRange(2, transferCols.statusCol, n, 4);
  const trVals = trRange.getValues();

  const studentsToAppend = [];
  const phdToAppend = [];
  const staffToAppend = [];

  // Fetch cache for recovery check
  const userCache = fetchCompleteUserCache_();

  let movedStudents = 0;
  let movedPhd = 0;
  let movedStaff = 0;
  let errors = 0;

  const processedDetails = [];

  for (let i = 0; i < n; i++) {
    const row = data[i];
    const tr = trVals[i];

    const trStatus = String(tr[0] || '').trim().toUpperCase();
    if (trStatus === 'TRANSFERRED') continue;

    const roleRaw = String(row[APP_CONFIG.FORM_COL_ROLE - 1] || '').trim();
    const roleU = roleRaw.toUpperCase();

    const nameUa = String(row[APP_CONFIG.FORM_COL_NAME_UA - 1] || '').trim();
    const surnameUa = String(row[APP_CONFIG.FORM_COL_SURNAME_UA - 1] || '').trim();
    const patrUa = String(row[APP_CONFIG.FORM_COL_PATR_UA - 1] || '').trim();
    const nameEn = String(row[APP_CONFIG.FORM_COL_NAME_EN - 1] || '').trim();
    const surnameEn = String(row[APP_CONFIG.FORM_COL_SURNAME_EN - 1] || '').trim();
    const personalEmail = String(row[APP_CONFIG.FORM_COL_PERSONAL_EMAIL - 1] || '').trim();
    const recoveryHint = String(row[APP_CONFIG.FORM_COL_RECOVERY_HINT - 1] || '').trim(); // C
    const groupVal = String(row[APP_CONFIG.FORM_COL_GROUP - 1] || '').trim(); // J
    const deptVal = String(row[APP_CONFIG.FORM_COL_DEPT - 1] || '').trim();  // K

    // Check for PhD based on 'ф' in group (e.g. ЗК-41ф, ЗФ-31ф), ignoring Dept code prefixes (ЗФ-...)
    const isPhdGroup = /\d.*[фФ]/.test(groupVal);

    // [UPDATED LOGIC]
    // If role is "Студент / аспірант", it contains BOTH words.
    // We should only force PhD if it's explicitly "Аспірант" (no "Student") OR if the group matches the regex.
    const isPhD = (roleU.includes('АСПІРАНТ') && !roleU.includes('СТУДЕНТ')) || isPhdGroup;

    // If matched PhD (by group or explicit role), then NOT a student.
    const isStudent = (roleU.includes('СТУДЕНТ')) && !isPhD;

    const isStaff = roleU.includes('ВИКЛАДАЧ') || roleU.includes('СПІВРОБІТ');

    if (!isStudent && !isPhD && !isStaff) {
      trVals[i] = ['ERROR', '', '', `Unknown role in column B: "${roleRaw}"`];
      errors++;
      continue;
    }

    const groupOrDept = (isStudent || isPhD) ? groupVal : deptVal;

    // Validate
    const missing = [];
    if (!nameUa) missing.push('name_ua');
    if (!surnameUa) missing.push('surname_ua');
    if (!personalEmail) missing.push('personal_email');
    if (!groupOrDept) missing.push((isStudent || isPhD) ? 'group' : 'кафедра');

    if (missing.length) {
      trVals[i] = ['ERROR', '', '', `Missing required fields: ${missing.join(', ')}`];
      errors++;
      continue;
    }

    let targetName = APP_CONFIG.STAFF_SHEET_NAME;
    if (isStudent) targetName = APP_CONFIG.STUDENTS_SHEET_NAME;
    else if (isPhD) targetName = APP_CONFIG.PHD_SHEET_NAME;

    // Build target row (A..O = 15 cols)
    const out = new Array(15).fill('');

    out[SHEET_COLS.mode - 1] = 'PREVIEW';

    // Recovery Check
    if (recoveryHint && looksLikeEmail_(recoveryHint) && looksLikeEmail_(personalEmail)) {
      const hint = recoveryHint.toLowerCase();
      const pEmail = personalEmail.toLowerCase();

      // Does hint exist as an active user/alias?
      if (userCache.corpMap.has(hint)) {
        const primary = userCache.corpMap.get(hint);

        // Does personal email link to this same primary?
        const linkedByRecovery = userCache.recoveryMap.get(pEmail);
        const linkedByWork = userCache.workMap.get(pEmail);

        if (linkedByRecovery === primary || linkedByWork === primary) {
          out[SHEET_COLS.mode - 1] = 'RECOVERY';
          // Optionally add info to status/error?
          // The user asked for "ACCOUNT_STATUS as RECOVERY". 
          // We are setting MODE to RECOVERY. Let's set Status too just in case.
          out[SHEET_COLS.status - 1] = 'RECOVERY_CANDIDATE';
          out[SHEET_COLS.error - 1] = `Matched existing user: ${primary}`;
        } else {
          // Mismatch: Hint exists, but personal email is not linked
          out[SHEET_COLS.error - 1] = `RECOVERY MISMTACH: Claimed ${hint}, but personal email (${pEmail}) is not linked to it.`;
        }
      } else {
        // Hint provided but not found
        out[SHEET_COLS.error - 1] = `RECOVERY HINT NOT FOUND: ${hint} is not an active user.`;
      }
    }

    out[SHEET_COLS.nameUa - 1] = nameUa;
    out[SHEET_COLS.surnameUa - 1] = surnameUa;
    out[SHEET_COLS.patrUa - 1] = patrUa;
    out[SHEET_COLS.nameEn - 1] = nameEn;
    out[SHEET_COLS.surnameEn - 1] = surnameEn;
    out[SHEET_COLS.groupOrDept - 1] = groupOrDept;
    out[SHEET_COLS.personalEmail - 1] = personalEmail;

    if (isStudent) {
      studentsToAppend.push(out);
      movedStudents++;
    } else if (isPhD) {
      phdToAppend.push(out);
      movedPhd++;
    } else {
      staffToAppend.push(out);
      movedStaff++;
    }

    trVals[i] = ['TRANSFERRED', new Date(), targetName, ''];

    // Add to details
    processedDetails.push({
      role: isStudent ? 'Student' : (isPhD ? 'PhD' : 'Staff'),
      name: `${nameUa} ${surnameUa}`,
      email: personalEmail,
      mode: out[SHEET_COLS.mode - 1], // PREVIEW vs RECOVERY
      hint: recoveryHint, // if any
      error: out[SHEET_COLS.error - 1] // if any (e.g. mismatch)
    });
  }

  // Append in batch
  if (studentsToAppend.length) {
    const studentsSheet = ensureSheetWithHeaders_(ss, APP_CONFIG.STUDENTS_SHEET_NAME, STUDENTS_HEADERS);
    const startRow = studentsSheet.getLastRow() + 1;
    appendRows15_(studentsSheet, studentsToAppend);
    setAccountModeDropdown_(studentsSheet, startRow, studentsToAppend.length);
  }
  if (phdToAppend.length) {
    const phdSheet = ensureSheetWithHeaders_(ss, APP_CONFIG.PHD_SHEET_NAME, PHD_HEADERS);
    const startRow = phdSheet.getLastRow() + 1;
    appendRows15_(phdSheet, phdToAppend);
    setAccountModeDropdown_(phdSheet, startRow, phdToAppend.length);
  }
  if (staffToAppend.length) {
    const staffSheet = ensureSheetWithHeaders_(ss, APP_CONFIG.STAFF_SHEET_NAME, STAFF_HEADERS);
    const startRow = staffSheet.getLastRow() + 1;
    appendRows15_(staffSheet, staffToAppend);
    setAccountModeDropdown_(staffSheet, startRow, staffToAppend.length);
  }

  // Write back transfer statuses
  trRange.setValues(trVals);

  return { movedStudents, movedPhd, movedStaff, errors, details: processedDetails };
}

function setAccountModeDropdown_(sheet, startRow, numRows) {
  // 1. Data Validation (Specific Range)
  const range = sheet.getRange(startRow, SHEET_COLS.mode, numRows, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['PREVIEW', 'DEPLOY', 'EXECUTED', 'REJECT', 'RECOVERY'], true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);

  // 2. Conditional Formatting (Global Row Highlighting)
  applyGlobalFormattingRules_(sheet);
}

function applyGlobalFormattingRules_(sheet) {
  // Ensure we cover the whole data area (up to max rows)
  const maxRows = sheet.getMaxRows();
  if (maxRows < 2) return;

  const fullRange = sheet.getRange(2, 1, maxRows - 1, 15); // A2:O<Max>

  // Clear OLD rules to prevent duplication/mess
  // We filter out any rule that might conflict or just reset all?
  // Since we manage this sheet, resetting is cleaner.
  sheet.clearConditionalFormatRules();

  const rules = [];

  // Helper to build row rule
  const addRule = (formula, color) => {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground(color)
      .setRanges([fullRange])
      .build();
  };

  // absolute column A, relative row starting at 2 -> $A2
  rules.push(addRule('=$A2="REJECT"', '#ffcdd2'));   // Red
  rules.push(addRule('=$A2="EXECUTED"', '#c8e6c9')); // Green
  rules.push(addRule('=$A2="DEPLOY"', '#bbdefb'));   // Blue
  rules.push(addRule('=$A2="RECOVERY"', '#ffe0b2')); // Orange
  rules.push(addRule('=$A2="PREVIEW"', '#fff9c4'));  // Yellow

  sheet.setConditionalFormatRules(rules);
}

function ensureFormTransferColumns_(formSheet) {
  const headerRow = 1;
  const lastCol = Math.max(1, formSheet.getLastColumn());
  const headers = formSheet.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());

  const need = ['TRANSFER_STATUS', 'TRANSFERRED_AT', 'TRANSFER_TARGET', 'TRANSFER_ERROR'];
  const idxs = {};

  for (let k = 0; k < need.length; k++) {
    const name = need[k];
    let idx = headers.findIndex(h => h === name);
    if (idx === -1) {
      const newCol = formSheet.getLastColumn() + 1;
      formSheet.getRange(headerRow, newCol).setValue(name);
      headers.push(name);
      idx = newCol - 1;
    }
    idxs[name] = idx + 1; // 1-based
  }

  return {
    statusCol: idxs['TRANSFER_STATUS'], // block start
  };
}


/** =========================
 *  COMMON UTILS
 *  ========================= */

function normalizeStatus_(v) {
  return String(v || '').trim().toUpperCase();
}

function hasNonEmptyValue_(v) {
  if (v === null || v === undefined) return false;
  if (v instanceof Date) return !isNaN(v.getTime());
  return String(v).trim() !== '';
}

function isRowFinalizedFixed_(row) {
  const status = normalizeStatus_(row[SHEET_COLS.status - 1]);
  const createdAt = row[SHEET_COLS.createdAt - 1];
  return (APP_CONFIG.FINAL_STATUSES || []).includes(status) || hasNonEmptyValue_(createdAt);
}

function writeRowResultFixed_(sheet, rowIndex, r) {
  if (r.status !== undefined) sheet.getRange(rowIndex, SHEET_COLS.status).setValue(r.status);
  if (r.createdAt !== undefined && r.createdAt) sheet.getRange(rowIndex, SHEET_COLS.createdAt).setValue(r.createdAt);
  if (r.error !== undefined) sheet.getRange(rowIndex, SHEET_COLS.error).setValue(r.error);
}

function normalizeNameTokenForEmail_(latin) {
  let s = String(latin || '').toLowerCase();
  s = s.replace(/[^a-z]/g, '');
  return s || 'x';
}

function looksLikeEmail_(s) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s || '').trim());
}

function escapeQueryValue_(value) {
  return String(value || '').replace(/'/g, "\\'");
}

function generateTempPassword_(length) {
  const len = Math.max(12, Number(length) || 14);
  const lower = 'abcdefghjkmnpqrstuvwxyz';
  const upper = 'ABCDEFGHJKMNPQRSTUVWXYZ';
  const digits = '23456789';
  const symbols = '!@#$%&*?';

  function pick(chars) {
    const idx = Math.floor(Math.random() * chars.length);
    return chars.charAt(idx);
  }

  const chars = [pick(lower), pick(upper), pick(digits), pick(symbols)];
  const all = lower + upper + digits + symbols;

  while (chars.length < len) chars.push(pick(all));

  for (let i = chars.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [chars[i], chars[j]] = [chars[j], chars[i]];
  }
  return chars.join('');
}

function sendCredentialsEmail_(toEmail, loginEmail, tempPassword, givenName, familyName, accountType) {
  const SENDER_NAME = 'Адміністратор edu.kpi.ua від ФБМІ';
  const LOGIN_LINK = 'https://myaccount.google.com';
  const LOGO_URL = 'https://fbmi.kpi.ua/wp-content/uploads/2025/09/fbmi_blue_logo.png';
  const SUBJECT = 'КПІ ім. Ігоря Сікорського: Дані для входу в Google Workspace for Education';

  const fullName = [givenName, familyName].filter(Boolean).join(' ').trim();
  const storageLimit = (accountType === 'staff') ? '50 ГБ' : '10 ГБ';

  // Plain Text Body (Fallback)
  const bodyPlain =
    fullName + ', вітаємо Вас у корпоративному середовищі Google Workspace for Education КПІ ім. Ігоря Сікорського!\n\n' +
    'Надаємо дані Вашого нового облікового запису:\n' +
    'Логін: ' + loginEmail + '\n' +
    'Пароль: ' + tempPassword + '\n' +
    'Увійти до облікового запису: ' + LOGIN_LINK + '\n\n' +
    'Що далі?\n' +
    '1) Під час першого входу змініть пароль на власний.\n' +
    '2) Базовий ліміт хмарного сховища Google Диску: ' + storageLimit + '.\n' +
    '3) Відповідно до внутрішніх правил Вашого закладу освіти, обмежені права на зміну інформації у профілі (фото та найменування). Для внесення змін звертайтесь до адміністратора.\n\n' +
    'Питання: звертайтесь до адміністраторів через Telegram-бот t.me/SikorskyDistance_bot\n' +
    'або до відповідальних на факультетах: https://medium.com/@s.admin-viasop/контакти-адміністраторів-корпоративного-середовища-google-workspace-for-educational-f80c0b8b9a93';

  // HTML Body
  const logoHtml = LOGO_URL
    ? '<div style="text-align:center;margin-top:12px;"><img src="' + LOGO_URL + '" alt="Лого університету" style="max-width:200px;height:auto;"></div>'
    : '';

  const htmlBody =
    '<!doctype html>' +
    '<html><body style="margin:0;padding:0;background:#ffffff;">' +
    '  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#ffffff;padding:24px 0;">' +
    '    <tr><td align="center">' +
    '      <table role="presentation" width="640" cellspacing="0" cellpadding="0" style="background:#ffffff;overflow:hidden;font-family:Arial,system-ui,-apple-system,\'Segoe UI\',Roboto,\'Noto Sans\',sans-serif;color:#111">' +
    '        <tr><td style="padding:24px 32px 8px;font-size:20px;font-weight:600;">' +
    '          КПІ ім. Ігоря Сікорського — Google Workspace for Education' +
    '        </td></tr>' +

    '        <tr><td style="padding:24px 32px 0;font-size:16px;line-height:1.6;">' +
    '          <p style="margin:0 0 12px 0;"><strong>' + escapeHtml_(fullName) + '</strong>, вітаємо Вас у корпоративному середовищі Google Workspace for Education КПІ ім. Ігоря Сікорського!</p>' +
    '          <p style="margin:0 0 12px 0;">Вам було створено новий обліковий запис:</p>' +
    '        </td></tr>' +

    '        <tr><td style="padding:8px 32px 0">' +
    '          <table width="100%" cellspacing="0" cellpadding="0" style="border:1px solid #e6e6e6;border-radius:8px;">' +
    '            <tr>' +
    '              <td style="padding:12px 16px;font-size:14px;background:#fafafa;width:40%;"><strong>Логін</strong></td>' +
    '              <td style="padding:12px 16px;font-size:14px;">' + escapeHtml_(loginEmail) + '</td>' +
    '            </tr>' +
    '            <tr>' +
    '              <td style="padding:12px 16px;font-size:14px;background:#fafafa;"><strong>Пароль</strong></td>' +
    '              <td style="padding:12px 16px;font-size:14px;">' + escapeHtml_(tempPassword) + '</td>' +
    '            </tr>' +
    '          </table>' +
    '        </td></tr>' +

    '        <tr><td style="padding:20px 32px 0" align="center">' +
    '          <a href="' + LOGIN_LINK + '" target="_blank" ' +
    '             style="display:inline-block;padding:12px 20px;border-radius:8px;background:#1a73e8;color:#ffffff;text-decoration:none;font-weight:600;border:1px solid #1a73e8;">' +
    '            Увійти до облікового запису' +
    '          </a>' +
    '        </td></tr>' +

    '        <tr><td style="padding:20px 32px 0;font-size:14px;line-height:1.6;color:#333">' +
    '          <p style="margin:0 0 10px 0;"><strong>Звертаємо Вашу увагу:</strong></p>' +
    '          <ol style="margin:0 0 0 18px;padding:0;">' +
    '            <li>Під час першого входу змініть пароль на власний.</li>' +
    '            <li>Базовий ліміт хмарного сховища Google Диску: ' + storageLimit + '.</li>' +
    '            <li>Відповідно до внутрішніх правил вашого закладу освіти, обмежені права на зміну інформації у профілі (фотографія та найменування). Для внесення змін звертайтесь до адміністратора.</li>' +
    '          </ol>' +
    '        </td></tr>' +

    '        <tr><td style="padding:16px 32px 16px;font-size:13px;color:#666;line-height:1.6;">' +
    '          Якщо у Вас виникнуть питання, звертайтесь до адміністраторів через ' +
    '          <a href="https://t.me/SikorskyDistance_bot" target="_blank" style="text-decoration:none;">бот</a> ' +
    '          або до <a href="https://medium.com/@s.admin-viasop/контакти-адміністраторів-корпоративного-середовища-google-workspace-for-educational-f80c0b8b9a93" target="_blank" style="text-decoration:none;">відповідальних адміністраторів на факультетах</a>.' +
    '        </td></tr>' +

    '        <tr><td style="padding:16px 32px 0;font-size:12px;color:#999;border-top:1px solid #eee">' +
    '          Це повідомлення згенеровано автоматично. Будь ласка, не відповідайте на нього.' +
    '        </td></tr>' +

    '        <tr><td style="padding:8px 32px 16px;">' + logoHtml + '</td></tr>' +
    '      </table>' +
    '    </td></tr>' +
    '  </table>' +
    '</body></html>';

  MailApp.sendEmail({
    to: toEmail,
    subject: SUBJECT,
    name: SENDER_NAME,
    htmlBody: htmlBody,
    body: bodyPlain
  });
}

function escapeHtml_(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function translitUaToLat_(text) {
  let s = String(text || '').trim();
  if (!s) return '';

  s = s.toLowerCase();
  s = s.replace(/[’'ʼ`ь]/gu, '');

  s = s.replace(/(^|[аеиіоуюяєїі])є/gu, '$1ye');
  s = s.replace(/(^|[аеиіоуюяєїі])ї/gu, '$1yi');
  s = s.replace(/(^|[аеиіоуюяєїі])ю/gu, '$1yu');
  s = s.replace(/(^|[аеиіоуюяєїі])я/gu, '$1ya');
  s = s.replace(/^й/gu, 'y');
  s = s.replace(/зг/gu, 'zgh');

  const map = {
    'є': 'ie', 'ї': 'i', 'й': 'i', 'ю': 'iu', 'я': 'ia',
    'ж': 'zh', 'х': 'kh', 'ц': 'ts', 'ч': 'ch', 'ш': 'sh', 'щ': 'shch',
    'а': 'a', 'б': 'b', 'в': 'v', 'г': 'h', 'ґ': 'g', 'д': 'd',
    'е': 'e', 'з': 'z', 'и': 'y', 'і': 'i', 'к': 'k', 'л': 'l',
    'м': 'm', 'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r', 'с': 's',
    'т': 't', 'у': 'u', 'ф': 'f',
  };

  let out = '';
  for (const ch of s) out += (map[ch] !== undefined) ? map[ch] : ch;

  out = out.replace(/(^|\s)([a-z])/g, (_, sep, letter) => sep + letter.toUpperCase());
  return out;
}

/** =========================
 *  API & CHECKING
 *  ========================= */
function userExistsByUserKey_(userKey) {
  try {
    AdminDirectory.Users.get(userKey);
    return true;
  } catch (e) {
    return !isNotFound_(e);
  }
}

function userExistsByDomainEmailOrAlias_(email) {
  try {
    const q = `email='${escapeQueryValue_(email)}'`;
    const res = AdminDirectory.Users.list({
      domain: APP_CONFIG.DOMAIN,
      query: q,
      maxResults: 1,
      orderBy: 'email',
    });
    return Array.isArray(res.users) && res.users.length > 0;
  } catch (e) {
    Logger.log('Users.list failed: %s', String(e));
    return false;
  }
}

/**
 * Checks if the given email exists as a Recovery Email (Personal) for any user in the domain.
 * This ensures no two users share the same personal email.
 */
/**
 * Fetches ALL users in the domain to build a lookup map:
 *   Map< personalEmail(lowercase) => primaryEmail >
 *
 * This is required because 'Users.list(query="recoveryEmail=...")'
 * is NOT supported by the API (returns Invalid Input or 0 results).
 */
/**
 * Builds a comprehensive cache of all users in the domain:
 * 1) Recovery Emails (Personal inputs -> Existing User)
 * 2) Work/Secondary Emails (Personal inputs -> Existing User)
 * 3) All Corporate Emails (Primary + Aliases) -> To check GEN_EMAIL collisions
 *
 * Returns:
 * {
 *   recoveryMap: Map<email, primaryEmail>,
 *   workMap: Map<email, primaryEmail>,
 *   corpMap: Map<email, primaryEmail> // Checks collision for GEN_EMAIL
 * }
 */
function fetchCompleteUserCache_() {
  const cache = {
    recoveryMap: new Map(),
    workMap: new Map(),
    corpMap: new Map() // includes primary and aliases
  };

  let pageToken = null;

  // Optional: show toast
  // SpreadsheetApp.getActiveSpreadsheet().toast('Завантаження бази користувачів...', 'Cache', 10);

  do {
    try {
      const res = AdminDirectory.Users.list({
        customer: 'my_customer',
        maxResults: 500,
        pageToken: pageToken,
        viewType: 'admin_view',
        fields: 'users(primaryEmail,recoveryEmail,emails,aliases),nextPageToken'
      });

      const users = res.users || [];
      for (const u of users) {
        const prim = u.primaryEmail.toLowerCase();

        // 1. Corporate (Primary)
        cache.corpMap.set(prim, prim);

        // 2. Corporate (Aliases)
        if (u.aliases && Array.isArray(u.aliases)) {
          for (const alias of u.aliases) {
            cache.corpMap.set(alias.toLowerCase(), prim);
          }
        }

        // 3. Recovery
        if (u.recoveryEmail) {
          cache.recoveryMap.set(u.recoveryEmail.trim().toLowerCase(), prim);
        }

        // 4. Other Emails (Work, Home, Custom)
        // These are often "working emails" added in Admin Console
        if (u.emails && Array.isArray(u.emails)) {
          for (const e of u.emails) {
            // e = { address: "...", type: "work"|"home"|"custom"|"primary" }
            if (e.type === 'primary') continue; // already handled

            const addr = e.address.trim().toLowerCase();
            cache.workMap.set(addr, prim);

            // Note: If an 'alias' is listed in emails[], it might be corporate. 
            // Usually aliases are in u.aliases, but sometimes in emails with type='alias'.
            // To be safe, if type exists and is 'alias', add to corpMap too.
            if (e.type === 'alias') {
              cache.corpMap.set(addr, prim);
            }
          }
        }
      }

      pageToken = res.nextPageToken;
    } catch (e) {
      Logger.log('fetchCompleteUserCache_ partial error: ' + e);
      break;
    }
  } while (pageToken);

  return cache;
}

// Deprecated/Removed single-search function as it doesn't work reliably
// function findUserByRecoveryEmail_(...) {}

function debugTestSearch_() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const response = ui.prompt('Debug Search Strategies', 'Введіть особисту пошту (recovery) для тесту:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const email = response.getResponseText().trim();
  if (!email) return;

  ss.toast(`Тестуємо пошук для '${email}'...`, 'DEBUG', 60);

  const results = [];
  const safeEmail = email.replace(/'/g, "\\'");

  // Strategy 1: Standard Quoted
  try {
    const q1 = `recoveryEmail='${safeEmail}'`;
    const res1 = AdminDirectory.Users.list({ domain: APP_CONFIG.DOMAIN, query: q1, viewType: 'admin_view' });
    results.push(`1) [${q1}] -> Found: ${res1.users ? res1.users.length : 0}`);
  } catch (e) { results.push(`1) Error: ${e}`); }

  // Strategy 2: Unquoted (if simple email)
  try {
    const q2 = `recoveryEmail:${email}`; // syntax recoveryEmail:value
    const res2 = AdminDirectory.Users.list({ domain: APP_CONFIG.DOMAIN, query: q2, viewType: 'admin_view' });
    results.push(`2) [${q2}] -> Found: ${res2.users ? res2.users.length : 0}`);
  } catch (e) { results.push(`2) Error: ${e}`); }

  // Strategy 3: Full Text (just email)
  try {
    const q3 = `${safeEmail}`;
    const res3 = AdminDirectory.Users.list({ domain: APP_CONFIG.DOMAIN, query: q3, viewType: 'admin_view' });
    let count = 0;
    if (res3.users) {
      // Filter manually to see if it's in recoveryEmail
      const exact = res3.users.filter(u => u.recoveryEmail === email);
      count = exact.length;
    }
    results.push(`3) [${q3}] -> API returned ${res3.users ? res3.users.length : 0}, Match recovery: ${count}`);
  } catch (e) { results.push(`3) Error: ${e}`); }

  ui.alert(`RESULTS for '${email}':\n\n` + results.join('\n\n'));
}

function debugShowUserInfo_() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Debug User Info', 'Введіть КОРПОРАТИВНУ пошту користувача (login@edu.kpi.ua), для якого перевіряємо дані:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const email = response.getResponseText().trim();
  if (!email) return;

  try {
    const user = AdminDirectory.Users.get(email, { viewType: 'admin_view' });
    const rec = user.recoveryEmail || '(не встановлено)';
    const aliases = (user.aliases || []).join(', ');
    const emails = (user.emails || []).map(e => `${e.address} (${e.type})`).join(', ');

    ui.alert(
      `INFO for: ${user.primaryEmail}\n\n` +
      `Recovery Email: ${rec}\n` +
      `Aliases: ${aliases}\n` +
      `All Emails: ${emails}\n` +
      `ID: ${user.id}`
    );
  } catch (e) {
    ui.alert(`ERROR: ${e}`);
  }
}

function isNotFound_(e) {
  const s = String(e || '');
  return s.includes('Resource Not Found') || s.includes('notFound') || s.includes('404');
}


/** =========================
 *  OTHER MENU ACTIONS
 *  ========================= */
function testSendCredentialsEmailToAdmin_() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const sheetName = sheet.getName();
  const allowedSheets = [APP_CONFIG.STUDENTS_SHEET_NAME, APP_CONFIG.PHD_SHEET_NAME, APP_CONFIG.STAFF_SHEET_NAME];
  if (!allowedSheets.includes(sheetName)) {
    ui.alert(`❌ Перейди на аркуш "${APP_CONFIG.STUDENTS_SHEET_NAME}", "${APP_CONFIG.PHD_SHEET_NAME}" або "${APP_CONFIG.STAFF_SHEET_NAME}" і вибери рядок з GEN_*.`);
    return;
  }

  const accountType = (sheetName === APP_CONFIG.STAFF_SHEET_NAME) ? 'staff' : 'student';

  const adminTo =
    String(APP_CONFIG.ADMIN_TEST_EMAIL || '').trim() ||
    String(Session.getActiveUser().getEmail() || '').trim();

  if (!adminTo) {
    ui.alert('Не задано ADMIN_TEST_EMAIL і не вдалося визначити email активного користувача.');
    return;
  }

  let rowIndex = 0;
  const ar = sheet.getActiveRange();
  if (ar && ar.getRow() >= 2) rowIndex = ar.getRow();
  if (!rowIndex) rowIndex = findFirstRowWithGeneratedFixed_(sheet);

  if (!rowIndex) {
    ui.alert('Не знайшов рядків із заповненими GEN_* (GEN_EMAIL, GEN_PASSWORD, GEN_RECOVERY_EMAIL). Спочатку запусти PREVIEW.');
    return;
  }

  const row = sheet.getRange(rowIndex, 1, 1, 15).getValues()[0];

  const genEmail = String(row[SHEET_COLS.genEmail - 1] || '').trim();
  const genPassword = String(row[SHEET_COLS.genPassword - 1] || '').trim();
  const genRecovery = String(row[SHEET_COLS.genRecovery - 1] || '').trim();

  const nameUa = String(row[SHEET_COLS.nameUa - 1] || '').trim();
  const surnameUa = String(row[SHEET_COLS.surnameUa - 1] || '').trim();

  if (!genEmail || !genPassword || !genRecovery) {
    ui.alert(`У рядку ${rowIndex} немає повних GEN_* даних. Запусти PREVIEW для цього рядка.`);
    return;
  }

  try {
    sendCredentialsEmail_(adminTo, genEmail, genPassword, nameUa, surnameUa, accountType);

    sheet.getRange(rowIndex, SHEET_COLS.status).setValue('ADMIN_EMAIL_TEST_SENT');
    sheet.getRange(rowIndex, SHEET_COLS.error).setValue('');

    ui.alert(
      `✅ Тестовий лист відправлено на: ${adminTo}\n\n` +
      `Рядок: ${rowIndex}\n` +
      `Логін: ${genEmail}\n` +
      `Recovery (довідково): ${genRecovery}`
    );
  } catch (e) {
    sheet.getRange(rowIndex, SHEET_COLS.error).setValue(String(e));
    ui.alert('❌ Помилка під час відправки: ' + String(e));
  }
}

function findFirstRowWithGeneratedFixed_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const n = lastRow - 1;
  const values = sheet.getRange(2, 1, n, 15).getValues();

  for (let i = 0; i < n; i++) {
    const row = values[i];
    const genEmail = String(row[SHEET_COLS.genEmail - 1] || '').trim();
    const genPassword = String(row[SHEET_COLS.genPassword - 1] || '').trim();
    const genRecovery = String(row[SHEET_COLS.genRecovery - 1] || '').trim();
    if (genEmail && genPassword && genRecovery) return i + 2;
  }
  return 0;
}

function debugTestEmailCache_() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Debug Cache', 'Enter email to check in cache maps:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const email = resp.getResponseText().trim().toLowerCase();

  ui.alert('Loading cache... please wait.');

  const cache = fetchCompleteUserCache_();

  let msg = `Results for "${email}":\n\n`;

  // 1. Recovery
  if (cache.recoveryMap.has(email)) {
    msg += `✅ FOUND in Recovery Map -> User: ${cache.recoveryMap.get(email)}\n`;
  } else {
    msg += `❌ NOT found in Recovery Map\n`;
  }

  // 2. Work
  if (cache.workMap.has(email)) {
    msg += `✅ FOUND in Work Map -> User: ${cache.workMap.get(email)}\n`;
  } else {
    msg += `❌ NOT found in Work Map\n`;
  }

  // 3. Corp
  if (cache.corpMap.has(email)) {
    msg += `✅ FOUND in Corp Map (Primary/Alias) -> User: ${cache.corpMap.get(email)}\n`;
  } else {
    msg += `❌ NOT found in Corp Map\n`;
  }

  ui.alert(msg);
}




