/**
 * ============================================================================
 *  FILE: Students.gc
 *  CONTAINS:
 *   - Logic specific for STUDENTS sheet
 *   - Preview (email generation, OU math)
 *   - Processing
 * ============================================================================
 */

/** =========================
 *  2) STUDENTS WRAPPERS
 *  ========================= */
function runPreviewStudents_() {
  const { count } = runPreviewStudentsSilent_();
  SpreadsheetApp.getUi().alert(`✅ PREVIEW завершено для "${APP_CONFIG.STUDENTS_SHEET_NAME}": оброблено ${count} рядків.`);
}

function runPreviewStudentsSilent_() {
  const sheetName = APP_CONFIG.STUDENTS_SHEET_NAME;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return 0;

    // Ensure headers
    ensureSheetWithHeaders_(ss, sheetName, STUDENTS_HEADERS);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return 0;

    const n = lastRow - 1;
    const values = sheet.getRange(2, 1, n, 15).getValues();
    let processedCount = 0;

    const rangesToHighlight = [];

    // Fetch all users once
    const userCache = fetchCompleteUserCache_();

    const results = [];

    for (let i = 0; i < n; i++) {
      const rowIndex = i + 2;
      const row = values[i];

      if (isRowFinalizedFixed_(row)) continue;

      let mode = String(row[SHEET_COLS.mode - 1] || '').trim().toUpperCase();
      if (!mode) {
        mode = 'PREVIEW';
        sheet.getRange(rowIndex, SHEET_COLS.mode).setValue('PREVIEW');
      }
      if (mode !== 'PREVIEW') continue;

      // Clear previous error
      sheet.getRange(rowIndex, SHEET_COLS.error).setValue('');

      const res = previewStudentRow_(sheet, rowIndex, row, userCache);
      if (res) results.push(res);

      processedCount++;

      rangesToHighlight.push(rowIndex);
    }

    // Highlight previewed rows
    if (rangesToHighlight.length > 0) {
      for (const rIdx of rangesToHighlight) {
        sheet.getRange(rIdx, 1, 1, 15).setBackground('#fff9c4'); // Light yellow for whole row
      }
    }

    return { count: processedCount, results: results };
  } finally {
    lock.releaseLock();
  }
}

function processStudentsByMode_() {
  const sheetName = APP_CONFIG.STUDENTS_SHEET_NAME;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);

    // Ensure headers
    ensureSheetWithHeaders_(ss, sheetName, STUDENTS_HEADERS);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const n = lastRow - 1;
    const rows = sheet.getRange(2, 1, n, 15).getValues();

    for (let i = 0; i < n; i++) {
      const rowIndex = i + 2;
      const row = rows[i];

      if (isRowFinalizedFixed_(row)) continue;

      let mode = String(row[SHEET_COLS.mode - 1] || '').trim().toUpperCase();
      if (!mode) {
        sheet.getRange(rowIndex, SHEET_COLS.mode).setValue('PREVIEW');
        mode = 'PREVIEW';
      }

      // Clear previous error
      sheet.getRange(rowIndex, SHEET_COLS.error).setValue('');

      if (mode === 'PREVIEW') {
        previewStudentRow_(sheet, rowIndex, row);
        continue;
      }

      if (mode === 'RECOVERY' || mode === 'REJECT') continue;

      if (mode !== 'DEPLOY') {
        writeRowResultFixed_(sheet, rowIndex, {
          status: 'SKIP',
          error: `Unknown ACCOUNT_MODE="${mode}". Use PREVIEW or DEPLOY.`,
        });
        continue;
      }

      // DEPLOY
      const genEmail = String(row[SHEET_COLS.genEmail - 1] || '').trim();
      const genOu = String(row[SHEET_COLS.genOu - 1] || '').trim();
      const genRecovery = String(row[SHEET_COLS.genRecovery - 1] || '').trim();
      const genPassword = String(row[SHEET_COLS.genPassword - 1] || '').trim();

      if (!genEmail || !genOu || !genRecovery || !genPassword) {
        writeRowResultFixed_(sheet, rowIndex, {
          status: 'ERROR',
          error: 'DEPLOY requires GEN_EMAIL, GEN_OU, GEN_RECOVERY_EMAIL, GEN_PASSWORD. Run PREVIEW first.',
        });
        continue;
      }

      // Existence checks
      if (userExistsByUserKey_(genEmail)) {
        sheet.getRange(rowIndex, SHEET_COLS.mode).setValue('REJECT');
        writeRowResultFixed_(sheet, rowIndex, {
          status: 'EXISTS',
          createdAt: new Date(),
          error: 'User already exists for generated email (primary or alias).',
        });
        continue;
      }

      if (looksLikeEmail_(genRecovery) && userExistsByDomainEmailOrAlias_(genRecovery)) {
        sheet.getRange(rowIndex, SHEET_COLS.mode).setValue('REJECT');
        writeRowResultFixed_(sheet, rowIndex, {
          status: 'EXISTS',
          createdAt: new Date(),
          error: 'Recovery/personal email already used as a user email/alias in this Workspace domain.',
        });
        continue;
      }

      const nameUa = String(row[SHEET_COLS.nameUa - 1] || '').trim();
      const surnameUa = String(row[SHEET_COLS.surnameUa - 1] || '').trim();

      if (!nameUa || !surnameUa) {
        writeRowResultFixed_(sheet, rowIndex, {
          status: 'ERROR',
          error: 'Missing name/surname in input columns for profile.',
        });
        continue;
      }

      const userResource = {
        primaryEmail: genEmail.toLowerCase(),
        name: { givenName: nameUa, familyName: surnameUa },
        password: genPassword,
        orgUnitPath: genOu,
        recoveryEmail: genRecovery,
        changePasswordAtNextLogin: true,
        emails: [
          { address: genRecovery, type: 'work', primary: false }
        ]
      };

      try {
        const created = AdminDirectory.Users.insert(userResource);

        writeRowResultFixed_(sheet, rowIndex, {
          status: 'CREATED',
          createdAt: new Date(),
          error: '',
        });

        // Mark as EXECUTED and clear background (success)
        sheet.getRange(rowIndex, SHEET_COLS.mode).setValue('EXECUTED');
        sheet.getRange(rowIndex, 1, 1, 15).setBackground(null);

        if (APP_CONFIG.SEND_CREDENTIALS_EMAIL) {
          sendCredentialsEmail_(genRecovery, created.primaryEmail || genEmail, genPassword, nameUa, surnameUa, 'student');
        }
      } catch (e) {
        sheet.getRange(rowIndex, SHEET_COLS.mode).setValue('REJECT');
        writeRowResultFixed_(sheet, rowIndex, {
          status: 'ERROR',
          error: String(e),
        });
      }

      Utilities.sleep(APP_CONFIG.SLEEP_MS_BETWEEN_CALLS);
    }

    SpreadsheetApp.getUi().alert(`✅ Обробка завершена для "${sheetName}". Створено лише рядки з ACCOUNT_MODE=DEPLOY.`);
  } finally {
    lock.releaseLock();
  }
}

/** =========================
 *  STUDENT PREVIEW LOGIC
 *  ========================= */
function previewStudentRow_(sheet, rowIndex, row, userCache) {
  const nameUa = String(row[SHEET_COLS.nameUa - 1] || '').trim();
  const surnameUa = String(row[SHEET_COLS.surnameUa - 1] || '').trim();
  const patrUa = String(row[SHEET_COLS.patrUa - 1] || '').trim();
  const groupOrDept = String(row[SHEET_COLS.groupOrDept - 1] || '').trim();
  const personalEmail = String(row[SHEET_COLS.personalEmail - 1] || '').trim();

  const missing = [];
  if (!nameUa) missing.push('name');
  if (!surnameUa) missing.push('surname');
  if (!personalEmail) missing.push('personal email');
  if (!groupOrDept) missing.push('group');

  if (missing.length) {
    const errorMsg = `Missing required fields: ${missing.join(', ')}.`;
    writeRowResultFixed_(sheet, rowIndex, {
      status: 'SKIP',
      error: errorMsg,
    });
    return { personalEmail, mode: 'SKIP', error: errorMsg };
  }

  // 1) Identity Check: Existing user by personal/recovery OR work email
  if (looksLikeEmail_(personalEmail)) {
    const pEmail = personalEmail.toLowerCase();

    // Check Recovery Map
    if (userCache && userCache.recoveryMap.has(pEmail)) {
      const existingUser = userCache.recoveryMap.get(pEmail);
      const errorMsg = `User already exists (matched recovery email): ${existingUser}`;
      sheet.getRange(rowIndex, SHEET_COLS.mode).setValue('REJECT');
      writeRowResultFixed_(sheet, rowIndex, {
        status: 'SKIP',
        error: errorMsg,
      });
      return { personalEmail, mode: 'REJECT', error: errorMsg };
    }

    // Check Work Map
    if (userCache && userCache.workMap.has(pEmail)) {
      const existingUser = userCache.workMap.get(pEmail);
      const errorMsg = `User already exists (matched work/secondary email): ${existingUser}`;
      sheet.getRange(rowIndex, SHEET_COLS.mode).setValue('REJECT');
      writeRowResultFixed_(sheet, rowIndex, {
        status: 'SKIP',
        error: errorMsg,
      });
      return { personalEmail, mode: 'REJECT', error: errorMsg };
    }
  }

  const ouObj = buildStudentOrgUnitPath_(groupOrDept);
  const genEmail = buildStudentPrimaryEmail_(surnameUa, nameUa, patrUa, ouObj.groupForEmailPart);
  const orgUnitPath = ouObj.orgUnitPath;

  // 2) Collision Check: Is GEN_EMAIL already taken?
  if (userCache && userCache.corpMap.has(genEmail)) {
    const errorMsg = `GEN_EMAIL CONFLICT: ${genEmail} is already active/alias. Manual intervention needed.`;
    sheet.getRange(rowIndex, SHEET_COLS.mode).setValue('REJECT');
    writeRowResultFixed_(sheet, rowIndex, {
      status: 'SKIP',
      error: errorMsg,
    });

    sheet.getRange(rowIndex, SHEET_COLS.genEmail).setValue(genEmail);
    sheet.getRange(rowIndex, SHEET_COLS.genOu).setValue(orgUnitPath);
    sheet.getRange(rowIndex, SHEET_COLS.genRecovery).setValue(personalEmail);
    return { personalEmail, mode: 'REJECT', error: errorMsg };
  }

  let genPassword = String(row[SHEET_COLS.genPassword - 1] || '').trim();
  if (!genPassword) genPassword = generateTempPassword_(14);

  sheet.getRange(rowIndex, SHEET_COLS.genEmail).setValue(genEmail);
  sheet.getRange(rowIndex, SHEET_COLS.genOu).setValue(orgUnitPath);
  sheet.getRange(rowIndex, SHEET_COLS.genRecovery).setValue(personalEmail);
  sheet.getRange(rowIndex, SHEET_COLS.genPassword).setValue(genPassword);

  writeRowResultFixed_(sheet, rowIndex, {
    status: 'PREVIEWED',
    error: '',
  });

  return { personalEmail, mode: 'PREVIEW', error: '' };
}

/** =========================
 *  STUDENT OU & EMAIL
 *  ========================= */
function buildStudentOrgUnitPath_(groupRaw) {
  const g = String(groupRaw || '').trim();
  const gUpper = g.toUpperCase();
  const gLower = g.toLowerCase();

  const isMaster = (gLower.includes('мп') || gLower.includes('мн'));
  const track = isMaster ? '2. Магістратура' : '1. Бакалаврат';

  const isBMI = (gUpper.includes('БМ') || gUpper.includes('ЗМ'));
  const isTMBI = (gUpper.includes('БФ') || gUpper.includes('ЗФ'));
  const isBMK = (gUpper.includes('БС') || gUpper.includes('ЗК'));
  const isBBZL = (gUpper.includes('БР') || gUpper.includes('ЗР'));

  let dept = 'Інше';
  if (isBMI) dept = 'БМІ';
  else if (isTMBI) dept = 'ТМБІ';
  else if (isBMK) dept = 'БМК';
  else if (isBBZL) dept = 'ББЗЛ';

  const groupForOu = formatGroupForOu_(g, isMaster);
  const orgUnitPath = `${APP_CONFIG.STUDENTS_BASE_OU}/${dept}/${track}/${groupForOu}`;

  const groupForEmailPart = normalizeGroupForEmail_(g);

  return { orgUnitPath, groupForEmailPart };
}

function formatGroupForOu_(group, isMaster) {
  let s = String(group || '').trim().toUpperCase();
  if (isMaster) {
    s = s.replace(/МП/g, 'мп').replace(/МН/g, 'мн');
  }
  return s;
}

function normalizeGroupForEmail_(groupUa) {
  let s = translitUaToLat_(groupUa).toLowerCase();
  s = s.replace(/-/g, '');
  s = s.replace(/[^a-z0-9]/g, '');
  return s || 'group';
}

function buildStudentPrimaryEmail_(surnameUa, nameUa, patrUa, groupEmailPart) {
  const surnameLat = normalizeNameTokenForEmail_(translitUaToLat_(surnameUa));
  const nameLat = normalizeNameTokenForEmail_(translitUaToLat_(nameUa));
  const patrLat = normalizeNameTokenForEmail_(translitUaToLat_(patrUa || ''));

  const i = (nameLat[0] || 'x');
  const p = (patrLat[0] || 'x');

  const local = `${surnameLat}.${i}.${p}.-${groupEmailPart}`;
  return `${local}@${APP_CONFIG.DOMAIN}`.toLowerCase();
}
