/**
 * ============================================================================
 *  FILE: Staff.gc
 *  CONTAINS:
 *   - Logic specific for STAFF sheet
 *   - Preview (email generation, OU math)
 *   - Processing
 * ============================================================================
 */

/** =========================
 *  3) STAFF WRAPPERS
 *  ========================= */
function runPreviewStaff_() {
  const { count } = runPreviewStaffSilent_();
  SpreadsheetApp.getUi().alert(`✅ PREVIEW завершено для "${APP_CONFIG.STAFF_SHEET_NAME}": оброблено ${count} рядків.`);
}

function runPreviewStaffSilent_() {
  const sheetName = APP_CONFIG.STAFF_SHEET_NAME;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return 0;

    // Ensure headers
    ensureSheetWithHeaders_(ss, sheetName, STAFF_HEADERS);

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return 0;

    const n = lastRow - 1;
    const values = sheet.getRange(2, 1, n, 15).getValues();
    let processedCount = 0;

    const rangesToHighlight = [];

    // [CACHE STRATEGY] Fetch all users once
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

      const res = previewStaffRow_(sheet, rowIndex, row, userCache);
      if (res) results.push(res);

      processedCount++;

      rangesToHighlight.push(rowIndex);
    }

    // Apply Highlight (Light Yellow)
    if (rangesToHighlight.length > 0) {
      for (const rIdx of rangesToHighlight) {
        sheet.getRange(rIdx, 1, 1, 15).setBackground('#fff9c4');
      }
    }

    return { count: processedCount, results: results };
  } finally {
    lock.releaseLock();
  }
}

function processStaffByMode_() {
  const sheetName = APP_CONFIG.STAFF_SHEET_NAME;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);

    // Ensure headers
    ensureSheetWithHeaders_(ss, sheetName, STAFF_HEADERS);

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
        previewStaffRow_(sheet, rowIndex, row);
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
        sheet.getRange(rowIndex, SHEET_COLS.mode).setValue('REJECT'); // [UPDATED]
        writeRowResultFixed_(sheet, rowIndex, {
          status: 'EXISTS',
          createdAt: new Date(),
          error: 'User already exists for generated email (primary or alias).',
        });
        continue;
      }

      if (looksLikeEmail_(genRecovery) && userExistsByDomainEmailOrAlias_(genRecovery)) {
        sheet.getRange(rowIndex, SHEET_COLS.mode).setValue('REJECT'); // [UPDATED]
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
          sendCredentialsEmail_(genRecovery, created.primaryEmail || genEmail, genPassword, nameUa, surnameUa, 'staff');
        }
      } catch (e) {
        sheet.getRange(rowIndex, SHEET_COLS.mode).setValue('REJECT'); // [UPDATED]
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
 *  STAFF PREVIEW LOGIC
 *  ========================= */
function previewStaffRow_(sheet, rowIndex, row, userCache) {
  const nameUa = String(row[SHEET_COLS.nameUa - 1] || '').trim();
  const surnameUa = String(row[SHEET_COLS.surnameUa - 1] || '').trim();
  const groupOrDept = String(row[SHEET_COLS.groupOrDept - 1] || '').trim();
  const personalEmail = String(row[SHEET_COLS.personalEmail - 1] || '').trim();

  const missing = [];
  if (!nameUa) missing.push('name');
  if (!surnameUa) missing.push('surname');
  if (!personalEmail) missing.push('personal email');
  if (!groupOrDept) missing.push('кафедра');

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
      sheet.getRange(rowIndex, SHEET_COLS.mode).setValue('REJECT'); // [UPDATED]
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
      sheet.getRange(rowIndex, SHEET_COLS.mode).setValue('REJECT'); // [UPDATED]
      writeRowResultFixed_(sheet, rowIndex, {
        status: 'SKIP',
        error: errorMsg,
      });
      return { personalEmail, mode: 'REJECT', error: errorMsg };
    }
  }

  const orgUnitPath = buildStaffOrgUnitPath_(groupOrDept);
  const genEmail = buildStaffPrimaryEmail_(nameUa, surnameUa);

  // 2) Collision Check: Is GEN_EMAIL already taken?
  if (userCache && userCache.corpMap.has(genEmail)) {
    const errorMsg = `GEN_EMAIL CONFLICT: ${genEmail} is already active/alias. Manual intervention needed.`;
    sheet.getRange(rowIndex, SHEET_COLS.mode).setValue('REJECT'); // [UPDATED]
    writeRowResultFixed_(sheet, rowIndex, {
      status: 'SKIP',
      error: errorMsg,
    });
    // We populate GEN_* but mark as skipped/error so user sees why
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
 *  STAFF OU & EMAIL
 *  ========================= */
function buildStaffOrgUnitPath_(deptRaw) {
  const d = String(deptRaw || '').trim().toUpperCase();

  if (d.includes('ББЗЛ')) return '/2. Факультети/ФБМІ/ББЗЛ/0. Співробітники';
  if (d.includes('БМІ')) return '/2. Факультети/ФБМІ/БМІ/0. Співробітники';
  if (d.includes('БМК')) return '/2. Факультети/ФБМІ/БМК/0. Співробітники';
  if (d.includes('ТМБ')) return '/2. Факультети/ФБМІ/ТМБІ/0. Співробітники';
  if (d.includes('ТОС')) return '/2. Факультети/ФБМІ/ТОС/0. Співробітники';

  return '/2. Факультети/ФБМІ/';
}

function buildStaffPrimaryEmail_(nameUa, surnameUa) {
  const nameLat = normalizeNameTokenForEmail_(translitUaToLat_(nameUa));
  const surnameLat = normalizeNameTokenForEmail_(translitUaToLat_(surnameUa));
  const local = `${nameLat}.${surnameLat}`;
  return `${local}@${APP_CONFIG.DOMAIN}`.toLowerCase();
}
