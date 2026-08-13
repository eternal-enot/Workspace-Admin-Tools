/**
 * ============================================================================
 *  FILE: Triggers.gc
 *  CONTAINS:
 *   - Form submission trigger setup
 *   - Notification logic
 * ============================================================================
 */

/**
 * Run this function ONCE (via Menu) to install the installable trigger.
 * It ensures only one trigger exists for 'onFormSubmitAction_'.
 */
function setupFormSubmitTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Remove existing triggers for this function to avoid duplicates
  const triggers = ScriptApp.getUserTriggers(ss);
  let count = 0;
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'onFormSubmitAction_') {
      ScriptApp.deleteTrigger(t);
      count++;
    }
  }

  // 2. Create new trigger
  ScriptApp.newTrigger('onFormSubmitAction_')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  let msg = '✅ Тригер на нові відповіді встановлено успішно.';
  if (count > 0) {
    msg += ` (Оновлено попередніх тригерів: ${count})`;
  }
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * TRIGGERED FUNCTION
 * Sends notification email when a new row is added via Form or manual submit.
 * 
 * @param {Object} e - Event object from onFormSubmit
 */
function onFormSubmitAction_(e) {
  const emailTo = APP_CONFIG.ADMIN_TEST_EMAIL;
  const subject = '🔔 [Admin Tools] Нові відповіді форми';

  // 1. Run Transfer (Silent)
  let trRes = null;
  try {
    trRes = transferFromFormResponsesSilent_();
  } catch (err) {
    console.error('Transfer failed in trigger:', err);
    trRes = { error: String(err) };
  }

  // 2. Run Preview (Silent)
  let prevStudentsCount = 0;
  let prevPhdCount = 0;
  let prevStaffCount = 0;

  if (trRes && !trRes.error) {
    try {
      let studentResults = [];
      let phdResults = [];
      let staffResults = [];

      if (trRes.movedStudents > 0) {
        const sRes = runPreviewStudentsSilent_();
        if (typeof sRes === 'object') {
          prevStudentsCount = sRes.count;
          studentResults = sRes.results || [];
        } else {
          prevStudentsCount = sRes;
        }
      }

      if (trRes.movedPhd > 0) {
        const sRes = runPreviewPhdSilent_();
        if (typeof sRes === 'object') {
          prevPhdCount = sRes.count;
          phdResults = sRes.results || [];
        } else {
          prevPhdCount = sRes;
        }
      }

      if (trRes.movedStaff > 0) {
        const sRes = runPreviewStaffSilent_();
        if (typeof sRes === 'object') {
          prevStaffCount = sRes.count;
          staffResults = sRes.results || [];
        } else {
          prevStaffCount = sRes;
        }
      }

      // Merge preview results into transfer details
      if (trRes.details && trRes.details.length > 0) {
        const allResults = [...studentResults, ...phdResults, ...staffResults];
        const resMap = new Map();
        for (const r of allResults) {
          if (r.personalEmail) resMap.set(r.personalEmail.toLowerCase(), r);
        }

        for (const d of trRes.details) {
          if (!d.email) continue;
          const match = resMap.get(d.email.toLowerCase());
          if (match) {
            if (match.mode) d.mode = match.mode;
            if (match.error) d.error = match.error;
          }
        }
      }

    } catch (err) {
      console.error('Preview failed in trigger:', err);
    }
  }

  // 3. AUTO-DEPLOY (If enabled and safe)
  let autoDeployResults = null;
  if (APP_CONFIG.AUTO_DEPLOY_ENABLED && trRes && trRes.details) {
    autoDeployResults = runAutoDeployLogic_(trRes.details);
  }

  // 4. Notification #1: New Response Received (Standard)
  sendNewResponseNotification_(emailTo, subject, trRes, prevStudentsCount, prevPhdCount, prevStaffCount);

  // 5. Notification #2: Auto-Deploy Report (If attempted)
  if (autoDeployResults) {
    sendAutoDeployNotification_(emailTo, autoDeployResults);
  }
}

/** 
 * HELPER: Filter and Run Auto Deploy 
 */
function runAutoDeployLogic_(details) {
  // 1. Identify candidates: PREVIEW mode, NO error, NO Hint (safety)
  const candidates = details.filter(d => {
    return d.mode === 'PREVIEW' && !d.error && !d.hint;
  });

  if (!candidates.length) return null;

  console.log('Auto-Deploy candidates:', candidates.length);

  // Mark matching PREVIEW rows as DEPLOY, then run process*ByMode_.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resReport = {
    processed: 0,
    success: 0,
    errors: [],
    details: [] // { name, email, status, error }
  };

  const sheetsToProcess = new Set();

  // Helper to mark DEPLOY
  const markDeploy = (sheetName, emailsToDeploy) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();

    // Set of personal emails
    const targetSet = new Set(emailsToDeploy.map(e => e.toLowerCase()));

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const mode = String(row[SHEET_COLS.mode - 1] || '').trim().toUpperCase();
      const pEmail = String(row[SHEET_COLS.personalEmail - 1] || '').trim().toLowerCase();

      if (mode === 'PREVIEW' && targetSet.has(pEmail)) {
        sheet.getRange(i + 2, SHEET_COLS.mode).setValue('DEPLOY');
      }
    }
    sheetsToProcess.add(sheetName);
  };

  // Group candidates by Role -> Sheet
  const students = candidates.filter(d => d.role === 'Student').map(d => d.email);
  const phd = candidates.filter(d => d.role === 'PhD').map(d => d.email);
  const staff = candidates.filter(d => d.role === 'Staff').map(d => d.email);

  if (students.length) markDeploy(APP_CONFIG.STUDENTS_SHEET_NAME, students);
  if (phd.length) markDeploy(APP_CONFIG.PHD_SHEET_NAME, phd);
  if (staff.length) markDeploy(APP_CONFIG.STAFF_SHEET_NAME, staff);

  // Run deploy processors (skip non-DEPLOY rows)
  if (sheetsToProcess.has(APP_CONFIG.STUDENTS_SHEET_NAME)) processStudentsByMode_();
  if (sheetsToProcess.has(APP_CONFIG.PHD_SHEET_NAME)) processPhdByMode_();
  if (sheetsToProcess.has(APP_CONFIG.STAFF_SHEET_NAME)) processStaffByMode_();

  // Gather final status from the sheet for the report
  const gatherResults = (sheetName, emailList) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
    const targetSet = new Set(emailList.map(e => e.toLowerCase()));

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const pEmail = String(row[SHEET_COLS.personalEmail - 1] || '').trim().toLowerCase();

      if (targetSet.has(pEmail)) {
        const mode = String(row[SHEET_COLS.mode - 1] || '').trim().toUpperCase();
        const status = String(row[SHEET_COLS.status - 1] || '').trim();
        const err = String(row[SHEET_COLS.error - 1] || '').trim();
        const name = `${row[SHEET_COLS.nameUa - 1]} ${row[SHEET_COLS.surnameUa - 1]}`;

        resReport.processed++;
        if (status === 'CREATED' || mode === 'EXECUTED') {
          resReport.success++;
          resReport.details.push({ name, email: pEmail, status: 'CREATED', error: '' });
        } else {
          resReport.errors.push(`${pEmail}: ${err || status}`);
          resReport.details.push({ name, email: pEmail, status: status || 'ERROR', error: err });
        }
      }
    }
  };

  if (students.length) gatherResults(APP_CONFIG.STUDENTS_SHEET_NAME, students);
  if (phd.length) gatherResults(APP_CONFIG.PHD_SHEET_NAME, phd);
  if (staff.length) gatherResults(APP_CONFIG.STAFF_SHEET_NAME, staff);

  return resReport;
}

function sendNewResponseNotification_(emailTo, subject, trRes, prevStudentsCount, prevPhdCount, prevStaffCount) {
  // 3. Prepare HTML Email (Original Logic moved here)
  let html = '<div style="font-family:Arial,sans-serif;color:#333;">';
  html += `<h3>Оброблено нові відповіді</h3>`;
  html += `<p style="font-size:12px;color:#666;">Час: ${new Date().toLocaleString('uk-UA')}</p>`;

  if (trRes && trRes.error) {
    html += `<p style="color:red;font-weight:bold;">❌ Помилка при перенесенні: ${trRes.error}</p>`;
  } else if (trRes && trRes.details && trRes.details.length > 0) {
    // Summary
    html += `<p><strong>Перенесено:</strong> Студентів: ${trRes.movedStudents}, Аспірантів: ${trRes.movedPhd}, Співробітників: ${trRes.movedStaff}</p>`;
    // Table
    html += `<table style="width:100%;border-collapse:collapse;margin-top:10px;font-size:13px;">`;
    html += `<tr style="background:#f0f0f0;text-align:left;">`;
    html += `<th style="border:1px solid #ddd;padding:8px;">Роль</th>`;
    html += `<th style="border:1px solid #ddd;padding:8px;">ПІБ</th>`;
    html += `<th style="border:1px solid #ddd;padding:8px;">Email</th>`;
    html += `<th style="border:1px solid #ddd;padding:8px;">Дія</th>`;
    html += `<th style="border:1px solid #ddd;padding:8px;">Результат/Помилка</th>`;
    html += `</tr>`;

    for (const d of trRes.details) {
      let bg = '#fff';
      let statusStyle = 'color:#333;';

      if (d.mode === 'RECOVERY') {
        bg = '#fff3e0'; // Orange tint
        statusStyle = 'color:#e65100;font-weight:bold;';
      } else if (d.mode === 'REJECT' || d.mode === 'SKIP') {
        bg = '#ffebee'; // Red tint
        statusStyle = 'color:#c62828;font-weight:bold;';
      }

      // If we have a hint mismatch error, highlight error
      const errText = d.error ? `<span style="color:red;">${d.error}</span>` : 'OK';
      const hintText = d.hint ? `<br><small style="color:#666;">(Hint: ${d.hint})</small>` : '';

      html += `<tr style="background:${bg};">`;
      html += `<td style="border:1px solid #ddd;padding:8px;">${d.role}</td>`;
      html += `<td style="border:1px solid #ddd;padding:8px;">${d.name}</td>`;
      html += `<td style="border:1px solid #ddd;padding:8px;">${d.email}${hintText}</td>`;
      html += `<td style="border:1px solid #ddd;padding:8px;${statusStyle}">${d.mode}</td>`;
      html += `<td style="border:1px solid #ddd;padding:8px;">${errText}</td>`;
      html += `</tr>`;
    }
    html += `</table>`;

  } else {
    html += `<p>Немає нових рядків для обробки (або всі вже перенесені).</p>`;
  }

  // Preview Stats
  if (prevStudentsCount > 0 || prevPhdCount > 0 || prevPhdCount > 0) {
    html += `<p style="margin-top:20px;"><strong>Preview Run:</strong> Сформовано профілів для ${prevStudentsCount + prevPhdCount + prevStaffCount} користувачів.</p>`;
  }

  html += `<p style="margin-top:20px;font-size:11px;color:#999;">Це автоматичне повідомлення Admin Tools.</p>`;
  html += '</div>';

  try {
    MailApp.sendEmail({
      to: emailTo,
      subject: subject,
      htmlBody: html
    });
  } catch (err) {
    console.error('Error sending notification email:', err);
  }
}

function sendAutoDeployNotification_(emailTo, res) {
  const subject = `🚀 [Admin Tools] Результат автоматичного деплою`;
  let html = '<div style="font-family:Arial,sans-serif;color:#333;">';
  html += `<h3>Автоматичне створення акаунтів (Test Mode)</h3>`;
  html += `<p><strong>Успішно:</strong> ${res.success} / ${res.processed} <br>`;

  if (res.errors.length > 0) {
    html += `<span style="color:red">Помилок: ${res.errors.length}</span></p>`;
  } else {
    html += `<span style="color:green">Помилок: 0</span></p>`;
  }

  if (res.details.length > 0) {
    html += `<table style="width:100%;border-collapse:collapse;margin-top:10px;font-size:13px;">`;
    html += `<tr style="background:#e3f2fd;text-align:left;">`;
    html += `<th style="border:1px solid #ddd;padding:8px;">ПІБ</th>`;
    html += `<th style="border:1px solid #ddd;padding:8px;">Email</th>`;
    html += `<th style="border:1px solid #ddd;padding:8px;">Статус</th>`;
    html += `<th style="border:1px solid #ddd;padding:8px;">Помилка</th>`;
    html += `</tr>`;

    for (const d of res.details) {
      let bg = d.status === 'CREATED' ? '#e8f5e9' : '#ffebee';
      html += `<tr style="background:${bg};">`;
      html += `<td style="border:1px solid #ddd;padding:8px;">${d.name}</td>`;
      html += `<td style="border:1px solid #ddd;padding:8px;">${d.email}</td>`;
      html += `<td style="border:1px solid #ddd;padding:8px;font-weight:bold;">${d.status}</td>`;
      html += `<td style="border:1px solid #ddd;padding:8px;">${d.error}</td>`;
      html += `</tr>`;
    }
    html += `</table>`;
  }

  html += `<p style="margin-top:20px;font-size:11px;color:#999;">Автоматичний деплой активовано в налаштуваннях.</p></div>`;

  try {
    MailApp.sendEmail({
      to: emailTo,
      subject: subject,
      htmlBody: html
    });
  } catch (err) {
    console.error('Error sending auto-deploy email:', err);
  }
}
