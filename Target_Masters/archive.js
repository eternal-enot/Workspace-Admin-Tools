/**
 * Helper to send an alumni notification email.
 */
function sendAlumniEmail_(mainEmail, recipients, deletionDateStr) {
    const SENDER_NAME = 'Адміністрація Google Workspace for Education';
    const subject = "Повідомлення щодо корпоративного облікового запису випускника edu.kpi.ua (lll.kpi.ua)";
    
    const bodyPlain = `Шановний користувачу,\n\nЩиро вітаємо Вас із успішним завершенням навчання!\n\nЦим листом ми хочемо повідомити про подальший статус Вашого корпоративного облікового запису:\n${mainEmail}\n\nЯкщо Ви плануєте вступати до магістратури або аспірантури КПІ ім. Ігоря Сікорського, обліковий запис буде збережено, а його назву оновлено відповідно до нової академічної групи після зарахування.\n\nЯкщо ж Ви не плануєте продовжувати навчання, відповідно до внутрішньої політики управління обліковими записами, цей акаунт буде деактивовано та остаточно видалено через шість місяців від дати цього листа — ${deletionDateStr}.\n\nУ зв'язку з цим просимо Вас завчасно вжити наступних заходів:\n- завантажити та зберегти всі необхідні файли, документи та дані, що розміщені на Google Диску або в інших сервісах, пов'язаних із цим обліковим записом;\n- переконатися у відсутності активних підписок чи зовнішніх служб, прив'язаних до акаунту.\n\nУ разі якщо Ви маєте додаткові запитання — будь ласка, зверніться до адміністратора Вашого підрозділу, відповівши на цей лист.\n\nЗ повагою,\nАдміністрація Google Workspace for Education\nКПІ ім. Ігоря Сікорського\nФакультет біомедичної інженерії`;

    const logoHtml = '<div style="text-align:center;margin-top:24px;"><img src="https://fbmi.kpi.ua/wp-content/uploads/2025/09/fbmi_blue_logo.png" alt="Лого університету" style="max-width:200px;height:auto;"></div>';

    const htmlBody = '<!doctype html>' +
        '<html><body style="margin:0;padding:0;background:#f6f7f9;">' +
        '  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#f6f7f9;padding:24px 0;">' +
        '    <tr><td align="center">' +
        '      <table role="presentation" width="640" cellspacing="0" cellpadding="0" style="background:#ffffff;border-radius:12px;overflow:hidden;font-family:Arial,system-ui,-apple-system,\'Segoe UI\',Roboto,\'Noto Sans\',sans-serif;color:#111">' +
        '        <tr><td style="padding:24px 32px 0;font-size:16px;line-height:1.6;">' +
        '          <p style="margin:0 0 12px 0;">Шановний користувачу,</p>' +
        '          <p style="margin:0 0 12px 0;color:#1a73e8;font-weight:bold;">Щиро вітаємо Вас із успішним завершенням навчання!</p>' +
        '          <p style="margin:0 0 12px 0;">Цим листом ми хочемо повідомити про подальший статус Вашого корпоративного облікового запису:</p>' +
        '        </td></tr>' +
        '        <tr><td style="padding:8px 32px 0">' +
        '          <table width="100%" cellspacing="0" cellpadding="0" style="border:1px solid #e6e6e6;border-radius:8px;">' +
        '            <tr>' +
        '              <td style="padding:12px 16px;font-size:16px;background:#fafafa;text-align:center;font-weight:bold;color:#1a73e8;">' + mainEmail + '</td>' +
        '            </tr>' +
        '          </table>' +
        '        </td></tr>' +
        '        <tr><td style="padding:24px 32px 0;font-size:16px;line-height:1.6;">' +
        '          Якщо Ви плануєте вступати до магістратури або аспірантури КПІ ім. Ігоря Сікорського, обліковий запис буде збережено, а його назву оновлено відповідно до нової академічної групи після зарахування.' +
        '        </td></tr>' +
        '        <tr><td style="padding:20px 32px 0;font-size:16px;line-height:1.6;color:#333">' +
        '          Якщо ж Ви не плануєте продовжувати навчання, відповідно до внутрішньої політики управління обліковими записами, цей акаунт буде деактивовано та остаточно видалено через шість місяців від дати цього листа — <strong>' + deletionDateStr + '</strong>.' +
        '        </td></tr>' +
        '        <tr><td style="padding:20px 32px 0;font-size:16px;line-height:1.6;color:#333">' +
        '          <p style="margin:0 0 10px 0;"><strong>У зв\'язку з цим просимо Вас завчасно вжити наступних заходів:</strong></p>' +
        '          <ul style="margin:0 0 0 20px;padding:0;">' +
        '            <li style="margin-bottom:8px;">завантажити та зберегти всі необхідні файли, документи та дані, що розміщені на Google Диску або в інших сервісах, пов\'язаних із цим обліковим записом;</li>' +
        '            <li>переконатися у відсутності активних підписок чи зовнішніх служб, прив\'язаних до акаунту.</li>' +
        '          </ul>' +
        '        </td></tr>' +
        '        <tr><td style="padding:24px 32px 24px;font-size:14px;color:#333;line-height:1.6;">' +
        '          У разі якщо Ви маєте додаткові запитання — будь ласка, зверніться до адміністратора Вашого підрозділу, відповівши на цей лист.' +
        '        </td></tr>' +
        '        <tr><td style="padding:24px 32px 24px;font-size:14px;color:#333;line-height:1.5;border-top:1px solid #eee;">' +
        '          З повагою,<br>Адміністрація Google Workspace for Education<br>КПІ ім. Ігоря Сікорського<br>Факультет біомедичної інженерії' +
        '        </td></tr>' +
        '        <tr><td style="padding:0 32px 24px;">' + logoHtml + '</td></tr>' +
        '      </table>' +
        '    </td></tr>' +
        '  </table>' +
        '</body></html>';
                 
    MailApp.sendEmail({
        to: recipients,
        subject: subject,
        name: SENDER_NAME,
        body: bodyPlain,
        htmlBody: htmlBody
    });
}

/**
 * Helper to send a deletion warning email.
 */
function sendDeletionEmail_(mainEmail, recipients, deletionDateStr) {
    const SENDER_NAME = 'Адміністрація Google Workspace for Education';
    const subject = "Повідомлення про заплановане видалення облікового запису edu.kpi.ua (lll.kpi.ua)";
    
    const bodyPlain = `Шановний користувачу,\n\nЦим листом повідомляємо Вас про видалення корпоративного облікового запису:\n${mainEmail}\n\nВідповідно до внутрішньої політики управління обліковими записами, зазначений акаунт буде деактивовано та остаточно видалено через шість місяців від дати цього листа — ${deletionDateStr}.\n\nПросимо Вас завчасно вжити наступних заходів:\n- завантажити та зберегти всі необхідні файли, документи та дані, що розміщені на Google Диску або в інших сервісах, пов'язаних із цим обліковим записом;\n- переконатися у відсутності активних підписок чи зовнішніх служб, прив'язаних до акаунту.\n\nУ разі поновлення навчання ми відновимо Ваш обліковий запис і оновимо його назву відповідно до нової академічної групи. Вам не потрібно нічого робити додатково.\n\nУ разі якщо Ви вважаєте, що це повідомлення надіслано помилково, або маєте додаткові запитання — будь ласка, зверніться до адміністратора Вашого підрозділу, відповівши на цей лист.\n\nЗ повагою,\nАдміністрація Google Workspace for Education\nКПІ ім. Ігоря Сікорського\nФакультет біомедичної інженерії`;

    const logoHtml = '<div style="text-align:center;margin-top:24px;"><img src="https://fbmi.kpi.ua/wp-content/uploads/2025/09/fbmi_blue_logo.png" alt="Лого університету" style="max-width:200px;height:auto;"></div>';

    const htmlBody = '<!doctype html>' +
        '<html><body style="margin:0;padding:0;background:#f6f7f9;">' +
        '  <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#f6f7f9;padding:24px 0;">' +
        '    <tr><td align="center">' +
        '      <table role="presentation" width="640" cellspacing="0" cellpadding="0" style="background:#ffffff;border-radius:12px;overflow:hidden;font-family:Arial,system-ui,-apple-system,\'Segoe UI\',Roboto,\'Noto Sans\',sans-serif;color:#111">' +
        '        <tr><td style="padding:24px 32px 0;font-size:16px;line-height:1.6;">' +
        '          <p style="margin:0 0 12px 0;">Шановний користувачу,</p>' +
        '          <p style="margin:0 0 12px 0;">Цим листом повідомляємо Вас про видалення корпоративного облікового запису:</p>' +
        '        </td></tr>' +
        '        <tr><td style="padding:8px 32px 0">' +
        '          <table width="100%" cellspacing="0" cellpadding="0" style="border:1px solid #e6e6e6;border-radius:8px;">' +
        '            <tr>' +
        '              <td style="padding:12px 16px;font-size:16px;background:#fafafa;text-align:center;font-weight:bold;color:#1a73e8;">' + mainEmail + '</td>' +
        '            </tr>' +
        '          </table>' +
        '        </td></tr>' +
        '        <tr><td style="padding:24px 32px 0;font-size:16px;line-height:1.6;">' +
        '          Відповідно до внутрішньої політики управління обліковими записами, зазначений акаунт буде деактивовано та остаточно видалено через шість місяців від дати цього листа — <strong>' + deletionDateStr + '</strong>.' +
        '        </td></tr>' +
        '        <tr><td style="padding:20px 32px 0;font-size:16px;line-height:1.6;color:#333">' +
        '          <p style="margin:0 0 10px 0;"><strong>Просимо Вас завчасно вжити наступних заходів:</strong></p>' +
        '          <ul style="margin:0 0 0 20px;padding:0;">' +
        '            <li style="margin-bottom:8px;">завантажити та зберегти всі необхідні файли, документи та дані, що розміщені на Google Диску або в інших сервісах, пов\'язаних із цим обліковим записом;</li>' +
        '            <li>переконатися у відсутності активних підписок чи зовнішніх служб, прив\'язаних до акаунту.</li>' +
        '          </ul>' +
        '        </td></tr>' +
        '        <tr><td style="padding:24px 32px 0;font-size:16px;line-height:1.6;color:#333">' +
        '          У разі поновлення навчання ми відновимо Ваш обліковий запис і оновимо його назву відповідно до нової академічної групи. Вам не потрібно нічого робити додатково.' +
        '        </td></tr>' +
        '        <tr><td style="padding:24px 32px 24px;font-size:14px;color:#333;line-height:1.6;">' +
        '          У разі якщо Ви вважаєте, що це повідомлення надіслано помилково, або маєте додаткові запитання — будь ласка, зверніться до адміністратора Вашого підрозділу, відповівши на цей лист.' +
        '        </td></tr>' +
        '        <tr><td style="padding:24px 32px 24px;font-size:14px;color:#333;line-height:1.5;border-top:1px solid #eee;">' +
        '          З повагою,<br>Адміністрація Google Workspace for Education<br>КПІ ім. Ігоря Сікорського<br>Факультет біомедичної інженерії' +
        '        </td></tr>' +
        '        <tr><td style="padding:0 32px 24px;">' + logoHtml + '</td></tr>' +
        '      </table>' +
        '    </td></tr>' +
        '  </table>' +
        '</body></html>';
                 
    MailApp.sendEmail({
        to: recipients,
        subject: subject,
        name: SENDER_NAME,
        body: bodyPlain,
        htmlBody: htmlBody
    });
}

/**
 * Formats a date in Ukrainian: "DD <month> YYYY року"
 */
function formatUkrainianDate_(date) {
    const tz = Session.getScriptTimeZone() || "Europe/Kyiv";
    const day = Utilities.formatDate(date, tz, "dd");
    const year = Utilities.formatDate(date, tz, "yyyy");
    const m = parseInt(Utilities.formatDate(date, tz, "MM"), 10);
    const months = [
        "січня", "лютого", "березня", "квітня", "травня", "червня",
        "липня", "серпня", "вересня", "жовтня", "листопада", "грудня"
    ];
    return `${day} ${months[m - 1]} ${year} року`;
}

/**
 * Checks and creates missing organizational units (OUs) along the provided path.
 */
function archiveEnsureOrgUnit_(ouPath) {
    if (!ouPath || ouPath === "/") return;
    
    let s = String(ouPath).trim();
    if (!s.startsWith("/")) s = "/" + s;
    while (s.length > 1 && s.endsWith("/")) s = s.slice(0, -1);
    
    let existingOUs = [];
    let pageToken;
    do {
        let res = archiveCallWithRetry_(() => AdminDirectory.Orgunits.list("my_customer", { type: "all", pageToken: pageToken }));
        if (res.organizationUnits) {
            existingOUs = existingOUs.concat(res.organizationUnits);
        }
        pageToken = res.nextPageToken;
    } while (pageToken);
    
    if (existingOUs.some(ou => ou.orgUnitPath === s)) {
        return; 
    }
    
    const parts = s.split("/").filter(p => p.trim().length > 0);
    let parentPath = "";
    
    for (let i = 0; i < parts.length; i++) {
        const name = parts[i];
        const currentPath = parentPath === "" ? "/" + name : parentPath + "/" + name;
        
        let found = existingOUs.find(ou => ou.orgUnitPath === currentPath);
        
        if (!found) {
            let parentData = parentPath === "" ? { orgUnitPath: "/" } : existingOUs.find(ou => ou.orgUnitPath === parentPath);
            
            if (!parentData) {
                throw new Error("Батьківська OU не знайдена: " + parentPath);
            }
            
            let inserted = archiveCallWithRetry_(() => AdminDirectory.Orgunits.insert({
                name: name,
                parentOrgUnitPath: parentData.orgUnitPath
            }, "my_customer"));
            
            existingOUs.push(inserted);
        }
        
        parentPath = currentPath;
    }
}

/**
 * Helper to safely call Google APIs using an exponential backoff retry pattern.
 */
function archiveCallWithRetry_(fn, maxRetries = 3) {
    let attempt = 0;
    while (true) {
        try {
            return fn();
        } catch (e) {
            const msg = (e && e.message) ? e.message : String(e);
            const transient = /Rate Limit|Quota|quota|Too many|Service invoked too many times|429|500|503|Backend Error/i.test(msg);
            
            if (!transient || attempt >= maxRetries) throw e;
            
            Utilities.sleep(1000 * Math.pow(2, attempt));
            attempt++;
        }
    }
}
