const SUPPORT_CFG = {
    SHEET_NAME: "Manager",
    START_ROW: 2,
    COL_FIRST_NAME: 1, // A
    COL_LAST_NAME: 2,  // B
    COL_EMAIL: 3,      // C (INPUT)
    COL_STATUS: 4,     // D
    COL_EXTRA_EMAILS: 5, // E
    COL_PHONES: 6,       // F
    COL_LAST_LOGIN: 7,   // G
    COL_OU: 8            // H
};

function checkSupportInfo() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SUPPORT_CFG.SHEET_NAME);
    if (!sheet) {
        sheet = ss.getActiveSheet();
        if (sheet.getName() !== SUPPORT_CFG.SHEET_NAME) {
            SpreadsheetApp.getUi().alert(`❌ Цей скрипт призначено для аркуша "${SUPPORT_CFG.SHEET_NAME}".`);
            return;
        }
    }

    const lastRow = Math.max(sheet.getLastRow(), SUPPORT_CFG.START_ROW);
    setupHeaders_(sheet);

    if (lastRow < SUPPORT_CFG.START_ROW) return;

    const tz = ss.getSpreadsheetTimeZone();
    const dataRange = sheet.getRange(SUPPORT_CFG.START_ROW, 1, lastRow - SUPPORT_CFG.START_ROW + 1, SUPPORT_CFG.COL_OU);
    const data = dataRange.getValues();
    
    let updatedCount = 0;

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const email = String(row[SUPPORT_CFG.COL_EMAIL - 1] || "").trim();
        const currentStatus = String(row[SUPPORT_CFG.COL_STATUS - 1] || "").trim().toUpperCase();

        if (!email) continue;
        if (currentStatus === "DELETE" || currentStatus === "DELETED") continue;

        try {
            const user = AdminDirectory.Users.get(email, {
                projection: "full",
                viewType: "admin_view"
            });

            const firstName = safe_(user?.name?.givenName);
            const lastName = safe_(user?.name?.familyName);
            const extraEmail = buildExtraEmail_(user);
            const phones = buildPhones_(user);
            const orgUnitPath = safe_(user?.orgUnitPath);

            const rawLogin = user?.lastLoginTime || "";
            let lastLoginLocal = "";
            if (!rawLogin || rawLogin === "1970-01-01T00:00:00.000Z") {
                lastLoginLocal = "Never logged in";
            } else {
                lastLoginLocal = Utilities.formatDate(new Date(rawLogin), tz, "yyyy-MM-dd HH:mm:ss");
            }

            row[SUPPORT_CFG.COL_FIRST_NAME - 1] = firstName;
            row[SUPPORT_CFG.COL_LAST_NAME - 1] = lastName;
            row[SUPPORT_CFG.COL_STATUS - 1] = "FOUND";
            row[SUPPORT_CFG.COL_EXTRA_EMAILS - 1] = extraEmail;
            row[SUPPORT_CFG.COL_PHONES - 1] = phones;
            row[SUPPORT_CFG.COL_LAST_LOGIN - 1] = lastLoginLocal;
            row[SUPPORT_CFG.COL_OU - 1] = orgUnitPath;

            updatedCount++;

        } catch (e) {
            const msg = (e && e.message) ? e.message : String(e);

            if (/(notFound|Resource Not Found|404)/i.test(msg)) {
                row[SUPPORT_CFG.COL_STATUS - 1] = "NOT FOUND";
                row[SUPPORT_CFG.COL_FIRST_NAME - 1] = "";
                row[SUPPORT_CFG.COL_LAST_NAME - 1] = "";
                row[SUPPORT_CFG.COL_EXTRA_EMAILS - 1] = "";
                row[SUPPORT_CFG.COL_PHONES - 1] = "";
                row[SUPPORT_CFG.COL_LAST_LOGIN - 1] = "";
                row[SUPPORT_CFG.COL_OU - 1] = "";
            } else {
                row[SUPPORT_CFG.COL_STATUS - 1] = "ERROR: " + msg;
            }
            updatedCount++;
        }
    }

    if (updatedCount > 0) {
        dataRange.setValues(data);
    }
}

function buildPhones_(user) {
    const res = [];
    const recoveryPhone = safe_(user?.recoveryPhone);
    if (recoveryPhone) {
        res.push(recoveryPhone + " (recovery)");
    }
    
    const phones = Array.isArray(user?.phones) ? user.phones : [];
    for (const p of phones) {
        const val = safe_(p?.value);
        if (!val) continue;
        const type = safe_(p?.type) || "other";
        const entry = `${val} (${type})`;
        if (!res.includes(entry) && val !== recoveryPhone) {
            res.push(entry);
        }
    }
    return res.join("; ");
}

function buildExtraEmail_(user) {
    const res = [];
    const primary = (user?.primaryEmail || "").toLowerCase();

    const recovery = safe_(user?.recoveryEmail);
    if (recovery) res.push(recovery);

    const emails = Array.isArray(user?.emails) ? user.emails : [];
    for (const e of emails) {
        const addr = safe_(e?.address);
        if (!addr) continue;
        const low = addr.toLowerCase();
        if (low && low !== primary) res.push(addr);
    }

    const aliases = Array.isArray(user?.aliases) ? user.aliases : [];
    for (const a of aliases) {
        const addr = safe_(a);
        if (!addr) continue;
        const low = addr.toLowerCase();
        if (low && low !== primary) res.push(addr);
    }

    const seen = new Set();
    const uniq = [];
    for (const x of res) {
        const k = x.toLowerCase();
        if (!seen.has(k)) {
            seen.add(k);
            uniq.push(x);
        }
    }
    return uniq.join("; ");
}

function safe_(v) {
    return (v === null || v === undefined) ? "" : String(v).trim();
}

function setupHeaders_(sheet) {
    const headers = [
        "First Name",
        "Last Name",
        "Email (Input)",
        "Status",
        "Extra/Recovery Emails",
        "Phone Numbers",
        "Last Login",
        "Org Unit Path"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Add conditional formatting for Status column (D)
    const maxRows = Math.max(sheet.getMaxRows(), 100);
    const colD = sheet.getRange(2, SUPPORT_CFG.COL_STATUS, maxRows - 1, 1);
    
    sheet.clearConditionalFormatRules();
    const rules = [];
    
    const statusColors = {
        "FOUND": "#c8e6c9", // green
        "NOT FOUND": "#ef9a9a", // red
        "DELETE": "#ffcc80", // orange
        "DELETED": "#e0e0e0" // grey
    };
    
    for (const st in statusColors) {
        rules.push(SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(st)
            .setBackground(statusColors[st])
            .setRanges([colD])
            .build());
    }
    
    // Any text starting with ERROR:
    rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextStartsWith("ERROR")
        .setBackground("#ef9a9a")
        .setRanges([colD])
        .build());
        
    sheet.setConditionalFormatRules(rules);
}
