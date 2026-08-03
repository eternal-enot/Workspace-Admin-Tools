/************** CONFIG (оновлено під нові колонки) **************/
/**
 * Після вашої операції:
 *  - стара F -> нова A   (тут тепер Status / dropdown)
 *  - стара K -> нова B
 *  - нові C,D — порожні
 *  - стара B (User key/email) -> нова F
 *  - результати (ім'я/прізвище/емейли/логін/коментар/OU) логічно писати в G..L
 */
const CONFIG = {
    TARGET_SHEET_NAME: "БР-21",

    START_ROW: 2,

    // INPUT: стара B -> тепер F
    INPUT_COL: 6, // F

    // STATUS: стара F -> тепер A
    STATUS_COL: 1, // A

    // OUTPUT (без статусу): пишемо 6 колонок
    // G: First name
    // H: Last name
    // I: Recovery/extra email
    // J: Last login (local)
    // K: Comment
    // L: Org unit path
    OUTPUT_START_COL: 7, // G
    OUTPUT_NUM_COLS: 6,  // G..L

    SET_HEADERS: true,

    // Якщо true — пропускати рядки, де в A (Status) або G..L вже є якісь значення,
    // АЛЕ рядки зі статусом PENDING у колонці A НЕ пропускати.
    SKIP_IF_OUTPUT_EXISTS: false,

    PENDING_VALUE: "PENDING"
};

function checkUsersActiveSheet() {
    checkUsersAndLastLogin_(false);
}

function checkUsersAllSheets() {
    checkUsersAndLastLogin_(true);
}

function checkUsersAndLastLogin_(processAll) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();

    let sheets = [];
    if (processAll) {
        sheets = ss.getSheets().filter(s => 
            s.getName() !== "Архів" && 
            s.getName() !== "До МАГІСТРІВ" &&
            s.getName() !== "Відповіді форми (1)"
        );
    } else {
        sheets = [ss.getActiveSheet()];
        const name = sheets[0].getName();
        if (name === "Відповіді форми (1)") {
            SpreadsheetApp.getUi().alert(`❌ Cannot execute this action on "${name}".`);
            return;
        }
    }

    for (const sheet of sheets) {
        const lastRow = sheet.getLastRow();
        if (lastRow < CONFIG.START_ROW) continue;

        if (CONFIG.SET_HEADERS) setHeaders_(sheet);

        const n = lastRow - CONFIG.START_ROW + 1;

        const keys = sheet
            .getRange(CONFIG.START_ROW, CONFIG.INPUT_COL, n, 1)
            .getValues()
            .map(r => String(r[0] ?? "").trim());

        const existingStatus = sheet
            .getRange(CONFIG.START_ROW, CONFIG.STATUS_COL, n, 1)
            .getValues()
            .map(r => String(r[0] ?? "").trim());

        const existingOut = sheet
            .getRange(CONFIG.START_ROW, CONFIG.OUTPUT_START_COL, n, CONFIG.OUTPUT_NUM_COLS)
            .getValues(); // G..L (6 cols)

        const statusOut = [];
        const out = [];

        for (let i = 0; i < n; i++) {
            const key = keys[i];
            const currentStatusRaw = existingStatus[i] || "";
            const currentOut = existingOut[i]; // 6 cols

            if (!key) {
                statusOut.push([""]);
                out.push(["", "", "", "", "", ""]);
                continue;
            }

            const currentStatus = currentStatusRaw.toUpperCase();
            const isPending = currentStatus === String(CONFIG.PENDING_VALUE).trim().toUpperCase();

            const rowHasAny = rowHasAnyValue_([currentStatusRaw, ...currentOut]);

            if (CONFIG.SKIP_IF_OUTPUT_EXISTS && rowHasAny && !isPending) {
                // Пропускаємо — залишаємо як є
                statusOut.push([currentStatusRaw]);
                out.push(currentOut);
                continue;
            }

            try {
                const user = AdminDirectory.Users.get(key, {
                    projection: "full",
                    viewType: "admin_view"
                });

                const firstName = safe_(user?.name?.givenName);
                const lastName = safe_(user?.name?.familyName);
                const extraEmail = buildExtraEmail_(user);
                const orgUnitPath = safe_(user?.orgUnitPath);

                const rawLogin = user?.lastLoginTime || "";
                let lastLoginLocal = "";
                let note = "";

                if (!rawLogin || rawLogin === "1970-01-01T00:00:00.000Z") {
                    note = "Never logged in";
                } else {
                    lastLoginLocal = Utilities.formatDate(new Date(rawLogin), tz, "yyyy-MM-dd HH:mm:ss");
                }

                statusOut.push(["FOUND"]);
                out.push([firstName, lastName, extraEmail, lastLoginLocal, note, orgUnitPath]);

            } catch (e) {
                const msg = (e && e.message) ? e.message : String(e);

                if (/(notFound|Resource Not Found|404)/i.test(msg)) {
                    statusOut.push(["NOT FOUND"]);
                    out.push(["", "", "", "", "", ""]);
                } else if (/(forbidden|403|insufficient|not authorized)/i.test(msg)) {
                    statusOut.push(["ERROR"]);
                    out.push(["", "", "", "", "No permission / not authorized", ""]);
                } else {
                    statusOut.push(["ERROR"]);
                    out.push(["", "", "", "", msg, ""]);
                }
            }
        }

        // Запис: статус окремо (A), решта полів — блоком (G..L)
        sheet.getRange(CONFIG.START_ROW, CONFIG.STATUS_COL, n, 1).setValues(statusOut);
        sheet.getRange(CONFIG.START_ROW, CONFIG.OUTPUT_START_COL, n, CONFIG.OUTPUT_NUM_COLS).setValues(out);
    }
}



/************** HELPERS **************/

function setHeaders_(sheet) {
    // Status (A)
    sheet.getRange(1, CONFIG.STATUS_COL).setValue("Status");

    // Input key (F)
    sheet.getRange(1, CONFIG.INPUT_COL).setValue("User key / email (input)");

    // Outputs (G..L)
    const headers = [
        "First name",
        "Last name",
        "Recovery/extra email",
        "Last login (local)",
        "Comment",
        "Org unit path"
    ];
    sheet.getRange(1, CONFIG.OUTPUT_START_COL, 1, headers.length).setValues([headers]);
}

function rowHasAnyValue_(rowArr) {
    return rowArr.some(v => String(v ?? "").trim() !== "");
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
