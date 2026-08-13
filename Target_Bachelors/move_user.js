/************** MOVE_CONFIG (adapted for new columns) **************/
/**
 * After column reorder:
 *  A = Status (was F)
 *  F = User key/email (was B)
 *  L = Org unit path (from checkUsersAndLastLogin; was I in the old layout)
 *
 * For move:
 *  - userKey from F
 *  - status from A
 *  - oldOU from L
 *  - newOU (manual) in M
 *  - result written to N (Move status) and O (Move note)
 */
const MOVE_CONFIG = {
    PROCESS_ALL_SHEETS: false,
    USE_ACTIVE_SHEET: true,
    TARGET_SHEET_NAME: "БР-21",
    START_ROW: 2,

    // Columns (1-based)
    COL_USERKEY: 6,   // F
    COL_STATUS: 1,    // A
    COL_OLD_OU: 12,   // L
    COL_NEW_OU: 13,   // M

    // Where to write the result
    COL_MOVE_STATUS: 14, // N
    COL_MOVE_NOTE: 15,   // O
    SET_HEADERS: true,

    VERIFY_CURRENT_OU_BEFORE_MOVE: false,
    STRICT_OLD_OU_MATCH: false,

    MAX_RETRIES: 5
};

/**
 * Moves users (FOUND) from OU in L to OU in M.
 * Requires Advanced Google Service "Admin SDK API" (Directory v1) => AdminDirectory
 */
function moveFoundUsersToNewOU() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = getMoveSheetsToProcess_(ss);

    for (const sheet of sheets) {
        const lastRow = sheet.getLastRow();
        if (lastRow < MOVE_CONFIG.START_ROW) continue;

        if (MOVE_CONFIG.SET_HEADERS) {
            sheet.getRange(1, MOVE_CONFIG.COL_MOVE_STATUS).setValue("Move status");
            sheet.getRange(1, MOVE_CONFIG.COL_MOVE_NOTE).setValue("Move note");
        }

        const n = lastRow - MOVE_CONFIG.START_ROW + 1;

        // Read only the needed columns separately (more stable after reorders)
        const userKeys = sheet.getRange(MOVE_CONFIG.START_ROW, MOVE_CONFIG.COL_USERKEY, n, 1).getValues().map(r => String(r[0] ?? "").trim());
        const statuses = sheet.getRange(MOVE_CONFIG.START_ROW, MOVE_CONFIG.COL_STATUS, n, 1).getValues().map(r => String(r[0] ?? "").trim());
        const oldOUs = sheet.getRange(MOVE_CONFIG.START_ROW, MOVE_CONFIG.COL_OLD_OU, n, 1).getValues().map(r => String(r[0] ?? "").trim());
        const newOUs = sheet.getRange(MOVE_CONFIG.START_ROW, MOVE_CONFIG.COL_NEW_OU, n, 1).getValues().map(r => String(r[0] ?? "").trim());

        const out = [];

        for (let i = 0; i < n; i++) {
            const userKey = userKeys[i];
            const status = statuses[i];
            const oldOU = oldOUs[i];
            const newOU = newOUs[i];

            if (!userKey) {
                out.push(["SKIPPED", "Empty user key (F)"]);
                continue;
            }

            if (String(status).trim().toUpperCase() !== "FOUND") {
                out.push(["SKIPPED", `Status is "${status}"`]);
                continue;
            }

            if (!oldOU) {
                out.push(["SKIPPED", "Missing old OU path in column L"]);
                continue;
            }

            if (!newOU) {
                out.push(["ERROR", "Missing new OU path in column M"]);
                continue;
            }

            const oldOUNorm = normalizeOuPath_(oldOU);
            const newOUNorm = normalizeOuPath_(newOU);

            // Skip if unchanged
            if (oldOUNorm && newOUNorm && oldOUNorm === newOUNorm) {
                out.push(["SKIPPED", "Old OU equals new OU (L == M)"]);
                continue;
            }

            try {
                if (MOVE_CONFIG.VERIFY_CURRENT_OU_BEFORE_MOVE) {
                    const u = callWithRetry_(() =>
                        AdminDirectory.Users.get(userKey, { projection: "basic", viewType: "admin_view" }),
                        MOVE_CONFIG.MAX_RETRIES
                    );

                    const currentOU = normalizeOuPath_(String(u?.orgUnitPath ?? ""));

                    const match = MOVE_CONFIG.STRICT_OLD_OU_MATCH
                        ? (currentOU === oldOUNorm)
                        : (currentOU === oldOUNorm || currentOU.startsWith(oldOUNorm + "/"));

                    if (!match) {
                        out.push(["SKIPPED", `Current OU "${currentOU}" does not match old OU "${oldOUNorm}"`]);
                        continue;
                    }
                }

                // Move
                callWithRetry_(() =>
                    AdminDirectory.Users.update({ orgUnitPath: newOUNorm }, userKey),
                    MOVE_CONFIG.MAX_RETRIES
                );

                out.push(["MOVED", `→ ${newOUNorm}`]);

            } catch (e) {
                const msg = (e && e.message) ? e.message : String(e);
                if (/(notFound|Resource Not Found|404)/i.test(msg)) {
                    out.push(["ERROR", "User not found (404)"]);
                } else if (/(forbidden|403|insufficient|not authorized)/i.test(msg)) {
                    out.push(["ERROR", "No permission / not authorized (403)"]);
                } else {
                    out.push(["ERROR", msg]);
                }
            }
        }

        // Write results to N:O
        sheet.getRange(MOVE_CONFIG.START_ROW, MOVE_CONFIG.COL_MOVE_STATUS, out.length, 2).setValues(out);
    }
}

/************** HELPERS **************/
function getMoveSheetsToProcess_(ss) {
    if (MOVE_CONFIG.PROCESS_ALL_SHEETS) return ss.getSheets();

    if (MOVE_CONFIG.USE_ACTIVE_SHEET) return [ss.getActiveSheet()];

    const sh = ss.getSheetByName(MOVE_CONFIG.TARGET_SHEET_NAME);
    if (!sh) throw new Error(`Sheet not found: "${MOVE_CONFIG.TARGET_SHEET_NAME}"`);
    return [sh];
}

function normalizeOuPath_(p) {
    let s = String(p ?? "").trim();
    if (!s) return "";
    if (!s.startsWith("/")) s = "/" + s;
    while (s.length > 1 && s.endsWith("/")) s = s.slice(0, -1);
    return s;
}

function callWithRetry_(fn, maxRetries) {
    let attempt = 0;
    while (true) {
        try {
            return fn();
        } catch (e) {
            const msg = (e && e.message) ? e.message : String(e);
            const transient =
                /Rate Limit|Quota|quota|Too many|Service invoked too many times|429|500|503|Backend Error/i.test(msg);

            if (!transient || attempt >= maxRetries) throw e;

            const sleepMs = Math.min(30000, 500 * Math.pow(2, attempt));
            Utilities.sleep(sleepMs);
            attempt++;
        }
    }
}


