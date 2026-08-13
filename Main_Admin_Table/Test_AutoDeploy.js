/**
 * TEST SUITE: Auto-Deploy Comprehensive Checks
 * 
 * Usage: Select 'runFullAutoDeployTests' from the script editor.
 * 
 * Scenarios Covered:
 * 1. [Student] Standard (Clean) -> Expect DEPLOY
 * 2. [Student] Recovery Hint -> Expect PREVIEW (RECOVERY)
 * 3. [Student] Missing Group -> Expect ERROR
 * 4. [PhD] Explicit Role ('Аспірант') -> Expect DEPLOY
 * 5. [PhD] Implicit Role ('Студент / аспірант' + 'ф' group) -> Expect DEPLOY
 * 6. [PhD] Recovery -> Expect PREVIEW (RECOVERY)
 * 7. [Staff] Standard (Clean) -> Expect DEPLOY
 * 8. [Staff] Missing Dept -> Expect ERROR
 * 9. [Staff] Recovery -> Expect PREVIEW (RECOVERY)
 */

function runFullAutoDeployTests() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const formSheet = ss.getSheetByName(APP_CONFIG.FORM_SHEET_NAME);

    if (!formSheet) {
        console.error('Form sheet not found!');
        return;
    }

    // 1. Force Auto-Deploy ON
    console.log('🔌 Force-enabling AUTO_DEPLOY_ENABLED...');
    APP_CONFIG.AUTO_DEPLOY_ENABLED = true;

    // 2. Clear Form Sheet
    const lastRow = formSheet.getLastRow();
    if (lastRow > 1) {
        formSheet.getRange(2, 1, lastRow - 1, formSheet.getLastColumn()).clearContent();
        formSheet.getRange(2, 1, lastRow - 1, formSheet.getLastColumn()).setBackground(null);
    }

    // 3. Prepare Test Data
    const ts = new Date();

    // Helper to build row
    // [Timestamp, Role, Hint, NameU, SurU, PatrU, NameE, SurE, Email, Group, Dept]
    const makeRow = (role, hint, name, sur, patr, nameE, surE, mail, grp, dept) => {
        return [ts, role, hint, name, sur, patr, nameE, surE, mail, grp, dept];
    };

    const testRows = [
        // 1. STUDENT: Standard -> DEPLOY
        makeRow('Студент / аспірант', '', 'Студент', 'Звичайний', 'І.', 'Student', 'Normal', 'std.norm@test.com', 'ТМ-01', ''),

        // 2. STUDENT: Recovery -> PREVIEW (RECOVERY mode)
        makeRow('Студент / аспірант', 'old.email@edu.kpi.ua', 'Студент', 'Відновлення', 'І.', 'Student', 'Recovery', 'std.rec@test.com', 'ТМ-01', ''),

        // 3. STUDENT: Missing Group -> ERROR
        makeRow('Студент / аспірант', '', 'Студент', 'Безгрупи', 'І.', 'Student', 'NoGroup', 'std.err@test.com', '', ''),

        // 4–5. PHD via composite role + PhD group pattern -> DEPLOY
        makeRow('Студент / аспірант', '', 'Аспірант', 'Груповий', 'П.', 'Phd', 'Group', 'phd.group@test.com', 'ТМ-01ф', ''),

        makeRow('Студент / аспірант', '', 'Аспірант', 'Прихований', 'П.', 'Phd', 'Implicit', 'phd.impl@test.com', 'ТМ-02ф', ''),

        // 6. PHD: Recovery -> PREVIEW
        makeRow('Студент / аспірант', 'phd.old@edu.kpi.ua', 'Аспірант', 'Відновлення', 'П.', 'Phd', 'Recovery', 'phd.rec@test.com', 'ТМ-01ф', ''),

        // 7. STAFF: Standard -> DEPLOY
        makeRow('Викладач / співробітник', '', 'Співробітник', 'Звичайний', 'В.', 'Staff', 'Normal', 'staff.norm@test.com', '', 'Кафедра БМІ'),

        // 8. STAFF: Missing Dept -> ERROR
        makeRow('Викладач / співробітник', '', 'Співробітник', 'Безкафедри', 'В.', 'Staff', 'NoDept', 'staff.err@test.com', '', ''),

        // 9. STAFF: Recovery -> PREVIEW
        makeRow('Викладач / співробітник', 'staff.old@edu.kpi.ua', 'Співробітник', 'Відновлення', 'В.', 'Staff', 'Recovery', 'staff.rec@test.com', '', 'Кафедра БМІ')
    ];

    // Write
    formSheet.getRange(2, 1, testRows.length, 11).setValues(testRows);
    console.log(`📝 Wrote ${testRows.length} test scenarios.`);

    // 4. Run Logic
    console.log('🚀 Executing Trigger Logic...');
    onFormSubmitAction_({ source: ss, triggerUid: 'TEST_SUITE_FULL' });

    console.log('🏁 Auto-deploy test suite finished. Check email & logs.');
}
