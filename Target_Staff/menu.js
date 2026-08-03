/**
 * Main menu for the application.
 */
function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu("Admin tools")
        .addItem("🚀 РАЗОВА МІГРАЦІЯ (Оновити таблицю)", "runOneTimeMigration")
        .addSeparator()
        .addItem("Update Status (Active Sheet)", "checkUsersActiveSheet")
        .addItem("Update Status (All Sheets)", "checkUsersAllSheets")
        .addSeparator()
        .addItem("Execute group actions (Active Sheet)", "executeGroupActionsActiveSheet")
        .addItem("Execute group actions (All Sheets)", "executeGroupActionsAllSheets")
        .addToUi();
}
