function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu("Support Tools")
        .addItem("Check Status & Info", "checkSupportInfo")
        .addItem("Remove Duplicates (Check sheet)", "removeDuplicatesInCheck")
        .addSeparator()
        .addItem("Execute Deletions", "executeSupportDeletions")
        .addToUi();
}
