function removeDuplicatesInCheck() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = "Check";
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
        ui.alert(`❌ Аркуш "${sheetName}" не знайдено.`);
        return;
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { // Assuming row 1 is headers
        ui.alert("🤷‍♂️ Немає даних для перевірки.");
        return;
    }
    
    // Columns A, B, C... Column C is index 3
    const colCIndex = 3;
    const data = sheet.getRange(1, colCIndex, lastRow, 1).getValues();
    
    const seenEmails = new Set();
    let rowsToDelete = [];
    
    // Start from the second row (index 1), skip header
    for (let i = 1; i < data.length; i++) {
        const email = String(data[i][0] || "").trim().toLowerCase();
        
        if (!email) continue;
        
        if (seenEmails.has(email)) {
            // Sheet row number = array index + 1
            rowsToDelete.push(i + 1);
        } else {
            seenEmails.add(email);
        }
    }
    
    if (rowsToDelete.length === 0) {
        ui.alert("✅ Дублікатів не знайдено.");
        return;
    }
    
    // Delete from bottom to top so indices stay valid
    rowsToDelete.reverse();
    
    for (const rowIndex of rowsToDelete) {
        sheet.deleteRow(rowIndex);
    }
    
    ui.alert(`✅ Успішно знайдено та видалено рядків з дублікатами: ${rowsToDelete.length} (залишено тільки перші входження).`);
}
