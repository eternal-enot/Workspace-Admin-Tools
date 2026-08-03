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
    
    // Колонки A, B, C... Колонка C має індекс 3
    const colCIndex = 3;
    const data = sheet.getRange(1, colCIndex, lastRow, 1).getValues();
    
    const seenEmails = new Set();
    let rowsToDelete = [];
    
    // Починаємо з другого рядка (індекс 1), пропускаючи заголовок
    for (let i = 1; i < data.length; i++) {
        const email = String(data[i][0] || "").trim().toLowerCase();
        
        if (!email) continue;
        
        if (seenEmails.has(email)) {
            // Номер рядка в таблиці — це індекс масиву + 1
            rowsToDelete.push(i + 1);
        } else {
            seenEmails.add(email);
        }
    }
    
    if (rowsToDelete.length === 0) {
        ui.alert("✅ Дублікатів не знайдено.");
        return;
    }
    
    // Видаляємо рядки з кінця до початку, щоб не збилися індекси
    rowsToDelete.reverse();
    
    for (const rowIndex of rowsToDelete) {
        sheet.deleteRow(rowIndex);
    }
    
    ui.alert(`✅ Успішно знайдено та видалено рядків з дублікатами: ${rowsToDelete.length} (залишено тільки перші входження).`);
}
