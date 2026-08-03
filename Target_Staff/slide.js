/**
 * Переносить:
 *  - колонку K -> A (тимчасово),
 *  - потім колонку F (яка після першого переносу стане колонкою G) -> A,
 *  - результат: F стане A, K стане B
 *  - далі вставляє 2 порожні колонки після B (тобто нові C та D)
 *
 * ВАЖЛИВО: moveColumns переносить колонку разом із форматами та Data validation,
 * тож випадаючі списки для F та K НЕ мають “злітати”.
 */
function reorderColumns_F_to_A_K_to_B_addCD_() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const maxRows = sheet.getMaxRows();

    const lastCol = sheet.getLastColumn();
    if (lastCol < 11) {
        throw new Error(`На аркуші "${sheet.getName()}" менше 11 колонок (K не існує).`);
    }

    // 1) Перенести K -> A
    sheet.moveColumns(sheet.getRange(1, 11, maxRows, 1), 1);

    // 2) Після цього "стара" F стає колонкою 7 (G), бо перед нею додалась K
    sheet.moveColumns(sheet.getRange(1, 7, maxRows, 1), 1);

    // 3) Додати 2 порожні колонки після B (отримаємо нові C та D)
    sheet.insertColumnsAfter(2, 2);

    SpreadsheetApp.flush();
}
