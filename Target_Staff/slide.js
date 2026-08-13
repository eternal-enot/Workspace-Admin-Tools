/**
 * Moves:
 *  - column K -> A (temporarily),
 *  - then column F (which becomes G after the first move) -> A,
 *  - result: F becomes A, K becomes B
 *  - then inserts 2 empty columns after B (new C and D)
 *
 * IMPORTANT: moveColumns moves the column with formats and data validation,
 * so dropdowns on F and K should not break.
 */
function reorderColumns_F_to_A_K_to_B_addCD_() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const maxRows = sheet.getMaxRows();

    const lastCol = sheet.getLastColumn();
    if (lastCol < 11) {
        throw new Error(`На аркуші "${sheet.getName()}" менше 11 колонок (K не існує).`);
    }

    // 1) Move K -> A
    sheet.moveColumns(sheet.getRange(1, 11, maxRows, 1), 1);

    // 2) After that, old F becomes column 7 (G) because K was inserted before it
    sheet.moveColumns(sheet.getRange(1, 7, maxRows, 1), 1);

    // 3) Insert 2 empty columns after B (new C and D)
    sheet.insertColumnsAfter(2, 2);

    SpreadsheetApp.flush();
}
