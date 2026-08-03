/************** MASTERS TRANSFER CONFIG **************/
const MASTERS_CFG = {
    SHEET_NAME: "Архів", // Об'єднано з архівом
    CHECK_COL: 3,    // C — якщо "master" → переносити
    RECORD_TYPE_COL: 3, // C — тип запису на цільовому аркуші
    DROPDOWN_COL: 4, // D — actions
    SOURCE_COL: 2,   // B — Source
    TIMESTAMP_COL: 14, // N — Надіслано
    START_ROW: 2,    // пропускаємо заголовок
    COL_A: 1,
    COL_B: 2,
    COL_C: 3,
    MAX_COLUMNS_TO_COPY: 12, // Копіюємо з A по L
    DROPDOWN_VALUES: ["IDLE", "Move to OU", "Notify Deletion", "Delete Account", "Restore"],
    DROPDOWN_COLORS: {
        "IDLE":            "#fff9c4",  // жовтий
        "Move to OU":      "#ffcc80",  // оранжевий
        "Notify Deletion": "#ce93d8",  // фіолетовий
        "Delete Account":  "#ef9a9a",  // червоний
        "Restore":         "#c8e6c9"   // зелений
    }
};

function helperMoveToMasters_(ss, toMasters) {
    if (toMasters.length === 0) return 0;
    
    const targetSheet = ensureMastersSheet_(ss, toMasters[0].data.length);
    
    // Sort forward
    const sortedRows = [...toMasters].reverse();
    
    const targetLastRow = targetSheet.getLastRow();
    const appendStart = Math.max(targetLastRow + 1, 2);
    const width = Math.max(toMasters[0].data.length, MASTERS_CFG.TIMESTAMP_COL);
    
    const values = sortedRows.map(r => {
        let d = r.data.slice(0, MASTERS_CFG.MAX_COLUMNS_TO_COPY);
        while (d.length < width) d.push("");
        
        d[MASTERS_CFG.RECORD_TYPE_COL - 1] = "master"; // C
        d[MASTERS_CFG.DROPDOWN_COL - 1] = "IDLE"; // D
        d[MASTERS_CFG.SOURCE_COL - 1] = r.sourceName; // B
        
        return d.slice(0, width);
    });
    
    targetSheet.getRange(appendStart, 1, values.length, width).setValues(values);
    
    const dropdownRange = targetSheet.getRange(appendStart, MASTERS_CFG.DROPDOWN_COL, values.length, 1);
    const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(MASTERS_CFG.DROPDOWN_VALUES, true)
        .setAllowInvalid(false)
        .build();
    dropdownRange.setDataValidation(rule);
    dropdownRange.setBackground(MASTERS_CFG.DROPDOWN_COLORS["IDLE"]);
    
    applyMastersFormatting_(targetSheet);
    
    return values.length;
}

/**
 * Створює аркуш "До МАГІСТРІВ", якщо його ще немає.
 */
function ensureMastersSheet_(ss, lastCol) {
    let targetSheet = ss.getSheetByName(MASTERS_CFG.SHEET_NAME);

    if (!targetSheet) {
        targetSheet = ss.insertSheet(MASTERS_CFG.SHEET_NAME);
        
        // Потрібно гарантувати щонайменше 14 колонок (до N)
        const headerWidth = Math.max(lastCol, MASTERS_CFG.TIMESTAMP_COL);
        if (targetSheet.getMaxColumns() < headerWidth) {
            targetSheet.insertColumnsAfter(targetSheet.getMaxColumns(),
                headerWidth - targetSheet.getMaxColumns());
        }
        targetSheet.getRange(1, MASTERS_CFG.RECORD_TYPE_COL).setValue("record_type");
        targetSheet.getRange(1, MASTERS_CFG.DROPDOWN_COL).setValue("actions");
        targetSheet.getRange(1, MASTERS_CFG.SOURCE_COL).setValue("Source");
        targetSheet.getRange(1, MASTERS_CFG.TIMESTAMP_COL).setValue("Надіслано");
        targetSheet.setFrozenRows(1);

        applyMastersFormatting_(targetSheet);
    }

    return targetSheet;
}

/**
 * Conditional formatting для колонки D на аркуші "До МАГІСТРІВ".
 */
function applyMastersFormatting_(targetSheet) {
    const maxRows = targetSheet.getMaxRows();
    if (maxRows < 2) return;

    const colD = targetSheet.getRange(2, MASTERS_CFG.DROPDOWN_COL, maxRows - 1, 1);

    // Очищаємо старі правила
    targetSheet.clearConditionalFormatRules();

    const rules = [];

    for (const val of MASTERS_CFG.DROPDOWN_VALUES) {
        rules.push(SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(val)
            .setBackground(MASTERS_CFG.DROPDOWN_COLORS[val])
            .setRanges([colD])
            .build());
    }

    targetSheet.setConditionalFormatRules(rules);
}

function copyColValidationMasters_(srcSheet, dstSheet, col, dstStartRow, numRows) {
    const srcValidation = srcSheet.getRange(MASTERS_CFG.START_ROW, col).getDataValidation();
    if (!srcValidation) return;

    const dstRange = dstSheet.getRange(dstStartRow, col, numRows, 1);
    dstRange.setDataValidation(srcValidation);
}


