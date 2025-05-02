// setupDropdowns.gs

/**
 * Dropdown config: add new columns/options by expanding this array.
 * [sheetName, column letter, options array, keyColumn (default: 3 = C)]
 */
const DROPDOWN_CONFIG = [
  ['Form responses 1', "V", ['Yes', 'No', 'Tidak Jadi Tes']],
  ['Form responses 1', "AG", ['Sent', 'Confirmed', 'Sent-No Answer']],
  ['Form responses 1', "AX", ['LUNAS', 'OKE', '😡', 'CEK', 'Nama Beda', 'Tidak Ada Nama', 'PALSU', 'SALAH BUKTI', 'Jumlah Salah', 'Pindah Pelatihan']]
];

function setupAllDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  DROPDOWN_CONFIG.forEach(cfg => {
    const [sheetName, colA, options, keyColC] = cfg;
    const keyCol = keyColC || 3; // default: column C
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const col = toColNum(colA);
    const startRow = 2;
    const lastRow = getLastNonEmptyRow(sheet, keyCol);

    const rowCount = Math.max(1, lastRow - startRow + 1); // Never negative
    const range = sheet.getRange(startRow, col, rowCount);
    const values = sheet.getRange(startRow, keyCol, rowCount).getValues();

    // Prepare dropdown rule
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(options)
      .setAllowInvalid(false)
      .build();

    // Apply validation only where C is not empty
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] !== "") {
        range.getCell(i + 1, 1).setDataValidation(rule);
      } else {
        range.getCell(i + 1, 1).clearDataValidations();
      }
    }
  });
}

// Helper: Convert 'AX' => 50, etc
function toColNum(colA) {
  let base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  let num = 0;
  for (let i = 0; i < colA.length; i++) {
    num = num * 26 + (base.indexOf(colA.charAt(i)) + 1);
  }
  return num;
}

// Helper: get last row with data in given col (defaults to C)
function getLastNonEmptyRow(sheet, col = 3) {
  const values = sheet.getRange(2, col, sheet.getLastRow() - 1, 1).getValues().map(r => r[0]);
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i] != "") return i + 2;
  }
  return 2;
}