// styling.gs

const STYLING_TARGET_SHEETS = [
  'Form responses 1',
  // ...other sheets...
];

const FORM_RESPONSES_1_COLOR_BANDS = [
  'V-AH',
  'AI-AL',
  'AM-AR',
  'AS-AY',
  'AZ-BC',
  'BD-BI',
  'BJ-BT',
  'BU-BY',
  'BZ-CF',
  'CH-CI'
];

// Pick 10 contrasting color pairs (dark for header, light for data), can use more!
const COLOR_PALETTES = [
  {header:'#1565c0', body:'#90caf9'},    // blue
  {header:'#2e7d32', body:'#a5d6a7'},    // green
  {header:'#ad1457', body:'#f8bbd0'},    // pink
  {header:'#6d4c41', body:'#bcaaa4'},    // brown
  {header:'#ff8f00', body:'#ffe082'},    // amber
  {header:'#c62828', body:'#ef9a9a'},    // red
  {header:'#4527a0', body:'#b39ddb'},    // purple
  {header:'#00838f', body:'#80deea'},    // teal
  {header:'#607d8b', body:'#cfd8dc'},    // blueish grey
  {header:'#689f38', body:'#dcedc8'},    // lime
];

// Helper: convert A1 or "BZ" etc to number
function colAtoNum(colA) {
  let n=0; for(let i=0; i<colA.length; i++) n=n*26 + (colA.charCodeAt(i)-64);
  return n;
}

// Helper: decide text color (white for dark, black for light backgrounds)
function getAutoFontColor(bg) {
  // bg: "#rrggbb"
  if(!bg || !bg.match(/^#[0-9a-f]{6}$/i)) return "#000000";
  let r = parseInt(bg.substr(1,2),16);
  let g = parseInt(bg.substr(3,2),16);
  let b = parseInt(bg.substr(5,2),16);
  // Simple luminance algorithm
  let luma = 0.2126*r + 0.7152*g + 0.0722*b;
  return luma < 150 ? "#ffffff" : "#212121";
}

function applyAllStyling() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  STYLING_TARGET_SHEETS.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    headerRange.setFontWeight("bold");

    if (sheetName === 'Form responses 1') {
      const lastRow = Math.max(2, sheet.getLastRow());
      const lastCol = sheet.getLastColumn();
      sheet.getRange(1, 1, lastRow, lastCol).setBackground(null).setFontColor("#212121");

      FORM_RESPONSES_1_COLOR_BANDS.forEach((band, idx) => {
        let [colStart, colEnd] = band.split('-').map(colAtoNum);

        // --- PALETTE OVERRIDE for AS-AY
        let dark, light;
        if (band === "AS-AY") {
          dark = "#2e7d32";
          light = "#a5d6a7";
        } else {
          const palette = COLOR_PALETTES[idx % COLOR_PALETTES.length];
          dark = palette.header;
          light = palette.body;
        }

        let headerFont = getAutoFontColor(dark), bodyFont = getAutoFontColor(light);

        // Header
        sheet.getRange(1, colStart, 1, colEnd - colStart + 1).setBackground(dark).setFontColor(headerFont);
        // Data
        if (lastRow > 1)
          sheet.getRange(2, colStart, lastRow - 1, colEnd - colStart + 1).setBackground(light).setFontColor(bodyFont);
      });
    }
  });
}