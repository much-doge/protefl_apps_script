// autoCounters.gs

/**
 * Protects column R (Original Schedule) in Form responses 1 for all except the owner.
 * Run this once or periodically in your admin script.
 */
function protectOriginalScheduleColumn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Form responses 1');
  if (!sheet) return;
  var col = 18; // R = column 18
  // Remove previous protection of column R (if any)
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protections.forEach(function (protection) {
    var rng = protection.getRange();
    if (rng && rng.getColumn() === col && rng.getNumColumns() === 1) protection.remove();
  });
  // Apply new protection
  var range = sheet.getRange(1, col, sheet.getMaxRows()); // entire R col
  var protection = range.protect().setDescription('Original ProTEFL Schedule');
  // Only allow owner to edit:
  var me = Session.getEffectiveUser();
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors().filter(e => e.getEmail() !== me.getEmail()));
  protection.setWarningOnly(false); // Prevent other editors
}


/**
 * onEdit trigger: When V or W is changed, update reschedule log (column X) and count (column Y).
 */
function onEditLogReschedule(e) {
  var range = e.range;
  var sheet = range.getSheet();
  if (sheet.getName() !== 'Form responses 1') return;
  var editedCol = range.getColumn();
  var row = range.getRow();
  if (row < 2) return; // Skip headers

  var colV = 22, colW = 23, colR = 18, colX = 24, colY = 25;

  // Only handle edits in V or W columns
  if (editedCol === colV || editedCol === colW) {
    var valC = sheet.getRange(row, 3).getValue();
    var valV = sheet.getRange(row, colV).getValue();
    var valW = sheet.getRange(row, colW).getValue();
    var valR = sheet.getRange(row, colR).getValue();
    var valX = sheet.getRange(row, colX).getValue();

    if (!valC) return; // don't log if row isn't used

    // If V != "Yes", do not log reschedule, always use R (reset X and Y if not empty)
    if (valV !== "Yes") {
      if (valX !== "") {
        sheet.getRange(row, colX).setValue("");
        sheet.getRange(row, colY).setValue("");
      }
      return;
    }

    // If V == Yes, log the value in W (if not already present in X or if changed)
    var logArr = valX ? valX.split(/\s*;\s*/) : [];
    if (valW && (logArr.length === 0 || logArr[logArr.length - 1] !== valW)) {
      logArr.push(valW);
      // Remove empty from trailing (could happen with manual clear)
      logArr = logArr.filter(x=>x && x.trim()!=="");
      sheet.getRange(row, colX).setValue(logArr.join("; "));
      // Count reschedules in Y
      sheet.getRange(row, colY).setValue(logArr.length);
    }
  }

  // Also: If X is changed manually, count again
  if (editedCol === colX) {
    var valX2 = sheet.getRange(row, colX).getValue();
    var count = valX2 ? valX2.split(/\s*;\s*/).filter(a=>a && a.trim()!=="").length : "";
    sheet.getRange(row, colY).setValue(count);
  }
}

/**
 * (Recommended)
 * Optionally run onOpen to make sure all reschedule counts (column Y) match log (column X) for all filled rows.
 */
function syncRescheduleCounts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Form responses 1');
  var lastRow = sheet.getLastRow();
  var valsX = sheet.getRange(2, 24, lastRow-1, 1).getValues();
  for (var i = 0; i < valsX.length; i++) {
    var x = valsX[i][0];
    var count = x ? x.split(/\s*;\s*/).filter(a=>a && a.trim()!=="").length : "";
    sheet.getRange(i+2, 25).setValue(count);
  }
}