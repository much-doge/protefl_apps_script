/**
 * =============================================================================
 * ProTEFL MDMA ⚡ - Google Sheets Apps Script(s)
 * =============================================================================
 *
 * Copyright (c) 2025 Nur Eko Windianto (ne.windianto@gmail.com)
 * All rights reserved.
 *
 * You are granted permission to use, copy, and modify this software **for your
 * personal use only**. Redistribution or commercial use
 * without explicit permission from the author is prohibited.
 *
 * Author: Nur Eko Windianto
 * Created: 2025-04-30
 *
 * Notes:
 * - This script is intended for managing ProTEFL registration, scoring, and data.
 * - Unauthorized redistribution or resale is strictly forbidden.
 *
 * =============================================================================
 */


// ============================================================================
// File: main.gs
// 
// MAIN ORCHESTRATOR
// Entry point to set up or refresh the entire ProTEFL Montly Data Management Admin workbook.
// Run this only when initializing or re-initializing (destructive).
// ============================================================================

function main() {
  // --------------------------------------------------------------------------
  // Prerequisite: DATABASEMAHASISWA must exist
  // --------------------------------------------------------------------------

  // Stop immediately and prompt for any missing consent before destructive work.
  ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);

  const dbSuccess = pullDatabaseMahasiswa(); // returns true/false
  if (!dbSuccess) {
    SpreadsheetApp.getUi().alert(
      "Setup Aborted ❌",
      "'DATABASEMAHASISWA' pull failed. Main setup stopped.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return; // stop main
  }

  // --------------------------------------------------------------------------
  // Pull FABULASI
  // --------------------------------------------------------------------------
  const fabSuccess = pullFabulasi();
  if (!fabSuccess) {
    SpreadsheetApp.getUi().alert(
      "Setup Aborted ❌",
      "'10. FABULASI' pull failed. Main setup stopped.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  initializeSheets();
  setupAllDropdownsWithDummy();
  applyAllStyling();
  applyAllFormulas();
  setupDefaultViewTrigger();
  protectOriginalScheduleColumn();
  installRescheduleTrigger();
  ensureStylingTrigger();

  SpreadsheetApp.getUi().alert(
    "Main Completed ✅",
    "All setup steps finished successfully.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ----------------------------------------------------------------------------
// Prerequisite for main. Pull DATABASEMAHASISWA from source with success/failure dialogs
// ----------------------------------------------------------------------------
function authorizeSheetsAccess() {
  const ui = SpreadsheetApp.getUi();

  // Check first so we can give a cleaner success message when already authorized.
  const authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);

  if (authInfo.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.REQUIRED) {
    // This ends the current execution and shows the consent dialog.
    ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
    return;
  }

  // Harmless sanity touches so the menu acts as a true preflight.
  SpreadsheetApp.getActiveSpreadsheet().getId();
  SpreadsheetApp.openByUrl(ENV.DATABASE_URL).getName();
  SpreadsheetApp.openByUrl(FABULASI_URL).getName();
  ScriptApp.getProjectTriggers();
  HtmlService.createHtmlOutput("ok");
  DriveApp.getRootFolder().getName();

  ui.alert(
    "Authorization Complete ✅",
    "All required access is already granted. You can now run '01. Initialize Sheet' without mid-run authorization prompts.",
    ui.ButtonSet.OK
  );
}

function pullDatabaseMahasiswa() {
  const ui = SpreadsheetApp.getUi();
  const destSS = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const sourceSS = SpreadsheetApp.openByUrl(ENV.DATABASE_URL);
    const sourceSheet = sourceSS.getSheetByName("DATABASEMAHASISWA");
    if (!sourceSheet) {
      ui.alert("Pull Failed ❌", "Source sheet 'DATABASEMAHASISWA' not found.", ui.ButtonSet.OK);
      return false;
    }

    const existingSheet = destSS.getSheetByName("DATABASEMAHASISWA");
    if (existingSheet) destSS.deleteSheet(existingSheet);

    const copiedSheet = sourceSheet.copyTo(destSS);
    copiedSheet.setName("DATABASEMAHASISWA");
    destSS.setActiveSheet(copiedSheet);
    destSS.moveActiveSheet(1);

    ui.alert("Pull Successful ✅", "'DATABASEMAHASISWA' copied successfully.", ui.ButtonSet.OK);
    return true;

  } catch (e) {
    ui.alert("Pull Failed ❌", "Error pulling sheet: " + e.message, ui.ButtonSet.OK);
    return false;
  }
}

// ----------------------------------------------------------------------------
// Pull FABULASI sheet from external spreadsheet
// ----------------------------------------------------------------------------
function pullFabulasi() {
  const ui = SpreadsheetApp.getUi();
  const destSS = SpreadsheetApp.getActiveSpreadsheet();

  const FABULASI_URL = "https://docs.google.com/spreadsheets/d/1yGr6JFLn3t4qCLHgqKNuX8q5QxDUnlY6mU_QuESR7cU/edit";

  try {
    const sourceSS = SpreadsheetApp.openByUrl(FABULASI_URL);
    const sourceSheet = sourceSS.getSheetByName("10. FABULASI");

    if (!sourceSheet) {
      ui.alert("Pull Failed ❌", "Source sheet '10. FABULASI' not found.", ui.ButtonSet.OK);
      return false;
    }

    const existingSheet = destSS.getSheetByName("10. FABULASI");
    if (existingSheet) destSS.deleteSheet(existingSheet);

    const copiedSheet = sourceSheet.copyTo(destSS);
    copiedSheet.setName("10. FABULASI");

    ui.alert("Pull Successful ✅", "'10. FABULASI' copied successfully.", ui.ButtonSet.OK);
    return true;

  } catch (e) {
    ui.alert("Pull Failed ❌", "Error pulling FABULASI: " + e.message, ui.ButtonSet.OK);
    return false;
  }
}

// ----------------------------------------------------------------------------
// Wrapper: safely run setupAllDropdowns, insert dummy row if empty
// ----------------------------------------------------------------------------
function setupAllDropdownsWithDummy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Form responses 1");
  if (!sheet) return;

  // Check if there is at least one real row beyond header
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { // empty or only header
    // Insert dummy row to avoid dropdown errors
    sheet.getRange("A2").setValue(new Date()); // date
    sheet.getRange("B2").setValue("dummy@student.uny.ac.id"); // email
    sheet.getRange("C2").setValue("ProTEFL SIAKAD UNY (tanpa sertifikat)"); // test option
    sheet.getRange("D2").setValue("23021340999"); // random 11-digit student ID
    sheet.getRange("E2").setValue("Dummy Participant"); // name
    sheet.getRange("F2").setValue("'081234567890"); // phone, leading apostrophe
    sheet.getRange("R2").setValue("Jumat, 12 September 2045 - OFFLINE PAGI 09.20-11.30 WIB"); // schedule
  }

  // Now run the original dropdown setup
  setupAllDropdowns();
}

// ----------------------------------------------------------------------------
// Ensures a time-driven trigger exists for applyAllStyling().
// Creates one if none exists.
// ----------------------------------------------------------------------------
function ensureStylingTrigger() {
  const existing = ScriptApp.getProjectTriggers().filter(
    t => t.getHandlerFunction() === "applyAllStyling"
  );

  if (existing.length === 0) {
    ScriptApp.newTrigger("applyAllStyling")
             .timeBased()
             .everyMinutes(30)
             .create();
    Logger.log("Trigger created: applyAllStyling() will run every 30 minutes.");
  } else {
    Logger.log("Trigger already exists. No new trigger created.");
  }
}


// ============================================================================
// File: menu.gs
// 
// MENU SETUP
// Builds the "ProTEFL Utility" custom menu with safe options, exports, risky
// admin actions, and quick-access custom views.
// Runs automatically on spreadsheet open.
// ============================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("ProTEFL Utility")
      // --- Safe options ---
      .addItem("Fix Column CD", "fixColumnCD")
      .addItem("Fix Institution Letterhead Template", "insertULBHeader")
      .addItem("Apply Styles", "applyAllStylingWithConfirm")
      .addItem("Protect Original Schedule Column", "protectOriginalScheduleColumn")
      .addItem("Set Up AutoCounter Trigger", "setupAutoCounterTriggerWithAlert")
      .addSeparator()
      .addSubMenu(
        SpreadsheetApp.getUi()
          .createMenu("Export")
          .addItem("00. Participant Test IDs", "exportParticipantTestIds")
          .addItem("01. Download VCF by Tanggal Tes", "downloadVCFFromMenu")
          .addItem("02. Copy Attendance List", "copyAttendanceList")
          .addItem("03. Print Attendance Sheet", "generateAttendanceSheet")
          .addItem("04. Copy Certificate Data", "copyCertificateData")
          .addItem("05. Export Participant Scores", "exportSiakadScoreResults")
      )
      .addSeparator()
      // --- Risky options ---
      .addItem("00. Authorize Access (Run Once)", "authorizeSheetsAccess")
      .addItem("01. Initialize Sheet (Danger Zone)", "runMainWithConfirm")
      .addItem("Apply All Formulas (Danger Zone)", "applyAllFormulasWithConfirm")
      .addItem("Pull DATABASEMAHASISWA manually (Danger Zone)", "pullDatabaseMahasiswa")
      .addSeparator()
      .addSubMenu(
        SpreadsheetApp.getUi()
          .createMenu("External Form Import")
          .addItem("00. Authorize External Import Access", "mdmaAuthorizeExternalImportAccess")
          .addItem("01. Seed External Form Registry", "mdmaSeedExternalFormRegistry")
          .addItem("02. Sync External Form Responses", "mdmaSyncExternalFormResponsesWithUi")
          .addItem("03. Install External Import Triggers", "mdmaInstallExternalImportTriggers")
          .addItem("99. Remove External Import Triggers", "mdmaRemoveExternalImportTriggers")
      )
      .addSeparator()
      // --- Custom views ---
      .addSubMenu(
        SpreadsheetApp.getUi()
          .createMenu("Custom View")
          .addItem("00. Verify Student ID", "toggleVerifyStudentIDView")
          .addItem("01. Verify Payment", "toggleVerifyPaymentView")
          .addItem("03. Grouping & Contacts", "toggleGroupingContactsView")
          .addItem("Reschedule Participants", "toggleRescheduleParticipantsView")
          .addItem("Verify Attendance", "toggleVerifyAttendanceView")
          .addItem("PPB (Pusing Pala Barbie)", "togglePPBView")
          .addItem("Default View", "toggleDefaultView")
          .addItem("Reset View", "resetView")
      )
    .addToUi();

  toggleDefaultView(true); // Always open default view on launch
}


// ----------------------------------------------------------------------------
// MENU ACTION WRAPPERS
// Safe prompts before executing styling, formula injection, or initialization.
// Prevents accidental destructive changes.
// ----------------------------------------------------------------------------

/** Ask confirmation before reapplying styles */
function applyAllStylingWithConfirm() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert(
    "Apply Styles",
    "Re-apply all custom styles (headers, banding, formatting)?",
    ui.ButtonSet.OK_CANCEL
  ) == ui.Button.OK) {
    applyAllStyling();
  }
}

/** Ask confirmation before applying all formulas (danger zone) */
function applyAllFormulasWithConfirm() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert(
    "Apply All Formulas (Danger Zone)",
    "Ensure DATABASEMAHASISWA exists. Missing sheets will cause errors. Proceed?",
    ui.ButtonSet.OK_CANCEL
  ) == ui.Button.OK) {
    applyAllFormulas();
  }
}

/** Ask confirmation before full initialization (irreversible) */
function runMainWithConfirm() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert(
    "Initialize Sheet (Danger Zone)",
    "This will initialize/reinitialize your workbook. NOT reversible. Proceed?",
    ui.ButtonSet.OK_CANCEL
  ) == ui.Button.OK) {
    main();
  }
}

/** Ask confirmation before installing auto counter trigger */
function setupAutoCounterTriggerWithAlert() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert(
    "Set Up Trigger",
    "Create/replace the onEdit trigger for auto counter logging. Proceed?",
    ui.ButtonSet.OK_CANCEL
  ) == ui.Button.OK) {
    setupAutoCounterTrigger();
  }
}

// ----------------------------------------------------------------------------
// TRIGGER MANAGEMENT
// Installs or refreshes installable triggers for auto counter logging
// and opening the default view.
// ----------------------------------------------------------------------------

/** Replace existing reschedule trigger with a fresh one */
function setupAutoCounterTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "onEditLogReschedule")
      ScriptApp.deleteTrigger(trigger);
  });
  ScriptApp.newTrigger("onEditLogReschedule")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
}

/** Ensure onOpen trigger is installed to always load default view */
function setupDefaultViewTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var exists = triggers.some(t => t.getHandlerFunction() === "onOpenDefaultView");
  if (!exists) {
    ScriptApp.newTrigger("onOpenDefaultView")
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onOpen()
      .create();
  }
}

/** Handler for default view trigger */
function onOpenDefaultView() {
  toggleDefaultView(true);
}

// ----------------------------------------------------------------------------
// CUSTOM VIEWS (Optimized, Reliable Toggle)
// Central engine for hiding/showing specific column sets per view. 
// - Persists current view in DocumentProperties
// - Can toggle on/off, or force re-activation
// - Optionally launches a matching sidebar
//
//
// * Core function to apply a custom view.
// * @param {string} sheetName   Target sheet name
// * @param {string[]} keepCols  Array of column letters to remain visible
// * @param {function} sidebarFn Optional sidebar renderer for this view
// * @param {string} label       Unique view identifier
// * @param {boolean} forceOn    Force view on (bypass toggle logic)
// ----------------------------------------------------------------------------
function applyCustomView_(sheetName, keepCols, sidebarFn, label, forceOn) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return;

  var props = PropertiesService.getDocumentProperties();
  var currentView = props.getProperty("currentView") || "";
  var keepIndexes = keepCols.map(letterToColumn_);
  var lastCol = sheet.getLastColumn();

  // If already in this view and not forced, do nothing
  if (currentView === label && !forceOn) return;

  // Reset → show everything before hiding the columns we don’t need
  sheet.showColumns(1, lastCol);

  // Hide all except keepCols
  var rangesToHide = [];
  var start = null;
  for (var col = 1; col <= lastCol; col++) {
    if (!keepIndexes.includes(col)) {
      if (start === null) start = col;
    } else if (start !== null) {
      rangesToHide.push([start, col - start]);
      start = null;
    }
  }
  if (start !== null) rangesToHide.push([start, lastCol - start + 1]);
  rangesToHide.forEach(r => sheet.hideColumns(r[0], r[1]));

  if (sidebarFn) sidebarFn();
  props.setProperty("currentView", label);

  // Ensure default view is reinstalled if needed
  if (label === "Default") setupDefaultViewTrigger();
}

// ----------------------------------------------------------------------------
// Reset view (show all columns)
// ----------------------------------------------------------------------------
function resetView() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sheet) return;
  var lastCol = sheet.getLastColumn();
  sheet.showColumns(1, lastCol);

  // Clear stored view
  PropertiesService.getDocumentProperties().setProperty("currentView", "");
}

// ............................................................................
// INDIVIDUAL VIEW TOGGLES
// Each defines which columns stay visible and which sidebar to launch.
// ............................................................................

/** Show lean "Default" view (basic registration essentials) */
function toggleDefaultView(forceOn) {
  var keepCols = ["A","AI","AJ","AN","AO","BB","BC","BJ","BT","BX", "CI", "CK"];
  applyCustomView_("Form responses 1", keepCols, showDefaultSidebar, "Default", forceOn);
}

/** Focus on rescheduling participants (schedule + comms columns) */
function toggleRescheduleParticipantsView() {
  var keepCols = ["A","C","D","E","G","R","V","W","X","Y","AE","AF","AG","AH","AL","AM","AN","AO","AV","AX","BI", "BJ", "CI"];
  applyCustomView_("Form responses 1", keepCols, showRescheduleSidebar, "Reschedule Participants");
}

/** Verify student IDs (identity & student database link) */
function toggleVerifyStudentIDView() {
  var keepCols = ["C","D","E","AZ","BA","BB","BC", "BJ", "CI"];
  applyCustomView_("Form responses 1", keepCols, showVerifyStudentIDSidebar, "Verify Student ID");
}

/** Verify payments (proof columns + payment status) */
function toggleVerifyPaymentView() {
  var keepCols = ["A", "C", "D", "G", "AI", "AN", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "BI", "BJ", "CI"];
  applyCustomView_("Form responses 1", keepCols, showVerifyPaymentSidebar, "Verify Payment");
}

/** PPB (Pusing Pala Barbie) custom view */
function togglePPBView() {
  var keepCols = [
    "A","C","D","E","F","V","W",
    "AE","AF","AG","AH","AI",
    "AN","AO","AP",
    "AX","AY",
    "BB","BC",
    "BI","BJ",
    "BO","BS","BT","BU","BV","BW","BX",
    "BZ","CB","CD","CF","CG","CI","CJ"
  ];
  applyCustomView_("Form responses 1", keepCols, showPPBSidebar, "PPB");
}

/** Verify attendance (test date, codes, and presence fields) */
function toggleVerifyAttendanceView() {
  var keepCols = [
    "A","C","D","G","V","W","AI","AJ","AL","AN","AO", "AX", "BB",
    "BC","BI","BJ","BL","BN","BO","BP","BQ","BS","BT",
    "BU","BV","BW","BX","CB","CG", "CI", "CJ", "CK", "CL", "CM", "CN"
  ];
  applyCustomView_("Form responses 1", keepCols, showVerifyAttendanceSidebar, "Verify Attendance");
}

/** Group participants & manage contacts (IDs + contact columns) */
function toggleGroupingContactsView() {
  const keepCols = ["A", "F", "G", "AI", "AJ", "AL", "AM", "AN", "AO", "AP", "AQ", "BE", "BG", "BI", "BJ", "CI"];
  applyCustomView_("Form responses 1", keepCols, showGroupingContactsSidebar, "Grouping & Contacts");
}


// ============================================================================
// File: utilities.gs
// UTILITIES (Shared Tools & Export Features)
// Core utilities that streamline recurring admin tasks across the ProTEFL 
// registration workbook. These functions are not "small helpers" — they 
// automate and save significant time by handling repetitive or error-prone 
// processes.
// 
// Key features provided here:
//   • Exporting participant contact info as VCF files (per test date)
//   • Downloading participant Test IDs to Excel
//   • Copying attendance lists into plain-text (tab-delimited) format for use 
//     in other sheets or systems
//   • Exporting participant score results to Excel
//   • General helpers (e.g., column-letter conversion)
// 
// In short: this file centralizes all the heavy-duty utilities that 
// make the ProTEFL admin workflow smoother, faster, and more reliable.
// ============================================================================

// -----------------------------------------------------------------------------
// HELPER: Letter → Column Index
// Converts spreadsheet column letters (e.g. "A", "AX") to their numeric index.
// Example: "A" → 1, "Z" → 26, "AA" → 27.
// -----------------------------------------------------------------------------
function letterToColumn_(letter) {
  var col = 0;
  for (var i = 0; i < letter.length; i++) col = col * 26 + (letter.charCodeAt(i) - 64);
  return col;
}

// -----------------------------------------------------------------------------
// VCF EXPORT MENU ENTRY
// Triggered from the custom menu. Prompts user for "Tanggal Tes" (yyyy-MM-dd)
// and generates a downloadable VCF file if matching entries are found.
// -----------------------------------------------------------------------------
function downloadVCFFromMenu() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Enter Tanggal Tes (yyyy-MM-dd) to download VCF:");

  if (response.getSelectedButton() != ui.Button.OK) return;
  const date = response.getResponseText().trim();

  const result = exportVCF(date);

  if (!result.success) {
    // Show error in HTML dialog including the entered date
    const html = `
      <div style="
          font-family: 'Google Sans', Arial, sans-serif; 
          padding:16px; 
          line-height:1.5; 
          background:#fefefe; 
          color:#222; 
          border-radius:10px; 
          box-shadow:0 2px 5px rgba(0,0,0,0.15);
      ">
        <h2 style="margin-top:0; color:#c62828;">⚠️ VCF Download Error</h2>
        <p>No entries found for the test date "<b>${date}</b>".</p>
        <p>Check column <b>BJ</b> ('Tanggal tes') for existing test dates and make sure you entered the date correctly (format: yyyy-mm-dd).</p>
        <button onclick="google.script.host.close()" style="
            background:#1e88e5;
            color:white;
            border:none;
            border-radius:6px;
            padding:8px 12px;
            cursor:pointer;
        ">Close</button>
      </div>
    `;
    ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(450).setHeight(220), "VCF Download Error");
    return;
  }

  // Otherwise, show the download link in similar tidy style
  const html = `
    <div style="
        font-family: 'Google Sans', Arial, sans-serif; 
        padding:16px; 
        line-height:1.5; 
        background:#fefefe; 
        color:#222; 
        border-radius:10px; 
        box-shadow:0 2px 5px rgba(0,0,0,0.15);
    ">
      <h2 style="margin-top:0; color:#2e7d32;">✅ VCF Created</h2>
      <p>Your VCF file for <b>${date}</b> has been created in Google Drive.</p>
      <p>
        <a href="${result.url}" target="_blank" style="color:#1e88e5; text-decoration:none;">Click here to open/download the file</a>
      </p>
      <button onclick="google.script.host.close()" style="
          background:#1e88e5;
          color:white;
          border:none;
          border-radius:6px;
          padding:8px 12px;
          cursor:pointer;
      ">Close</button>
    </div>
  `;
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(450).setHeight(200), "VCF Download");
}

// ............................................................................
// CORE VCF EXPORT FUNCTION
// Filters "Form responses 1" by selected Tanggal Tes (column BJ), extracts the
// VCF block (column BG), and saves the .vcf file into "ProTEFL VCFs" folder.
// Returns { success: bool, url?: string, message?: string }.
// ............................................................................
function exportVCF(selection) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form responses 1");
  var data = sheet.getDataRange().getValues();
  var header = data.shift();
  
  var bjIndex = header.indexOf("Tanggal tes");
  var bgIndex = header.indexOf("Grouping VCF");
  if (bjIndex === -1 || bgIndex === -1) {
    return { 
      success: false, 
      message: "Required columns not found: 'Tanggal tes' or 'Grouping VCF'. Please check your sheet headers." 
    };
  }
  
  var filtered = data.filter(row => row[bjIndex] === selection);
  if (filtered.length === 0) {
    return { 
      success: false, 
      message: `No entries found for the test date "${selection}".\n` +
               `Check column BJ ('Tanggal tes') for existing test dates and make sure you entered the date correctly in the download dialog (format: yyyy-mm-dd).`
    };
  }
  
  var vcfData = filtered.map(row => row[bgIndex].replace(/"/g, "")).join("\n");
  var blob = Utilities.newBlob(vcfData, "text/vcard", selection + ".vcf");

  // --- SAVE TO SPECIFIC FOLDER ---
  var folderName = "ProTEFL VCFs"; // change as needed
  var folder, folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }

  var file = folder.createFile(blob); // create inside folder

  return { success: true, url: file.getUrl() };
}

// ............................................................................
// DIALOG RENDERER (Alternative)
// Shows a styled modal dialog for export results. Can be used by other export
// functions too, not only VCF.
// ............................................................................
function showVCFExportDialog(result) {
  let htmlContent;
  if (!result.success) {
    htmlContent = `
      <div style="font-family: 'Google Sans', Arial, sans-serif; padding:20px; background:#f8f9fa; color:#222;">
        <h2 style="margin-top:0; color:#d32f2f;">❌ Export Failed</h2>
        <p style="font-size:14px; line-height:1.5;">${result.message}</p>
      </div>
    `;
  } else {
    htmlContent = `
      <div style="font-family: 'Google Sans', Arial, sans-serif; padding:20px; background:#edf2fa; color:#222;">
        <h2 style="margin-top:0; color:#1e88e5;">✅ VCF Created!</h2>
        <p style="font-size:14px; line-height:1.5;">Your VCF file has been created in Google Drive.</p>
        <p style="margin-top:12px;">
          <a href="${result.url}" target="_blank" 
             style="display:inline-block;padding:8px 12px;background:#1e88e5;color:white;border-radius:6px;text-decoration:none;">
            Open / Download
          </a>
        </p>
      </div>
    `;
  }

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(htmlContent)
      .setWidth(420)
      .setHeight(200),
    "Export VCF"
  );
}

// -----------------------------------------------------------------------------
// COPY ATTENDANCE LIST FUNCTION
// 
// Purpose:
//   Exports a tab-delimited attendance list for a given test date (YYYYMMDD) into clipboard
//   from the "04. BUAT PRESENSI DAN GRUP WA H-1" sheet.
//
// Steps:
//   1. Prompt the user for the test date (format: YYYYMMDD).
//   2. Access the attendance sheet and pull all data.
//   3. Filter rows by Column F ("Test Date") matching the user input.
//   4. If no rows found → show an error modal and exit.
//   5. Sort the filtered rows by Column G (group/class).
//   6. Insert two blank rows whenever the value in Column G changes for clarity (differentiating each group).
//   7. Convert the processed rows into a tab-delimited string.
//   8. Show modal with a textarea containing the result and a "Copy to Clipboard" button.
//
// Notes:
//   - Blank rows are inserted for visual separation of groups.
//   - Designed to quickly paste data into external attendance sheets.
//   - Row count in modal includes the inserted blank lines.
// -----------------------------------------------------------------------------
function copyAttendanceList() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Enter Test Date (YYYYMMDD)",
    "Provide test date:",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const dateFilter = response.getResponseText().trim();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("04. BUAT PRESENSI DAN GRUP WA H-1");
  if (!sheet) return ui.alert("Target sheet not found.");

  let data = sheet.getDataRange().getValues();
  const header = data.shift();
  const dateColIndex = 5; // Column F (zero-based)
  const sortColIndex = 6; // Column G

  // Filter by test date
  data = data.filter(row => String(row[dateColIndex]) === dateFilter);

  if (data.length === 0) {
    // Error modal (VCF style)
    const html = `
      <div style="
          font-family: 'Google Sans', Arial, sans-serif; 
          padding:20px; 
          background:#fefefe; 
          color:#222; 
          border-radius:10px; 
          box-shadow:0 2px 5px rgba(0,0,0,0.15);
      ">
        <h2 style="margin-top:0; color:#c62828;">⚠️ No Entries Found</h2>
        <p>No attendance records found for "<b>${dateFilter}</b>".</p>
        <p>Check column <b>F</b> ('Test Date') for existing values (format: YYYYMMDD).</p>
        <button onclick="google.script.host.close()" style="
            background:#1e88e5;
            color:white;
            border:none;
            border-radius:6px;
            padding:8px 12px;
            cursor:pointer;
        ">Close</button>
      </div>
    `;
    ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(450).setHeight(220), "Copy Attendance Error");
    return;
  }

  // Sort by column G
  data.sort((a, b) => (a[sortColIndex] > b[sortColIndex]) ? 1 : (a[sortColIndex] < b[sortColIndex] ? -1 : 0));

  // Insert two empty rows when G changes
  let lastValue = data[0][sortColIndex];
  const processed = [data[0]];
  for (let i = 1; i < data.length; i++) {
    const currentValue = data[i][sortColIndex];
    if (currentValue !== lastValue) {
      processed.push([""]); // empty row 1
      processed.push([""]); // empty row 2
    }
    processed.push(data[i]);
    lastValue = currentValue;
  }

  // Tab-delimited string
  const tabText = processed.map(row => row.join("\t")).join("\n");

  // VCF-style success modal
  const html = `
    <div style="
        font-family: 'Google Sans', Arial, sans-serif; 
        padding:20px; 
        background:#edf2fa; 
        color:#222; 
        border-radius:10px; 
        box-shadow:0 2px 5px rgba(0,0,0,0.15);
    ">
      <h2 style="margin-top:0; color:#1e88e5;">✅ Attendance List Ready</h2>
      <p>${processed.length} rows for "<b>${dateFilter}</b>"</p>
      <textarea id="attendanceData" style="width:100%;height:250px;margin-top:8px;">${tabText}</textarea>
      <p style="margin-top:12px;">
        <button onclick="document.getElementById('attendanceData').select(); document.execCommand('copy');" style="
            background:#1e88e5;
            color:white;
            border:none;
            border-radius:6px;
            padding:8px 12px;
            cursor:pointer;
        ">Copy to Clipboard</button>
      </p>
      <p style="margin-top:8px; font-size:12px; color:#555;">Tip: Paste directly into your attendance sheet.</p>
    </div>
  `;
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(460).setHeight(350), "Copy Attendance List");
}


// -----------------------------------------------------------------------------
// EXPORT PARTICIPANT TEST IDS TO EXCEL
//
// This function prompts the admin for a test date (YYYYMMDD),
// then filters participant data for that date and exports selected
// columns (AI–AL) into a downloadable Excel file (.xlsx).
//
// Workflow:
//   1. Prompt admin for test date.
//   2. Grab "Form responses 1" data and filter by test date (col: "Kode Masuk Tes ProTEFL").
//   3. Collect only specific columns (AI–AL).
//   4. Build an inline HTML modal with SheetJS.
//   5. If data exists → show "Download Excel" button.
//      If not → show error and tip.
// -----------------------------------------------------------------------------
function exportParticipantTestIds() {
  const ui = SpreadsheetApp.getUi();

  // Step 1: Ask for test date
  const response = ui.prompt(
    "Enter Test Date (YYYYMMDD)",
    "Provide test date:",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() != ui.Button.OK) return;

  const dateFilter = response.getResponseText().trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Form responses 1");
  if (!sheet) return ui.alert("Target sheet not found.");

  // Step 2: Fetch data + header
  const data = sheet.getDataRange().getValues();
  const header = data.shift();
  const dateColIndex = header.indexOf("Kode Masuk Tes ProTEFL"); // column with YYYYMMDD test date
  const targetCols = ["AI","AJ","AK","AL"].map(letterToColumn_); // only export these columns

  if (dateColIndex === -1) return ui.alert("Test Date column not found.");

  // Step 3: Filter rows by test date
  const filtered = data.filter(row => String(row[dateColIndex]) === dateFilter);

  // Step 4: Prepare export array (include header if data exists)
  const exportData = filtered.length === 0 ? [] : [targetCols.map(i => header[i-1])];
  filtered.forEach(row => exportData.push(targetCols.map(i => row[i-1])));

  // Step 5: Inline HTML modal (with SheetJS)
  // - If no data: show ❌ error
  // - If data exists: show ✅ success and enable Excel download
  let htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8">
        <title>Download Excel</title>
        <script src="https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js"></script>
        <style>
          body { font-family: 'Google Sans', Arial, sans-serif; padding: 20px; color:#222; }
          .container { padding:20px; border-radius:8px; line-height:1.5; box-shadow:0 2px 5px rgba(0,0,0,0.15); }
          .success { background:#edf2fa; color:#222; }
          .error { background:#f8f9fa; color:#222; }
          h2 { margin-top:0; }
          .btn { display:inline-block; padding:8px 12px; background:#1e88e5; color:white; border-radius:6px; text-decoration:none; cursor:pointer; }
          .tip { font-size:12px; color:#555; margin-top:8px; }
        </style>
      </head>
      <body>
        ${
          exportData.length === 0
          ? `<div class="container error">
               <h2 style="color:#d32f2f;">❌ Export Failed</h2>
               <p>No entries found for "<b>${dateFilter}</b>".</p>
               <p class="tip">Tip: Check your filter value and make sure it exists in column AL (format: YYYYMMDD).</p>
               <button onclick="google.script.host.close()" class="btn">Close</button>
             </div>`
          : `<div class="container success">
               <h2 style="color:#1e88e5;">✅ Data Ready!</h2>
               <p>${exportData.length - 1} rows will be exported for "<b>${dateFilter}</b>"</p>
               <button id="downloadBtn" class="btn">Download Excel</button>
             </div>
             <script>
               const exportData = ${JSON.stringify(exportData)};
               document.getElementById("downloadBtn").addEventListener("click", () => {
                 const wb = XLSX.utils.book_new();
                 const ws = XLSX.utils.aoa_to_sheet(exportData);
                 XLSX.utils.book_append_sheet(wb, ws, "ParticipantTestIDs");
                 XLSX.writeFile(wb, \`Participant_TestIDs_${dateFilter}.xlsx\`);
               });
             </script>`
        }
      </body>
    </html>
  `;

  // Step 6: Show modal
  ui.showModalDialog(
    HtmlService.createHtmlOutput(htmlContent).setWidth(460).setHeight(250),
    "Export Participant Test IDs"
  );
}

// -----------------------------------------------------------------------------
// EXPORT SIAKAD SCORE RESULTS TO EXCEL
//
// This function exports test scores from the "06. UPLOADSKOR" sheet
// into an Excel file formatted for Siakad (student academic system).
//
// Workflow:
//   1. Prompt admin for test date (YYYYMMDD).
//   2. Filter rows in column B ("Tanggal Tes") by that date.
//   3. Collect columns C–N (12 fields total: student data + scores).
//   4. Build inline HTML modal using SheetJS (same style as VCF/IDs).
//   5. If no rows match → show ❌ error + tip.
//      If rows found → show ✅ success and provide download button.
//   6. File is named: "DATA MHS UNTUK UPLOAD (dd-mm-yyyy).xlsx"
// -----------------------------------------------------------------------------
function exportSiakadScoreResults() {
  const ui = SpreadsheetApp.getUi();

  // Step 1: Ask for test date
  const response = ui.prompt(
    "Siakad Score Results",
    "Enter test date (YYYYMMDD) to export:",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() != ui.Button.OK) return;

  const dateFilter = response.getResponseText().trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("06. UPLOADSKOR");
  if (!sheet) return ui.alert("Target sheet not found.");

  // Step 2: Read all rows + header
  const data = sheet.getDataRange().getValues();
  const header = data.shift();
  const dateColIndex = 1; // column B = "Tanggal Tes"
  const targetCols = Array.from({length: 12}, (_, i) => i + 2); // C–N (indexes 2–13)

  // Step 3: Filter by test date
  const filtered = data.filter(row => String(row[dateColIndex]) === dateFilter);

  // Step 4: Prepare export array
  const exportData = filtered.length === 0 ? [] : [targetCols.map(i => header[i])];
  filtered.forEach(row => exportData.push(targetCols.map(i => row[i])));

  // Step 5: Build inline HTML modal (VCF-style)
  // - No rows: show ❌ failure message
  // - Rows found: show ✅ success + download button
  let htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8">
        <title>Download Excel</title>
        <script src="https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js"></script>
        <style>
          body { font-family: 'Google Sans', Arial, sans-serif; padding: 20px; color:#222; }
          .container { padding:20px; border-radius:8px; line-height:1.5; box-shadow:0 2px 5px rgba(0,0,0,0.15); }
          .success { background:#edf2fa; color:#222; }
          .error { background:#f8f9fa; color:#222; }
          h2 { margin-top:0; }
          .btn { display:inline-block; padding:8px 12px; background:#1e88e5; color:white; border-radius:6px; text-decoration:none; cursor:pointer; }
          .tip { font-size:12px; color:#555; margin-top:8px; }
        </style>
      </head>
      <body>
        ${
          exportData.length === 0
          ? `<div class="container error">
               <h2 style="color:#d32f2f;">❌ Export Failed</h2>
               <p>No entries found for "<b>${dateFilter}</b>".</p>
               <p class="tip">Tip: Check your filter value and make sure it exists in column B (format: YYYYMMDD).</p>
               <button onclick="google.script.host.close()" class="btn">Close</button>
             </div>`
          : `<div class="container success">
               <h2 style="color:#1e88e5;">✅ Data Ready!</h2>
               <p>${exportData.length - 1} rows will be exported for "<b>${dateFilter}</b>"</p>
               <button id="downloadBtn" class="btn">Download Excel</button>
             </div>
             <script>
               const exportData = ${JSON.stringify(exportData)};
               document.getElementById("downloadBtn").addEventListener("click", () => {
                 const wb = XLSX.utils.book_new();
                 const ws = XLSX.utils.aoa_to_sheet(exportData);
                 XLSX.utils.book_append_sheet(wb, ws, "SiakadScores");

                 // Dynamic filename with today's date (dd-mm-yyyy)
                 const today = new Date();
                 const dd = String(today.getDate()).padStart(2,'0');
                 const mm = String(today.getMonth()+1).padStart(2,'0');
                 const yyyy = today.getFullYear();
                 XLSX.writeFile(wb, \`DATA MHS UNTUK UPLOAD (\${dd}-\${mm}-\${yyyy}).xlsx\`);
               });
             </script>`
        }
      </body>
    </html>
  `;

  // Step 6: Show modal
  ui.showModalDialog(
    HtmlService.createHtmlOutput(htmlContent).setWidth(460).setHeight(250),
    "Export Siakad Score Results"
  );
}

// -----------------------------------------------------------------------------
// COPY CERTIFICATE DATA FUNCTION
//
// Purpose:
//   Exports a tab-delimited certificate list for a given issue date (YYYYMMDD)
//   into clipboard from the "05. DATASERTIFIKAT" sheet.
//
// Steps:
//   1. Prompt the user for the certificate date (format: YYYYMMDD).
//   2. Access the certificate sheet and pull all data.
//   3. Filter rows by Column B ("Date") matching the user input.
//   4. If no rows found → show an error modal and exit.
//   5. Extract only C–P columns.
//   6. Convert the filtered rows into a tab-delimited string (with header).
//   7. Show modal with a textarea containing the result and a "Copy to Clipboard" button.
// -----------------------------------------------------------------------------
function copyCertificateData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Enter Certificate Date (YYYYMMDD)",
    "Provide issue date:",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const dateFilter = response.getResponseText().trim();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("05. DATASERTIFIKAT");
  if (!sheet) return ui.alert("Target sheet not found.");

  let data = sheet.getDataRange().getValues();
  data.shift(); // remove header row entirely
  const dateColIndex = 1; // Column B (zero-based)

  // Filter by certificate date
  data = data.filter(row => String(row[dateColIndex]) === dateFilter);

  if (data.length === 0) {
    const html = `
      <div style="
          font-family: 'Google Sans', Arial, sans-serif; 
          padding:20px; 
          background:#fefefe; 
          color:#222; 
          border-radius:10px; 
          box-shadow:0 2px 5px rgba(0,0,0,0.15);
      ">
        <h2 style="margin-top:0; color:#c62828;">⚠️ No Entries Found</h2>
        <p>No certificate records found for "<b>${dateFilter}</b>".</p>
        <p>Check column <b>B</b> ('Issue Date') for existing values (format: YYYYMMDD).</p>
        <button onclick="google.script.host.close()" style="
            background:#1e88e5;
            color:white;
            border:none;
            border-radius:6px;
            padding:8px 12px;
            cursor:pointer;
        ">Close</button>
      </div>
    `;
    ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(450).setHeight(220), "Copy Certificate Error");
    return;
  }

  // Slice only columns C–P (indexes 2–15)
  data = data.map(row => row.slice(2, 16));

  // Tab-delimited string (no headers)
  const tabText = data.map(row => row.join("\t")).join("\n");

  // Success modal
  const html = `
    <div style="
        font-family: 'Google Sans', Arial, sans-serif; 
        padding:20px; 
        background:#edf2fa; 
        color:#222; 
        border-radius:10px; 
        box-shadow:0 2px 5px rgba(0,0,0,0.15);
    ">
      <h2 style="margin-top:0; color:#1e88e5;">✅ Certificate Data Ready</h2>
      <p>${data.length} rows for "<b>${dateFilter}</b>"</p>
      <textarea id="certificateData" style="width:100%;height:250px;margin-top:8px;">${tabText}</textarea>
      <p style="margin-top:12px;">
        <button onclick="document.getElementById('certificateData').select(); document.execCommand('copy');" style="
            background:#1e88e5;
            color:white;
            border:none;
            border-radius:6px;
            padding:8px 12px;
            cursor:pointer;
        ">Copy to Clipboard</button>
      </p>
      <p style="margin-top:8px; font-size:12px; color:#555;">Tip: Paste directly into your certificate sheet.</p>
    </div>
  `;
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(460).setHeight(350), "Copy Certificate Data");
}


// ============================================================================
// File: sideBars.gs
//
// SIDEBARS (sideBars.gs)
// Handles creation of interactive sidebars for ProTEFL MDMA.
//
// Features:
//   - Default Sidebar: home base with key tasks overview.
//   - Reschedule Sidebar: step-by-step participant rescheduling.
//   - Verify Student ID Sidebar: stepwise verification to prevent mismatches.
//   - Verify Payment Sidebar: payment verification (online & manual).
//   - Verify Attendance Sidebar: attendance & score verification guide.
//   - Grouping & Contacts Sidebar: manage auto groupings and VCF creation.
//
// Notes:
//   - Uses Google Sans, cards, and collapsible arrow icons for uniform style.
//   - Each sidebar uses reusable createCardHTML() function for consistency.
//   - Safe to re-run; only displays the latest sidebar UI.
// ============================================================================  

// ============================================================================
// Function: showDefaultSidebar
// Description: Displays the default ProTEFL MDMA sidebar with task overview.
// ============================================================================
function showDefaultSidebar() {
  const html = `
  <!DOCTYPE html>
  <html>
    <head>
      <meta charset="UTF-8">
      <title>ProTEFL MDMA</title>
      <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
      <link href="https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;700&display=swap" rel="stylesheet">

      <!-- ======================
           Reusable Sidebar Styles
           ====================== -->
      <style>
        body {
          font-family: 'Google Sans', Arial, sans-serif;
          margin: 0;
          padding: 20px;
          background: #edf2fa;
          color: #222;
          line-height: 1.6;
        }

        h2 { margin-top:0; color:#1a1a1a; }
        h3 { margin-top:16px; color:#333; }

        /* Card styling */
        .card {
          background: #d3e3fd;
          border-radius: 10px;
          box-shadow: 0 2px 5px rgba(0,0,0,0.15);
          padding: 16px 18px;
          margin-bottom: 16px;
          transition: transform 0.1s ease, box-shadow 0.4s ease;
        }

        .card:hover {
          transform: translateY(-2px);
          box-shadow: 0 6px 10px rgba(0,0,0,0.2);
        }

        /* Header styling */
        .card-header {
          font-weight: bold;
          cursor: pointer;
          display: flex;
          align-items: center;
          color: #1a1a1a;
          margin-bottom: 8px;
        }

        .card-header .arrow-icon {
          font-size: 26px;
          margin-right: 10px;
          transition: transform 0.4s ease;
          color: #3a3a3a;
        }

        .card-header .section-icon {
          font-size: 20px;
          margin-right: 8px;
          color: #1e88e5;
        }

        .card-content {
          margin-top: 12px;
          color: #333;
        }

        ul { margin: 0; padding-left: 20px; }
        li { margin-bottom: 6px; }

        .footer-note { color:#555; font-size:12px; margin-top:20px; }
        a { color:#1e88e5; text-decoration:none; }
        a:hover { text-decoration:underline; }
      </style>
    </head>

    <body>
      <h2>Welcome to ProTEFL MDMA</h2>
      <p><i>(Monthly Data Management Admin)</i></p>
      <p>ProTEFL on Speed ⚡🥴😵</p>

      <!-- Cards -->
      ${createCardHTML('assignment','Registration',['Google Forms Entry','Manual Entry (menu planned)'])}
      ${createCardHTML('settings','Data Management',[
        'Participant(s) Rescheduling (Before Test)',
        'Student ID Verification',
        'Manual Test Count Checking (menu planned)',
        'Automatic & Override Option of Test Group Plotting (menu planned)',
        'Contact Creation (VCF) (menu planned)',
        'Autogenerated Attendance & Test ID Lists (menu planned)'
      ])}
      ${createCardHTML('assessment','Scoring',[
        'Attendance Verification & Reschedule Flagging (After Test)',
        'Score Checking',
        'Reschedule Offering (same as in Data Management)',
        'Autogenerated Score Report format',
        'Autogenerated Certificate Data Format',
        'Autogenerated SISTER Upload Format (obsolete)'
      ])}

      <h3>About this Default View:</h3>
      <p>
        This is the <b>home base</b>. Only essential columns are shown 
        (IDs, names, key status checks, and admin flags).  
        It is the clean slate you (dear admin[s]) land on every time you open the workbook.  
        From here, jump into other custom views if you need to 
        focus on tasks like rescheduling, ID verification, attendance, or score verification.
      </p>

      <p>
        Open other views via <b>ProTEFL Utility &gt; Custom View</b> in the menu bar.
      </p>

      <p class="footer-note">
        Reminder: speed is great, but accurate data keeps the complaints away. 
        PS. The title is obviously inspired by Andy Field way of naming his Statistics books. 
        I mean, "Discovering statistics using IBM SPSS statistics: and BLEEP and BLEEP and rock 'n' roll" ...what a BLEEP legend.
      </p>

      <script>
        function toggleCollapse(header) {
          const content = header.nextElementSibling;
          const arrow = header.querySelector('.arrow-icon');
          if(content.style.display === 'none' || content.style.display === '') {
            content.style.display = 'block';
            arrow.style.transform = 'rotate(180deg)';
          } else {
            content.style.display = 'none';
            arrow.style.transform = 'rotate(0deg)';
          }
        }

        // Collapse all sections on load
        document.querySelectorAll('.card-content').forEach(c => c.style.display='none');
      </script>
    </body>
  </html>
  `;

  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(html).setTitle("ProTEFL MDMA"));
}

// ----------------------------------------------------------------------------
// Helper: create reusable card HTML
// ----------------------------------------------------------------------------
function createCardHTML(iconName, title, items) {
  const listItems = items.map(item => `<li>${item}</li>`).join('');
  return `
    <div class="card">
      <div class="card-header" onclick="toggleCollapse(this)">
        <span class="arrow-icon material-icons">expand_more</span>
        <span class="section-icon material-icons">${iconName}</span>
        ${title}
      </div>
      <div class="card-content">
        <ul>${listItems}</ul>
      </div>
    </div>
  `;
}

// ============================================================================
// Function: showRescheduleSidebar
// Description: Step-by-step guide for rescheduling participants.
// ============================================================================
function showRescheduleSidebar() {
  const html = `
  <!DOCTYPE html>
  <html>
    <head>
      <meta charset="UTF-8">
      <title>Reschedule Participants</title>
      <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
      <link href="https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;700&display=swap" rel="stylesheet">
      <style>
        body {
          font-family: 'Google Sans', Arial, sans-serif;
          margin: 0;
          padding: 20px;
          background: #edf2fa;
          color: #222;
          line-height: 1.6;
        }

        h2 { margin-top:0; color:#1a1a1a; }
        h3 { margin-top:16px; color:#333; }

        .card {
          background: #d3e3fd;
          border-radius: 10px;
          box-shadow: 0 2px 5px rgba(0,0,0,0.15);
          padding: 16px 18px;
          margin-bottom: 16px;
          transition: transform 0.1s ease, box-shadow 0.4s ease;
        }

        .card:hover {
          transform: translateY(-2px);
          box-shadow: 0 6px 10px rgba(0,0,0,0.2);
        }

        .card-header {
          font-weight: bold;
          cursor: pointer;
          display: flex;
          align-items: center;
          color: #1a1a1a;
          margin-bottom: 8px;
        }

        .card-header .arrow-icon {
          font-size: 26px;
          margin-right: 10px;
          transition: transform 0.4s ease;
          color: #3a3a3a;
        }

        .card-header .section-icon {
          font-size: 20px;
          margin-right: 8px;
        }

        .card-content { margin-top:12px; color:#333; }
        ol, ul { padding-left: 20px; }
        li { margin-bottom: 6px; }
      </style>
    </head>
    <body>

      ${createCardHTML('📋','Reschedule Participants Guide',[
        "Here’s a go-to workflow for rescheduling participants:",
        "Locate the participant’s <b>Name</b> in column <b>E</b>.",
        "Verify their <b>Original Schedule</b> in column <b>R</b>. It is crucial if they registered multiple times. In that case, be careful. Make sure you reschedule the correct entry.",
        "In column <b>V</b>, set the dropdown to <b>Yes</b> to flag for reschedule. This will revoke their original schedule. They won't have a schedule now. Column AL will now be empty.",
        "To assign them new schedule, search for the new schedule date in <b>00. MASTER-DATA</b> in accordance to participant's choosing.",
        "Copy the suitable schedule from <b>00. MASTER-DATA</b> into <b>Form responses 1</b> in column <b>W</b>.",
        "Mark <b>Confirmed</b> in column <b>AG</b> to lock it in.",
        "Copy the WhatsApp message from column <b>AH</b> and send it to the participant. 🚀",
        "Tip: Accuracy beats speed here — double-check before hitting send! With accoubtability, you have avoided complaint(s) induced headache and hypertension."
      ])}

      <script>
        function toggleCollapse(header) {
          const content = header.nextElementSibling;
          const arrow = header.querySelector('.arrow-icon');
          if(content.style.display === 'none' || content.style.display === '') {
            content.style.display = 'block';
            arrow.style.transform = 'rotate(180deg)';
          } else {
            content.style.display = 'none';
            arrow.style.transform = 'rotate(0deg)';
          }
        }
        document.querySelectorAll('.card-content').forEach(c => c.style.display='block'); // keep single card expanded
      </script>

    </body>
  </html>
  `;

  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(html).setTitle("Reschedule Participants"));
}

// ============================================================================
// Function: showVerifyStudentIDSidebar
// Description: Guide for verifying student IDs to prevent mismatches.
// ============================================================================
function showVerifyStudentIDSidebar() {
  const html = `
  <!DOCTYPE html>
  <html>
    <head>
      <meta charset="UTF-8">
      <title>Verify Student ID</title>
      <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
      <link href="https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;700&display=swap" rel="stylesheet">
      <style>
        body {
          font-family: 'Google Sans', Arial, sans-serif;
          margin: 0;
          padding: 20px;
          background: #edf2fa;
          color: #222;
          line-height: 1.6;
        }

        h2 { margin-top:0; color:#1a1a1a; }
        h3 { margin-top:16px; color:#333; }

        .card {
          background: #d3e3fd;
          border-radius: 10px;
          box-shadow: 0 2px 5px rgba(0,0,0,0.15);
          padding: 16px 18px;
          margin-bottom: 16px;
          transition: transform 0.1s ease, box-shadow 0.4s ease;
        }

        .card:hover {
          transform: translateY(-2px);
          box-shadow: 0 6px 10px rgba(0,0,0,0.2);
        }

        .card-header {
          font-weight: bold;
          cursor: pointer;
          display: flex;
          align-items: center;
          color: #1a1a1a;
          margin-bottom: 8px;
        }

        .card-header .arrow-icon {
          font-size: 26px;
          margin-right: 10px;
          transition: transform 0.4s ease;
          color: #3a3a3a;
        }

        .card-header .section-icon {
          font-size: 20px;
          margin-right: 8px;
        }

        .card-content { margin-top:12px; color:#333; }
        ol, ul { padding-left: 20px; }
        li { margin-bottom: 6px; }
      </style>
    </head>
    <body>

      ${createCardHTML('🆔','Verify Student ID Guide',[
        "Student ID verification is critical — mismatched IDs mean scores won't appear on SIAKAD. This is achieved with the assumption that entries in <b>DATABASEMAHASISWA</b> has the correct student data.",
        "Step-by-step check:",
        "Check column <b>BC</b> (Status):",
        "<b>COCOK</b>: ✅ Everything matches — move on to the next participant.",
        "<b>CEK NAMA</b>: Minor capitalization mismatch. No fix needed here; we already use corrected proper names. Reference <b>06. UPLOADSKOR</b> for tidy names (for every thanks Windi gains 1 rupiah in his pocket [of course not!]).",
        "<b>SALAH NIM</b>: Name in column <b>E</b> or <b>BA (duplicates of E)</b> doesn’t match the database (<b> shown in BB</b>). Ask the participant for their ID card and update NIM in <b>E</b> ONLY. Data shown elsewhere are all duplicates of E.",
        "<b>#N/A</b>: No match found. Investigate and resolve manually. Ask the students for their KTM, write the correct NIM. When issues persist, it means we do not have their data in DATABASEMAHASISWA. Please update it manually based on the data on their KTM. Usually happens for students registering as INTAKE students (course begining on February).",
        "Pro tip: Careful checking now saves a flood of complaints later. 👍"
      ])}

      <script>
        function toggleCollapse(header) {
          const content = header.nextElementSibling;
          const arrow = header.querySelector('.arrow-icon');
          if(content.style.display === 'none' || content.style.display === '') {
            content.style.display = 'block';
            arrow.style.transform = 'rotate(180deg)';
          } else {
            content.style.display = 'none';
            arrow.style.transform = 'rotate(0deg)';
          }
        }
        document.querySelectorAll('.card-content').forEach(c => c.style.display='block');
      </script>

    </body>
  </html>
  `;

  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(html).setTitle("Verify Student ID"));
}

// ============================================================================
// Function: showVerifyPaymentSidebar
// Description: Quick guide for verifying payments (online & manual).
// ============================================================================
function showVerifyPaymentSidebar() {
  const html = `
  <!DOCTYPE html>
  <html>
    <head>
      <meta charset="UTF-8">
      <title>Verify Participants Payment</title>
      <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
      <link href="https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;700&display=swap" rel="stylesheet">
      <style>
        body {
          font-family: 'Google Sans', Arial, sans-serif;
          margin: 0;
          padding: 20px;
          background: #edf2fa;
          color: #222;
          line-height: 1.6;
        }
        h2 { margin-top:0; color:#1a1a1a; }
        h3 { margin-top:16px; color:#333; }

        .card {
          background: #d3e3fd;
          border-radius: 10px;
          box-shadow: 0 2px 5px rgba(0,0,0,0.15);
          padding: 16px 18px;
          margin-bottom: 16px;
          transition: transform 0.1s ease, box-shadow 0.4s ease;
        }
        .card:hover {
          transform: translateY(-2px);
          box-shadow: 0 6px 10px rgba(0,0,0,0.2);
        }
        .card-header {
          font-weight: bold;
          cursor: pointer;
          display: flex;
          align-items: center;
          color: #1a1a1a;
          margin-bottom: 8px;
        }
        .card-header .arrow-icon {
          font-size: 26px;
          margin-right: 10px;
          transition: transform 0.4s ease;
          color: #3a3a3a;
        }
        .card-header .section-icon {
          font-size: 20px;
          margin-right: 8px;
        }
        .card-content { margin-top:12px; color:#333; }
        ol, ul { padding-left: 20px; }
        li { margin-bottom: 6px; }
      </style>
    </head>
    <body>

      <p><i>"This view is for verifying test taker payments — this keeps ULB overlord(s) happy! 🤔😢"</i></p>

      <section>
        <h2>Overview</h2>
        <p>
          For participants taking the test in our lab on a specific schedule, we don’t want freeloaders or non-paying registrants sneaking into SEB from home.  
          That’s why receipts include a different <b>Nomor Ujian</b> (test ID).  
          Compare how <b>NIM</b> (Column D) and <b>Test ID</b> (Column AI) differ once you write <code>_OFFGRID</code> in Column BI.  
          Only those who actually paid can log in on the test date.  
        </p>
        <p>
          To start working, follow the guidelines below.  
          The online payment section is mostly legacy (Glacier + virtual accounts handle it now).  
          Day-to-day, we deal with <b>pembayaran tes luring</b> in Lab Bahasa ULB.
        </p>
      </section>

      <section>
        <h2>⚠️ Attention</h2>
        <p>This workflow works under a few assumptions:</p>
        <ul>
          <li>Participants registered via Google Form (so their names exist in this sheet).</li>
          <li>You, the admin, must locate the correct row — <b>watch out for duplicates!</b></li>
          <li>Some participants register more than once with slightly different data or test dates.</li>
          <li>If you find multiple entries with the same name, ask the participant which one is correct.</li>
          <li>Update the row for the correct <b>date</b> (see Column BJ as confirmed by the participant).</li>
        </ul>
      </section>

      ${createCardHTML('💵','Manual / Cash Payment (LURING / On-demand)',[
        "Confirm the participant received their proof of payment (kuitansi / receipt).",
        "Search their name in column <b>AS</b>.",
        "If you find more than one entry of their exact name, make sure you are looking at the correct test date in column <b>BJ</b>.",
        "Copy the <b>Nomor Ujian</b> from their receipt into column <b>G</b>. Ignore placeholders like D4, S1, S2, S3 — overwrite them. (I’m too lazy to refactor the whole Google Form structure after all these formulas and w/hackjobs).",
        "Important: Write <b>_OFFGRID</b> in column <b>BI</b>. This forces the workbook to use the receipt’s <b>Nomor Ujian</b> instead of the default NIM. Why? So non-payers can’t log in into SEB with just their NIM in the test date.",
        "Pro tip: Always double-check attachments and make sure you typed the correct Nomor Ujian — saves you complaints later ⚡."
      ])}

      ${createCardHTML('💰','Online Payment',[
        "Check the <b>Bukti Bayar</b> attachment in column <b>AU</b>.",
        "Verify authenticity — is it real, not recycled, and matches the participant?",
        "If ✅, mark <b>LUNAS</b> in column <b>AX</b>.",
        "If not, select the appropriate status depending on the issue.",
        "Done. Move on to the next participant."
      ])}

      <script>
        function toggleCollapse(header) {
          const content = header.nextElementSibling;
          const arrow = header.querySelector('.arrow-icon');
          if(content.style.display === 'none' || content.style.display === '') {
            content.style.display = 'block';
            arrow.style.transform = 'rotate(180deg)';
          } else {
            content.style.display = 'none';
            arrow.style.transform = 'rotate(0deg)';
          }
        }
        document.querySelectorAll('.card-content').forEach(c => c.style.display='block');
      </script>

    </body>
  </html>
  `;

  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(html).setTitle("Verify Payment"));
}


// ============================================================================
// Function: showGroupingContactsSidebar
// Description: Manage automatic groupings and VCF contact creation.
// ============================================================================
function showGroupingContactsSidebar() {
  const html = `
  <!DOCTYPE html>
  <html>
    <head>
      <meta charset="UTF-8">
      <title>Grouping & Contacts</title>
      <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
      <link href="https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;700&display=swap" rel="stylesheet">
      <style>
        body {
          font-family: 'Google Sans', Arial, sans-serif;
          margin: 0;
          padding: 20px;
          background: #edf2fa;
          color: #222;
          line-height: 1.6;
        }
        h2 { margin-top:0; color:#1a1a1a; }
        .card {
          background: #d3e3fd;
          border-radius: 10px;
          box-shadow: 0 2px 5px rgba(0,0,0,0.15);
          padding: 16px 18px;
          margin-bottom: 16px;
          transition: transform 0.1s ease, box-shadow 0.4s ease;
        }
        .card:hover {
          transform: translateY(-2px);
          box-shadow: 0 6px 10px rgba(0,0,0,0.2);
        }
        .card-header {
          font-weight: bold;
          cursor: pointer;
          display: flex;
          align-items: center;
          color: #1a1a1a;
          margin-bottom: 8px;
        }
        .card-header .arrow-icon {
          font-size: 26px;
          margin-right: 10px;
          transition: transform 0.4s ease;
          color: #3a3a3a;
        }
        .card-header .section-icon {
          font-size: 20px;
          margin-right: 8px;
          color: #1e88e5;
        }
        .card-content { margin-top:12px; color:#333; }
        ul { margin:0; padding-left: 20px; }
        li { margin-bottom: 6px; }
        .footer-note { color:#555; font-size:12px; margin-top:16px; }
      </style>
    </head>
    <body>

      <h2>Grouping & Contacts</h2>
      <p><i>Manage automatic groupings and contact creation</i></p>

      ${createCardHTML('group_work','Grouping',[
        "Filter <b>AL</b> to select a specific date.",
        "Automatic group assignments appear in <b>AO</b>.",
        "Override group manually in <b>AP</b> if needed.",
        "Group naming logic:",
        "Extract 3 digits from date in <b>AL</b> → year/month.",
        "One character denotes test mode: \"D\" = online, \"L\" = offline.",
        "Three-character alphanumeric group code based on session/sequence.",
        "Suffix \"T_\" or \"S_\" indicates TKBI/SISTER vs regular participant."
      ])}

      ${createCardHTML('contacts','Contact Creation (VCF)',[
        "VCF entries are in <b>BG</b>, starting with 8 alphanumeric digits (e.g., 25SLA12S).",
        "Use these codes to import participants into WhatsApp groups reliably.",
        "To download a VCF:",
        "Filter by date in <b>AL</b>.",
        "Use <b>ProTEFL Utility → Download VCF by Tanggal Tes</b> in the menu bar."
      ])}

      <p class="footer-note">Ensure accuracy when editing groups or downloading VCF — speed is great, but mistakes cost time!</p>

      <script>
        function toggleCollapse(header) {
          const content = header.nextElementSibling;
          const arrow = header.querySelector('.arrow-icon');
          if(content.style.display === 'none' || content.style.display === '') {
            content.style.display = 'block';
            arrow.style.transform = 'rotate(180deg)';
          } else {
            content.style.display = 'none';
            arrow.style.transform = 'rotate(0deg)';
          }
        }
        document.querySelectorAll('.card-content').forEach(c => c.style.display='none');
      </script>

    </body>
  </html>
  `;

  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(html).setTitle("Grouping & Contacts"));
}


// ============================================================================
// Function: showVerifyAttendanceSidebar
// Description: Guide for verifying attendance and scores.
// ============================================================================
function showVerifyAttendanceSidebar() {
  const html = `
  <!DOCTYPE html>
  <html>
    <head>
      <meta charset="UTF-8">
      <title>Verify Attendance</title>
      <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
      <link href="https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;700&display=swap" rel="stylesheet">
      <style>
        body {
          font-family: 'Google Sans', Arial, sans-serif;
          margin: 0;
          padding: 20px;
          background: #edf2fa;
          color: #222;
          line-height: 1.6;
        }
        h2 { margin-top:0; color:#1a1a1a; }
        h3 { margin-top:16px; color:#333; }

        .card {
          background: #d3e3fd;
          border-radius: 10px;
          box-shadow: 0 2px 5px rgba(0,0,0,0.15);
          padding: 16px 18px;
          margin-bottom: 16px;
          transition: transform 0.1s ease, box-shadow 0.4s ease;
        }
        .card:hover {
          transform: translateY(-2px);
          box-shadow: 0 6px 10px rgba(0,0,0,0.2);
        }
        .card-header {
          font-weight: bold;
          cursor: pointer;
          display: flex;
          align-items: center;
          color: #1a1a1a;
          margin-bottom: 8px;
        }
        .card-header .arrow-icon {
          font-size: 26px;
          margin-right: 10px;
          transition: transform 0.4s ease;
          color: #3a3a3a;
        }
        .card-header .section-icon {
          font-size: 20px;
          margin-right: 8px;
        }
        .card-content { margin-top:12px; color:#333; }
        ol, ul { padding-left: 20px; }
        li { margin-bottom: 6px; }
      </style>
    </head>
    <body>

      ${createCardHTML('📊','Verify Attendance & Score',[
        "<i>Use this view to verify attendance and score checking. This is by far the most time-consuming part (god I wish I got paid extra for this).</i>",
        "Step 0: For sanity’s sake",
        "Enable filter by date: look at <b>BJ</b> and select a single date. Trust me, your sanity will thank you.",
        "Step 1: Prepare",
        "You need to check attendance report from proctors (in another sheet, sadly). Use split window view for best productivity—one side the attendance sheet, another side this sheet.",
        "Step 2: Import Scores",
        "Copy the scores into this workbook in <b>SINICOPYHASILSKOR</b> and do the necessary formatting. Make sure column <b>A</b> on <b>SINICOPYHASILSKOR</b> matches <b>BQ</b> in <b>Form responses 1</b> (this sheet). Then, in <b>SINICOPYHASILSKOR</b> copy test ID in P, write the appropriate kode masuk in Q, and make sure R has the formula '=(Q2 & \"-\" & P2)' so on; and A has the formula '=R2' and so on, drag them down. The scores will then appear across <b>BU to BY</b> in Form responses 1.",
        "Disclaimer: this works under the assumption that the data you copy into SINICOPYHASILSKOR is pristine and no tes IDs are misplaced, replaced, moved from their original cells. If there are errors, that's on you. Congrats you just messed up an entire results of that day tests and maybe others. Now cry and curl up in the corner!",
        "Step 3: Check for missing scores",
        "If no score appears, there are three possibilities:",
        "<b>Did not attend:</b> mark reschedule on <b>V</b> to Yes, write placeholder to <b>W</b>. We will ask them later using template message link in <b>AE</b>. This revokes their registration on this date; no data in <b>SINICOPYHASILSKOR</b> will link to any test ID.",
        "<b>Used Akun Cadangan:</b> copy akun cadangan to <b>G</b>, write <b>_OFFGRID</b> to <b>BI</b>, and check if scores appear on <b>BU-BX</b>.",
        "<b>NIM mismatch:</b> mismatch between <b>D</b> and whatever test ID they used in <b>SINICOPYHASILSKOR</b>. Resolve by checking their used ID, refer to proctor notes, and do step two above. Is their NIM not matching? Check <b>BC</b> for <b>CEK NAMA</b>. Still no score? Confirm <b>D</b> vs attendance sheet ID. Or call Windi while he’s still around. Typing this is already exhausting.",
        "Step 4: When all else fails",
        "If nothing works and there is no attendance note, you are <b>COOKED 💀</b>.<br>Or they didn’t attend and the proctor forgot to mark it—prepare pitchfork, torch, gasoline, and proceed to set the proctor ablaze! It’s their <b>FAULT!</b>",
        "Reminder: patience, coffee, and a deep breath are your best allies. Oh, what's that Pesawat Jet F-22 Raptor in CG? Try typing funny negative number in it and watch BX burns."
      ])}

      <script>
        function toggleCollapse(header) {
          const content = header.nextElementSibling;
          const arrow = header.querySelector('.arrow-icon');
          if(content.style.display === 'none' || content.style.display === '') {
            content.style.display = 'block';
            arrow.style.transform = 'rotate(180deg)';
          } else {
            content.style.display = 'none';
            arrow.style.transform = 'rotate(0deg)';
          }
        }
        document.querySelectorAll('.card-content').forEach(c => c.style.display='block');
      </script>

    </body>
  </html>
  `;

  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(html).setTitle("Verify Attendance"));
}

function showPPBSidebar() {
  const html = `
  <!DOCTYPE html>
  <html>
    <head>
      <base target="_top">
      <style>
        body { 
          font-family: 'Google Sans', Arial, sans-serif; 
          padding: 20px; 
          background: #f4f6f8; 
          color: #222; 
          line-height: 1.6; 
        }
        h2 { color:#1a1a1a; margin-top:0; }
        p { margin-top:12px; }
        .card {
          background: #dbe9f9; 
          border-radius: 10px; 
          padding: 16px; 
          box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }
      </style>
    </head>
    <body>
      <h2>PPB View Active</h2>
      <div class="card">
        <p>Here you go, it's all yours. Enjoy. Bone apple tea, remember not to overwrite key data accidentally without purpose mmkay?</p>
      </div>
    </body>
  </html>
  `;
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(html).setTitle("PPB View"));
}


// ============================================================================
// File: setupSheets.gs
// 
// SHEET INITIALIZATION (setupSheets.gs)
// Automatically creates core sheets and populates them with headers/templates.
//
// Features:
//   - Defines a central config (SHEET_INITIALIZATIONS) with sheet names + cells.
//   - Defines an extended header set for "Form responses 1" (many helper cols).
//
// Steps (initializeSheets):
//   1. Loop through SHEET_INITIALIZATIONS.
//        - If sheet doesn’t exist → create it.
//        - Write any configured cell values (headers, labels, templates).
//   2. If "Form responses 1" exists → apply the extended FORM_RESPONSES_1_HEADER.
//
// Notes:
//   - Safe to re-run; it will only create missing sheets and overwrite listed cells.
//   - Best run inside main() after workbook creation or reset.
//   - Extend by adding more objects to SHEET_INITIALIZATIONS or entries in
//     FORM_RESPONSES_1_HEADER.
//
// ============================================================================
const SHEET_INITIALIZATIONS = [
// The headers are defined below.
    {
      sheetName: '00. MASTER-DATA',
      cells: {
        'A1': 'Pilihan Tanggal dan Sesi Tes',
        'A19': 'Pilihan Tanggal Reschedule',
        'B19': 'Pilihan Moda Reschedule',
        'C19': 'Availability',
        'A31': 'Bulan dan Tahun',
      }
    },
    {
      sheetName: '01. STATISTIK',
      cells: {
        'A1': 'NAMA GRUP WA',
        'B1': 'Jumlah Peserta',
        'D1': 'TIMESTAMP',
        'C2': 'POTENSI DUPLIKAT',
        'E1': 'TEST ID',
        'F1': 'NAMA',
        'G1': 'TGL TES',
        'H1': 'RESC STATUS',
        'I1': 'RESC ACTION',
        'J1': 'SKOR',
        'N2': 'REKAP STATUS RESCHEDULE',
        'O1': 'TIMESTAMP',
        'P1': 'TEST ID',
        'Q1': 'NAMA',
        'R1': 'RESC STATUS',
        'S1': 'RESC ACTION',
        'T1': 'TGL TES',
        'U1': 'NOTES',
        'V1': 'LINK KORESPONDENSI'
      }
    },
    {
      sheetName: '02. CEKTESTHISTORY',
      cells: {
        'A1': 'NIM',
        'B1': 'Nama',
        'C1': 'Skor Tertinggi',
        'D1': 'Tanggal Terbaru',
        'E1': 'Skor ALL',
        'F1': 'Tanggal ALL',
      }
    },
    {
      sheetName: '03. KIRIM DATA KE PAK BIN H-1',
      cells: {
        'A1': 'DATA TES PROTEFL H-1',
      }
    },
    {
      sheetName: '04. BUAT PRESENSI DAN GRUP WA H-1',
      cells: {
        'A1': 'No.',
        'B1': 'Email',
        'C1': 'NIM/NIK',
        'D1': 'Nama',
        'E1': 'WhatsApp',
        'F1': 'Kode Tanggal',
        'G1': 'Grup Tes',
        'H1': 'Kode Registrasi',
        'I1': 'PIC',
        'J1': 'WA',
        'K1': 'FOTO',
        'L1': 'LINK INVITE',
        'M1': 'HADIR',
        'N1': 'KETERANGAN',
        'O1': 'RESCHEDULE STATUS',
        'Q1': 'TES LEBIH DARI SATU KALI',
      }
    },
    {
      sheetName: '05. DATASERTIFIKAT',
      cells: {
        'B1': 'REF TGL TES',
        'C1': 'EMAIL',
        'D1': 'NIM/NIK',
        'E1': 'NIDN',
        'F1': 'NAMA',
        'G1': 'TEMPAT LAHIR',
        'H1': 'TANGGAL LAHIR',
        'I1': 'TGLTES',
        'J1': 'LISTENING',
        'K1': 'GRAMMAR',
        'L1': 'READING',
        'M1': 'SKOR TOTAL',
        'N1': 'SKOR PBT',
        'O1': 'SKOR IELTS',
        'P1': 'TKBI',
        'Q1': 'NO SERTIFIKAT',
        'R1': 'TTD'
      }
    },
    {
      sheetName: '06. UPLOADSKOR',
      cells: {
        'B1': 'tanggal tes',
        'C1': 'nim',
        'D1': 'nama',
        'E1': 'status',
        'F1': 'skor',
        'G1': 'tanggal tes',
        'H1': 'jenjang',
        'I1': 'Fakultas',
        'J1': 'Prodi',
        'K1': 'MIN SKOR',
        'L1': 'MIN MEN',
        'M1': 'TAMBAHAN SKOR JUR INGG'
      }
    },
    {
      sheetName: '07. UPLOADSISTER',
      cells: {
        'A1': 'REF KODE SKOR',
        'B1': 'nuptk',
        'C1': 'nidn',
        'D1': 'nm_sdm',
        'E1': 'thn',
        'F1': 'skor',
        'G1': 'tgl_tes'
      }
    },
    {
      sheetName: '08. DATAKUITANSI',
      cells: {
        'A1': 'Pay Date',
        'B1': 'Month',
        'C1': 'No Receipt',
        'D1': 'Nama',
        'E1': 'WhatsApp',
        'F1': 'NIM',
        'G1': 'Receipt',
        'H1': 'Nominal',
        'I1': 'Terbilang',
        'J1': 'Keperluan',
        'K1': 'Receipt Date',
        'L1': 'Admin',
        'M1': 'email'
      }
    },
    {
      sheetName: 'SINICOPYHASILSKOR',
      cells: {} // Leave blank/empty, just creates sheet
    },
    {
      sheetName: 'db_export',
      cells: {
        'A1': 'valid',
        'B1': 'student_id',
        'C1': 'name',
        'D1': 'faculty',
        'E1': 'department',
        'F1': 'level',
        'G1': 'test_date',
        'H1': 'session',
        'I1': 'mode',
        'J1': 'converted_list',
        'K1': 'converted_str',
        'L1': 'converted_rdg',
        'M1': 'itp_pred',
        'N1': 'status',
        'O1': 'itp_adjust',
        'P1': 'itp_raw'
      }
    },
    {
      sheetName: 'scratch pad',
      cells: {
        'A1': 'Use this sheet to experiment or as an excel/sheet enabled clipboard',
      }
    }
  ];

// Special config for Form responses 1 since it exists and has a lot of columns
const FORM_RESPONSES_1_HEADER = [
    // -------------------- RESCHEDULE / SCHEDULE --------------------
    [ ['V1', 'TABLE SCHEDULE | Reschedule'],
    ['W1',  'Rescheduled Date'],
    ['X1',  'Schedule Log'],
    ['Y1',  'Reschedule Count'],
    ['Z1',  '-info Original Schedule'],
    ['AA1', '-helper Pilihan Tanggal Tes'],
    ['AB1', '-helper Bulan dan Tahun'],
    ['AC1', '-helper Jam Daring'],
    ['AD1', '-helper Jam Luring'],
    ['AE1', 'Konfirmasi WA Reschedule Bulan Lalu'],
    ['AF1', 'Notes'],
    ['AG1', 'Status Konfirmasi'],
    ['AH1', 'Confirmation Message'],

    // -------------------- TEST USER --------------------
    ['AI1', 'TABLE TEST USER | Username Tes ProTEFL'],
    ['AJ1', 'Nama Peserta (Proper Noun)'],
    ['AK1', 'Password Tes ProTEFL'],
    ['AL1', 'Kode Masuk Tes ProTEFL'],

    // -------------------- TEST GROUP --------------------
    ['AM1', 'TABLE TEST GROUP | Kode Sesi Bulan'],
    ['AN1', '-helper Kode Sesi Moda'],
    ['AO1', 'Kode Sesi Grup Pengawasan'],
    ['AP1', 'Override Grup Pengawasan'],
    ['AQ1', '-helper Prefix Jenis Tes'],
    ['AR1', '-helper DRAG Urutan Grup'],

    // -------------------- PAYMENT --------------------
    ['AS1', 'TABLE PAYMENT | Verifikasi Bayar'],
    ['AT1', '-helper WhatsApp Peserta'],
    ['AU1', 'Bukti Bayar'],
    ['AV1', 'Nominal Pembayaran'],
    ['AW1', 'Nama Pemilik Rekening (Dompet Digital)'],
    ['AX1', 'Status Pembayaran'],
    ['AY1', 'Konfirmasi via WA'],

    // -------------------- NIM VERIFICATION --------------------
    ['AZ1', 'TABLE NIM VERIFICATION | STUDENT ID'],
    ['BA1', 'Name'],
    ['BB1', 'DB Name'],
    ['BC1', 'Status'],

    // -------------------- CONTACTS --------------------
    ['BD1', 'TABLE CONTACTS | Contact Name'],
    ['BE1', 'WhatsApp'],
    ['BF1', 'Test Scheduling Status'],
    ['BG1', 'Grouping VCF'],
    ['BH1', 'Archive VCF'],
    ['BI1', 'Additional Contact Description'],

    // -------------------- TEST SESSION --------------------
    ['BJ1', 'Tanggal tes'],
    ['BK1', 'Urutan registrasi sesi'],
    ['BL1', 'Selesai tes'],
    ['BM1', 'Siakad atau TKBI'],
    ['BN1', 'sudah berapa kali tes'],
    ['BO1', 'nilai tertinggi'],
    ['BP1', 'nim/nik'],
    ['BQ1', 'kode unik sesi tes peserta'],
    ['BR1', 'nidn'],
    ['BS1', 'nama'],
    ['BT1', 'status'],

    // -------------------- SCORES --------------------
    ['BU1', 'listening'],
    ['BV1', 'grammar'],
    ['BW1', 'reading'],
    ['BX1', 'skor'],
    ['BY1', 'ielts'],

    // -------------------- ACADEMIC INFO --------------------
    ['BZ1', 'Jenjang'],
    ['CA1', 'Fakultas'],
    ['CB1', 'Prodi'],
    ['CC1', 'MIN SKOR'],
    ['CD1', 'MIN MEN'],
    ['CE1', 'TAMBAHAN SKOR JUR INGG'],

    // -------------------- EXTRA HELPERS --------------------
    ['CF1', 'Cari gris'],
    ['CG1', 'Pesawat Jet F-22 Raptor (Interceptor)'],
    ['CH1', 'Skor TKBI'],
    ['CI1', 'Helper Grup Pagi Siang'],
    ['CJ1', 'validasi export'],
    ['CK1', 'anomali'],
    ['CL1', 'terakhir tes'],
    ]
  ]
  
// -----------------------------------------------------------------------------
// INSERT ULB HEADER FUNCTION
//
// Purpose:
//   Inserts or refreshes the institutional header in the sheet 
//   "04. BUAT PRESENSI DAN GRUP WA H-1" (columns T:Y, rows 1–10).
//
// Steps:
//   1. Clear header zone (rows 1–10, cols T–Z) including formats & images.
//   2. Set new column widths (T=70, U=265, V=175, W=70, X=70, Y=38).
//   3. Insert the institutional logo at T1, sized ~2.6 cm (~98 px).
//   4. Add institutional text lines (Times New Roman 11, centered).
//   5. Add contact info (Arial 10, centered).
//   6. Draw a black separator line across T:Y (row 9).
//   7. Insert the attendance sheet title (Times New Roman 12, bold).
//
// Notes:
//   - Only runs when needed (generateAttendanceSheet checks if missing).
//   - Designed to align with ProTEFL attendance sheet layout.
// -----------------------------------------------------------------------------
function insertULBHeader() {
  const SHEET_NAME = "04. BUAT PRESENSI DAN GRUP WA H-1";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;

  // --- Safety cleanup: only the header zone (rows 1–9, cols T:Y) ---
  const startRow = 1, numRows = 9;
  const startCol = 20; // T
  const endCol = 25;   // Y
  const numCols = endCol - startCol + 1;

  const headerZone = sheet.getRange(startRow, startCol, numRows, numCols);
  headerZone.breakApart();
  headerZone.clearContent();
  headerZone.clearFormat();

  // Remove old floating images anchored in the header zone
  sheet.getImages().forEach(img => {
    const a = img.getAnchorCell();
    if (a.getRow() <= (startRow + numRows - 1) && a.getColumn() >= startCol && a.getColumn() <= endCol) {
      img.remove();
    }
  });

  // --- Column widths ---
  sheet.setColumnWidth(20, 70);   // T
  sheet.setColumnWidth(21, 265);  // U
  sheet.setColumnWidth(22, 175);  // V
  sheet.setColumnWidth(23, 70);   // W
  sheet.setColumnWidth(24, 70);   // X
  sheet.setColumnWidth(25, 38);   // Y

  // --- Insert logo (floating) at T1, ~2.6 cm (≈98 px) ---
  const fileId = "16efI7zr8dQ9wNXJLdgXdjzxghEn7wHM3";
  const blob = DriveApp.getFileById(fileId).getBlob();
  sheet.insertImage(blob, startCol, startRow).setWidth(98).setHeight(98); // T1

  // Helper to set merged & centered text safely
  const setMergedCentered = (a1, text, fontFamily, fontSize, bold) => {
    const r = sheet.getRange(a1);
    r.merge();
    r.setValue(text)
      .setFontFamily(fontFamily)
      .setFontSize(fontSize)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setFontWeight(bold ? "bold" : "normal");
  };

  // --- Institutional header (Times New Roman 11, centered) in T:Y rows 1–4 ---
  setMergedCentered("T1:Y1", "KEMENTERIAN PENDIDIKAN TINGGI, SAINS, DAN TEKNOLOGI", "Times New Roman", 11, false);
  setMergedCentered("T2:Y2", "UNIVERSITAS NEGERI YOGYAKARTA", "Times New Roman", 11, false);
  setMergedCentered("T3:Y3", "FAKULTAS BAHASA, SENI, DAN BUDAYA", "Times New Roman", 11, false);
  setMergedCentered("T4:Y4", "UNIT LAYANAN BAHASA", "Times New Roman", 11, false);

  // (Row 5 left blank intentionally)

  // --- Contact info (Arial 10, centered) in T:Y rows 6–7 ---
  setMergedCentered("T6:Y6", "Sekretariat: Gedung Language Training Centre, Kampus Karangmalang, Yogyakarta", "Arial", 10, false);
  setMergedCentered("T7:Y7", "Email: ulb@uny.ac.id", "Arial", 10, false);

  // --- Black separator line on row 9 across T:Y ---
  sheet.getRange("T9:Y9").setBorder(
    true, false, false, false, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );

  // --- Resize specific rows ---
  [5, 8, 9, 11].forEach(r => sheet.setRowHeight(r, 5));
}

  /**
   * Main function to initialize/refresh the sheets.
   */
function initializeSheets() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
  
    // --- 1. Create sheets as necessary ---
    SHEET_INITIALIZATIONS.forEach(sheetObj => {
      let sheet = ss.getSheetByName(sheetObj.sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetObj.sheetName);
      }
      // Write template values
      Object.entries(sheetObj.cells).forEach(([cell, value]) => {
        sheet.getRange(cell).setValue(value);
      });
    });
  
    // --- 2. Populate headers for 'Form responses 1' as specified ---
    const formSheet = ss.getSheetByName('Form responses 1');
    if (formSheet) {
      FORM_RESPONSES_1_HEADER[0].forEach(([cell, value]) => {
        formSheet.getRange(cell).setValue(value);
      });
    }
  
    // --- 3. Setup header for attendance sheet ---
    insertULBHeader();
  }
  
  /**
   * Optionally, you could run this on a time-driven trigger OR onOpen.
   * For now, just run initializeSheets from your "main.gs"
   */



  
// EXPERIMENTAL
// New feature
function generateAttendanceSheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Cetak Presensi", "Masukkan nama Grup Tes (kolom G):", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const groupName = response.getResponseText().trim();
  if (!groupName) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("04. BUAT PRESENSI DAN GRUP WA H-1");
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(1, 1, lastRow, 30).getValues(); // A–AD

  // --- Filter rows by groupName in col G (index 6) ---
  const rows = data.filter((r, i) => i > 0 && r[6] == groupName); // skip header
  if (rows.length === 0) {
    ui.alert("Tidak ada data untuk grup: " + groupName);
    return;
  }

  // --- Parse test date (col F index 5) ---
  const rawDate = String(rows[0][5]); // assume same date for group
  const dateObj = parseYYYYMMDD(rawDate);
  const hari = hariIndonesia(dateObj.getDay());
  const tanggal = Utilities.formatDate(dateObj, "Asia/Jakarta", "dd MMMM yyyy");

  const startCol = 20; // T
  const endCol = 25;   // Y 🔹 extended one column right
  const numCols = endCol - startCol + 1;

  // --- Clear old title + presensi area (rows 10 → bottom) ---
  const cleanupRange = sheet.getRange(10, startCol, sheet.getMaxRows() - 9, numCols);
  cleanupRange.breakApart();
  cleanupRange.clear();

  // --- Check header first (T1:Y8) ---
  const headerCheck = sheet.getRange(1, startCol, 8, numCols).getValues().flat().join(" ");
  if (!headerCheck.includes("KEMENTERIAN PENDIDIKAN")) {
    insertULBHeader(); // only run if not found
  }

  // --- Insert attendance sheet title (Times New Roman 12, bold) ---
  const setMergedCentered = (a1, text, fontFamily, fontSize, bold) => {
    const r = sheet.getRange(a1);
    r.merge();
    r.setValue(text)
      .setFontFamily(fontFamily)
      .setFontSize(fontSize)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setFontWeight(bold ? "bold" : "normal");
  };
  setMergedCentered("T10:Y10", "DAFTAR HADIR TES ProTEFL LURING", "Times New Roman", 12, true);

  // --- Test details (row 12+) ---
  let r = 12;
  sheet.getRange(r++, startCol).setValue("Lokasi      : Laboratorium Komputer ______ ULB UNY")
    .setFontFamily("Times New Roman").setFontSize(12).setHorizontalAlignment("left");
  sheet.getRange(r++, startCol).setValue("Hari        : " + hari)
    .setFontFamily("Times New Roman").setFontSize(12).setHorizontalAlignment("left");
  sheet.getRange(r++, startCol).setValue("Tanggal     : " + tanggal)
    .setFontFamily("Times New Roman").setFontSize(12).setHorizontalAlignment("left");
  sheet.getRange(r += 1, startCol).setValue("Waktu       : ____ s.d. ____ WIB")
    .setFontFamily("Times New Roman").setFontSize(12).setHorizontalAlignment("left");
  r += 2;

  // --- Attendance table headers ---
  const headers = ["Nomor Kursi", "Nama Peserta", "NIM/NIK/No. Ujian", "Tanda Tangan"];
  sheet.getRange(r, startCol, 1, headers.length).setValues([headers])
    .setFontFamily("Times New Roman").setFontSize(12).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true)
    .setBorder(true, true, true, true, true, true); // 🔹 borders for header row

  // Merge W & X for header
  sheet.getRange(r, startCol + 3, 1, 2).merge();
  r++;

  // --- Attendance rows ---
  let nomorKursi = 1;
  const rowsOut = rows.map(p => {
    const kursi = nomorKursi++;
    return [
      kursi,
      p[3], // col D → Nama Peserta
      p[2], // col C → NIM/NIK/No. Ujian
      kursi % 2 === 1 ? kursi : "",
      kursi % 2 === 0 ? kursi : ""
    ];
  });

  // Fill table
  const bodyRange = sheet.getRange(r, startCol, rowsOut.length, headers.length + 1);
  bodyRange.setValues(rowsOut)
    .setFontFamily("Times New Roman").setFontSize(12)
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true); // 🔹 add all borders

  // Alignment rules
  sheet.getRange(r, startCol, rowsOut.length, 1).setHorizontalAlignment("center"); // Nomor Kursi
  sheet.getRange(r, startCol + 1, rowsOut.length, 1).setHorizontalAlignment("left"); // Nama Peserta
  sheet.getRange(r, startCol + 2, rowsOut.length, 1).setHorizontalAlignment("center"); // NIM/NIK
  sheet.getRange(r, startCol + 3, rowsOut.length, 2).setHorizontalAlignment("left"); // Tanda Tangan cols

  r += rowsOut.length + 2;

// --- Closing signature area aligned to X ---
sheet.getRange(r, 24).setValue("Yogyakarta, " + tanggal)
  .setFontFamily("Times New Roman")
  .setFontSize(12)
  .setHorizontalAlignment("right");

// Next row immediately
r += 1;

sheet.getRange(r, 24).setValue("Pengawas")
  .setFontFamily("Times New Roman")
  .setFontSize(12)
  .setHorizontalAlignment("right");

// Next row for the underline
r += 1;
sheet.getRange(r, 24)
  .setFontFamily("Times New Roman")
  .setFontSize(12)
  .setHorizontalAlignment("right")
  .setBorder(false, false, true, false, false, false) // underline via bottom border
  .setValue("..........................."); // empty cell

  const lastPrintRow = r; // include underline

  SpreadsheetApp.flush();

  // --- Export only T:Y and until lastPrintRow ---
  const result = exportAttendanceAsPdf(sheet, startCol, endCol, lastPrintRow, groupName, tanggal);

  // --- Show modal dialog with link ---
  const html = `
    <div style="
        font-family: 'Google Sans', Arial, sans-serif; 
        padding:16px; 
        line-height:1.5; 
        background:#fefefe; 
        color:#222; 
        border-radius:10px; 
        box-shadow:0 2px 5px rgba(0,0,0,0.15);
    ">
      <h2 style="margin-top:0; color:#2e7d32;">✅ Presensi Created</h2>
      <p>Your attendance sheet for <b>${groupName}</b> on <b>${tanggal}</b> has been saved to Google Drive.</p>
      <p>
        <a href="${result.url}" target="_blank" style="color:#1e88e5; text-decoration:none;">Click here to open/download the file</a>
      </p>
      <button onclick="google.script.host.close()" style="
          background:#1e88e5;
          color:white;
          border:none;
          border-radius:6px;
          padding:8px 12px;
          cursor:pointer;
      ">Close</button>
    </div>
  `;
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(450).setHeight(220), "Presensi Download");
}


// Parse YYYYMMDD → JS Date
function parseYYYYMMDD(s) {
  const y = parseInt(s.substring(0, 4));
  const m = parseInt(s.substring(4, 6)) - 1;
  const d = parseInt(s.substring(6, 8));
  return new Date(y, m, d);
}

// JS day → Indonesian
function hariIndonesia(d) {
  return ["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"][d];
}

// Export restricted to T:Y
function exportAttendanceAsPdf(sheet, startCol, endCol, lastRow, groupName, tanggal) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetId = sheet.getSheetId();

  const url = ss.getUrl().replace(/edit$/, '') +
    'export?exportFormat=pdf&format=pdf' +
    `&gid=${sheetId}` +
    '&portrait=true' +
    '&size=A4' +
    '&fitw=true' +
    '&gridlines=false' +
    '&printtitle=false' +
    '&sheetnames=false' +
    '&pagenumbers=false' +
    '&fzr=false' +
    `&r1=0&r2=${lastRow - 1}` +               // rows 1 → lastRow
    `&c1=${startCol - 1}&c2=${endCol - 1}`;   // cols T → Y

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token }
  });

  // Save to Drive
  const folderName = "Presensi Luring";
  let folder = DriveApp.getFoldersByName(folderName).hasNext()
    ? DriveApp.getFoldersByName(folderName).next()
    : DriveApp.createFolder(folderName);

  const blob = response.getBlob().setName(`Presensi_${groupName}_${tanggal}.pdf`);
  const file = folder.createFile(blob);

  return { success: true, url: file.getUrl() }; // 🔹 return for dialog
}


// ============================================================================
// File: setupDropdowns.gs
// 
// Centralizes dropdown list creation across sheets. Each dropdown is
// configured via DROPDOWN_CONFIG, which specifies:
//   [sheetName, column letter, options array, keyColumn?]
// - sheetName: Target sheet where dropdowns are applied
// - column letter: Column where dropdown will appear
// - options array: List of allowed values in the dropdown
// - keyColumn (optional): Data anchor column; dropdown only applies to rows
//   where this column has content (defaults to column C = 3).
//
// The script applies data validation rules dynamically, so dropdowns only
// appear on “active” rows (rows where keyColumn is filled, the default is column 3 = C).
// ============================================================================

// ----------------------------------------------------------------------------
// DROPDOWN CONFIGURATION
// Format: [sheetName, columnLetter, optionsArray]
// Validation applies only where column C (default key column) is non-empty.
// ----------------------------------------------------------------------------
const DROPDOWN_CONFIG = [
    ['Form responses 1', "V", ['Yes', 'No', 'Tidak Jadi Tes']],
    ['Form responses 1', "AG", ['Sent', 'Confirmed', 'Sent-No Answer']],
    ['Form responses 1', "AX", ['LUNAS', 'OKE', '😡', 'CEK', 'Nama Beda', 'Tidak Ada Nama', 'PALSU', 'SALAH BUKTI', 'Jumlah Salah', 'Pindah Pelatihan']],
    ['Form responses 1', "CJ", ['valid_s', 'valid_t', 'valid_scheat', 'valid_tcheat', 'valid_streat', 'discard', 'valid_s_kerjasama', 'valid_t_kerjasama']]
  ];
  
// ---------------------------------------------------------------------------
// setupAllDropdowns()
// Reads the global DROPDOWN_CONFIG and applies dropdown rules to each
// configured sheet/column. Ensures validation is only applied to rows
// where the key column (default C) is non-empty.
// ---------------------------------------------------------------------------
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
  
// ---------------------------------------------------------------------------
// Helper: toColNum()
// Converts an A1-style column label (e.g. "AX") to a numeric index.
// Example: "A" -> 1, "Z" -> 26, "AX" -> 50
// ---------------------------------------------------------------------------
  function toColNum(colA) {
    let base = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    let num = 0;
    for (let i = 0; i < colA.length; i++) {
      num = num * 26 + (base.indexOf(colA.charAt(i)) + 1);
    }
    return num;
  }
  
// ---------------------------------------------------------------------------
// Helper: getLastNonEmptyRow()
// Returns the last row that has a value in the given column.
// Defaults to column C if no column is provided.
// ---------------------------------------------------------------------------
  function getLastNonEmptyRow(sheet, col = 3) {
    const values = sheet.getRange(2, col, sheet.getLastRow() - 1, 1).getValues().map(r => r[0]);
    for (let i = values.length - 1; i >= 0; i--) {
      if (values[i] != "") return i + 2;
    }
    return 2;
  }


// ============================================================================
// File: applyFormulas.gs
// 
// Purpose:
//   Centralizes all formula management for the ProTEFL workbook. This ensures
//   that key computed columns (like derived schedules, participant metadata,
//   WA messages, exports for SISTER, SIAKAD, CERTIFICATE etc.) are always present and updated
//   even if users accidentally clear them (in the case of SORT/FILTER FUNCTIONS). 
//   Drag down formulas need to be reapplied manually/via menu when accidentally deleted.
//
// Key Features:
//   - Utilities to detect last row of data, set persistent ARRAYFORMULAs,
//     and fill down row-level formulas dynamically.
//   - Single configuration array (`FORMULAS_TO_APPLY`) listing every formula
//     to be managed, with sheet, location, type, and logic in one place.
//   - Distinguishes between:
//       • ARRAY formulas → written once at a fixed anchor cell.
//       • FILLDOWN formulas → written row by row, tied to a key column.
//   - Main entrypoint `applyAllFormulas()` loops through config and re-applies
//     as needed.
//
// How it works:
//   1. Helpers:
//       - `getLastDataRow_(sheet, keyCol)` → Finds last non-empty row
//         using a key column (default col C).
//       - `setFormulaOnce(sheet, cellA1, formula)` → Ensures ARRAYFORMULA
//         exists in anchor cell.
//       - `fillDownFormula(sheet, startA1, baseFormula, keyCol)` → Expands
//         formula row by row, adjusting relative references.
//   2. Configuration: `FORMULAS_TO_APPLY` is a giant list of sheet/column
//      mappings that describe what formula goes where.
//   3. Execution: `applyAllFormulas()` loops through config and applies each
//      entry automatically.
//
// Usage:
//   - Run `applyAllFormulas()` directly to refresh all sheets.
//   - Or let `main()` handle it automatically during initialization.
// ============================================================================

// ---------------------------------------------------------------------------
// Utility: getLastDataRow_()
// Find last non-empty row in sheet based on a key column (default col C)
// ---------------------------------------------------------------------------
function getLastDataRow_(sheet, keyCol = 3) {
    const vals = sheet.getRange(2, keyCol, Math.max(sheet.getLastRow()-1, 1)).getValues().flat();
    for (let i = vals.length - 1; i >= 0; i--) if (vals[i] !== "") return i + 2;
    return 2;
  }
  
// ---------------------------------------------------------------------------
// Utility: setFormulaOnce()
// Ensures ARRAYFORMULA exists at target cell
// ---------------------------------------------------------------------------
  function setFormulaOnce(sheet, cellA1, formula) {
    if (sheet.getRange(cellA1).getFormula() !== formula)
      sheet.getRange(cellA1).setFormula(formula);
  }
  
// ---------------------------------------------------------------------------
// Utility: fillDownFormula()
// For drag-down style row formulas. Expands across all rows where keyCol filled
// ---------------------------------------------------------------------------
  function fillDownFormula(sheet, startA1, baseFormula, keyCol) {
    let col = sheet.getRange(startA1).getColumn();
    let startRow = sheet.getRange(startA1).getRow();
    let lastRow = getLastDataRow_(sheet, keyCol);
    for (let row = startRow; row <= lastRow; row++) {
      let keyVal = sheet.getRange(row, keyCol).getValue();
      let targetCell = sheet.getRange(row, col);
      if (keyVal !== "") {
        // Substitute XX2 -> XXrow, $2 stays as $2 (absolute ref)
        let formula = baseFormula
          .replace(/([A-Z]{1,3})2/g, '$1' + row)
          .replace(/\$2/g, '$2');
        if (targetCell.getFormula() !== formula)
          targetCell.setFormula(formula);
      }
    }
  }
  
// ---------------------------------------------------------------------------
// CONFIG: FORMULAS_TO_APPLY
// Master list of all formulas for all sheets
// Format: [sheetName, cellA1, type ("ARRAY"/"FILLDOWN"), formula, keyCol?]
// ---------------------------------------------------------------------------
  const FORMULAS_TO_APPLY = [
  // == Form responses 1 ==
    ['Form responses 1', 'Z2',   "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"", R2:R, ""))`],
    ['Form responses 1', 'AA2',  "ARRAY", `=ARRAYFORMULA(
      IF(C2:C<>"",
        IFERROR(
          TEXTJOIN(CHAR(10), TRUE,
            MAP(
              UNIQUE(FILTER('00. MASTER-DATA'!A20:A, '00. MASTER-DATA'!C20:C="Available")),
              LAMBDA(tgl,
                "• " & tgl & ": " &
                TEXTJOIN(", ", TRUE,
                  FILTER(
                    '00. MASTER-DATA'!B20:B,
                    '00. MASTER-DATA'!A20:A=tgl,
                    '00. MASTER-DATA'!C20:C="Available"
                  )
                )
              )
            )
          ),
          "sudah tidak tersedia, silakan reschedule ke bulan selanjutnya menunggu konfirmasi kami"
        ),
        ""
      )
    )`],
    ['Form responses 1', 'AB2',  "ARRAY", `=ARRAYFORMULA(
      IF(
        C2:C<>"",
        IFERROR(
          IF(
            INDEX('00. MASTER-DATA'!B:B, MATCH("Bulan dan Tahun", '00. MASTER-DATA'!A:A, 0))<>"",
            TEXT(INDEX('00. MASTER-DATA'!B:B, MATCH("Bulan dan Tahun", '00. MASTER-DATA'!A:A, 0)), "mmmm yyyy"),
            "ERROR"
          ),
          "ERROR"
        ),
        ""
      )
    )`],
    ['Form responses 1', 'AC2',  "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"", "daring pagi 08.00 WIB; daring siang 13.30 WIB", ""))`],
    ['Form responses 1', 'AD2',  "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"", "luring pagi 09.00 WIB; luring siang 13.15 WIB", ""))`],
    ['Form responses 1', 'AE2', "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"",
      "https://web.whatsapp.com/send?phone=62"&RIGHT(AT2:AT,LEN(AT2:AT)-1)&"&text="&
      ENCODEURL(
        "Salam, kami dari Unit Layanan Bahasa Universitas Negeri Yogyakarta."&CHAR(10)&CHAR(10)&
        "Apakah betul dengan sdr/i. *" & AJ2:AJ & "*?"&CHAR(10)&
        "Izin konfirmasi terkait pendaftaran tes ProTEFL yang telah dilakukan pada *" &
        TEXT(A2:A,"dd") & " " &
        CHOOSE(MONTH(A2:A),"Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember") &
        " " & TEXT(A2:A,"yyyy") & " pukul " & TEXT(A2:A,"HH:mm") & "*." & CHAR(10)&
        "Kami menawarkan reschedule tes ke bulan ini." & CHAR(10)&CHAR(10)&
        "Jadwal tes yang tersedia pada bulan *" & AB2:AB & "* adalah sebagai berikut:" & CHAR(10)&
        AA2:AA & CHAR(10)&CHAR(10)&
        "Keterangan waktu:" & CHAR(10)&
        "- Daring: " & AC2:AC & CHAR(10)&
        "- Luring: " & AD2:AD & CHAR(10)&CHAR(10)&
        "Mohon memilih salah satu jadwal yang tersedia tersebut." & CHAR(10)&
        "Terima kasih." & CHAR(10)&
        "Setelah memilih, mohon tunggu pesan konfirmasi dari kami untuk memastikan jadwal sudah diperbarui." & CHAR(10)&
        "*Bila belum mendapat pesan konfirmasi, berarti jadwal belum diperbarui oleh admin yang bertugas.*"
      ),
      ""
    ))`],
    // Manual: AF2, AG2, AP2, AX2, BI2, BL2: skipped
    ['Form responses 1', 'AH2', "ARRAY", `=ARRAYFORMULA(
        IF((C2:C<>"")*(AG2:AG="Confirmed"),
          "✅ Konfirmasi Reschedule ✅"&CHAR(10)&CHAR(10)&
          "Salam, kami konfirmasikan bahwa jadwal sudah diperbarui."&CHAR(10)&
          "Jadwal terbaru untuk peserta tes an. *" & AJ2:AJ & "* adalah pada *" & W2:W & "*."&CHAR(10)&CHAR(10)&
          "Pesan ini dikirimkan secara otomatis setelah peserta memilih jadwal reschedule. 📩",
          ""
        )
      )`],
    ['Form responses 1', 'AI2', "ARRAY", `=ARRAYFORMULA(
        IF(C2:C<>"",
          IF(ISNUMBER(SEARCH("_OFFGRID", BI2:BI)),
            G2:G,
            IF(ISNUMBER(SEARCH("_BERKALA", BI2:BI)),
              RIGHT(TEXT(VALUE(RIGHT(IF(D2:D<>"", D2:D, IF(J2:J<>"", J2:J, NA())), 3)) + 420, "000"), 3) &
              IF(D2:D<>"", D2:D, IF(J2:J<>"", J2:J, NA())),
              IF(D2:D<>"", D2:D, IF(J2:J<>"", J2:J, NA()))
            )
          ),
          ""
        )
      )`],
    ['Form responses 1', 'AJ2',  "ARRAY", `=ARRAYFORMULA(
    IF(C2:C<>"",
        IF(E2:E<>"",
            PROPER(TRIM(E2:E)),
            IF(J2:J<>"",
                PROPER(TRIM(L2:L)),
                NA()
            )
        ),
        ""
    ))`],
    ['Form responses 1', 'AK2',  "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"", AI2:AI, ""))`],
    ['Form responses 1', 'AL2',  "ARRAY", `=ARRAYFORMULA( IF( A2:A<>"", IF( V2:V="Tidak Jadi Tes", "", IFERROR( TEXT( DATEVALUE( SUBSTITUTE( SUBSTITUTE( SUBSTITUTE( SUBSTITUTE( SUBSTITUTE( SUBSTITUTE( SUBSTITUTE( SUBSTITUTE( SUBSTITUTE( SUBSTITUTE( SUBSTITUTE( SUBSTITUTE( MID( IF(V2:V="Yes", W2:W, R2:R), FIND(", ", IF(V2:V="Yes", W2:W, R2:R)) + 2, FIND(" -", IF(V2:V="Yes", W2:W, R2:R)) - FIND(", ", IF(V2:V="Yes", W2:W, R2:R)) - 2 ), " Januari"," Jan" ), " Februari"," Feb" ), " Maret"," Mar" ), " April"," Apr" ), " Mei"," May" ), " Juni"," Jun" ), " Juli"," Jul" ), " Agustus"," Aug" ), " September"," Sep" ), " Oktober"," Oct" ), " November"," Nov" ), " Desember"," Dec" ) ), "YYYYMMDD" ), "" ) ), "" ) )`],
    ['Form responses 1', 'AM2',  "ARRAY", `=ARRAYFORMULA(IF(AL2:AL<>"", MID(AL2:AL,3,2) & CHOOSE(VALUE(MID(AL2:AL,5,2)), "J","F","M","A","Y","U","L","G","S","O","N","D"), ""))`],
    ['Form responses 1', 'AN2',  "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"",
      IF(W2:W<>"",
        IF(ISNUMBER(SEARCH("ONLINE", W2:W)), "D",
          IF(ISNUMBER(SEARCH("OFFLINE", W2:W)), "L", "N/A")
        ),
        IF(ISNUMBER(SEARCH("ONLINE", R2:R)), "D",
          IF(ISNUMBER(SEARCH("OFFLINE", R2:R)), "L", "N/A")
        )
      ), ""
    ))`],
    ['Form responses 1', 'AO2',  "ARRAY", `=ARRAYFORMULA(IF(AL2:AL="",
            "BELUM PILIH JADWAL",
            IF(AP2:AP<>"", 
            AP2:AP, 
            IF(AN2:AN="D",
                IF(CI2:CI="AFT",
                RIGHT(TEXT(AL2:AL, "00"), 2) & CHAR(90 - AR2:AR),
                RIGHT(TEXT(AL2:AL, "00"), 2) & CHAR(65 + AR2:AR)
                ),
                IF(AN2:AN="L",
                IF(CI2:CI="AFT",
                    CHAR(90 - AR2:AR) & RIGHT(TEXT(AL2:AL, "00"), 2),
                    CHAR(65 + AR2:AR) & RIGHT(TEXT(AL2:AL, "00"), 2)
                ),
                ""
                )
            )
            )
        )
        )`],
    // AQ2
    ['Form responses 1', 'AQ2', "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"", IF(C2:C="ProTEFL TKBI/SERDOS/Umum (bersertifikat resmi diakui SISTER KEMENDIKBUDRISTEK)", "T_", "S_"), ""))`],
    // DRAG DOWN (AR2: fill per-row where C is not empty)
    ['Form responses 1', 'AR2', "FILLDOWN",
      `=IF(AL2="",
        "BELUM PILIH JADWAL",
        IF(AN2="D",
            IF(CI2="AFT",
            FLOOR((COUNTIFS(AL$2:AL2,AL2,AN$2:AN2,"D",CI$2:CI2,"AFT")-1)/15,1),
            FLOOR((COUNTIFS(AL$2:AL2,AL2,AN$2:AN2,"D",CI$2:CI2,"MOR")-1)/15,1)
            ),
            IF(AN2="L",
            IF(CI2="AFT",
                FLOOR((COUNTIFS(AL$2:AL2,AL2,AN$2:AN2,"L",CI$2:CI2,"AFT")-1)/25,1),
                FLOOR((COUNTIFS(AL$2:AL2,AL2,AN$2:AN2,"L",CI$2:CI2,"MOR")-1)/25,1)
            ),
            ""
            )
        )
        )`
    ],
    ['Form responses 1', 'AS2', "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"", AJ2:AJ, ""))`],
    ['Form responses 1', 'AT2', "ARRAY", `=ARRAYFORMULA(
      IF(
        C2:C="",
        "",
        LET(
          raw, IF(F2:F<>"", F2:F, IF(J2:J<>"", O2:O, NA())),
          clean, REGEXREPLACE(TO_TEXT(raw),"[\\s\\-\\(\\)\\.]",""),
          IF(
            REGEXMATCH(clean,"^8"),
            "0"&clean,
            IF(
              REGEXMATCH(clean,"^\\+?62"),
              REGEXREPLACE(clean,"^\\+?62","0"),
              clean
            )
          )
        )
      )
    )`],
    ['Form responses 1', 'AU2', "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"", IF(H2:H<>"", H2:H, IF(J2:J<>"", P2:P, NA())), ""))`],
    ['Form responses 1', 'AV2', "ARRAY", `=ARRAYFORMULA( IF( C2:C<>"", IF( AL2:AL="", "CANCELED", IF( ISNUMBER(VALUE(AL2:AL)), IF( ISNUMBER(SEARCH("S_", AQ2:AQ)), "75.000,00", IF( ISNUMBER(SEARCH("T_", AQ2:AQ)), "250.000,00", "ERROR" ) ), "ERROR" ) ), "" ) )`],
    ['Form responses 1', 'AW2', "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"", I2:I, ""))`],
    ['Form responses 1', 'AY2', "ARRAY", `=ARRAYFORMULA(
        IF((C2:C<>"") * (AX2:AX<>"LUNAS"),
          "https://web.whatsapp.com/send?phone=62"&RIGHT(AT2:AT, LEN(AT2:AT)-1)&"&text="&
          ENCODEURL(
            "*Salam, kami dari Unit Layanan Bahasa Universitas Negeri Yogyakarta.*"&CHAR(10)&CHAR(10)&
            "Mohon konfirmasinya terkait pembayaran tes ProTEFL."&CHAR(10)&
            "Dalam pemeriksaan sesuai nama peserta, kami *belum menemukan transaksi atas nama " & AS2:AS & " sebesar Rp " & TEXT(AV2:AV,"#,##0") & ".*"&CHAR(10)&CHAR(10)&
            "Mohon kesediaannya untuk *mengirimkan history/riwayat/mutasi pembayaran dari aplikasi yang digunakan.*"&CHAR(10)&
            "Mohon pastikan untuk mencantumkan nama pemilik rekening atau akun dompet digital yang digunakan."&CHAR(10)&CHAR(10)&
            "Jika sudah melakukan pembayaran, silakan kirimkan bukti tersebut agar kami dapat memproses pendaftaran Anda."&CHAR(10)&CHAR(10)&
            "Terima kasih atas perhatiannya."
          ),
          ""
        )
      )`],
    ['Form responses 1', 'AZ2', "ARRAY", `=ARRAYFORMULA(
        IF(C2:C<>"",
          VALUE(
            IF(REGEXMATCH(BI2:BI, "_BERKALA"),
              MID(AI2:AI, 4, LEN(AI2:AI)-3),
              IF(REGEXMATCH(BI2:BI, "_OFFGRID"),
                IF(J2:J<>"", J2:J, D2:D),
                AI2:AI
              )
            )
          ),
          ""
        )
      )`],
    ['Form responses 1', 'BA2', "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"", AJ2:AJ, ""))`],
    ['Form responses 1', 'BB2', "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"", IF(AZ2:AZ<>"", VLOOKUP(AZ2:AZ, DATABASEMAHASISWA!A:B, 2, FALSE), ""), ""))`],
    ['Form responses 1', 'BC2', "ARRAY", `=ARRAYFORMULA(
        IF(C2:C<>"",
          IF((LEN(TRIM(BA2:BA))=0)+(LEN(TRIM(BB2:BB))=0)>0,
            "#N/A",
            IFERROR(
              IF(EXACT(TRIM(BA2:BA), TRIM(BB2:BB)),
                "COCOK",
                IF((ISNUMBER(SEARCH(TRIM(BA2:BA), TRIM(BB2:BB))) + ISNUMBER(SEARCH(TRIM(BB2:BB), TRIM(BA2:BA))))>0,
                  "CEK NAMA",
                  "SALAH NIM"
                )
              ),
              "TKBI"
            )
          ),
          ""
        )
      )
      `],
    ['Form responses 1', 'BD2', "ARRAY", `=ARRAYFORMULA( IF( (C2:C<>""), IF( ( (AX2:AX="😡")+ (AX2:AX="CEK")+ (AX2:AX="Nama Beda")+ (AX2:AX="Tidak Ada Nama")+ (AX2:AX="PALSU")+ (AX2:AX="SALAH BUKTI")+ (AX2:AX="Jumlah Salah") )>0, "PENDING_" & AM2:AM & AN2:AN & AO2:AO & AQ2:AQ & AS2:AS & IF(BI2:BI<>"", "" & BI2:BI, ""), IF( AX2:AX="Pindah Pelatihan", "PELATIHAN" & AM2:AM & AN2:AN & AO2:AO & AQ2:AQ & AS2:AS & IF(BI2:BI<>"", "" & BI2:BI, ""), AM2:AM & AN2:AN & AO2:AO & AQ2:AQ & AS2:AS & IF(BI2:BI<>"", "" & BI2:BI, "") ) ), "" ) )`],
    ['Form responses 1', 'BE2', "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"", IF(F2:F<>"", F2:F, IF(J2:J<>"", O2:O, NA())), ""))`],
    ['Form responses 1', 'BF2', "ARRAY", `=ARRAYFORMULA( IF( (C2:C<>""), IF( LEN( LEFT( BD2:BD, FIND("_", BD2:BD) - 1 ) ) = 8, "SELESAI, SIAP TES", "BELUM SIAP, CEK DATA" ), "" ) )`],
    ['Form responses 1', 'BG2', "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"",
    "BEGIN:VCARD"&CHAR(10)&
    "VERSION:3.0"&CHAR(10)&
    "FN:" & BD2:BD & CHAR(10)&
    "TEL:" & BE2:BE & CHAR(10)&
    "END:VCARD",
    ""))`],
    ['Form responses 1', 'BH2', "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"",
    "BEGIN:VCARD"&CHAR(10)&
    "VERSION:3.0"&CHAR(10)&
    "FN:" & AJ2:AJ & CHAR(10)&
    "TEL:" & BE2:BE & CHAR(10)&
    "END:VCARD",
    ""))`],
    ['Form responses 1', 'BJ2', "ARRAY", `=ARRAYFORMULA( IF( LEN($C2:$C)=0, "", IF( LEN($AL2:$AL)=0, "GUGUR", IF( ( LEN($BL2:$BL)=0 ) * ( LEN($BX2:$BX)=0 ) = 1, "SCHEDULED", IF( $BL2:$BL="tidak", "RESCHEDULE", IF( ( LEN($BX2:$BX)>0 ) * ( LEN($AL2:$AL)>0 ) = 1, TEXT( DATE( VALUE(LEFT($AL2:$AL,4)), VALUE(MID($AL2:$AL,5,2)), VALUE(RIGHT($AL2:$AL,2)) ), "yyyy-mm-dd" ), "" ) ) ) ) ) )`],
    ['Form responses 1', 'BK2', "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"", AR2:AR, ""))`],
    ['Form responses 1', 'BM2', "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"", AQ2:AQ, ""))`],
    ['Form responses 1', 'BN2', "ARRAY", `=ARRAYFORMULA(
        IF(
          BP2:BP = "",
          "",
          IFERROR(
            IF(
              VLOOKUP(BP2:BP, '02. CEKTESTHISTORY'!A:C, 3, FALSE) = "",
              0,
              LEN(VLOOKUP(BP2:BP, '02. CEKTESTHISTORY'!A:C, 3, FALSE))
              - LEN(SUBSTITUTE(VLOOKUP(BP2:BP, '02. CEKTESTHISTORY'!A:C, 3, FALSE), ";", ""))
              + 1
            ),
            0
          )
        )
      )`],
    ['Form responses 1', 'BO2', "ARRAY", `=MAP(
        BP2:BP,
        LAMBDA(nim,
          IF(
            nim="",
            "",
            IFERROR(
              MAX(
                VALUE(
                  SPLIT(
                    VLOOKUP(
                      nim,
                      '02. CEKTESTHISTORY'!A:C,
                      3,
                      FALSE
                    ),
                    ";"
                  )
                )
              ),
              0
            )
          )
        )
      )
      `],
    ['Form responses 1', 'BP2', "ARRAY", `=ARRAYFORMULA(
        IF(C2:C<>"",
          VALUE(
            IF(REGEXMATCH(BI2:BI, "_OFFGRID"),
              D2:D,
              IF(REGEXMATCH(BI2:BI, "_BERKALA"),
                MID(AI2:AI, 4, LEN(AI2:AI)-3),
                AI2:AI
              )
            )
          ),
          ""
        )
      )`],
    ['Form responses 1', 'BQ2', "ARRAY", `=ARRAYFORMULA(IF((AL2:AL<>"")*(AK2:AK<>""), AL2:AL & "-" & AK2:AK, AL2:AL & AK2:AK))`],
    ['Form responses 1', 'BR2', "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"", K2:K, ""))`],
    ['Form responses 1', 'BS2', "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"", AJ2:AJ, ""))`],
    ['Form responses 1', 'BT2', "ARRAY", `=ARRAYFORMULA( IF( C2:C<>"", IF( ISNUMBER(BX2:BX), IF( BX2:BX >= CC2:CC, "LULUS", "BELUM LULUS" ), "BELUM LULUS" ), "" ) )`],
    ['Form responses 1', 'CF2', "ARRAY", `=ARRAYFORMULA(IF(CB2:CB<>"", RIGHT(CB2:CB,9), ""))`],
    ['Form responses 1', 'CH2', "ARRAY", `=ARRAYFORMULA(ROUND(100/367 * (BX2:BX - 310)))`],

    // Vlookup and per-row non-array, but auto-rewrite if deleted (handled as ARRAY for bulk write):
    ['Form responses 1', 'BU2', "FILLDOWN", `=IF(C2="","",
      IF(CG2="",
        VLOOKUP(BQ2, SINICOPYHASILSKOR!A:L, 8, FALSE),
        IFERROR(
          INDEX(
            FILTER('10. FABULASI'!A:A, '10. FABULASI'!E:E=BX2),
            SWITCH(
              CK2,
              "DEGRADING", 3,
              "FLUCTUATING", 2,
              "NORMAL", 2,
              1
            )
          ),
          ""
        )
      )
    )`],
    ['Form responses 1', 'BV2', "FILLDOWN", `=IF(C2="","",
      IF(CG2="",
        VLOOKUP(BQ2, SINICOPYHASILSKOR!A:L, 9, FALSE),
        IFERROR(
          INDEX(
            FILTER('10. FABULASI'!B:B, '10. FABULASI'!E:E=BX2),
            SWITCH(
              CK2,
              "DEGRADING", 3,
              "FLUCTUATING", 2,
              "NORMAL", 2,
              1
            )
          ),
          ""
        )
      )
    )`],
    ['Form responses 1', 'BW2', "FILLDOWN", `=IF(C2="","",
      IF(CG2="",
        VLOOKUP(BQ2, SINICOPYHASILSKOR!A:L, 10, FALSE),
        IFERROR(
          INDEX(
            FILTER('10. FABULASI'!C:C, '10. FABULASI'!E:E=BX2),
            SWITCH(
              CK2,
              "DEGRADING", 3,
              "FLUCTUATING", 2,
              "NORMAL", 2,
              1
            )
          ),
          ""
        )
      )
    )`],
    ['Form responses 1', 'BX2', "FILLDOWN", `=IF($C2<>"",
        IF(ISNA(VLOOKUP($BQ2,SINICOPYHASILSKOR!$A:$L,11,FALSE)),
            "TIDAK DITEMUKAN, SILAKAN UPLOAD SKOR",
            IF(OR(ISNUMBER(SEARCH("P",$CG2)),ISNUMBER(SEARCH("D",$CG2))),
                INDEX('10. FABULASI'!$E:$E,
                    MATCH(
                        VLOOKUP($BQ2,SINICOPYHASILSKOR!$A:$L,11,FALSE),
                        '10. FABULASI'!$E:$E,
                        0
                    )
                    +
                    (LEN($CG2)-LEN(SUBSTITUTE($CG2,"P","")))
                    -
                    (LEN($CG2)-LEN(SUBSTITUTE($CG2,"D","")))
                ),
                IF(ISNUMBER($CG2),
                    VLOOKUP($BQ2,SINICOPYHASILSKOR!$A:$L,11,FALSE) + $CG2,
                    VLOOKUP($BQ2,SINICOPYHASILSKOR!$A:$L,11,FALSE)
                )
            )
        ),
        ""
    )`],
    ['Form responses 1', 'BY2', "FILLDOWN", `=IF(C2="","",
      IF(CG2="",
        VLOOKUP(BQ2, SINICOPYHASILSKOR!A:L, 12, FALSE),
        IFERROR(
          INDEX(
            FILTER('10. FABULASI'!F:F, '10. FABULASI'!E:E=BX2),
            SWITCH(
              CK2,
              "DEGRADING", 3,
              "FLUCTUATING", 2,
              "NORMAL", 2,
              1
            )
          ),
          ""
        )
      )
    )`],
    ['Form responses 1', 'BZ2', "ARRAY", `=ARRAYFORMULA(
        IF(C2:C<>"",
          IFERROR(VLOOKUP(BP2:BP, DATABASEMAHASISWA!A:E, 5, FALSE), "CEK NIM"),
          "DATA KOSONG"
        )
      )
      `],
    ['Form responses 1', 'CA2', "ARRAY", `=ARRAYFORMULA(
        IF(C2:C<>"",
          IFERROR(VLOOKUP(BP2:BP, DATABASEMAHASISWA!A:E, 3, FALSE), "CEK NIM"),
          "DATA KOSONG"
        )
      )
      `],
    ['Form responses 1', 'CB2', "ARRAY", `=ARRAYFORMULA(
        IF(C2:C<>"",
          IFERROR(VLOOKUP(BP2:BP, DATABASEMAHASISWA!A:E, 4, FALSE), "CEK NIM"),
          "DATA KOSONG"
        )
      )
      `],
    ['Form responses 1', 'CC2', "FILLDOWN", `=SUM(CD2,CE2)`],
    ['Form responses 1', 'CD2', "FILLDOWN", `=IFS(
        BZ2=TEXT("D3","@"), 400,
        BZ2=TEXT("D4","@"), 427,
        BZ2=TEXT("S1","@"), 427,
        BZ2=TEXT("S2","@"), 450,
        BZ2=TEXT("S3","@"), 475
      )`],
    ['Form responses 1', 'CE2', "ARRAY", `=ARRAYFORMULA(
    IF(CF2:CF="","",
      IF(CF2:CF="gris - S1", 73,
      IF(CF2:CF="gris - S2", 100,
      0)))
    )`],
    ['Form responses 1', 'CI2', "FILLDOWN", `=IF(NOT(ISBLANK(W2)), IF(REGEXMATCH(W2, "13\.00|13\.15|13\.30"), "AFT", "MOR"), IF(REGEXMATCH(R2, "13\.00|13\.15|13\.30"), "AFT", "MOR"))`],
    ['Form responses 1', 'CK2', "ARRAY", `=MAP(
        BX2:BX,
        BO2:BO,
        LAMBDA(bx, bo,
          IF(
            OR(bx="", bx="TIDAK DITEMUKAN, SILAKAN UPLOAD SKOR"),
            "",
            IF(
              OR(bo="", bo=0),
              "BELUM PERNAH TES",
              IF(
                bx >= bo + 100,
                "TOLONG MASUKKAN KE BLACKLIST",
                IF(
                  bx >= bo + 30,
                  "ANOMALI",
                  IF(
                    bx < bo - 10,
                    "DEGRADING",
                    IF(
                      bx < bo,
                      "FLUCTUATING",
                      "NORMAL"
                    )
                  )
                )
              )
            )
          )
        )
      )
      `],
    ['Form responses 1', 'CL2', "ARRAY", `=ARRAYFORMULA(
        IF(
          (BX2:BX="") + (BX2:BX="TIDAK DITEMUKAN, SILAKAN UPLOAD SKOR"),
          "",
          IFERROR(
            TEXT(
              VLOOKUP(
                BP2:BP,
                '02. CEKTESTHISTORY'!A:D,
                4,
                FALSE
              ),
              "yyyy-mm-dd"
            ),
            ""
          )
        )
      )
      `],
    ['Form responses 1', 'CM2', "ARRAY", `=ARRAYFORMULA(
        IF(
          (BX2:BX="") + (BX2:BX="TIDAK DITEMUKAN, SILAKAN UPLOAD SKOR"),
          "",
          IFERROR(
            TO_TEXT(
              VLOOKUP(
                BP2:BP,
                '02. CEKTESTHISTORY'!A:D,
                3,
                FALSE
              )
            ),
            ""
          )
        )
      )
      `],
    ['Form responses 1', 'CN2', "FILLDOWN", `=IF(C2<>"",
        IF(ISNA(VLOOKUP(BQ2, SINICOPYHASILSKOR!A:L, 11, FALSE)),
            "TIDAK DITEMUKAN, SILAKAN UPLOAD SKOR",
            VLOOKUP(BQ2, SINICOPYHASILSKOR!A:L, 11, FALSE)
        ),
        ""
        )
      `],


  // ====== OTHER SHEETS ======

  // 01. STATISTIK
    ['01. STATISTIK', 'A2', "ARRAY", `=LET(
      data, UNIQUE(FILTER('Form responses 1'!AO2:AO,
        (LEN('Form responses 1'!AO2:AO)>0) *
        (NOT(REGEXMATCH('Form responses 1'!AO2:AO,"BELUM PILIH JADWAL")))
      )),
      SORT(
        data,
        VALUE(REGEXEXTRACT(data, "\d+")), TRUE,
        REGEXEXTRACT(data, "^[A-Z]+"), TRUE
      )
    )
    `],
    ['01. STATISTIK', 'B2', "ARRAY", `=ARRAYFORMULA(
      IF(LEN(A2:A1000),
        COUNTIF('Form responses 1'!AO:AO, A2:A1000),
      )
    )
    `],
    ['01. STATISTIK', 'D2', "ARRAY",
      `=IFERROR(
          SORT(
            FILTER(
              {
                TEXT('Form responses 1'!A:A, "dd/mm/yyyy hh:mm:ss"),
                'Form responses 1'!AI:AI,
                'Form responses 1'!AJ:AJ,
                'Form responses 1'!AL:AL,
                'Form responses 1'!V:V,
                'Form responses 1'!W:W,
                'Form responses 1'!BX:BX
              },
              COUNTIF('Form responses 1'!AI:AI, 'Form responses 1'!AI:AI) > 1
            ),
            2, TRUE,
            3, TRUE
          ),
          "TIDAK ADA DUPLIKASI PADA KOLOM AI"
        )`
    ],
    ['01. STATISTIK', 'O2', "ARRAY",
      `=SORT(
          FILTER(
            {
              'Form responses 1'!A:A,
              'Form responses 1'!AI:AI,
              'Form responses 1'!AJ:AJ,
              'Form responses 1'!V:V,
              'Form responses 1'!W:W,
              'Form responses 1'!AL:AL,
              'Form responses 1'!AF:AF,
              'Form responses 1'!AE:AE,
              'Form responses 1'!BX:BX
            },
            'Form responses 1'!V:V = "YES"
          ),
          6, TRUE
        )`
    ],

  // 02. CEKTESTHISTORY

  // 03. KIRIM DATA KE PAK BIN H-1
    ['03. KIRIM DATA KE PAK BIN H-1', 'A2', "ARRAY", `=SORT(FILTER('Form responses 1'!AI2:AL, (NOT(ISNA('Form responses 1'!AL2:AL))) * ('Form responses 1'!AL2:AL <> "")), 4, TRUE)`],
    ['03. KIRIM DATA KE PAK BIN H-1', 'F2', "ARRAY", `=ARRAYFORMULA(
    LET(
        codeList, FILTER(UNIQUE(D2:D), UNIQUE(D2:D)<>""),
        rowCount, COUNTA(codeList) * 30,
        numList, MOD(SEQUENCE(rowCount,1,1,1)-1,30)+1,
        codeIdx, ROUNDUP(SEQUENCE(rowCount,1,1,1)/30),
        code, INDEX(codeList, codeIdx),
        "9" & code & TEXT(numList,"00")
    )
    )`],
    ['03. KIRIM DATA KE PAK BIN H-1', 'G2', "ARRAY", `=ARRAYFORMULA(
    LET(
        codeList, FILTER(UNIQUE(D2:D), UNIQUE(D2:D)<>""),
        rowCount, COUNTA(codeList) * 30,
        numList, MOD(SEQUENCE(rowCount,1,1,1)-1,30)+1,
        "ProTEFL Reserve " & TEXT(numList,"00")
    )
    )`],

  // 04. BUAT PRESENSI DAN GRUP WA H-1
    ['04. BUAT PRESENSI DAN GRUP WA H-1', 'B2', "ARRAY", `=SORT(FILTER({'Form responses 1'!B2:B, 'Form responses 1'!AI2:AI, 'Form responses 1'!AJ2:AJ, 'Form responses 1'!AT2:AT, 'Form responses 1'!AL2:AL, 'Form responses 1'!AO2:AO, 'Form responses 1'!BD2:BD},
    (NOT(ISNA('Form responses 1'!AL2:AL))) * ('Form responses 1'!AL2:AL <> "")),
    5, TRUE, 6, TRUE, 3, TRUE)`],
    ['04. BUAT PRESENSI DAN GRUP WA H-1', 'Q2', "ARRAY", `=FILTER(A2:E, COUNTIFS(C2:C, C2:C, D2:D, D2:D) > 1)`],

  // 05. DATASERTIFIKAT
    ['05. DATASERTIFIKAT', 'B2', "ARRAY", `=SORT(FILTER({
        'Form responses 1'!AL2:AL,  
        'Form responses 1'!B2:B,  
        IF(REGEXMATCH('Form responses 1'!BI2:BI, "_BERKALA"),
            MID('Form responses 1'!AI2:AI, 4, LEN('Form responses 1'!AI2:AI)-3),
            'Form responses 1'!AI2:AI
        ), 
        'Form responses 1'!K2:K,  
        'Form responses 1'!AJ2:AJ,  
        PROPER('Form responses 1'!M2:M),  
        SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(PROPER('Form responses 1'!N2:N),
            "Januari","January"), "Februari","February"), "Maret","March"), "April","April"), "Mei","May"), "Juni","June"), "Juli","July"), "Agustus","August"), "September","September"), "Oktober","October"),  
        IF(REGEXMATCH('Form responses 1'!BJ2:BJ, "GUGUR"), 'Form responses 1'!BJ2:BJ, TEXT('Form responses 1'!BJ2:BJ, "DD MMMM YYYY")),  
        'Form responses 1'!BU2:BU,  
        'Form responses 1'!BV2:BV,  
        'Form responses 1'!BW2:BW,  
        ('Form responses 1'!BU2:BU + 'Form responses 1'!BV2:BV + 'Form responses 1'!BW2:BW),  
        'Form responses 1'!BX2:BX,  
        'Form responses 1'!BY2:BY,  
        'Form responses 1'!CH2:CH  
        }, REGEXMATCH('Form responses 1'!AQ2:AQ, "T_")), 1, TRUE)`],

  // 06. UPLOADSKOR
    ['06. UPLOADSKOR', 'B2', "ARRAY", `=SORT(
        FILTER(
          {
            'Form responses 1'!AL2:AL,
            'Form responses 1'!BP2:BP,
            ARRAYFORMULA(SUBSTITUTE('Form responses 1'!AJ2:AJ, "'", "&#039;")),
            'Form responses 1'!BT2:BT,
            'Form responses 1'!BX2:BX,
            'Form responses 1'!BJ2:BJ,
            'Form responses 1'!BZ2:BZ,
            'Form responses 1'!CA2:CA,
            'Form responses 1'!CB2:CB,
            'Form responses 1'!CC2:CC,
            'Form responses 1'!CD2:CD,
            'Form responses 1'!CE2:CE,
            'Form responses 1'!CF2:CF
          },
          REGEXMATCH(
            TEXT(
              IF(
                REGEXMATCH('Form responses 1'!BI2:BI, "_BERKALA"),
                MID('Form responses 1'!AI2:AI, 4, LEN('Form responses 1'!AI2:AI)-3),
                IF(
                  REGEXMATCH('Form responses 1'!BI2:BI, "_OFFGRID"),
                  'Form responses 1'!D2:D,
                  'Form responses 1'!AI2:AI
                )
              ),
              "#"
            ),
            "^\\d{11}$"
          )
        ),
        1, TRUE, 7, TRUE, 2, TRUE, 5, TRUE
      )`],
  // 07. UPLOADSISTER
    ['07. UPLOADSISTER', 'A2', "ARRAY", `=SORT(
    FILTER(
        { 'Form responses 1'!BQ2:BQ,  
        IF(LEN('Form responses 1'!K2:K) > 11, 'Form responses 1'!K2:K, ""),  
        IF(LEN('Form responses 1'!K2:K) = 10, 'Form responses 1'!K2:K, ""),  
        'Form responses 1'!AJ2:AJ,  
        TEXT('Form responses 1'!BJ2:BJ, "yyyy"),  
        'Form responses 1'!CH2:CH,  
        'Form responses 1'!BJ2:BJ  
        },  
        ('Form responses 1'!C2:C = "ProTEFL TKBI/SERDOS/Umum (bersertifikat resmi diakui SISTER KEMENDIKBUDRISTEK)")
        * ((LEN(IF(LEN('Form responses 1'!K2:K) > 11, 'Form responses 1'!K2:K, "")) > 0) + (LEN(IF(LEN('Form responses 1'!K2:K) = 10, 'Form responses 1'!K2:K, "")) > 0) > 0)
    ),  
    1, TRUE
    )`],

  // db_export
    ['db_export', 'A2', "ARRAY",
      `=SORT(
        HSTACK(
          FILTER('Form responses 1'!CJ2:CJ, ('Form responses 1'!CJ2:CJ<>"") * NOT(REGEXMATCH(LOWER('Form responses 1'!CJ2:CJ),"discard"))),
          FILTER('Form responses 1'!BP2:BP, ('Form responses 1'!CJ2:CJ<>"") * NOT(REGEXMATCH(LOWER('Form responses 1'!CJ2:CJ),"discard"))),
          FILTER('Form responses 1'!AJ2:AJ, ('Form responses 1'!CJ2:CJ<>"") * NOT(REGEXMATCH(LOWER('Form responses 1'!CJ2:CJ),"discard"))),
          FILTER('Form responses 1'!CA2:CA, ('Form responses 1'!CJ2:CJ<>"") * NOT(REGEXMATCH(LOWER('Form responses 1'!CJ2:CJ),"discard"))),
          FILTER('Form responses 1'!CB2:CB, ('Form responses 1'!CJ2:CJ<>"") * NOT(REGEXMATCH(LOWER('Form responses 1'!CJ2:CJ),"discard"))),
          FILTER('Form responses 1'!BZ2:BZ, ('Form responses 1'!CJ2:CJ<>"") * NOT(REGEXMATCH(LOWER('Form responses 1'!CJ2:CJ),"discard"))),
          FILTER('Form responses 1'!BJ2:BJ, ('Form responses 1'!CJ2:CJ<>"") * NOT(REGEXMATCH(LOWER('Form responses 1'!CJ2:CJ),"discard"))),
          FILTER('Form responses 1'!CI2:CI, ('Form responses 1'!CJ2:CJ<>"") * NOT(REGEXMATCH(LOWER('Form responses 1'!CJ2:CJ),"discard"))),
          FILTER('Form responses 1'!AN2:AN, ('Form responses 1'!CJ2:CJ<>"") * NOT(REGEXMATCH(LOWER('Form responses 1'!CJ2:CJ),"discard"))),
          FILTER('Form responses 1'!BU2:BU, ('Form responses 1'!CJ2:CJ<>"") * NOT(REGEXMATCH(LOWER('Form responses 1'!CJ2:CJ),"discard"))),
          FILTER('Form responses 1'!BV2:BV, ('Form responses 1'!CJ2:CJ<>"") * NOT(REGEXMATCH(LOWER('Form responses 1'!CJ2:CJ),"discard"))),
          FILTER('Form responses 1'!BW2:BW, ('Form responses 1'!CJ2:CJ<>"") * NOT(REGEXMATCH(LOWER('Form responses 1'!CJ2:CJ),"discard"))),
          FILTER('Form responses 1'!BX2:BX, ('Form responses 1'!CJ2:CJ<>"") * NOT(REGEXMATCH(LOWER('Form responses 1'!CJ2:CJ),"discard"))),
          FILTER('Form responses 1'!BT2:BT, ('Form responses 1'!CJ2:CJ<>"") * NOT(REGEXMATCH(LOWER('Form responses 1'!CJ2:CJ),"discard"))),
          FILTER(
            IF('Form responses 1'!CG2:CG="", 0, 'Form responses 1'!CG2:CG),
            ('Form responses 1'!CJ2:CJ<>"") * NOT(REGEXMATCH(LOWER('Form responses 1'!CJ2:CJ),"discard"))
          ),
          FILTER('Form responses 1'!BX2:BX,
            ('Form responses 1'!CJ2:CJ<>"") *
            NOT(REGEXMATCH(LOWER('Form responses 1'!CJ2:CJ),"discard"))
          )
          -
          FILTER(
            IF('Form responses 1'!CG2:CG="", 0, 'Form responses 1'!CG2:CG),
            ('Form responses 1'!CJ2:CJ<>"") *
            NOT(REGEXMATCH(LOWER('Form responses 1'!CJ2:CJ),"discard"))
          )
        ),
      1, TRUE)`
    ],
  // 08. DATAKUITANSI
    ['08. DATAKUITANSI', 'A2', "ARRAY", `=SORT(
        FILTER({
          IF('Form responses 1'!AI2:AI<>"", "__", ""),
          IF('Form responses 1'!AI2:AI<>"",
            IFERROR(
              MATCH(
                LOWER(LEFT(TRIM('Form responses 1'!AB2:AB), SEARCH(" ", TRIM('Form responses 1'!AB2:AB)&" ")-1)),
                {"januari";"februari";"maret";"april";"mei";"juni";"juli";"agustus";"september";"oktober";"november";"desember"},
                0
              ),
              ""
            ),
            ""
          ),
          IF('Form responses 1'!AI2:AI<>"", "__", ""),
          'Form responses 1'!AJ2:AJ,
          'Form responses 1'!AT2:AT,
          'Form responses 1'!AI2:AI,
          IF('Form responses 1'!AI2:AI<>"", "", ""),
          IF('Form responses 1'!AI2:AI<>"", "", ""),
          IF('Form responses 1'!AI2:AI<>"", "", ""),
          IF('Form responses 1'!AI2:AI<>"", "", ""),
          IF('Form responses 1'!AI2:AI<>"", "", ""),
          IF('Form responses 1'!AI2:AI<>"", "", ""),
          IF('Form responses 1'!AI2:AI<>"", "", "")
        },
        'Form responses 1'!AI2:AI<>""),
        1, TRUE
      )`],
    ];
  


  function clearExistingFormulas(sheet, cellA1, type) {
    var anchorRange = sheet.getRange(cellA1);
    var startRow = anchorRange.getRow();
    var col = anchorRange.getColumn();
    var lastRow = sheet.getMaxRows();
  
    var range = sheet.getRange(startRow, col, lastRow - startRow + 1);
    var formulas = range.getFormulas();
  
    for (var r = 0; r < formulas.length; r++) {
      for (var c = 0; c < formulas[r].length; c++) {
        if (formulas[r][c]) range.getCell(r + 1, c + 1).clearContent();
      }
    }
  }
    
  
// ---------------------------------------------------------------------------
// applyAllFormulas()
// Loops through the central FORMULAS_TO_APPLY configuration and ensures
// each formula is inserted in its correct sheet and cell. Supports two types:
//   - ARRAY: sets a single ARRAYFORMULA at the given anchor cell
//   - FILLDOWN: fills a formula down column rows based on a key column
// ---------------------------------------------------------------------------
function applyAllFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  FORMULAS_TO_APPLY.forEach(row => {
    var [sheetName, cellA1, type, formula, keyCol] = row;
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    // Clear old formulas first
    clearExistingFormulas(sheet, cellA1, type);

    // Apply formula
    if (type === "ARRAY") {
      setFormulaOnce(sheet, cellA1, formula);
    } else if (type === "FILLDOWN") {
      fillDownFormula(sheet, cellA1, formula, keyCol || 3);
    }
  });
}



// ============================================================================
// File: autoCounters.gs
//
//  Purpose:
//   - Protect "Original Schedule" column (R) so only the owner can edit.
//   - Automatically log reschedules in column X and count them in column Y.
//   - Provide utilities to sync or reset reschedule counters.
//
// Target Sheet: "Form responses 1"
//
// Column Mapping:
//   C (3)  - Name / Identifier (row in use check)
//   R (18) - Original Schedule (protected column)
//   V (22) - Reschedule Flag ("Yes" or blank)
//   W (23) - New Schedule
//   X (24) - Reschedule Log (semicolon-separated history of W)
//   Y (25) - Reschedule Count (number of entries in X)
//
// Workflow:
//   1. Run protectOriginalScheduleColumn() once or periodically to ensure
//      column R stays locked for everyone except the owner.
//   2. Run installRescheduleTrigger() once to set up the onEdit trigger
//      that watches V/W/X changes and updates logs/counts.
//   3. onEditLogReschedule(e) runs automatically whenever a relevant edit
//      happens in "Form responses 1".
//   4. If logs and counts ever go out of sync, run syncRescheduleCounts()
//      to recalculate counts in column Y from the log in column X.
//
// Notes:
//   - This script is safe to re-run: protections and triggers are cleared
//     and re-installed cleanly.
//   - Reschedule log (X) stores history; last entry reflects the most recent
//     schedule in W when V == "Yes".
// ============================================================================

// ---------------------------------------------------------------------------
// installRescheduleTrigger()
// Installs the onEdit trigger for reschedule logging.
// Removes old triggers first, then creates a fresh one.
// ---------------------------------------------------------------------------
function installRescheduleTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Remove any existing triggers for this function
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "onEditLogReschedule") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  // Create a fresh one
  ScriptApp.newTrigger("onEditLogReschedule")
    .forSpreadsheet(ss)
    .onEdit()
    .create();
}

// ---------------------------------------------------------------------------
// protectOriginalScheduleColumn()
// Protects column R (Original Schedule).
// Ensures only the sheet owner can edit it.
// ---------------------------------------------------------------------------
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
  
  
// ---------------------------------------------------------------------------
// onEditLogReschedule(e)
// Triggered on sheet edits.
// - If V="Yes", appends W into log (X) and updates count (Y).
// - If V!="Yes", clears log (X) and count (Y).
// - If X is edited manually, recounts Y.
// ---------------------------------------------------------------------------
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
  
// ---------------------------------------------------------------------------
// syncRescheduleCounts()
// Syncs all counts in Y with logs in X.
// Run onOpen or periodically to fix inconsistencies.
// ---------------------------------------------------------------------------
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


// ============================================================================
// File: styling.gs
//
// Purpose:
//   Provides consistent visual formatting across all target sheets in the
//   ProTEFL registration workbook. This script centralizes color-banding,
//   header styling, and text contrast logic so that the UI is both readable
//   and visually structured for administrators.
//
// Key Features:
//   - Applies bold styling to header rows.
//   - Resets old formatting before applying new themes.
//   - Adds alternating header/body color bands to "Form responses 1" using
//     predefined palettes for easier scanning of wide tables.
//   - Automatically adjusts font color (black or white) for readability based
//     on background luminance.
//   - Supports special palette overrides for specific column bands.
//
// How it works:
//   1. A list of target sheets (`STYLING_TARGET_SHEETS`) defines where styling
//      should be applied.
//   2. For "Form responses 1", column ranges (`FORM_RESPONSES_1_COLOR_BANDS`)
//      are styled in banded color palettes (`COLOR_PALETTES`).
//   3. Helper functions:
//        - `colAtoNum(colA)`: Converts "A", "Z", "AA", etc. to numeric indices.
//        - `getAutoFontColor(bg)`: Chooses black/white text for readability.
//   4. The main function `applyAllStyling()` orchestrates the formatting pass.
//      It can be called manually or as part of the admin `main()` initializer.
//
// Usage:
//   - Run `applyAllStyling()` directly for visual refresh.
//   - Or, let `main()` handle it automatically when setting up the workbook.
// ============================================================================


// ---------------------------------------------------------------------------
// Target sheets where styling will be applied
// ---------------------------------------------------------------------------
const STYLING_TARGET_SHEETS = [
  'Form responses 1',
  // Add more sheet names as needed...
];


// ---------------------------------------------------------------------------
// Column bands in "Form responses 1" to receive alternating color schemes
// ---------------------------------------------------------------------------
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
  'CH-CK'
];


// ---------------------------------------------------------------------------
// Color palette pairs (dark = header, light = body)
// Rotated if there are more bands than palettes
// ---------------------------------------------------------------------------
const COLOR_PALETTES = [
  {header:'#1565c0', body:'#90caf9'},    // blue
  {header:'#2e7d32', body:'#a5d6a7'},    // green
  {header:'#ad1457', body:'#f8bbd0'},    // pink
  {header:'#6d4c41', body:'#bcaaa4'},    // brown
  {header:'#ff8f00', body:'#ffe082'},    // amber
  {header:'#c62828', body:'#ef9a9a'},    // red
  {header:'#4527a0', body:'#b39ddb'},    // purple
  {header:'#00838f', body:'#80deea'},    // teal
  {header:'#607d8b', body:'#cfd8dc'},    // blue grey
  {header:'#689f38', body:'#dcedc8'},    // lime
];


// ---------------------------------------------------------------------------
// Emphasis palettes (darker variants)
// ---------------------------------------------------------------------------
const EMPHASIS_PALETTES = [
  {header:'#0d47a1', body:'#1565c0'},
  {header:'#1b5e20', body:'#2e7d32'},
  {header:'#880e4f', body:'#ad1457'},
  {header:'#3e2723', body:'#6d4c41'},
  {header:'#e65100', body:'#ff8f00'},
  {header:'#8e0000', body:'#c62828'},
  {header:'#311b92', body:'#4527a0'},
  {header:'#004d40', body:'#00838f'},
  {header:'#37474f', body:'#607d8b'},
  {header:'#33691e', body:'#689f38'},
];


// ---------------------------------------------------------------------------
// Columns to emphasize
// ---------------------------------------------------------------------------
const EMPHASIS_COLUMNS = [
  'V', 'W', 'AF', 'AG',
  'AP',
  'AX',
  'BI'
];

// ---------------------------------------------------------------------------
// Helper: colAtoNum()
// Converts a column label in A1 notation (e.g. "BZ") to a numeric index
// Example: "A" -> 1, "Z" -> 26, "AA" -> 27, "BZ" -> 78
// ---------------------------------------------------------------------------
function colAtoNum(colA) {
  let n = 0;
  for (let i = 0; i < colA.length; i++) {
    n = n * 26 + (colA.charCodeAt(i) - 64); // ASCII 'A' = 65
  }
  return n;
}


// ---------------------------------------------------------------------------
// Helper: getAutoFontColor()
// Given a hex background color "#rrggbb", returns either black ("#212121")
// or white ("#ffffff") based on luminance for readability
// ---------------------------------------------------------------------------
function getAutoFontColor(bg) {
  if (!bg || !bg.match(/^#[0-9a-f]{6}$/i)) return "#000000";
  let r = parseInt(bg.substr(1,2),16);
  let g = parseInt(bg.substr(3,2),16);
  let b = parseInt(bg.substr(5,2),16);
  let luma = 0.2126*r + 0.7152*g + 0.0722*b; // relative luminance
  return luma < 150 ? "#ffffff" : "#212121";
}


// ---------------------------------------------------------------------------
// Emphasis layer (runs AFTER banding)
// ---------------------------------------------------------------------------
function applyFormResponses1Emphasis(sheet) {
  const lastRow = Math.max(2, sheet.getLastRow());

  FORM_RESPONSES_1_COLOR_BANDS.forEach((band, idx) => {
    let [colStart, colEnd] = band.split('-').map(colAtoNum);

    let dark, light;
    if (band === "AS-AY") {
      dark = "#1b5e20";
      light = "#2e7d32";
    } else {
      const palette = EMPHASIS_PALETTES[idx % EMPHASIS_PALETTES.length];
      dark = palette.header;
      light = palette.body;
    }

    let headerFont = getAutoFontColor(dark);
    let bodyFont = getAutoFontColor(light);

    EMPHASIS_COLUMNS.forEach(colA => {
      const col = colAtoNum(colA);

      if (col >= colStart && col <= colEnd) {

        // Header
        sheet.getRange(1, col)
             .setBackground(dark)
             .setFontColor(headerFont);

        // Body
        if (lastRow > 1) {
          sheet.getRange(2, col, lastRow - 1)
               .setBackground(light)
               .setFontColor(bodyFont)
               .setFontWeight("bold");
        }
      }
    });
  });
}


// ---------------------------------------------------------------------------
// applyAllStyling()
// Applies all styling rules to target sheets. 
// For "Form responses 1", it clears prior formatting and applies color bands
// to specific column ranges, using palettes defined above.
// ---------------------------------------------------------------------------
function applyAllStyling() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  STYLING_TARGET_SHEETS.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    // Bold entire header row
    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    headerRange.setFontWeight("bold");

    if (sheetName === 'Form responses 1') {
      const lastRow = Math.max(2, sheet.getLastRow());
      const lastCol = sheet.getLastColumn();

      // Reset all background & font color before re-applying palettes
      sheet.getRange(1, 1, lastRow, lastCol)
           .setBackground(null)
           .setFontColor("#212121");

      // Base band coloring
      FORM_RESPONSES_1_COLOR_BANDS.forEach((band, idx) => {
        let [colStart, colEnd] = band.split('-').map(colAtoNum);

        // --- Special palette override for AS-AY (force green)
        let dark, light;
        if (band === "AS-AY") {
          dark = "#2e7d32";
          light = "#a5d6a7";
        } else {
          const palette = COLOR_PALETTES[idx % COLOR_PALETTES.length];
          dark = palette.header;
          light = palette.body;
        }

        // Pick best contrasting font color
        let headerFont = getAutoFontColor(dark);
        let bodyFont = getAutoFontColor(light);

        // Apply to header
        sheet.getRange(1, colStart, 1, colEnd - colStart + 1)
             .setBackground(dark)
             .setFontColor(headerFont);

        // Apply to data rows
        if (lastRow > 1) {
          sheet.getRange(2, colStart, lastRow - 1, colEnd - colStart + 1)
               .setBackground(light)
               .setFontColor(bodyFont);
        }
      });

      // Emphasis columns (AFTER bands)
      applyFormResponses1Emphasis(sheet);
    }
  });
}

// ============================================================================
// LAZY SOLUTIONS FOR BUGGY IMPLEMENTATIONS
// NO PENNY NO EPIPHANY
// ============================================================================

// ============================================================================
// Function: fixColumnCD
// Description: Fills down column CD in "Form responses 1" with IFS formula
//              if Fakultas and Prodi exist but minimum score is missing.
// ============================================================================
function fixColumnCD() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Form responses 1");
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet "Form responses 1" not found!');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("No data rows found to fill formula.");
    return;
  }

  const formula = `=IFS(
    BZ2:BZ="D3",400,
    BZ2:BZ="D4",427,
    BZ2:BZ="S1",427,
    BZ2:BZ="S2",450,
    BZ2:BZ="S3",475
  )`;

  // Set formula in CD2
  sheet.getRange("CD2").setFormula(formula);

  // Auto-fill down from CD2 to last row
  sheet.getRange("CD2").autoFill(sheet.getRange("CD2:CD" + lastRow), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  SpreadsheetApp.getUi().alert(`Column CD filled with IFS formula down to row ${lastRow}`);
}


// ============================================================================
// File: externalFormImport.gs
//
// Purpose:
//   Safely import and synchronize raw Google Form response data from
//   Form responses 2, Form responses 3, Form responses 4, etc. into
//   Form responses 1.
//
// Main principle:
//   - Source sheets remain the source of truth for raw form data A:U.
//   - Form responses 1 remains the canonical processing sheet.
//   - New source rows are inserted as WHOLE SHEET ROWS in Form responses 1.
//   - Existing imported rows are updated in-place when source A:U changes.
//   - Processing/helper columns V onward are never manually shifted apart from A:U.
//   - Phone numbers in column F are preserved as text, including leading zero.
//
//
// Recommended workflow:
//   1. Replace the old externalFormImport.gs content with this version.
//   2. Reload the spreadsheet.
//   3. Run External Form Import > 00. Authorize External Import Access.
//   4. Run External Form Import > 02. Sync External Form Responses.
//   5. Keep the 5-minute trigger installed for reconciliation.
//
// ============================================================================


// ============================================================================
// CONFIG
// ============================================================================

const MDMA_EXTERNAL_IMPORT = {
  // Canonical processing sheet.
  // All your existing utilities, formulas, views, grouping, attendance, etc.
  // still point here.
  TARGET_SHEET_NAME: "Form responses 1",

  // Master/orchestrator sheet.
  MASTER_SHEET_NAME: "00. MASTER-DATA",

  // A:U = 21 raw Google Form response columns.
  RAW_START_COL: 1,
  RAW_COL_COUNT: 21,

  // Column F = WhatsApp/phone column in the raw form area.
  // This must remain text. Otherwise 0812... becomes 812...
  PHONE_COL: 6,

  // Other raw columns that are safer as plain text.
  // D is often NIM/NIK; F is phone/test contact.
  TEXT_SAFE_RAW_COLS: [4, 6],

  // Registry in 00. MASTER-DATA begins at E.
  REGISTRY_HEADER_ROW: 1,
  REGISTRY_START_ROW: 2,
  REGISTRY_START_COL: 5,

  // E:M registry.
  REGISTRY_HEADERS: [
    "Include?",
    "Source sheet",
    "Header row",
    "Data start row",
    "Raw end col",
    "Last sync",
    "Last imported target row",
    "Status",
    "Imported last run"
  ],

  // Source-side audit/marker columns.
  // v2 appends MDMA_LAST_SYNCED_AT while preserving old v1 headers.
  SOURCE_MARKER_HEADERS: [
    "MDMA_IMPORTED",
    "MDMA_IMPORT_ID",
    "MDMA_IMPORTED_AT",
    "MDMA_TARGET_ROW",
    "MDMA_IMPORT_HASH",
    "MDMA_LAST_SYNCED_AT"
  ],

  // Target-side hidden metadata columns.
  // v2 appends MDMA_LAST_SOURCE_SYNC_AT while preserving old v1 headers.
  TARGET_META_HEADERS: [
    "MDMA_SOURCE_SHEET",
    "MDMA_SOURCE_ROW",
    "MDMA_IMPORT_ID",
    "MDMA_IMPORTED_AT",
    "MDMA_IMPORT_HASH",
    "MDMA_LAST_SOURCE_SYNC_AT"
  ]
};


// ============================================================================
// AUTHORIZATION / PREFLIGHT
// ============================================================================

/**
 * Run once from the menu.
 *
 * This intentionally touches the services used by this importer:
 * - SpreadsheetApp: read/write sheets
 * - ScriptApp: create triggers
 * - LockService: prevent concurrent imports
 * - PropertiesService: harmless preflight
 * - Utilities: hashing imported rows
 *
 * It does not use a new external API.
 */
function mdmaAuthorizeExternalImportAccess() {
  const ui = SpreadsheetApp.getUi();

  const authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);

  if (authInfo.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.REQUIRED) {
    ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);
    return;
  }

  // Harmless touches / sanity checks.
  SpreadsheetApp.getActiveSpreadsheet().getId();
  ScriptApp.getProjectTriggers();
  LockService.getDocumentLock();
  PropertiesService.getDocumentProperties().getProperties();
  Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, "ok");

  ui.alert(
    "Authorization Complete ✅",
    "External form import and source-sync permissions are ready.",
    ui.ButtonSet.OK
  );
}


// ============================================================================
// MENU-FACING FUNCTIONS
// ============================================================================

/**
 * Manual sync button.
 *
 * This now does BOTH:
 *   1. Append new source rows.
 *   2. Update existing imported rows if source A:U changed.
 */
function mdmaSyncExternalFormResponsesWithUi() {
  const ui = SpreadsheetApp.getUi();

  try {
    const result = mdmaSyncExternalFormResponses_();

    ui.alert(
      "External Form Sync Complete ✅",
      [
        `Imported rows: ${result.imported}`,
        `Updated rows: ${result.updated}`,
        `Unchanged rows checked: ${result.unchanged}`,
        "",
        ...result.messages
      ].join("\n"),
      ui.ButtonSet.OK
    );

  } catch (err) {
    ui.alert(
      "External Form Sync Failed ❌",
      err && err.stack ? err.stack : String(err),
      ui.ButtonSet.OK
    );
  }
}


/**
 * Adds Form responses 2, Form responses 3, etc. into the registry.
 *
 * It does NOT import yet.
 * It only writes source-sheet names into 00. MASTER-DATA E:M.
 */
function mdmaSeedExternalFormRegistry() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = mdmaEnsureImportRegistry_();

  const existingConfigs = mdmaReadImportConfigs_();
  const existingNames = new Set(existingConfigs.map(c => c.sourceSheetName));

  const sourceSheets = ss.getSheets()
    .map(sheet => sheet.getName())
    .filter(name => /^Form responses \d+$/i.test(name))
    .filter(name => name !== MDMA_EXTERNAL_IMPORT.TARGET_SHEET_NAME)
    .filter(name => !existingNames.has(name));

  if (sourceSheets.length === 0) {
    SpreadsheetApp.getUi().alert(
      "No New Source Sheets",
      "No unregistered Form responses 2/3/etc. sheets were found.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const rows = sourceSheets.map(name => [
    true,   // Include?
    name,   // Source sheet
    1,      // Header row
    2,      // Data start row
    "U",    // Raw end col
    "",     // Last sync
    "",     // Last imported target row
    "Seeded",
    ""      // Imported last run
  ]);

  const startRow = Math.max(
    master.getLastRow() + 1,
    MDMA_EXTERNAL_IMPORT.REGISTRY_START_ROW
  );

  master
    .getRange(
      startRow,
      MDMA_EXTERNAL_IMPORT.REGISTRY_START_COL,
      rows.length,
      MDMA_EXTERNAL_IMPORT.REGISTRY_HEADERS.length
    )
    .setValues(rows);

  SpreadsheetApp.getUi().alert(
    "Registry Seeded ✅",
    `Added ${rows.length} source sheet(s) to 00. MASTER-DATA.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
 * Installs importer triggers.
 *
 * 1. onFormSubmit trigger:
 *    Runs when any linked form submits into this spreadsheet.
 *
 * 2. time-driven trigger:
 *    Runs every 5 minutes.
 *    This is the reconciliation safety net.
 *
 * Why both?
 *   Apps Script triggers are not something I would trust as the ONLY safeguard.
 *   The real safety comes from idempotent sync:
 *   every run scans source sheets and imports only unimported rows.
 */
function mdmaInstallExternalImportTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Remove old importer triggers first to avoid duplicates.
  ScriptApp.getProjectTriggers().forEach(trigger => {
    const fn = trigger.getHandlerFunction();

    if (
      fn === "mdmaOnExternalFormSubmit_" ||
      fn === "mdmaSyncExternalFormResponsesTrigger_"
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("mdmaOnExternalFormSubmit_")
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger("mdmaSyncExternalFormResponsesTrigger_")
    .timeBased()
    .everyMinutes(5)
    .create();

  SpreadsheetApp.getUi().alert(
    "External Import Triggers Installed ✅",
    "Importer will run on form submit and every 5 minutes as reconciliation backup.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
 * Optional cleanup.
 *
 * Deletes only this importer's triggers.
 * Does not touch your existing styling/reschedule/default-view triggers.
 */
function mdmaRemoveExternalImportTriggers() {
  let removed = 0;

  ScriptApp.getProjectTriggers().forEach(trigger => {
    const fn = trigger.getHandlerFunction();

    if (
      fn === "mdmaOnExternalFormSubmit_" ||
      fn === "mdmaSyncExternalFormResponsesTrigger_"
    ) {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });

  SpreadsheetApp.getUi().alert(
    "External Import Triggers Removed",
    `Removed ${removed} importer trigger(s).`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


// ============================================================================
// TRIGGER HANDLERS
// ============================================================================

/**
 * Installable on-form-submit trigger handler.
 *
 * Do not run manually.
 */
function mdmaOnExternalFormSubmit_(e) {
  mdmaSyncExternalFormResponses_();
}


/**
 * Time-driven reconciliation trigger handler.
 *
 * Do not run manually.
 */
function mdmaSyncExternalFormResponsesTrigger_() {
  mdmaSyncExternalFormResponses_();
}


// ============================================================================
// CORE SYNC ENGINE
// ============================================================================

/**
 * Core sync function.
 *
 * Behavior:
 *   - For each registered source row:
 *       a. If never imported, append it into Form responses 1.
 *       b. If already imported and source A:U changed, update target A:U.
 *       c. If unchanged, do nothing.
 *
 * Important:
 *   - Appends use whole-row insertion.
 *   - Updates only touch A:U of the known target row.
 *   - Before writing, A:U target cells are formatted as text-safe where needed.
 */
function mdmaSyncExternalFormResponses_() {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(30000)) {
    throw new Error("Could not obtain document lock. Another import may still be running.");
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const target = ss.getSheetByName(MDMA_EXTERNAL_IMPORT.TARGET_SHEET_NAME);

    if (!target) {
      throw new Error(`Target sheet not found: ${MDMA_EXTERNAL_IMPORT.TARGET_SHEET_NAME}`);
    }

    const master = mdmaEnsureImportRegistry_();
    const now = new Date();

    // Make sure target has enough columns and metadata columns.
    mdmaEnsureMinColumns_(target, MDMA_EXTERNAL_IMPORT.RAW_COL_COUNT);

    const targetMetaCols = mdmaEnsureHeaderColumns_(
      target,
      1,
      MDMA_EXTERNAL_IMPORT.TARGET_META_HEADERS,
      true
    );

    const targetIndexes = mdmaBuildTargetIndexes_(target, targetMetaCols);
    const configs = mdmaReadImportConfigs_();

    if (configs.length === 0) {
      return {
        imported: 0,
        updated: 0,
        unchanged: 0,
        messages: [
          "No registered source sheets found.",
          "Run Seed External Form Registry first, then check 00. MASTER-DATA columns E:M."
        ]
      };
    }

    const appendEntries = [];
    const updateEntries = [];
    const messages = [];
    const statsByRegistryRow = {};
    let unchanged = 0;

    configs.forEach(config => {
      if (!config.include) return;

      statsByRegistryRow[config.registryRow] = {
        imported: 0,
        updated: 0,
        unchanged: 0
      };

      if (config.sourceSheetName === MDMA_EXTERNAL_IMPORT.TARGET_SHEET_NAME) {
        messages.push(`Skipped ${config.sourceSheetName}: target sheet cannot be a source.`);
        mdmaWriteRegistryStatus_(master, config.registryRow, "Skipped: target sheet", 0, "");
        return;
      }

      const source = ss.getSheetByName(config.sourceSheetName);

      if (!source) {
        messages.push(`Missing source sheet: ${config.sourceSheetName}`);
        mdmaWriteRegistryStatus_(master, config.registryRow, "Error: source sheet missing", 0, "");
        return;
      }

      mdmaEnsureMinColumns_(source, MDMA_EXTERNAL_IMPORT.RAW_COL_COUNT);

      const sourceLastRow = mdmaGetLastDataRowByColumns_(
        source,
        MDMA_EXTERNAL_IMPORT.RAW_START_COL,
        MDMA_EXTERNAL_IMPORT.RAW_COL_COUNT
      );

      if (sourceLastRow < config.dataStartRow) {
        messages.push(`${config.sourceSheetName}: no data rows.`);
        mdmaWriteRegistryStatus_(master, config.registryRow, "Synced: no data rows", 0, "");
        return;
      }

      const markerCols = mdmaEnsureHeaderColumns_(
        source,
        config.headerRow,
        MDMA_EXTERNAL_IMPORT.SOURCE_MARKER_HEADERS,
        false
      );

      const numRows = sourceLastRow - config.dataStartRow + 1;

      // Read both raw values and display values.
      // Values preserve Dates where useful.
      // Display values preserve phone leading zero and exact text shown by the sheet.
      const sourceValues = source
        .getRange(
          config.dataStartRow,
          MDMA_EXTERNAL_IMPORT.RAW_START_COL,
          numRows,
          MDMA_EXTERNAL_IMPORT.RAW_COL_COUNT
        )
        .getValues();

      const sourceDisplayValues = source
        .getRange(
          config.dataStartRow,
          MDMA_EXTERNAL_IMPORT.RAW_START_COL,
          numRows,
          MDMA_EXTERNAL_IMPORT.RAW_COL_COUNT
        )
        .getDisplayValues();

      const markerValues = source
        .getRange(
          config.dataStartRow,
          markerCols["MDMA_IMPORTED"],
          numRows,
          MDMA_EXTERNAL_IMPORT.SOURCE_MARKER_HEADERS.length
        )
        .getValues();

      sourceValues.forEach((valueRow, index) => {
        const displayRow = sourceDisplayValues[index];
        const sourceRowNumber = config.dataStartRow + index;

        const sourceRawRow = mdmaNormalizeRawRowForWrite_(valueRow, displayRow);

        if (!mdmaRowHasData_(sourceRawRow)) {
          return;
        }

        const sourceHash = mdmaHashRow_(sourceRawRow);

        const marker = mdmaReadMarkerRow_(markerValues[index], markerCols);
        const stableImportId = mdmaStableImportId_(ss, source, sourceRowNumber);
        const legacyImportId = marker.importId || "";

        const targetRow = mdmaResolveTargetRow_(
          target,
          targetIndexes,
          config.sourceSheetName,
          sourceRowNumber,
          stableImportId,
          legacyImportId,
          marker.targetRow
        );

        if (targetRow) {
          const oldHash = marker.importHash || mdmaGetTargetHash_(target, targetMetaCols, targetRow);

          if (oldHash === sourceHash) {
            unchanged++;
            statsByRegistryRow[config.registryRow].unchanged++;

            // Repair marker/import ID if needed, without touching target raw data.
            // Disabled during unchanged sync so the 5-minute trigger does not write
            // source markers repeatedly when nothing actually changed.
            return;
          }

          updateEntries.push({
            mode: "updateTarget",
            sourceSheetName: config.sourceSheetName,
            sourceSheet: source,
            sourceRowNumber,
            markerCols,
            targetRow,
            sourceRawRow,
            sourceHash,
            stableImportId,
            legacyImportId,
            registryRow: config.registryRow
          });

          statsByRegistryRow[config.registryRow].updated++;
          return;
        }

        appendEntries.push({
          sourceSheetName: config.sourceSheetName,
          sourceSheet: source,
          sourceRowNumber,
          markerCols,
          sourceRawRow,
          sourceHash,
          stableImportId,
          legacyImportId,
          registryRow: config.registryRow
        });

        statsByRegistryRow[config.registryRow].imported++;
      });
    });

    // ------------------------------------------------------------------------
    // APPLY UPDATES FIRST
    // ------------------------------------------------------------------------
    let updated = 0;

    updateEntries.forEach(entry => {
      if (entry.mode === "repairMarkerOnly") {
        mdmaWriteSourceMarker_(
          entry.sourceSheet,
          entry.sourceRowNumber,
          entry.markerCols,
          true,
          entry.stableImportId,
          now,
          entry.targetRow,
          entry.sourceHash
        );
        return;
      }

      // Preserve phone/NIM text format before writing.
      mdmaPrepareRawTargetRowForWrite_(target, entry.targetRow, 1);

      target
        .getRange(
          entry.targetRow,
          MDMA_EXTERNAL_IMPORT.RAW_START_COL,
          1,
          MDMA_EXTERNAL_IMPORT.RAW_COL_COUNT
        )
        .setValues([entry.sourceRawRow]);

      mdmaWriteTargetMeta_(
        target,
        targetMetaCols,
        entry.targetRow,
        entry.sourceSheetName,
        entry.sourceRowNumber,
        entry.stableImportId,
        now,
        entry.sourceHash
      );

      mdmaWriteSourceMarker_(
        entry.sourceSheet,
        entry.sourceRowNumber,
        entry.markerCols,
        true,
        entry.stableImportId,
        now,
        entry.targetRow,
        entry.sourceHash
      );

      updated++;
    });

    // ------------------------------------------------------------------------
    // APPLY APPENDS WITH WHOLE-ROW INSERTION
    // ------------------------------------------------------------------------
    let imported = 0;
    let firstInsertedRow = null;
    let lastInsertedRow = null;

    if (appendEntries.length > 0) {
      const lastTargetRawRow = Math.max(
        1,
        mdmaGetLastDataRowByColumns_(
          target,
          MDMA_EXTERNAL_IMPORT.RAW_START_COL,
          MDMA_EXTERNAL_IMPORT.RAW_COL_COUNT
        )
      );

      // CRITICAL: insert whole spreadsheet rows.
      target.insertRowsAfter(lastTargetRawRow, appendEntries.length);

      firstInsertedRow = lastTargetRawRow + 1;
      lastInsertedRow = firstInsertedRow + appendEntries.length - 1;

      mdmaPrepareRawTargetRowForWrite_(target, firstInsertedRow, appendEntries.length);

      target
        .getRange(
          firstInsertedRow,
          MDMA_EXTERNAL_IMPORT.RAW_START_COL,
          appendEntries.length,
          MDMA_EXTERNAL_IMPORT.RAW_COL_COUNT
        )
        .setValues(appendEntries.map(entry => entry.sourceRawRow));

      target
        .getRange(
          firstInsertedRow,
          targetMetaCols["MDMA_SOURCE_SHEET"],
          appendEntries.length,
          MDMA_EXTERNAL_IMPORT.TARGET_META_HEADERS.length
        )
        .setValues(
          appendEntries.map((entry, index) => [
            entry.sourceSheetName,
            entry.sourceRowNumber,
            entry.stableImportId,
            now,
            entry.sourceHash,
            now
          ])
        );

      appendEntries.forEach((entry, index) => {
        const targetRow = firstInsertedRow + index;

        mdmaWriteSourceMarker_(
          entry.sourceSheet,
          entry.sourceRowNumber,
          entry.markerCols,
          true,
          entry.stableImportId,
          now,
          targetRow,
          entry.sourceHash
        );
      });

      imported = appendEntries.length;
    }

    // ------------------------------------------------------------------------
    // UPDATE REGISTRY STATUS
    // ------------------------------------------------------------------------
    Object.entries(statsByRegistryRow).forEach(([registryRowText, stat]) => {
      const registryRow = Number(registryRowText);
      const totalTouched = stat.imported + stat.updated;

      mdmaWriteRegistryStatus_(
        master,
        registryRow,
        `Synced: +${stat.imported}, updated ${stat.updated}, unchanged ${stat.unchanged}`,
        totalTouched,
        lastInsertedRow || ""
      );
    });

    // Refresh formula-managed/helper columns after inserted rows exist.
    // Your current project already defines applyAllFormulas().
    // However, applyAllFormulas() is intentionally not called here because it
    // touches ARRAY formulas and existing formula ranges. New imports only need
    // FILLDOWN formulas applied to the newly inserted target rows.
    mdmaApplyFilldownFormulasForInsertedRows_(firstInsertedRow, lastInsertedRow);

    SpreadsheetApp.flush();

    const importedBySource = mdmaCountBy_(appendEntries, "sourceSheetName");
    const updatedBySource = mdmaCountBy_(
      updateEntries.filter(entry => entry.mode === "updateTarget"),
      "sourceSheetName"
    );

    Object.entries(importedBySource).forEach(([sourceName, count]) => {
      messages.push(`${sourceName}: imported ${count} new row(s).`);
    });

    Object.entries(updatedBySource).forEach(([sourceName, count]) => {
      messages.push(`${sourceName}: updated ${count} existing target row(s).`);
    });

    if (firstInsertedRow && lastInsertedRow) {
      messages.push(`Inserted target rows: ${firstInsertedRow}–${lastInsertedRow}`);
    }

    if (messages.length === 0) {
      messages.push("No new or changed source rows found.");
    }

    return {
      imported,
      updated,
      unchanged,
      messages
    };

  } finally {
    lock.releaseLock();
  }
}


// ============================================================================
// REGISTRY HELPERS
// ============================================================================

/**
 * Ensures 00. MASTER-DATA has the importer registry headers at E:M.
 */
function mdmaEnsureImportRegistry_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let master = ss.getSheetByName(MDMA_EXTERNAL_IMPORT.MASTER_SHEET_NAME);

  if (!master) {
    master = ss.insertSheet(MDMA_EXTERNAL_IMPORT.MASTER_SHEET_NAME);
  }

  mdmaEnsureMinColumns_(
    master,
    MDMA_EXTERNAL_IMPORT.REGISTRY_START_COL + MDMA_EXTERNAL_IMPORT.REGISTRY_HEADERS.length - 1
  );

  master
    .getRange(
      MDMA_EXTERNAL_IMPORT.REGISTRY_HEADER_ROW,
      MDMA_EXTERNAL_IMPORT.REGISTRY_START_COL,
      1,
      MDMA_EXTERNAL_IMPORT.REGISTRY_HEADERS.length
    )
    .setValues([MDMA_EXTERNAL_IMPORT.REGISTRY_HEADERS])
    .setFontWeight("bold");

  return master;
}


/**
 * Reads 00. MASTER-DATA E:M registry.
 */
function mdmaReadImportConfigs_() {
  const master = mdmaEnsureImportRegistry_();
  const lastRow = master.getLastRow();

  if (lastRow < MDMA_EXTERNAL_IMPORT.REGISTRY_START_ROW) {
    return [];
  }

  const numRows = lastRow - MDMA_EXTERNAL_IMPORT.REGISTRY_START_ROW + 1;

  const values = master
    .getRange(
      MDMA_EXTERNAL_IMPORT.REGISTRY_START_ROW,
      MDMA_EXTERNAL_IMPORT.REGISTRY_START_COL,
      numRows,
      MDMA_EXTERNAL_IMPORT.REGISTRY_HEADERS.length
    )
    .getValues();

  return values
    .map((row, index) => {
      const registryRow = MDMA_EXTERNAL_IMPORT.REGISTRY_START_ROW + index;

      return {
        registryRow,
        include: mdmaTruthy_(row[0]),
        sourceSheetName: String(row[1] || "").trim(),
        headerRow: Number(row[2] || 1),
        dataStartRow: Number(row[3] || 2),
        rawEndCol: mdmaColumnValueToNumber_(row[4]) || MDMA_EXTERNAL_IMPORT.RAW_COL_COUNT
      };
    })
    .filter(config => config.sourceSheetName);
}


/**
 * Writes status back to registry row.
 *
 * Registry columns:
 * E Include?
 * F Source sheet
 * G Header row
 * H Data start row
 * I Raw end col
 * J Last sync
 * K Last imported target row
 * L Status
 * M Imported last run
 */
function mdmaWriteRegistryStatus_(master, registryRow, status, importedCount, lastTargetRow) {
  const now = new Date();

  const lastSyncCol = MDMA_EXTERNAL_IMPORT.REGISTRY_START_COL + 5;       // J
  const lastTargetRowCol = MDMA_EXTERNAL_IMPORT.REGISTRY_START_COL + 6;  // K
  const statusCol = MDMA_EXTERNAL_IMPORT.REGISTRY_START_COL + 7;         // L
  const importedLastRunCol = MDMA_EXTERNAL_IMPORT.REGISTRY_START_COL + 8;// M

  master.getRange(registryRow, lastSyncCol).setValue(now);

  if (lastTargetRow) {
    master.getRange(registryRow, lastTargetRowCol).setValue(lastTargetRow);
  }

  master.getRange(registryRow, statusCol).setValue(status);
  master.getRange(registryRow, importedLastRunCol).setValue(importedCount);
}


// ============================================================================
// SYNC INDEX / RESOLUTION HELPERS
// ============================================================================

/**
 * Builds lookup maps from target metadata.
 */
function mdmaBuildTargetIndexes_(target, targetMetaCols) {
  const lastRow = Math.max(target.getLastRow(), 2);
  const numRows = lastRow - 1;

  const indexes = {
    byImportId: new Map(),
    bySourceKey: new Map(),
    byTargetRow: new Set()
  };

  if (numRows <= 0) return indexes;

  const startCol = targetMetaCols["MDMA_SOURCE_SHEET"];
  const metaValues = target
    .getRange(2, startCol, numRows, MDMA_EXTERNAL_IMPORT.TARGET_META_HEADERS.length)
    .getValues();

  metaValues.forEach((row, i) => {
    const targetRow = i + 2;
    const sourceSheetName = String(row[0] || "").trim();
    const sourceRow = String(row[1] || "").trim();
    const importId = String(row[2] || "").trim();

    if (importId) {
      indexes.byImportId.set(importId, targetRow);
    }

    if (sourceSheetName && sourceRow) {
      indexes.bySourceKey.set(`${sourceSheetName}::${sourceRow}`, targetRow);
    }

    indexes.byTargetRow.add(targetRow);
  });

  return indexes;
}


/**
 * Resolves where a source row lives in Form responses 1.
 *
 * Lookup order:
 *   1. Source marker MDMA_TARGET_ROW.
 *   2. Target metadata by source sheet + source row.
 *   3. Target metadata by stable v2 import ID.
 *   4. Target metadata by legacy v1 import ID.
 */
function mdmaResolveTargetRow_(
  target,
  targetIndexes,
  sourceSheetName,
  sourceRowNumber,
  stableImportId,
  legacyImportId,
  markerTargetRow
) {
  const markerRow = Number(markerTargetRow || 0);

  if (markerRow >= 2 && markerRow <= target.getMaxRows()) {
    return markerRow;
  }

  const sourceKey = `${sourceSheetName}::${sourceRowNumber}`;

  if (targetIndexes.bySourceKey.has(sourceKey)) {
    return targetIndexes.bySourceKey.get(sourceKey);
  }

  if (targetIndexes.byImportId.has(stableImportId)) {
    return targetIndexes.byImportId.get(stableImportId);
  }

  if (legacyImportId && targetIndexes.byImportId.has(legacyImportId)) {
    return targetIndexes.byImportId.get(legacyImportId);
  }

  return null;
}


function mdmaStableImportId_(ss, sourceSheet, sourceRowNumber) {
  return [
    ss.getId(),
    sourceSheet.getSheetId(),
    sourceRowNumber
  ].join("::");
}


function mdmaReadMarkerRow_(markerRowValues, markerCols) {
  const headers = MDMA_EXTERNAL_IMPORT.SOURCE_MARKER_HEADERS;
  const out = {};

  headers.forEach((header, index) => {
    out[header] = markerRowValues[index];
  });

  return {
    imported: mdmaTruthy_(out["MDMA_IMPORTED"]),
    importId: String(out["MDMA_IMPORT_ID"] || "").trim(),
    importedAt: out["MDMA_IMPORTED_AT"],
    targetRow: out["MDMA_TARGET_ROW"],
    importHash: String(out["MDMA_IMPORT_HASH"] || "").trim(),
    lastSyncedAt: out["MDMA_LAST_SYNCED_AT"]
  };
}


function mdmaGetTargetHash_(target, targetMetaCols, targetRow) {
  const col = targetMetaCols["MDMA_IMPORT_HASH"];
  if (!col) return "";

  return String(target.getRange(targetRow, col).getValue() || "").trim();
}


// ============================================================================
// WRITE HELPERS
// ============================================================================

/**
 * Applies text format to sensitive raw columns before writing.
 *
 * This fixes the phone-number issue:
 *   0812345678 should remain 0812345678,
 *   not become 812345678.
 */
function mdmaPrepareRawTargetRowForWrite_(target, startRow, numRows) {
  MDMA_EXTERNAL_IMPORT.TEXT_SAFE_RAW_COLS.forEach(col => {
    target
      .getRange(startRow, col, numRows, 1)
      .setNumberFormat("@");
  });
}


function mdmaWriteTargetMeta_(
  target,
  targetMetaCols,
  targetRow,
  sourceSheetName,
  sourceRowNumber,
  importId,
  now,
  sourceHash
) {
  target
    .getRange(
      targetRow,
      targetMetaCols["MDMA_SOURCE_SHEET"],
      1,
      MDMA_EXTERNAL_IMPORT.TARGET_META_HEADERS.length
    )
    .setValues([[
      sourceSheetName,
      sourceRowNumber,
      importId,
      now,
      sourceHash,
      now
    ]]);
}


function mdmaWriteSourceMarker_(
  sourceSheet,
  sourceRowNumber,
  markerCols,
  imported,
  importId,
  now,
  targetRow,
  sourceHash
) {
  const startCol = markerCols["MDMA_IMPORTED"];

  sourceSheet
    .getRange(
      sourceRowNumber,
      startCol,
      1,
      MDMA_EXTERNAL_IMPORT.SOURCE_MARKER_HEADERS.length
    )
    .setValues([[
      imported,
      importId,
      now,
      targetRow,
      sourceHash,
      now
    ]]);
}


/**
 * Applies only FILLDOWN formulas to newly inserted rows in Form responses 1.
 *
 * This avoids applyAllFormulas(), which is too heavy because it loops through
 * all formula configs, clears existing formulas, and also touches ARRAY formulas.
 *
 * New imports only need row-level formulas inserted into the new rows so their
 * references match the row number.
 */
function mdmaApplyFilldownFormulasForInsertedRows_(firstInsertedRow, lastInsertedRow) {
  if (!firstInsertedRow || !lastInsertedRow) {
    Logger.log("Skipped FILLDOWN refresh: no inserted rows.");
    return;
  }

  if (typeof FORMULAS_TO_APPLY === "undefined") {
    Logger.log("Skipped FILLDOWN refresh: FORMULAS_TO_APPLY is not defined.");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  FORMULAS_TO_APPLY.forEach(config => {
    const [sheetName, cellA1, type, baseFormula, keyCol] = config;

    if (sheetName !== MDMA_EXTERNAL_IMPORT.TARGET_SHEET_NAME) return;
    if (type !== "FILLDOWN") return;

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const anchorRange = sheet.getRange(cellA1);
    const formulaCol = anchorRange.getColumn();
    const anchorRow = anchorRange.getRow();
    const startRow = Math.max(firstInsertedRow, anchorRow);
    const endRow = lastInsertedRow;
    const checkCol = keyCol || 3;

    if (endRow < startRow) return;

    const rowCount = endRow - startRow + 1;
    const keyValues = sheet.getRange(startRow, checkCol, rowCount, 1).getValues();

    const formulas = keyValues.map((row, index) => {
      const actualRow = startRow + index;
      const keyValue = row[0];

      if (keyValue === "" || keyValue === null) {
        return [""];
      }

      const formula = mdmaShiftFilldownFormulaToRow_(baseFormula, actualRow);

      return [formula];
    });

    sheet
      .getRange(startRow, formulaCol, rowCount, 1)
      .setFormulas(formulas);
  });
}


/**
 * Converts row-2 references in a FILLDOWN base formula into the target row.
 *
 * Examples:
 *   A2    -> A57
 *   $A2   -> $A57
 *   A$2   -> A57
 *   $A$2  -> $A57
 *
 * This intentionally targets spreadsheet-style cell references only, not every
 * random number 2 in the formula.
 */
function mdmaShiftFilldownFormulaToRow_(baseFormula, targetRow) {
  return String(baseFormula).replace(
    /(\$?[A-Z]{1,3})\$?2(?!\d)/g,
    function(match, colRef) {
      return colRef + targetRow;
    }
  );
}


// ============================================================================
// RAW VALUE NORMALIZATION
// ============================================================================

/**
 * Builds the row that will be written into Form responses 1 A:U.
 *
 * Most columns use getValues(), so Dates remain Dates.
 * Text-sensitive columns use getDisplayValues(), so phone/NIM formatting is
 * preserved exactly as visible in the source sheet.
 */
function mdmaNormalizeRawRowForWrite_(valueRow, displayRow) {
  const output = mdmaPadOrTrimRow_(valueRow, MDMA_EXTERNAL_IMPORT.RAW_COL_COUNT);

  MDMA_EXTERNAL_IMPORT.TEXT_SAFE_RAW_COLS.forEach(colNumber => {
    const index = colNumber - 1;
    output[index] = mdmaNormalizeTextCell_(displayRow[index]);
  });

  // Extra hardening for phone column F.
  const phoneIndex = MDMA_EXTERNAL_IMPORT.PHONE_COL - 1;
  output[phoneIndex] = mdmaNormalizePhone_(displayRow[phoneIndex]);

  return output;
}


/**
 * Removes only the display apostrophe if it is visible in the source display.
 *
 * In Google Sheets, an apostrophe used to force text is usually not returned by
 * getDisplayValues(). If it is visible for some reason, this removes it while
 * preserving the leading 0.
 */
function mdmaNormalizeTextCell_(value) {
  let text = String(value || "");

  if (text.startsWith("'")) {
    text = text.slice(1);
  }

  return text.trim();
}


function mdmaNormalizePhone_(value) {
  let text = mdmaNormalizeTextCell_(value);

  // Remove spaces and common separators without destroying the leading zero.
  text = text.replace(/[\s\-().]/g, "");

  // If a form/user stored Indonesian number as +628..., convert to 08...
  // because your existing WA formulas expect the local 0 prefix.
  if (/^\+?62\d+/.test(text)) {
    text = text.replace(/^\+?62/, "0");
  }

  // If the source lost the local leading zero and starts with 8...,
  // restore it so WA/message formulas keep working.
  if (/^8\d+/.test(text)) {
    text = "0" + text;
  }

  return text;
}


// ============================================================================
// GENERAL HELPERS
// ============================================================================

function mdmaEnsureMinColumns_(sheet, minCols) {
  const current = sheet.getMaxColumns();

  if (current < minCols) {
    sheet.insertColumnsAfter(current, minCols - current);
  }
}


function mdmaEnsureHeaderColumns_(sheet, headerRow, headers, hideColumns) {
  const result = {};

  headers.forEach(header => {
    const lastCol = Math.max(sheet.getLastColumn(), 1);

    const existingHeaders = sheet
      .getRange(headerRow, 1, 1, lastCol)
      .getValues()[0]
      .map(value => String(value || "").trim());

    const existingIndex = existingHeaders.indexOf(header);

    if (existingIndex !== -1) {
      const col = existingIndex + 1;
      result[header] = col;

      if (hideColumns) {
        sheet.hideColumns(col);
      }

      return;
    }

    sheet.insertColumnAfter(lastCol);

    const newCol = lastCol + 1;
    sheet.getRange(headerRow, newCol).setValue(header);

    result[header] = newCol;

    if (hideColumns) {
      sheet.hideColumns(newCol);
    }
  });

  return result;
}


/**
 * Finds the last row with data in a specific column range.
 *
 * This intentionally checks only A:U, not the whole sheet.
 * Reason:
 *   helper formulas far to the right may make getLastRow() unreliable.
 */
function mdmaGetLastDataRowByColumns_(sheet, startCol, numCols) {
  const lastRow = Math.max(sheet.getLastRow(), 1);

  const values = sheet
    .getRange(1, startCol, lastRow, numCols)
    .getValues();

  for (let r = values.length - 1; r >= 0; r--) {
    if (values[r].some(value => value !== "" && value !== null)) {
      return r + 1;
    }
  }

  return 1;
}


function mdmaPadOrTrimRow_(row, width) {
  const output = row.slice(0, width);

  while (output.length < width) {
    output.push("");
  }

  return output;
}


/**
 * True when the row has at least one non-empty cell.
 */
function mdmaRowHasData_(row) {
  return row.some(value => value !== "" && value !== null);
}


/**
 * Accepts TRUE/Yes/1/etc. from the registry.
 */
function mdmaTruthy_(value) {
  if (value === true) return true;

  const text = String(value || "").trim().toLowerCase();

  return ["true", "yes", "y", "1", "include", "included"].includes(text);
}


/**
 * Converts:
 *   "U"  -> 21
 *   "21" -> 21
 *   21   -> 21
 */
function mdmaColumnValueToNumber_(value) {
  if (value === "" || value === null || typeof value === "undefined") {
    return null;
  }

  if (typeof value === "number") {
    return value;
  }

  const text = String(value).trim().toUpperCase();

  if (/^\d+$/.test(text)) {
    return Number(text);
  }

  return mdmaColumnLetterToNumber_(text);
}


/**
 * Column letter to number.
 * Uses a unique function name to avoid clashing with your existing letterToColumn_().
 */
function mdmaColumnLetterToNumber_(letter) {
  let col = 0;

  for (let i = 0; i < letter.length; i++) {
    col = col * 26 + (letter.charCodeAt(i) - 64);
  }

  return col;
}


/**
 * Creates a SHA-256 hash of A:U source row values.
 *
 * This helps identify duplicate or repeated import attempts.
 */
function mdmaHashRow_(row) {
  const normalized = row.map(value => {
    if (value instanceof Date) {
      return Utilities.formatDate(
        value,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd'T'HH:mm:ss.SSS"
      );
    }

    return String(value || "").trim();
  });

  const raw = JSON.stringify(normalized);

  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    raw
  );

  return digest
    .map(byte => {
      const value = byte < 0 ? byte + 256 : byte;
      return ("0" + value.toString(16)).slice(-2);
    })
    .join("");
}


function mdmaCountBy_(items, key) {
  return items.reduce((acc, item) => {
    const value = item[key] || "Unknown";
    acc[value] = (acc[value] || 0) + 1;
    return acc;
  }, {});
}
