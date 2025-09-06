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
  initializeSheets();              // Create sheets and populate headers/templates
  setupAllDropdowns();             // Add dropdown validations
  protectOriginalScheduleColumn(); // Lock the "Original Schedule" column (R)
  applyAllStyling();               // Apply header fonts, widths, colors
  applyAllFormulas();              // Insert all ARRAY/FILLDOWN formulas
  setupDefaultViewTrigger();       // Ensure Default View trigger is installed
  installRescheduleTrigger();      // Ensure reschedule auto-counter trigger
}

// ============================================================================
// MENU SETUP
// Builds the "ProTEFL Utility" custom menu with safe options, exports, risky
// admin actions, and quick-access custom views.
// Runs automatically on spreadsheet open.
// ============================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("ProTEFL Utility")
      // --- Safe options ---
      .addItem("Apply Styles", "applyAllStylingWithConfirm")
      .addItem("Protect Original Schedule Column", "protectOriginalScheduleColumn")
      .addItem("Set Up AutoCounter Trigger", "setupAutoCounterTriggerWithAlert")
      .addSeparator()
      .addSubMenu(
        SpreadsheetApp.getUi()
          .createMenu("Export")
          .addItem("Participant Test IDs", "exportParticipantTestIds")
          .addItem("Download VCF by Tanggal Tes", "downloadVCFFromMenu")
          .addItem("Copy Attendance List", "copyAttendanceList")
          .addItem("Export Participant Scores", "exportSiakadScoreResults")
      )
      .addSeparator()
      // --- Risky options ---
      .addItem("Apply All Formulas (Danger Zone)", "applyAllFormulasWithConfirm")
      .addItem("Initialize Sheet (Danger Zone)", "runMainWithConfirm")
      .addSeparator()
      // --- Custom views ---
      .addSubMenu(
        SpreadsheetApp.getUi()
          .createMenu("Custom View")
          .addItem("Default View", "toggleDefaultView")
          .addItem("Reschedule Participants", "toggleRescheduleParticipantsView")
          .addItem("Verify Student ID", "toggleVerifyStudentIDView")
          .addItem("Verify Payment", "toggleVerifyPaymentView")
          .addItem("Verify Attendance", "toggleVerifyAttendanceView")
          .addItem("Grouping & Contacts", "toggleGroupingContactsView")
      )
    .addToUi();

  toggleDefaultView(true); // Always open default view on launch
}

// ============================================================================
// MENU ACTION WRAPPERS
// Safe prompts before executing styling, formula injection, or initialization.
// Prevents accidental destructive changes.
// ============================================================================

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

// ============================================================================
// TRIGGER MANAGEMENT
// Installs or refreshes installable triggers for auto counter logging
// and opening the default view.
// ============================================================================

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

// ======================
// CUSTOM VIEWS (Optimized, Reliable Toggle)
// ======================
function applyCustomView_(sheetName, keepCols, sidebarFn, label, forceOn) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return;

  var props = PropertiesService.getDocumentProperties();
  var currentView = props.getProperty("currentView") || "";
  var keepIndexes = keepCols.map(letterToColumn_);
  var lastCol = sheet.getLastColumn();

  var activateView = forceOn || currentView !== label;

  // Show all first
  sheet.showColumns(1, lastCol);

  if (activateView) {
    // Hide columns not in keepCols
    var rangesToHide = [];
    var start = null;
    for (var col = 1; col <= lastCol; col++) {
      if (!keepIndexes.includes(col)) {
        if (start === null) start = col;
      } else {
        if (start !== null) {
          rangesToHide.push([start, col - start]);
          start = null;
        }
      }
    }
    if (start !== null) rangesToHide.push([start, lastCol - start + 1]);
    rangesToHide.forEach(r => sheet.hideColumns(r[0], r[1]));

    if (sidebarFn) sidebarFn();
    props.setProperty("currentView", label);
  } else {
    // Deactivating view → show all
    props.setProperty("currentView", "");
  }

  // Install Default View trigger if applicable
  if (label === "Default") setupDefaultViewTrigger();
}

// === Individual View Functions ===
function toggleDefaultView(forceOn) {
  var keepCols = ["A","AI","AJ","AN","AO","BB","BC","BJ","BT","BX"];
  applyCustomView_("Form responses 1", keepCols, showDefaultSidebar, "Default", forceOn);
}

function toggleRescheduleParticipantsView() {
  var keepCols = ["A","C","D","E","G","R","V","W","X","Y","AE","AF","AG","AH","AL","AM","AN","AO","BI"];
  applyCustomView_("Form responses 1", keepCols, showRescheduleSidebar, "Reschedule Participants");
}

function toggleVerifyStudentIDView() {
  var keepCols = ["C","D","E","AZ","BA","BB","BC"];
  applyCustomView_("Form responses 1", keepCols, showVerifyStudentIDSidebar, "Verify Student ID");
}

function toggleVerifyPaymentView() {
  // Columns to keep visible: A, G, AS-AY, BI
  var keepCols = ["A", "G", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "BI"];
  applyCustomView_("Form responses 1", keepCols, showVerifyPaymentSidebar, "Verify Payment");
}

function toggleVerifyAttendanceView() {
  var keepCols = [
    "A","C","D","G","V","W","AI","AJ","AL","AN","AO",
    "BC","BI","BJ","BL","BN","BQ","BS",
    "BU","BV","BW","BX","CB","CG"
  ];
  applyCustomView_("Form responses 1", keepCols, showVerifyAttendanceSidebar, "Verify Attendance");
}

function toggleGroupingContactsView() {
  const keepCols = ["A", "AI", "AJ", "AL", "AM", "AN", "AO", "AP", "AQ", "BE", "BG", "BI", "BJ", "CI"];
  applyCustomView_("Form responses 1", keepCols, showGroupingContactsSidebar, "Grouping & Contacts");
}



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



// ======================
// UTILITY
// ======================
function letterToColumn_(letter) {
  var col = 0;
  for (var i = 0; i < letter.length; i++) col = col * 26 + (letter.charCodeAt(i) - 64);
  return col;
}

//EXP
// ======================
// DOWNLOAD VCF
// ======================
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

// ======================
// EXPORT PARTICIPANT TEST IDS TO EXCEL
// ======================
function exportParticipantTestIds() {
  const ui = SpreadsheetApp.getUi();
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

  const data = sheet.getDataRange().getValues();
  const header = data.shift();
  const dateColIndex = header.indexOf("Kode Masuk Tes ProTEFL");
  const targetCols = ["AI","AJ","AK","AL"].map(letterToColumn_);

  if (dateColIndex === -1) return ui.alert("Test Date column not found.");

  const filtered = data.filter(row => String(row[dateColIndex]) === dateFilter);
  const exportData = filtered.length === 0 ? [] : [targetCols.map(i => header[i-1])];
  filtered.forEach(row => exportData.push(targetCols.map(i => row[i-1])));

  // Inline HTML dialog (styled like VCF modal)
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

  ui.showModalDialog(HtmlService.createHtmlOutput(htmlContent).setWidth(460).setHeight(250), "Export Participant Test IDs");
}
//EXP


// ======================
// SIDEBARS (Optimized)
// ======================

function showDefaultSidebar() {
  const html = `
  <!DOCTYPE html>
  <html>
    <head>
      <meta charset="UTF-8">
      <title><span class="material-icons">storage</span>ProTEFL MDMA</title>
      <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">
      <!-- Google Sans -->
      <link href="https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;700&display=swap" rel="stylesheet">
      <style>
        body {
          font-family: 'Google Sans', Arial, sans-serif;
          margin: 0;
          padding: 16px;
          background: #edf2fa;   /* sidebar background */
          color: #222;
        }

        h2 { margin-top:0; color:#1a1a1a; }
        h3 { margin-top:12px; color:#333; }

        /* Card styling */
        .card {
          background: #d3e3fd;
          border-radius: 10px;
          box-shadow: 0 2px 5px rgba(0,0,0,0.15);
          padding: 12px 16px;
          margin-bottom: 12px;
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
        }

        .card-header .arrow-icon {
          font-size: 26px;       /* big arrow */
          margin-right: 8px;
          transition: transform 0.4s ease;
          color: #3a3a3a;
        }

        .card-header .section-icon {
          font-size: 20px;
          margin-right: 6px;
          color: #1e88e5;       /* blue icon accent */
        }

        .card-content {
          margin-top: 8px;
          color: #333;
        }

        ul { margin: 0; padding-left: 18px; }
        li { margin-bottom: 4px; }

        .footer-note { color:#555; font-size:12px; margin-top:16px; }
        a { color:#1e88e5; text-decoration:none; }
        a:hover { text-decoration:underline; }
      </style>
    </head>
    <body>
      <h2>Welcome to ProTEFL MDMA</h2>
      <p><i>(ProTEFL Monthly Data Management Admin)</i></p>
      <p>It's ProTEFL but on Speed ⚡</p>

      <!-- Registration Card -->
      <div class="card">
        <div class="card-header" onclick="toggleCollapse(this)">
          <span class="arrow-icon material-icons">expand_more</span>
          <span class="section-icon material-icons">assignment</span>
          Registration
        </div>
        <div class="card-content">
          <ul>
            <li>Google Forms Entry</li>
            <li>Manual Entry (menu planned)</li>          
          </ul>
        </div>
      </div>

      <!-- Data Management Card -->
      <div class="card">
        <div class="card-header" onclick="toggleCollapse(this)">
          <span class="arrow-icon material-icons">expand_more</span>
          <span class="section-icon material-icons">settings</span>
          Data Management
        </div>
        <div class="card-content">
          <ul>
            <li>Participant(s) Rescheduling (Before Test)</li>
            <li>Student ID Verification</li>
            <li>Manual Test Count Checking (menu planned)</li>
            <li>Automatic & Override Option of Test Group Plotting (menu planned)</li>
            <li>Contact Creation (VCF) (menu planned)</li>
            <li>Autogenerated Attendance & Test ID Lists (menu planned)</li>
          </ul>
        </div>
      </div>

      <!-- Scoring Card -->
      <div class="card">
        <div class="card-header" onclick="toggleCollapse(this)">
          <span class="arrow-icon material-icons">expand_more</span>
          <span class="section-icon material-icons">assessment</span>
          Scoring
        </div>
        <div class="card-content">
          <ul>
            <li>Attendance Verification & Reschedule Flagging (After Test)</li>
            <li>Score Checking</li>
            <li>Reschedule Offering (same as in Data Management)</li>
            <li>Autogenerated Score Report format</li>
            <li>Autogenerated Certificate Data Format</li>
            <li>Autogenerated SISTER Upload Format (obsolete)</li>
          </ul>
        </div>
      </div>

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
        PS. The title is obviously inspired by Andy Field way of naming his books. 
        I mean, "Discovering statistics using IBM SPSS statistics: and **x and d**** and rock 'n' roll" ...what a legend.
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

  SpreadsheetApp.getUi()
    .showSidebar(HtmlService.createHtmlOutput(html).setTitle("ProTEFL MDMA"));
}

function showRescheduleSidebar() {
  const html = `
    <div style="font-family:Arial,sans-serif;padding:16px;line-height:1.5;color:#222;">
      <h2 style="margin-top:0;">📋 Reschedule Participants Guide</h2>
      <p>Here’s a go-to workflow for rescheduling participants:</p>
      <ol style="padding-left:18px;">
        <li>Locate the participant’s <b>Name</b> in column <b>E</b>.</li>
        <li>Verify their <b>Original Schedule</b> in column <b>R</b>. It is crucial if they registered multiple times. In that case, be careful. Make sure you reschedule the correct entry.</li>
        <li>In column <b>V</b>, set the dropdown to <b>Yes</b> to flag for reschedule. This will revoke their original schedule. They won't have a schedule now. Column AL will now be empty.</li>
        <li>To assign them new schedule, search for the new schedule date in <b>00. MASTER-DATA</b> in accordance to participant's choosing.</li>
        <li>Copy the suitable schedule from <b>00. MASTER-DATA</b> into <b>Form responses 1</b> in column <b>W</b>.</li>
        <li>Mark <b>Confirmed</b> in column <b>AG</b> to lock it in.</li>
        <li>Copy the WhatsApp message from column <b>AH</b> and send it to the participant. 🚀</li>
      </ol>
      <p style="margin-top:12px; font-size:12px; color:#555;">
        Tip: Accuracy beats speed here — double-check before hitting send! With accoubtability, you have avoided complaint(s) induced headache and hypertension.
      </p>
    </div>
  `;
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(html).setTitle("Reschedule Participants"));
}

function showVerifyStudentIDSidebar() {
  const html = `
    <div style="font-family:Arial,sans-serif;padding:16px;line-height:1.5;color:#222;">
      <h2 style="margin-top:0;">🆔 Verify Student ID Guide</h2>
      <p>Student ID verification is critical — mismatched IDs mean scores won't appear on SIAKAD. This is achieved with the assumption that entries in <b>DATABASEMAHASISWA</b> has the correct student data.</p>
      <h3 style="margin-top:12px;">Step-by-step check:</h3>
      <ol style="padding-left:18px;">
        <li>Check column <b>BC</b> (Status):</li>
        <ul>
          <li><b>COCOK</b>: ✅ Everything matches — move on to the next participant.</li>
          <li><b>CEK NAMA</b>: Minor capitalization mismatch. No fix needed here; we already use corrected proper names. Reference <b>06. UPLOADSKOR</b> for tidy names (say thanks Windi right now 😒).</li>
          <li><b>SALAH NIM</b>: Name in column <b>E</b> or <b>BA (duplicates of E)</b> doesn’t match the database (<b> shown in BB</b>). Ask the participant for their ID card and update NIM in <b>E</b> ONLY. Data shown elsewhere are all duplicates of E.</li>
          <li><b>#N/A</b>: No match found. Investigate and resolve manually. Ask the students for their KTM, write the correct NIM. When issues persist, it means we do not have their data in DATABASEMAHASISWA. Please update it manually based on the data on their KTM. Usually happens for students registering as INTAKE students (course begining on February).</li>
        </ul>
      </ol>
      <p style="margin-top:12px; font-size:12px; color:#555;">
        Pro tip: Careful checking now saves a flood of complaints later. 👍
      </p>
    </div>
  `;
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(html).setTitle("Verify Student ID"));
}

function showVerifyPaymentSidebar() {
  const html = `
    <div style="font-family:Arial,sans-serif;padding:16px;line-height:1.5;color:#222;">
      <h2 style="margin-top:0;">💰 Verify Payment Quick Guide</h2>
      <p>This view is for verifying test taker payments — this keeps ULB overlord(s) happy!</p>

      <h3>Online payment via transfer:</h3>
      <ol style="padding-left:18px;">
        <li>Check the <b>Bukti Bayar</b> attachment in column <b>AU</b>.</li>
        <li>Verify: is it authentic? Not fake? Matches participant? </li>
        <li>If everything is ✅, select <b>LUNAS</b> in column <b>AX</b>.</li>
        <li>If any issue arises, select the other status(es) in accordance with the problem.</li>
        <li>Done! Move on to the next participant.</li>
      </ol>

      <h3>Manual payment (e.g. LURING / on-demand):</h3>
      <ol style="padding-left:18px;">
        <li>Ensure the participant received their proof of payment / kuitansi / receipt.</li>
        <li>Search their name in column <b>AS</b>.</li>
        <li>Copy the <b>Nomor Ujian</b> from their receipt into column <b>G</b>. Ignore other text like D4, S1, S2, S3 — overwrite them, those are just placeholder (I am too lazy to restructure the whole Google Form structure after all these formulas and magic).</li>
        <li>Important: write <b>_OFFGRID</b> in column <b>BI</b>. This forces the workbook to use the receipt’s <b>Nomor Ujian</b> instead of the default NIM. Why? To make sure that non-paying registrants cannot sneak in/log in to ProTEFL SEB using their NIM.</li>
      </ol>

      <p style="margin-top:12px; font-size:12px; color:#555;">
        Pro tip: Always double-check attachments or make sure you write the correct Nomor Ujian to avoid complaints later on. ⚡
      </p>
    </div>
  `;
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(html).setTitle("Verify Payment"));
}

function showVerifyAttendanceSidebar() {
  var htmlContent = `
    <div style="font-family:Arial, sans-serif; padding:16px; line-height:1.5;">
      <h2 style="margin-top:0;">📊 Verify Attendance & Score</h2>
      <p><i>Use this view to verify attendance and score checking. This is by far the most time-consuming part (god I wish I got paid extra for this).</i></p>

      <h3>Step 0: For sanity’s sake</h3>
      <p>Enable filter by date: look at <b>BJ</b> and select a single date. Trust me, your sanity will thank you.</p>

      <h3>Step 1: Prepare</h3>
      <p>You need to check attendance report from proctors (in another sheet, sadly). Use split window view for best productivity—one side the attendance sheet, another side this sheet.</p>

      <h3>Step 2: Import Scores</h3>
      <p>Copy the scores into this workbook in <b>SINICOPYHASILSKOR</b> and do the necessary formatting. Make sure column <b>A</b> on <b>SINICOPYHASILSKOR</b> matches <b>BQ</b> in <b>Form responses 1</b> (this sheet). Then, in <b>SINICOPYHASILSKOR</b> copy test ID in P, write the appropriate kode masuk in Q, and make sure R has the formula "=(Q2 & "-" & P2)" so on; and A has the formula "=R2" and so on, drag them down. The scores will then appear across <b>BU to BY</b> in Form responses 1.</p>

      <p>Disclaimer: this works under the assumption that the data you copy into SINICOPYHASILSKOR is pristine and no tes IDs are misplaced, replaced, moved from their original cells. If there are errors, that's on you. Congrats you just messed up an entire results of that day tests and maybe others. Now cry and curl up in the corner!</p>

      <h3>Step 3: Check for missing scores</h3>
      <p>If no score appears, there are three possibilities:</p>
      <ol style="padding-left:18px;">
        <li>
          <b>Did not attend:</b> mark reschedule on <b>V</b> to Yes, write placeholder to <b>W</b>. We will ask them later using template message link in <b>AE</b>. This revokes their registration on this date; no data in <b>SINICOPYHASILSKOR</b> will link to any test ID.
        </li>
        <li>
          <b>Used Akun Cadangan:</b> copy akun cadangan to <b>G</b>, write <b>_OFFGRID</b> to <b>BI</b>, and check if scores appear on <b>BU-BX</b>.
        </li>
        <li>
          <b>NIM mismatch:</b> mismatch between <b>D</b> and whatever test ID they used in <b>SINICOPYHASILSKOR</b>. Resolve by checking their used ID, refer to proctor notes, and do step two above. 
          Is their NIM not matching? Check <b>BC</b> for <b>CEK NAMA</b>. Still no score? Confirm <b>D</b> vs attendance sheet ID. Or call Windi while he’s still around. Typing this is already exhausting.
        </li>
      </ol>

      <h3>Step 4: When all else fails</h3>
      <p>
        If nothing works and there is no attendance note, you are <b>COOKED 💀</b>.<br>
        Or they didn’t attend and the proctor forgot to mark it—prepare pitchfork, torch, gasoline, and proceed to set the proctor ablaze! It’s their <b>FAULT!</b>
      </p>

      <p style="color:#555; font-size:12px; margin-top:12px;">
        Reminder: patience, coffee, and a deep breath are your best allies. Oh, what's that God Mode in CG? Try typing funny negative number in it and watch BX burns.
      </p>
    </div>
  `;
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutput(htmlContent).setTitle("Verify Attendance"));
}

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
        body { font-family:'Google Sans', Arial, sans-serif; margin:0; padding:16px; background:#edf2fa; color:#222; }
        h2 { margin-top:0; color:#1a1a1a; }
        .card { background:#d3e3fd; border-radius:10px; box-shadow:0 2px 5px rgba(0,0,0,0.15); padding:12px 16px; margin-bottom:12px; transition: transform 0.1s ease, box-shadow 0.4s ease; }
        .card:hover { transform:translateY(-2px); box-shadow:0 6px 10px rgba(0,0,0,0.2); }
        .card-header { font-weight:bold; cursor:pointer; display:flex; align-items:center; color:#1a1a1a; }
        .card-header .arrow-icon { font-size:26px; margin-right:8px; transition: transform 0.4s ease; color:#3a3a3a; }
        .card-header .section-icon { font-size:20px; margin-right:6px; color:#1e88e5; }
        .card-content { margin-top:8px; color:#333; }
        ul { margin:0; padding-left:18px; }
        li { margin-bottom:4px; }
        .footer-note { color:#555; font-size:12px; margin-top:16px; }
      </style>
    </head>
    <body>
      <h2>Grouping & Contacts</h2>
      <p><i>Manage automatic groupings and contact creation</i></p>

      <div class="card">
        <div class="card-header" onclick="toggleCollapse(this)">
          <span class="arrow-icon material-icons">expand_more</span>
          <span class="section-icon material-icons">group_work</span>
          Grouping
        </div>
        <div class="card-content">
          <ul>
            <li>Filter <b>AL</b> to select a specific date.</li>
            <li>Automatic group assignments appear in <b>AO</b>.</li>
            <li>Override group manually in <b>AP</b> if needed.</li>
            <li>Group naming logic:
              <ul>
                <li>Extract 3 digits from date in <b>AL</b> → year/month.</li>
                <li>One character denotes test mode: "D" = online, "L" = offline.</li>
                <li>Three-character alphanumeric group code based on session/sequence.</li>
                <li>Suffix "T_" or "S_" indicates TKBI/SISTER vs regular participant.</li>
              </ul>
            </li>
          </ul>
        </div>
      </div>

      <div class="card">
        <div class="card-header" onclick="toggleCollapse(this)">
          <span class="arrow-icon material-icons">expand_more</span>
          <span class="section-icon material-icons">contacts</span>
          Contact Creation (VCF)
        </div>
        <div class="card-content">
          <ul>
            <li>VCF entries are in <b>BG</b>, starting with 8 alphanumeric digits (e.g., 25SLA12S).</li>
            <li>Use these codes to import participants into WhatsApp groups reliably.</li>
            <li>To download a VCF:
              <ul>
                <li>Filter by date in <b>AL</b>.</li>
                <li>Use <b>ProTEFL Utility → Download VCF by Tanggal Tes</b> in the menu bar.</li>
              </ul>
            </li>
          </ul>
        </div>
      </div>

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

// ======================
// COPY ATTENDANCE LIST FUNCTION
// ======================
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


//
// EXPERIMENTAL
//
function exportSiakadScoreResults() {
  const ui = SpreadsheetApp.getUi();
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

  const data = sheet.getDataRange().getValues();
  const header = data.shift();
  const dateColIndex = 1; // column B = tanggal tes
  const targetCols = Array.from({length: 12}, (_, i) => i + 2); // C–N = index 2–13

  const filtered = data.filter(row => String(row[dateColIndex]) === dateFilter);
  const exportData = filtered.length === 0 ? [] : [targetCols.map(i => header[i])];
  filtered.forEach(row => exportData.push(targetCols.map(i => row[i])));

  // Inline HTML dialog (VCF-style)
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

  ui.showModalDialog(HtmlService.createHtmlOutput(htmlContent).setWidth(460).setHeight(250), "Export Siakad Score Results");
}
//
// EXPERIMENTAL
//



// These are used to automatically populate the headers/titles inside each sheet.
/**
 * Sheet setup config: easy to extend!
 * For each sheet, provide
 *   - sheetName: the name to create
 *   - cells: { range: value }
 */
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
        'A40': 'LINK GRUP WA',
        'F40': 'Jumlah Peserta',
      }
    },
    {
      sheetName: '02. CEKTESTHISTORY',
      cells: {
        'A1': 'Name',
        'B1': 'Student ID',
        'C1': 'Whatever it broke if I delete C',
        'D1': 'Urutan Cek (Helper)',
        'E1': 'Test Taken',
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
    }
  ];

// Special config for Form responses 1 since it exists and has a lot of columns (~AC1=… hundreds)
  const FORM_RESPONSES_1_HEADER = [
    // For example, col(22 = V1) = ...
    [ ['V1', 'TABLE SCHEDULE | Reschedule'],
      ['W1', 'Rescheduled Date'],
      ['X1', 'Schedule Log'],
      ['Y1', 'Reschedule Count'],
      ['Z1', '-info Original Schedule'],
      ['AA1', '-helper Pilihan Tanggal Tes'],
      ['AB1', '-helper Bulan dan Tahun'],
      ['AC1', '-helper Jam Daring'],
      ['AD1', '-helper Jam Luring'],
      ['AE1', 'Konfirmasi WA Reschedule Bulan Lalu'],
      ['AF1', 'Notes'],
      ['AG1', 'Status Konfirmasi'],
      ['AH1', 'Confirmation Message'],
      ['AI1', 'TABLE TEST USER | Username Tes ProTEFL'],
      ['AJ1', 'Nama Peserta (Proper Noun)'],
      ['AK1', 'Password Tes ProTEFL'],
      ['AL1', 'Kode Masuk Tes ProTEFL'],
      ['AM1', 'TABLE TEST GROUP | Kode Sesi Bulan'],
      ['AN1', '-helper Kode Sesi Moda'],
      ['AO1', 'Kode Sesi Grup Pengawasan'],
      ['AP1', 'Override Grup Pengawasan'],
      ['AQ1', '-helper Prefix Jenis Tes'],
      ['AR1', '-helper DRAG Urutan Grup'],
      ['AS1', 'TABLE PAYMENT | Verifikasi Bayar'],
      ['AT1', '-helper WhatsApp Peserta'],
      ['AU1', 'Bukti Bayar'],
      ['AV1', 'Nominal Pembayaran'],
      ['AW1', 'Nama Pemilik Rekening (Dompet Digital)'],
      ['AX1', 'Status Pembayaran'],
      ['AY1', 'Konfirmasi via WA'],
      ['AZ1', 'TABLE NIM VERIFICATION | STUDENT ID'],
      ['BA1', 'Name'],
      ['BB1', 'DB Name'],
      ['BC1', 'Status'],
      ['BD1', 'TABLE CONTACTS | Contact Name'],
      ['BE1', 'WhatsApp'],
      ['BF1', 'Test Scheduling Status'],
      ['BG1', 'Grouping VCF'],
      ['BH1', 'Archive VCF'],
      ['BI1', 'Additional Contact Description'],
      ['BJ1', 'Tanggal tes'],
      ['BK1', 'Urutan registrasi sesi'],
      ['BL1', 'Selesai tes'],
      ['BM1', 'Siakad atau TKBI'],
      ['BN1', 'sudah berapa kali tes | MANUAL CEK SIAKAD OLD'],
      ['BO1', 'cek angkatan'],
      ['BP1', 'nim/nik'],
      ['BQ1', 'kode unik sesi tes peserta'],
      ['BR1', 'nidn'],
      ['BS1', 'nama'],
      ['BT1', 'status'],
      ['BU1', 'listening'],
      ['BV1', 'grammar'],
      ['BW1', 'reading'],
      ['BX1', 'skor'],
      ['BY1', 'ielts'],
      ['BZ1', 'Jenjang'],
      ['CA1', 'Fakultas'],
      ['CB1', 'Prodi'],
      ['CC1', 'MIN SKOR'],
      ['CD1', 'MIN MEN'],
      ['CE1', 'TAMBAHAN SKOR JUR INGG'],
      ['CF1', 'Cari gris'],
      ['CG1', 'God Mode'],
      ['CH1', 'Skor TKBI'],
      ['CI1', 'Helper Grup Pagi Siang']
    ]
  ]
  
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
  }
  
  /**
   * Optionally, you could run this on a time-driven trigger OR onOpen.
   * For now, just run initializeSheets from your "main.gs"
   */
  


// ============================================================================
// setupDropdowns.gs
// ---------------------------------------------------------------------------
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

// ============================================================================
// DROPDOWN CONFIGURATION
// Format: [sheetName, columnLetter, optionsArray]
// Validation applies only where column C (default key column) is non-empty.
// ============================================================================
const DROPDOWN_CONFIG = [
    ['Form responses 1', "V", ['Yes', 'No', 'Tidak Jadi Tes']],
    ['Form responses 1', "AG", ['Sent', 'Confirmed', 'Sent-No Answer']],
    ['Form responses 1', "AX", ['LUNAS', 'OKE', '😡', 'CEK', 'Nama Beda', 'Tidak Ada Nama', 'PALSU', 'SALAH BUKTI', 'Jumlah Salah', 'Pindah Pelatihan']]
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
// ---------------------------------------------------------------------------
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
      IF(C2:C<>"", IFERROR(
        TEXTJOIN(", ", TRUE, FILTER('00. MASTER-DATA'!A20:A & " " & '00. MASTER-DATA'!B20:B, '00. MASTER-DATA'!C20:C="Available")),
        "sudah tidak tersedia, silakan reschedule ke bulan selanjutnya menunggu konfirmasi kami"
      ), "")
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
    ['Form responses 1', 'AC2',  "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"", "daring pukul 09.30 WIB (kecuali Jumat, dimulai pukul 08.30 WIB)", ""))`],
    ['Form responses 1', 'AD2',  "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"", "luring pukul 09.00 WIB", ""))`],
    ['Form responses 1', 'AE2', "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"",
        "https://web.whatsapp.com/send?phone=62"&RIGHT(AT2:AT,LEN(AT2:AT)-1)&"&text="&
        ENCODEURL(
          "Salam, kami dari Unit Layanan Bahasa Universitas Negeri Yogyakarta."&CHAR(10)&CHAR(10)&
          "Apakah betul dengan sdr/i. *" & AJ2:AJ & "*?"&CHAR(10)&
          "Izin konfirmasi terkait pendaftaran tes ProTEFL yang telah dilakukan pada *" & TEXT(A2:A, "dd") & " " & CHOOSE(MONTH(A2:A),"Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember") & " " & TEXT(A2:A,"yyyy") & " pukul " & TEXT(A2:A,"HH:mm") & "*." &CHAR(10)&
          "Kami menawarkan reschedule tes ke bulan ini."&CHAR(10)&CHAR(10)&
          "Jadwal tes yang tersedia pada bulan *" & AB2:AB & "* adalah pada tanggal:"&CHAR(10)&
          "*" & AA2:AA & "* dengan pilihan moda *" & AC2:AC & "* dan *" & AD2:AD & "*."&CHAR(10)&CHAR(10)&
          "Mohon memilih salah satu jadwal yang tersedia tersebut."&CHAR(10)&
          "Terima kasih."&CHAR(10)&
          "Setelah memilih, mohon tunggu pesan konfirmasi dari kami untuk memastikan jadwal sudah diperbarui."&CHAR(10)&
          "*Bila belum mendapat pesan konfirmasi, berarti jadwal belum diperbaharui oleh admin yang bertugas.*"
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
    ['Form responses 1', 'AJ2',  "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"", IF(E2:E<>"", PROPER(E2:E), IF(J2:J<>"", PROPER(L2:L), NA())), ""))`],
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
    ['Form responses 1', 'AT2', "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"", IF(F2:F<>"", F2:F, IF(J2:J<>"", O2:O, NA())), ""))`],
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
    ['Form responses 1', 'BN2', "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"", '02. CEKTESTHISTORY'!E2:E, ""))`],
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
    ['Form responses 1', 'BU2', "FILLDOWN", `=IF(C2<>"", VLOOKUP(BQ2, SINICOPYHASILSKOR!A:L, 8, FALSE), "")`],
    ['Form responses 1', 'BV2', "FILLDOWN", `=IF(C2<>"", VLOOKUP(BQ2, SINICOPYHASILSKOR!A:L, 9, FALSE), "")`],
    ['Form responses 1', 'BW2', "FILLDOWN", `=IF(C2<>"", VLOOKUP(BQ2, SINICOPYHASILSKOR!A:L, 10, FALSE), "")`],
    ['Form responses 1', 'BX2', "FILLDOWN", `=IF(C2<>"",
    IF(ISNA(VLOOKUP(BQ2, SINICOPYHASILSKOR!A:L, 11, FALSE)),
        "TIDAK DITEMUKAN, SILAKAN UPLOAD SKOR",
        VLOOKUP(BQ2, SINICOPYHASILSKOR!A:L, 11, FALSE) + CG2
    ),
    ""
    )`],
    ['Form responses 1', 'BY2', "FILLDOWN", `=IF(C2<>"", VLOOKUP(BQ2, SINICOPYHASILSKOR!A:L, 12, FALSE), "")`],
    ['Form responses 1', 'BZ2', "FILLDOWN", `=IF(C2<>"", VLOOKUP(BP2,DATABASEMAHASISWA!$A:$E,5,FALSE), "")`],
    ['Form responses 1', 'CA2', "FILLDOWN", `=IF(C2<>"", VLOOKUP(BP2,DATABASEMAHASISWA!$A:$E,3,FALSE), "")`],
    ['Form responses 1', 'CB2', "FILLDOWN", `=IF(C2<>"", VLOOKUP(BP2,DATABASEMAHASISWA!$A:$E,4,FALSE), "")`],
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
    ['Form responses 1', 'CI2', "FILLDOWN", `=IF(NOT(ISBLANK(W2)), IF(REGEXMATCH(W2, "13\.00|13\.15"), "AFT", "MOR"), IF(REGEXMATCH(R2, "13\.00|13\.15"), "AFT", "MOR"))`],

  // ====== OTHER SHEETS ======

  // 02. CEKTESTHISTORY
    ['02. CEKTESTHISTORY', 'A2', "ARRAY", `=ARRAYFORMULA('Form responses 1'!AJ2:AJ)`],
    ['02. CEKTESTHISTORY', 'B2', "ARRAY", `=ARRAYFORMULA( IF( 'Form responses 1'!BO2:BO="ijazah/test x5/angkatan lama", IF( LEN('Form responses 1'!AI2:AI)=11, 'Form responses 1'!AI2:AI, IF( 'Form responses 1'!C2:C="ProTEFL TKBI/SERDOS/Umum (bersertifikat resmi diakui SISTER KEMENDIKBUDRISTEK)", "Peserta TKBI/UMUM", IF( 'Form responses 1'!C2:C="ProTEFL SIAKAD UNY (tanpa sertifikat)", "Error: CEK NIM", "" ) ) ), "" ) )`],
    ['02. CEKTESTHISTORY', 'D2', "ARRAY", `=ARRAYFORMULA( IF( B2:B<>"", SCAN( 0, B2:B, LAMBDA(acc, x, IF(x<>"", acc+1, acc) ) ), "" ) )`],

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
            IF(
              REGEXMATCH('Form responses 1'!BI2:BI, "_BERKALA"),
              MID('Form responses 1'!AI2:AI, 4, LEN('Form responses 1'!AI2:AI)-3),
              IF(
                REGEXMATCH('Form responses 1'!BI2:BI, "_OFFGRID"),
                'Form responses 1'!D2:D,
                'Form responses 1'!AI2:AI
              )
            ),
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
        1, TRUE, 2, TRUE
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
  
// ---------------------------------------------------------------------------
// applyAllFormulas()
// Loops through the central FORMULAS_TO_APPLY configuration and ensures
// each formula is inserted in its correct sheet and cell. Supports two types:
//   - ARRAY: sets a single ARRAYFORMULA at the given anchor cell
//   - FILLDOWN: fills a formula down column rows based on a key column
// ---------------------------------------------------------------------------
function applyAllFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Iterate through all configured formulas
  FORMULAS_TO_APPLY.forEach(row => {
    var [sheetName, cellA1, type, formula, keyCol] = row;

    // Skip if target sheet does not exist
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    // Dispatch based on formula type
    if (type === "ARRAY") {
      // Insert ARRAYFORMULA once into the anchor cell
      setFormulaOnce(sheet, cellA1, formula);

    } else if (type === "FILLDOWN") {
      // Fill the formula down to all non-empty rows
      // Default key column = 3 (column C) if none provided
      fillDownFormula(sheet, cellA1, formula, keyCol || 3);
    }
  });
}


// ============================================================================
// File: autoCounters.gs
// Purpose:
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
// ---------------------------------------------------------------------------
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
  'CH-CI'
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

      // Apply color band styling
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
    }
  });
}
