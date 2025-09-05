// main.gs
/**
 * Main admin orchestrator for ProTEFL registration workbook.
 * Run this function to set up or refresh everything.
 */
function main() {
    initializeSheets();                 // Creates and populates sheets and templates (setupSheets.gs)
    setupAllDropdowns();                // Adds dropdowns to columns (setupDropdowns.gs)
    protectOriginalScheduleColumn();    // Protects the 'Original Schedule' column (autoCounters.gs)
    applyAllStyling();                  // Applies header and other styling (styling.gs)
    applyAllFormulas();                 // Applies all relevant formulas (applyFormulas.gs)
    // Optionally: syncRescheduleCounts(); // Recalculate reschedule count column (autoCounters.gs)
}

// Create ribbon or whatever you call it, menu. Oh it's menu bar. Yes, this creates menu bar called "ProTEFL Utility."
function onOpen() {
    SpreadsheetApp.getUi()
      .createMenu("ProTEFL Utility")
      .addItem("Search Test Taker", "showSearchSidebar")
      .addItem("Apply Styles", "applyAllStylingWithConfirm")
      .addItem("Protect Original Schedule Column", "protectOriginalScheduleColumn")
      .addItem("Set Up AutoCounter Trigger", "setupAutoCounterTriggerWithAlert")
      .addSeparator()
      .addItem("Apply All Formulas (Danger Zone)", "applyAllFormulasWithConfirm")
      .addItem("Initialize Sheet (Danger Zone)", "runMainWithConfirm")
      .addToUi();
  }
    function applyAllStylingWithConfirm() {
      var ui = SpreadsheetApp.getUi();
      var resp = ui.alert(
        "Apply Styles",
        "Do you want to re-apply all custom styles to your registration sheets? This will reset header colors, column banding, and other visual formatting in all managed sheets. Proceed?",
        ui.ButtonSet.OK_CANCEL
      );
      if (resp == ui.Button.OK) {
        applyAllStyling();
        ui.alert("Styling applied!");
      } else {
        ui.alert("Cancelled. No changes made.");
      }
    }
    function applyAllFormulasWithConfirm() {
      var ui = SpreadsheetApp.getUi();
      var resp = ui.alert(
        "Apply All Formulas (Danger Zone)",
        "Have you copied or imported the DATABASEMAHASISWA sheet into this workbook? Applying ALL formulas may cause errors if lookup sheets aren't present. Are you SURE you want to rerun ALL formulas and array/dynamic columns?",
        ui.ButtonSet.OK_CANCEL
      );
      if (resp == ui.Button.OK) {
        applyAllFormulas();
        ui.alert("All formulas applied!");
      } else {
        ui.alert("Cancelled. No changes made.");
      }
    }
    function runMainWithConfirm() {
      var ui = SpreadsheetApp.getUi();
      var resp = ui.alert(
        "Initialize Sheet (Danger Zone)",
        "This will initialize/reinitialize your registration workbook (create sheets, headers, values, styles, etc). Have you copied the DATABASEMAHASISWA sheet to this workbook? This process is NOT reversible. Proceed?",
        ui.ButtonSet.OK_CANCEL
      );
      if (resp == ui.Button.OK) {
        main();
        ui.alert("Sheet initialization complete!");
      } else {
        ui.alert("Cancelled. No changes made.");
      }
    }
    function setupAutoCounterTriggerWithAlert() {
        var ui = SpreadsheetApp.getUi();
        var resp = ui.alert(
          "Set Up Trigger",
          "This will create or replace the installable onEdit trigger for the auto counter/logging script on this spreadsheet.\n\nProceed?",
          ui.ButtonSet.OK_CANCEL
        );
        if (resp == ui.Button.OK) {
          setupAutoCounterTrigger();
          ui.alert("AutoCounter Trigger is now set up! (If it was already present, it has been replaced.)");
        } else {
          ui.alert("Cancelled. No changes made.");
        }
    }
    function setupAutoCounterTrigger() {
        var triggers = ScriptApp.getProjectTriggers();
        // Remove previous onEditLogReschedule triggers for this spreadsheet
        triggers.forEach(function(trigger) {
          if (trigger.getHandlerFunction() === 'onEditLogReschedule') {
            ScriptApp.deleteTrigger(trigger);
          }
        });
        ScriptApp.newTrigger('onEditLogReschedule')
          .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
          .onEdit()
          .create();
    }

// EXPERIMENTAL FEATURE
/**
 * Show sidebar with dropdown of test takers
 */
function showSearchSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('SearchTestTaker')
    .setTitle('Search Test Taker');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Get list of test taker names (AJ) and their Test IDs (AI)
 * Returns array of objects [{name, testId}]
 */
function showSearchTestTakerSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('SearchTestTaker')
    .setTitle('Search Test Taker');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getTestTakers() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Form responses 1');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const numRows = lastRow - 1;
  const names = sheet.getRange(2, 36, numRows, 1).getValues().map(r => r[0]); // AJ
  const ids = sheet.getRange(2, 37, numRows, 1).getValues().map(r => r[0]);   // AK

  const result = [];
  for (let i = 0; i < numRows; i++) {
    if (names[i]) {
      result.push({ name: names[i], testId: ids[i] || '' });
    }
  }

  return result;
}
// EXPERIMENTAL FEATURE

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
      ['CG1', 'Treatment'],
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



// applyFormulas.gs

/** Utility: get last non-empty row based on 'keyColumn' (default = col C) */
function getLastDataRow_(sheet, keyCol = 3) {
    const vals = sheet.getRange(2, keyCol, Math.max(sheet.getLastRow()-1, 1)).getValues().flat();
    for (let i = vals.length - 1; i >= 0; i--) if (vals[i] !== "") return i + 2;
    return 2;
  }
  
  /** Utility: set ARRAYFORMULA at target cell if not present or formula deleted */
  function setFormulaOnce(sheet, cellA1, formula) {
    if (sheet.getRange(cellA1).getFormula() !== formula)
      sheet.getRange(cellA1).setFormula(formula);
  }
  
  /** Utility: for dragdown formulas (per-row), sets on every non-empty keyCol row */
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
  
  // ----- CONFIG SECTION: Add all formulas below (sheetName, cellA1, type, formula, optional: keyColumn) -----
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
                FLOOR((COUNTIFS(AL$2:AL2,AL2,AN$2:AN2,"L",CI$2:CI2,"AFT")-1)/30,1),
                FLOOR((COUNTIFS(AL$2:AL2,AL2,AN$2:AN2,"L",CI$2:CI2,"MOR")-1)/30,1)
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
    ['Form responses 1', 'CE2', "FILLDOWN", `=IFERROR(IFS(
        CF2=TEXT("gris - S1","@"), 73,
        CF2=TEXT("gris - S2","@"), 100
      ), "0")`],
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
  
  // === Main function ===
  function applyAllFormulas() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    FORMULAS_TO_APPLY.forEach(row => {
      var [sheetName, cellA1, type, formula, keyCol] = row;
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      if (type === "ARRAY") {
        setFormulaOnce(sheet, cellA1, formula);
      } else if (type === "FILLDOWN") {
        fillDownFormula(sheet, cellA1, formula, keyCol || 3);
      }
    });
  }


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