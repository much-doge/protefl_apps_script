/**
 * Sheet setup config: easy to extend!
 * For each sheet, provide
 *   - sheetName: the name to create
 *   - cells: { range: value }
 */
const SHEET_INITIALIZATIONS = [
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
