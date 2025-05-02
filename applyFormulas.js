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
          "Izin konfirmasi terkait tes ProTEFL yang belum terlaksana pada *" & R2:R & "*."&CHAR(10)&
          "Kami menawarkan reschedule tes ke bulan ini."&CHAR(10)&CHAR(10)&
          "Jadwal tes yang tersedia pada bulan *" & AB2:AB & "* adalah pada tanggal:"&CHAR(10)&
          "*" & AA2:AA & "* dengan pilihan moda *" & AC2:AC & "* dan *" & AD2:AD & "*."&CHAR(10)&CHAR(10)&
          "Mohon memilih salah satu jadwal yang tersedia tersebut."&CHAR(10)&
          "Terima kasih."&CHAR(10)&
          "Setelah memilih, mohon tunggu pesan konfirmasi dari kami untuk memastikan jadwal sudah diperbarui."
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
        IF(ISNUMBER(SEARCH("_BERKALA", BI2:BI)),
            RIGHT(TEXT(VALUE(RIGHT(IF(D2:D<>"", D2:D, IF(J2:J<>"", J2:J, NA())), 3)) + 420, "000"), 3) &
            IF(D2:D<>"", D2:D, IF(J2:J<>"", J2:J, NA())),
            IF(D2:D<>"", D2:D, IF(J2:J<>"", J2:J, NA()))
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
            AI2:AI
            )
        ),
        ""
        )
    )`],
    ['Form responses 1', 'BA2', "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"", AJ2:AJ, ""))`],
    ['Form responses 1', 'BB2', "ARRAY", `=ARRAYFORMULA(IF(C2:C<>"", IF(AZ2:AZ<>"", VLOOKUP(AZ2:AZ, DATABASEMAHASISWA!A:B, 2, FALSE), ""), ""))`],
    ['Form responses 1', 'BC2', "ARRAY", `=ARRAYFORMULA( IF( (C2:C<>""), IF( ( (LEN(BA2:BA)=0) + (LEN(BB2:BB)=0) )>0, "#N/A", IFERROR( IF( EXACT(BA2:BA, BB2:BB), "COCOK", IF( ( ISNUMBER(SEARCH(BA2:BA, BB2:BB)) + ISNUMBER(SEARCH(BB2:BB, BA2:BA)) )>0, "CEK NAMA", "GUNDULMU REK, SALAH NIM" ) ), "TKBI" ) ), "" ) )`],
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
            IF(REGEXMATCH(BI2:BI, "_BERKALA"),
            MID(AI2:AI, 4, LEN(AI2:AI)-3),
            AI2:AI
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
    ['Form responses 1', 'CD2', "FILLDOWN", `=IFS(BZ2="D3",400,BZ2="D4",427,BZ2="S1",427,BZ2="S2",450,BZ2="S3",475)`],
    ['Form responses 1', 'CE2', "FILLDOWN", `=IFERROR(IFS(CF2="gris - S1",73,CF2="gris - S2",100),"0")`],
    ['Form responses 1', 'CI2', "FILLDOWN", `=IF(REGEXMATCH(R2,"13\\.00"), "AFT", "MOR")`],

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
              'Form responses 1'!AI2:AI
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
                'Form responses 1'!AI2:AI
              ), "#"
            ),
            "^\\d{11}$"
          )
        ),
        1, TRUE
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