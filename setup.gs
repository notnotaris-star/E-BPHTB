/**
 * setup.gs - ZettBOT Safe Migrate Database Init
 * Memastikan sheet DATABASE dan header 17 kolom siap digunakan tanpa menghapus data yang ada.
 */
function initDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "DATABASE";
  var sheet = ss.getSheetByName(sheetName);

  var requiredHeaders = [
    "No Register",
    "Tgl Register",
    "Jenis",
    "WP",
    "LUAS TANAH",
    "HARGA",
    "BANGUNAN",
    "HARGA BANGUNAN",
    "NILAI VALIDASI",
    "BPHTB WARIS",
    "BPHTB",
    "PPH 2,5%",
    "NOP",
    "No Sertifikat",
    "NTPD",
    "No Va",
    "Status"
  ];

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(requiredHeaders);
    sheet.getRange(1, 1, 1, requiredHeaders.length).setFontWeight("bold").setBackground("#151b26").setFontColor("#00ffaa");
    Logger.log("Sheet DATABASE berhasil dibuat dengan header standar.");
  } else {
    // Safe Migrate: Jangan hapus data, cukup cek dan update baris header pertama
    var existingHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
    var isHeaderEmpty = existingHeaders.length === 0 || existingHeaders[0] === "";
    
    if (isHeaderEmpty) {
      sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
      sheet.getRange(1, 1, 1, requiredHeaders.length).setFontWeight("bold").setBackground("#151b26").setFontColor("#00ffaa");
      Logger.log("Header sheet DATABASE berhasil diisi.");
    } else {
      Logger.log("Sheet DATABASE dan data sudah ada (" + sheet.getLastRow() + " baris). Struktur aman tanpa perubahan destruktif.");
    }
  }

  // Set format teks pada kolom NOP, No Register, NTPD, dan No Va agar titik dan angka panjang tidak error
  sheet.getRange("A:A").setNumberFormat("@");
  sheet.getRange("M:M").setNumberFormat("@");
  sheet.getRange("O:P").setNumberFormat("@");
}