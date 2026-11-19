/**
 * code.gs - Versi Bersih & Stabil e-BPHTB
 * Database: Tab "DATABASE" (17 Kolom)
 * Password Admin: adminNOT
 */

var ADMIN_PIN = "adminNOT";

function doGet(e) {
  var template = HtmlService.createTemplateFromFile("index");
  return template.evaluate()
    .setTitle("e-BPHTB Dashboard & Management")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getDatabaseSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return null;
  
  var sheet = ss.getSheetByName("DATABASE");
  if (sheet) return sheet;
  
  var allSheets = ss.getSheets();
  for (var i = 0; i < allSheets.length; i++) {
    if (allSheets[i].getName().trim().toUpperCase() === "DATABASE") {
      return allSheets[i];
    }
  }
  return allSheets[0];
}

function cleanNumber(val) {
  if (!val) return 0;
  if (typeof val === "number") return val;
  var cleaned = val.toString().replace(/[^0-9]/g, "");
  return parseInt(cleaned, 10) || 0;
}

function formatRupiah(num) {
  return "Rp" + Number(num || 0).toLocaleString("id-ID");
}

function isAuthorizedAdmin(pin) {
  if (!pin) return false;
  return String(pin).trim() === ADMIN_PIN;
}

function calculateTaxes(jenis, lt, tt, lb, tb) {
  var hTanah = (lt || 0) * (tt || 0);
  var hBangunan = (lb || 0) * (tb || 0);
  var totalNpop = hTanah + hBangunan;

  var bphtb = 0;
  var bphtbWaris = 0;
  var pph = 0;

  var jLower = (jenis || "").toString().toLowerCase().trim();

  if (jLower === "waris") {
    var kenaPajakWaris = Math.max(0, totalNpop - 300000000);
    bphtbWaris = Math.round(kenaPajakWaris * 0.03);
    bphtb = 0;
    pph = 0;
  } else {
    var kenaPajakReguler = Math.max(0, totalNpop - 80000000);
    bphtb = Math.round(kenaPajakReguler * 0.04);
    bphtbWaris = 0;
    pph = Math.round(totalNpop * 0.025);
  }

  return {
    luasTanahStr: lt + " x " + tt,
    hargaTanah: formatRupiah(hTanah),
    bangunanStr: lb + " x " + tb,
    hargaBangunan: formatRupiah(hBangunan),
    nilaiValidasi: formatRupiah(totalNpop),
    bphtbWaris: bphtbWaris > 0 ? formatRupiah(bphtbWaris) : "-",
    bphtb: bphtb > 0 ? formatRupiah(bphtb) : "-",
    pph: pph > 0 ? formatRupiah(pph) : "-"
  };
}

function generateNextRegisterNo() {
  var prefix = "SMP3529";
  var year = Utilities.formatDate(new Date(), "Asia/Jakarta", "yy");
  var randomDigits = Math.floor(1000 + Math.random() * 9000);
  return prefix + year + randomDigits;
}

function getDashboardData(params) {
  try {
    var sheet = getDatabaseSheet();
    if (!sheet) throw new Error("Sheet DATABASE tidak ditemukan!");

    var lastRow = sheet.getLastRow();
    
    var statusCounts = {
      "pengajuan": 0,
      "pengecekan berkas": 0,
      "persetujuan berkas": 0,
      "verifikasi": 0,
      "persetujuan verifikasi": 0,
      "validasi": 0,
      "persetujuan validasi": 0,
      "menunggu pembayaran": 0,
      "selesai": 0,
      "di tolak / di gagalkan": 0
    };

    if (lastRow <= 1) {
      return {
        stats: { total: 0, selesai: 0, ditolak: 0, totalBphtb: "Rp0", totalPph: "Rp0" },
        statusCounts: statusCounts,
        items: [],
        totalFiltered: 0,
        page: 1,
        totalPages: 1
      };
    }

    var data = sheet.getRange(2, 1, lastRow - 1, 17).getDisplayValues();

    var searchQuery = (params && params.search ? String(params.search).toLowerCase().trim() : "");
    var statusFilter = (params && params.status ? String(params.status).toLowerCase().trim() : "");
    var jenisFilter = (params && params.jenis ? String(params.jenis).toLowerCase().trim() : "");
    var startDateStr = (params && params.startDate ? String(params.startDate).trim() : "");
    var endDateStr = (params && params.endDate ? String(params.endDate).trim() : "");

    var page = (params && params.page ? parseInt(params.page, 10) : 1) || 1;
    var pageSize = (params && params.pageSize ? parseInt(params.pageSize, 10) : 15) || 15;

    var totalCount = data.length;
    var selesaiCount = 0;
    var ditolakCount = 0;
    var sumBphtb = 0;
    var sumPph = 0;

    var filteredList = [];

    for (var i = 0; i < totalCount; i++) {
      var row = data[i];
      var rowIndex = i + 2;

      var noReg = row[0] || "";
      var tglReg = row[1] || "";
      var jenis = row[2] || "";
      var wp = row[3] || "";
      var luasTanahStr = row[4] || "";
      var hargaTanah = row[5] || "";
      var bangunanStr = row[6] || "";
      var hargaBangunan = row[7] || "";
      var nilaiValidasi = row[8] || "";
      var bphtbWaris = row[9] || "";
      var bphtb = row[10] || "";
      var pph = row[11] || "";
      var nop = row[12] || "";
      var noSertifikat = row[13] || "";
      var ntpd = row[14] || "";
      var noVa = row[15] || "";
      var status = (row[16] || "Pengajuan").trim();

      var stLower = status.toLowerCase();

      if (stLower.indexOf("tolak") !== -1 || stLower.indexOf("gagalkan") !== -1) {
        statusCounts["di tolak / di gagalkan"]++;
        ditolakCount++;
      } else if (stLower.indexOf("selesai") !== -1) {
        statusCounts["selesai"]++;
        selesaiCount++;
      } else if (stLower.indexOf("menunggu pembayaran") !== -1) {
        statusCounts["menunggu pembayaran"]++;
      } else if (stLower.indexOf("persetujuan validasi") !== -1) {
        statusCounts["persetujuan validasi"]++;
      } else if (stLower.indexOf("validasi") !== -1) {
        statusCounts["validasi"]++;
      } else if (stLower.indexOf("persetujuan verifikasi") !== -1) {
        statusCounts["persetujuan verifikasi"]++;
      } else if (stLower.indexOf("verifikasi") !== -1) {
        statusCounts["verifikasi"]++;
      } else if (stLower.indexOf("persetujuan berkas") !== -1) {
        statusCounts["persetujuan berkas"]++;
      } else if (stLower.indexOf("pengecekan berkas") !== -1) {
        statusCounts["pengecekan berkas"]++;
      } else {
        statusCounts["pengajuan"]++;
      }

      var numBphtb = cleanNumber(bphtb);
      var numBphtbWaris = cleanNumber(bphtbWaris);
      var numPph = cleanNumber(pph);

      sumBphtb += (numBphtb + numBphtbWaris);
      sumPph += numPph;

      var matchSearch = true;
      if (searchQuery) {
        var text = (noReg + " " + wp + " " + nop + " " + noSertifikat + " " + ntpd).toLowerCase();
        if (text.indexOf(searchQuery) === -1) matchSearch = false;
      }

      var matchStatus = true;
      if (statusFilter && statusFilter !== "all") {
        matchStatus = (stLower.indexOf(statusFilter) !== -1);
      }

      var matchJenis = true;
      if (jenisFilter && jenisFilter !== "all") {
        matchJenis = (jenis.toLowerCase() === jenisFilter);
      }

      var matchDate = true;
      if (startDateStr || endDateStr) {
        if (startDateStr && tglReg < startDateStr) matchDate = false;
        if (endDateStr && tglReg > endDateStr) matchDate = false;
      }

      if (matchSearch && matchStatus && matchJenis && matchDate) {
        filteredList.push({
          rowIndex: rowIndex,
          noReg: noReg,
          tglReg: tglReg,
          jenis: jenis,
          wp: wp,
          luasTanahStr: luasTanahStr,
          hargaTanah: hargaTanah,
          bangunanStr: bangunanStr,
          hargaBangunan: hargaBangunan,
          nilaiValidasi: nilaiValidasi,
          bphtbWaris: bphtbWaris,
          bphtb: bphtb,
          pph: pph,
          nop: nop,
          noSertifikat: noSertifikat,
          ntpd: ntpd,
          noVa: noVa,
          status: status
        });
      }
    }

    filteredList.reverse();

    var totalFiltered = filteredList.length;
    var totalPages = Math.ceil(totalFiltered / pageSize) || 1;
    if (page > totalPages) page = totalPages;

    var startIndex = (page - 1) * pageSize;
    var pagedItems = filteredList.slice(startIndex, startIndex + pageSize);

    return {
      stats: {
        total: totalCount,
        selesai: selesaiCount,
        ditolak: ditolakCount,
        totalBphtb: formatRupiah(sumBphtb),
        totalPph: formatRupiah(sumPph)
      },
      statusCounts: statusCounts,
      items: pagedItems,
      totalFiltered: totalFiltered,
      page: page,
      totalPages: totalPages
    };
  } catch (err) {
    throw new Error("Gagal mengambil data: " + err.message);
  }
}

function createBphtbRecord(payload) {
  try {
    var inputPin = (payload && payload.adminPin) ? String(payload.adminPin).trim() : "";
    if (!isAuthorizedAdmin(inputPin)) {
      throw new Error("Akses Ditolak: Anda tidak memiliki hak akses Admin untuk menambah berkas.");
    }

    var sheet = getDatabaseSheet();
    if (!sheet) throw new Error("Database sheet tidak ditemukan.");

    var tglRegister = payload.tglRegister || Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");
    var noRegister = payload.noRegister || generateNextRegisterNo();

    var calc = calculateTaxes(
      payload.jenis,
      cleanNumber(payload.luasTanah),
      cleanNumber(payload.tarifTanah),
      cleanNumber(payload.luasBangunan),
      cleanNumber(payload.tarifBangunan)
    );

    var newRow = [
      noRegister,
      tglRegister,
      payload.jenis,
      (payload.wp || "").toUpperCase(),
      calc.luasTanahStr,
      calc.hargaTanah,
      calc.bangunanStr,
      calc.hargaBangunan,
      calc.nilaiValidasi,
      calc.bphtbWaris,
      calc.bphtb,
      calc.pph,
      payload.nop || "",
      payload.noSertifikat || "",
      payload.ntpd || "",
      payload.noVa || "",
      payload.status || "Pengajuan"
    ];

    sheet.appendRow(newRow);
    SpreadsheetApp.flush();

    return {
      success: true,
      message: "Berkas berhasil didaftarkan!",
      noRegister: noRegister
    };
  } catch (err) {
    throw new Error(err.message);
  }
}

function updateBphtbRecord(payload) {
  try {
    if (!payload) throw new Error("Data kiriman kosong.");

    var inputPin = (payload && payload.adminPin) ? String(payload.adminPin).trim() : "";
    if (!isAuthorizedAdmin(inputPin)) {
      throw new Error("Akses Ditolak: Anda tidak memiliki izin Admin untuk mengubah data berkas ini.");
    }

    var sheet = getDatabaseSheet();
    if (!sheet) throw new Error("Database sheet tidak ditemukan.");

    var rowIndex = parseInt(payload.rowIndex, 10);
    if (isNaN(rowIndex) || rowIndex < 2) {
      throw new Error("Nomor baris data tidak valid (rowIndex: " + payload.rowIndex + ").");
    }

    var currentRow = sheet.getRange(rowIndex, 1, 1, 17).getValues()[0];
    var jenis = payload.jenis || currentRow[2] || "Jual Beli";

    if (payload.noRegister !== undefined && String(payload.noRegister).trim() !== "") {
      sheet.getRange(rowIndex, 1).setValue(String(payload.noRegister).trim());
    }

    var lt = cleanNumber(payload.luasTanah);
    var tt = cleanNumber(payload.tarifTanah);
    var lb = cleanNumber(payload.luasBangunan);
    var tb = cleanNumber(payload.tarifBangunan);

    var calc = calculateTaxes(jenis, lt, tt, lb, tb);

    sheet.getRange(rowIndex, 5, 1, 8).setValues([[
      calc.luasTanahStr,
      calc.hargaTanah,
      calc.bangunanStr,
      calc.hargaBangunan,
      calc.nilaiValidasi,
      calc.bphtbWaris,
      calc.bphtb,
      calc.pph
    ]]);

    if (payload.noSertifikat !== undefined) sheet.getRange(rowIndex, 14).setValue(String(payload.noSertifikat).trim());
    if (payload.ntpd !== undefined) sheet.getRange(rowIndex, 15).setValue(String(payload.ntpd).trim());
    if (payload.noVa !== undefined) sheet.getRange(rowIndex, 16).setValue(String(payload.noVa).trim());
    if (payload.status !== undefined) sheet.getRange(rowIndex, 17).setValue(String(payload.status).trim());

    SpreadsheetApp.flush();

    return {
      success: true,
      message: "Data berkas & No Register berhasil diperbarui!"
    };
  } catch (err) {
    throw new Error(err.message);
  }
}