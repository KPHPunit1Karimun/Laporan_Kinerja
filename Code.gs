/**
 * ==========================================
 * PORTFOLIO KINERJA KPHP - Backend Code.gs
 * ==========================================
 */

const SHEET_ID = "1QyNnTFX4UM17EgfVIcB82Y4wbkKFx79KUkOswc_UmZI"; // (Ini yang lama)
const SHEET_BMD_ID = "1veWPpg33oBlGJHdK6l1-kSxuD9AT_2ucK33gM59yjcE"; // (Ini yang baru)

// =======================================================================
// 1. FUNGSI GLOBAL & ROUTING
// =======================================================================
function getGeminiAPIKey() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
}

function doGet(e) {
  try {
    return HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle('Portfolio Kinerja')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    return HtmlService.createHtmlOutput("Terjadi kesalahan sistem: " + error.message);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getPage(filename) {
  try {
    return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
  } catch (e) {
    return "<h3 style='color:red;'>Error 404</h3><p>Halaman <b>" + filename + "</b> tidak ditemukan.</p>";
  }
}

function pancingIzinExternal() { UrlFetchApp.fetch("[https://www.google.com](https://www.google.com)"); }

// =======================================================================
// 2. MODUL OTENTIKASI & USER SESSION (LOGIN)
// =======================================================================
function loginUser(username, password) {
  var res = cekLoginDiSheet(username, password); 
  if (res.status_berhasil) {
    return { success: true, nama: res.nama_lengkap, role: res.role, username: res.nip };
  } else {
    return { success: false, message: res.pesan };
  }
}

function setLoginSession(r) {
  PropertiesService.getUserProperties().setProperty("username", r.username);
  return true;
}

function logoutUser() {
  try {
    PropertiesService.getUserProperties().deleteProperty("username");
    return true;
  } catch(e) {
    return false;
  }
}

// Ganti fungsi ini di Code.gs
function cekLoginDiSheet(username, password) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("DATA_USER");
    if (!sheet) return { status_berhasil: false, pesan: "Sheet DATA_USER tidak ditemukan!" };

    // Pakai getValues() agar lebih ringan & cepat daripada getDisplayValues()
    const data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var userDiSheet = data[i][1].toString().trim().toLowerCase(); 
      var passDiSheet = data[i][2].toString().trim();
      
      if (userDiSheet === String(username).toLowerCase().trim() && passDiSheet === password) {
        if(data[i][5].toString().trim().toUpperCase() !== "AKTIF") {
            return { status_berhasil: false, pesan: "Akun Anda sedang dinonaktifkan!" };
        }
        let rolePegawai = data[i][4].toString().trim(); 
        let sessionToken = Utilities.getUuid();
        let waktuLogin = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
        
        // PERCEPATAN: Menulis 2 kolom sekaligus dalam 1 perintah agar tidak lag
        sheet.getRange(i + 1, 7, 1, 2).setValues([[waktuLogin, sessionToken]]);

        return { 
          status_berhasil: true, pesan: "Berhasil login", nip: data[i][1].toString().trim(), 
          role: rolePegawai, nama_lengkap: data[i][3].toString().trim(), token: sessionToken 
        };
      }
    }
    return { status_berhasil: false, pesan: "Username atau Password Salah!" };
  } catch(err) { return { status_berhasil: false, pesan: "Error Server: " + err.toString() }; }
}

// Ganti fungsi ini di Code.gs
// PERCEPATAN: Terima nipUser langsung dari antarmuka agar sistem tidak perlu 
// membuka dan membaca Sheet DATA_USER untuk kedua kalinya.
function getMenuUser(nipUser) {
  try {
    if (!nipUser) return [];
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName("HAK_AKSES_MENU");
    if (!sh) return []; 
    var data = sh.getDataRange().getDisplayValues();
    var hasil = [];
    
    var targetNip = String(nipUser).replace(/'/g, "").trim().toLowerCase(); 
    
    for (var i = 1; i < data.length; i++) {
      var rowNip = String(data[i][1]).replace(/'/g, "").trim().toLowerCase();
      var aktif = String(data[i][4]).trim().toUpperCase(); 
      
      if (rowNip === targetNip && aktif === "YA") {
        hasil.push({ kode: String(data[i][2]).trim(), nama: String(data[i][3]).trim(), parent: "" });
      }
    }
    return hasil;
  } catch (e) { return []; }
}

function getCurrentUser() {
  try {
    var username = PropertiesService.getUserProperties().getProperty("username");
    if (!username) return null;
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName("DATA_USER");
    if (!sh) return null;
    
    var data = sh.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(username).trim()) {
        return { username: data[i][1], nama: data[i][3], role: data[i][4] };
      }
    }
    return null;
  } catch (e) { return null; }
}

function validasiOtoritas(nip, tokenReq, izinRoleArray) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("DATA_USER");
    const data = sheet.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim() === String(nip).trim()) {
        let dbToken = data[i][7]; 
        let dbRole = String(data[i][4]).toUpperCase(); 
        if (dbToken === tokenReq && izinRoleArray.includes(dbRole)) return true; 
      }
    }
    return false; 
  } catch(e) { return false; }
}

// =======================================================================
// 3. MODUL HAK AKSES MENU (SPESIFIK PER-PEGAWAI)
// =======================================================================
function getDaftarAkun() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName("DATA_USER");
    if (!sh) return [];
    var data = sh.getDataRange().getDisplayValues();
    var hasil = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && String(data[i][1]).trim() !== "") { 
        hasil.push({ nip: String(data[i][1]).trim(), nama: String(data[i][3]).trim() });
      }
    }
    return hasil;
  } catch (e) { return []; }
}

function simpanHakAksesUser(dataArray) {
  try {
    if (!dataArray || dataArray.length === 0) return { success: false, message: "Tidak ada data." };
    var targetNip = String(dataArray[0][1]).trim().toLowerCase();
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName("HAK_AKSES_MENU");
    
    if (!sh) {
      sh = ss.insertSheet("HAK_AKSES_MENU");
    }
    
    // Pastikan header benar
    sh.getRange("A1:H1").setValues([["ID", "USERNAME_PEGAWAI", "MENU", "NAMA_MENU", "LIHAT", "TAMBAH", "EDIT", "HAPUS"]]).setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
    
    var data = sh.getDataRange().getDisplayValues();
    for (var i = data.length - 1; i >= 1; i--) {
      var rowNip = String(data[i][1]).replace(/'/g, "").trim().toLowerCase();
      if (rowNip === targetNip) {
        sh.deleteRow(i + 1);
      }
    }
    
    SpreadsheetApp.flush(); 
    
    var rowsToAppend = [];
    var baseId = sh.getLastRow();
    for (var j = 0; j < dataArray.length; j++) {
       var status = dataArray[j][4]; 
       rowsToAppend.push([ baseId + j, "'" + targetNip, dataArray[j][2], dataArray[j][3], status, status, status, status ]);
    }
    if (rowsToAppend.length > 0) sh.getRange(sh.getLastRow() + 1, 1, rowsToAppend.length, 8).setValues(rowsToAppend);
    
    SpreadsheetApp.flush();
    return { success: true, message: "Hak akses pegawai berhasil disimpan!" };
  } catch (e) { return { success: false, message: e.message }; }
}

function getHakAksesUser(nip) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName("HAK_AKSES_MENU");
    if (!sh) return [];
    var data = sh.getDataRange().getDisplayValues();
    var hasil = [];
    var targetNip = String(nip).replace(/'/g, "").trim().toLowerCase();
    
    for (var i = 1; i < data.length; i++) {
      var rowNip = String(data[i][1]).replace(/'/g, "").trim().toLowerCase();
      if (rowNip === targetNip) {
        var aktif = String(data[i][4]).trim().toUpperCase(); // Kolom LIHAT
        hasil.push({ kode: String(data[i][2]).trim(), aktif: aktif });
      }
    }
    return hasil;
  } catch(e) { return []; }
}


// =======================================================================
// 4. MASTER DATA PENDUKUNG (Pegawai, Rekening, Pengawas)
// =======================================================================
function getMasterPegawai() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID); 
    var sheet = ss.getSheetByName("DATA_USER");
    if(!sheet) return [];
    var data = sheet.getDataRange().getDisplayValues();
    var arr = [];
    for(var i = 1; i < data.length; i++) {
      if (data[i][1] !== "" && data[i][5].toUpperCase() === "AKTIF") {
          arr.push({ id: data[i][0], nama: String(data[i][3]).trim(), nip: String(data[i][1]).trim(), peran: String(data[i][4]).trim() });
      }
    }
    return arr;
  } catch (e) { return []; }
}

function getPegawaiDariMaster() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID); 
    var sheet = ss.getSheetByName("Master_Pengawas");
    if(!sheet) return [];
    var data = sheet.getDataRange().getDisplayValues();
    var arr = [];
    for(var i = 1; i < data.length; i++) {
      if (data[i][1] !== "") {
          arr.push({ id: data[i][0], nama: String(data[i][1]).trim(), nip: String(data[i][2]).trim(), pangkat: String(data[i][3]).trim(), jabatan: String(data[i][4]).trim() });
      }
    }
    return arr;
  } catch (e) { return []; }
}

function simpanMasterPengawas(payload) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("Master_Pengawas");
    var id = "PGW-" + new Date().getTime();
    sheet.appendRow([id, String(payload.nama).toUpperCase(), payload.nip, payload.pangkat, payload.jabatan]);
    return {success: true, message: "Data Pengawas ditambahkan."};
  } catch (e) { return {success: false, message: e.message}; }
}

function hapusMasterPengawas(id) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("Master_Pengawas");
    var data = sheet.getDataRange().getDisplayValues();
    for(var i = data.length - 1; i >= 1; i--) {
      if(String(data[i][0]).trim() === String(id).trim()) {
        sheet.deleteRow(i + 1); return {success: true, message: "Data terhapus."};
      }
    }
    return {success: false, message: "ID tidak ditemukan."};
  } catch(e) { return {success: false, message: e.message}; }
}

function getMasterRekening() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_Rekening");
    if (!sheet) return [];
    var data = sheet.getDataRange().getDisplayValues(); 
    var result = [];
    for (var i = 1; i < data.length; i++) {
      if(data[i][0]) result.push({ Kode: String(data[i][0]), Uraian: String(data[i][1]) });
    }
    return result;
  } catch(e) { return []; }
}

function simpanMasterRekening(kode, uraian) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_Rekening");
    if (!sheet) {
      sheet = ss.insertSheet("DB_Rekening");
      sheet.getRange("A1:B1").setValues([["Kode_Rekening", "Uraian_Rekening"]]).setBackground("#1e293b").setFontColor("white");
    }
    sheet.appendRow(["'" + kode, uraian]);
    if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).sort(1);
    SpreadsheetApp.flush();
    return { success: true, message: "Kode Rekening berhasil ditambahkan." };
  } catch(e) { return { success: false, message: e.message }; }
}

function hapusMasterRekening(kode) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_Rekening");
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).trim() === String(kode).trim()) {
        sheet.deleteRow(i + 1); SpreadsheetApp.flush();
        return { success: true, message: "Kode Rekening dihapus." };
      }
    }
    return { success: false, message: "Kode tidak ditemukan." };
  } catch(e) { return { success: false, message: e.message }; }
}

function editMasterRekening(kodeLama, kodeBaru, uraianBaru) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_Rekening");
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).trim() === String(kodeLama).trim()) {
        sheet.getRange(i + 1, 1).setValue("'" + kodeBaru);
        sheet.getRange(i + 1, 2).setValue(uraianBaru);
        if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).sort(1);
        SpreadsheetApp.flush();
        return { success: true, message: "Data rekening berhasil diperbarui." };
      }
    }
    return { success: false, message: "Data rekening lama tidak ditemukan." };
  } catch(e) { return { success: false, message: e.message }; }
}

// =======================================================================
// 5. MODUL KEUANGAN & DPA
// =======================================================================
function perbaruiHeaderKeuangan() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName("Sheet_Keuangan");
  if (!sheet) sheet = ss.insertSheet("Sheet_Keuangan");
  var headers = ["ID_Transaksi", "Bulan", "Waktu_Pelaksanaan", "Program", "Kegiatan", "Sub_Kegiatan", "Jumlah_Anggaran", "Sifat_Kegiatan", "Target_Hasil", "Target_Manfaat", "Target_Dampak", "Kode_Lvl_1", "Uraian_Lvl_1", "Pagu_Lvl_1", "Kode_Lvl_2", "Uraian_Lvl_2", "Pagu_Lvl_2", "Kode_Lvl_3", "Uraian_Lvl_3", "Pagu_Lvl_3", "Kode_Lvl_4", "Uraian_Lvl_4", "Pagu_Lvl_4", "Kode_Lvl_5", "Uraian_Lvl_5", "Pagu_Lvl_5", "Kode_Lvl_6", "Uraian_Lvl_6", "Pagu_Lvl_6", "Pagu_Dana", "Target_Volume", "Satuan_Volume", "Target_SP2D_Rp", "Realisasi_SPJ_Rp", "Realisasi_Fisik_Vol", "Satuan_Fisik", "Sisa_Anggaran", "Data_Kontrak", "Keterangan", "Tanggal_Input", "Keterangan_Formula"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#0f172a").setFontColor("#00f3ff").setHorizontalAlignment("center");
  sheet.setFrozenRows(1);
}

function getTableData(sheetName) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { error: true, message: "Sheet tidak ditemukan" };
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return []; 
    var headers = data[0]; var result = [];
    for (var i = 1; i < data.length; i++) {
      var obj = {}; for (var j = 0; j < headers.length; j++) { obj[headers[j]] = data[i][j]; } result.push(obj);
    }
    return result;
  } catch (e) { return { error: true, message: e.message }; }
}

function simpanKeuanganDenganFile(payload, files) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("Sheet_Keuangan");
    if (!sheet) return { success: false, message: "Tab 'Sheet_Keuangan' tidak ditemukan!" };
    var fileUrls = [];
    if (files && files.length > 0) {
      var folderId = "1zbXFoNSxBX65qw3HaUIHCJlBTGg3atXg"; 
      var folder = DriveApp.getFolderById(folderId);
      for (var i = 0; i < files.length; i++) {
        var blob = Utilities.newBlob(Utilities.base64Decode(files[i].base64), files[i].mimeType, files[i].name);
        var file = folder.createFile(blob); fileUrls.push(file.getUrl());
      }
    }
    var urlLampiran = fileUrls.join("\n");
    var rowData = [ payload.ID_Transaksi, payload.Bulan, payload.Waktu_Pelaksanaan, payload.Program, payload.Kegiatan, payload.Sub_Kegiatan, payload.Jumlah_Anggaran, payload.Sifat_Kegiatan, payload.Target_Hasil, payload.Target_Manfaat, payload.Target_Dampak, payload.Kode_Lvl_1, payload.Uraian_Lvl_1, payload.Pagu_Lvl_1, payload.Kode_Lvl_2, payload.Uraian_Lvl_2, payload.Pagu_Lvl_2, payload.Kode_Lvl_3, payload.Uraian_Lvl_3, payload.Pagu_Lvl_3, payload.Kode_Lvl_4, payload.Uraian_Lvl_4, payload.Pagu_Lvl_4, payload.Kode_Lvl_5, payload.Uraian_Lvl_5, payload.Pagu_Lvl_5, payload.Kode_Lvl_6, payload.Uraian_Lvl_6, payload.Pagu_Lvl_6, payload.Pagu_Dana, payload.Target_Volume, payload.Satuan_Volume, payload.Target_SP2D_Rp, payload.Realisasi_SPJ_Rp, payload.Realisasi_Fisik_Vol, payload.Satuan_Fisik, payload.Sisa_Anggaran, payload.Data_Kontrak, payload.Keterangan, urlLampiran, payload.Tanggal_Input ];
    sheet.appendRow(rowData); SpreadsheetApp.flush(); recalculasiSisaSheet();
    return { success: true, message: "Data Keuangan berhasil disimpan." };
  } catch (error) { return { success: false, message: error.message }; }
}

function deleteDataKeuangan(id) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("Sheet_Keuangan");
    var data = sheet.getDataRange().getDisplayValues(); 
    var idIndex = data[0].indexOf('Tanggal_Input');
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][idIndex]).trim() === String(id).trim()) { 
        sheet.deleteRow(i + 1); SpreadsheetApp.flush(); recalculasiSisaSheet(); return { success: true, message: "Data Keuangan berhasil dihapus." }; 
      }
    }
    return { success: false, message: "Data tidak ditemukan." };
  } catch (e) { return { success: false, message: e.message }; }
}

function deleteDataKeuanganSpesifik(sub, kode, realisasi, tgl) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("Sheet_Keuangan");
    var data = sheet.getDataRange().getDisplayValues(); 
    var headers = data[0];
    var idxSub = headers.indexOf('Sub_Kegiatan'); var idxReal = headers.indexOf('Realisasi_SPJ_Rp'); var idxTgl = headers.indexOf('Tanggal_Input');

    function getDeepestRekDB(rowArr) {
        for (let j = 6; j >= 1; j--) { let k = rowArr[headers.indexOf('Kode_Lvl_' + j)]; if (k && String(k).trim() !== '' && String(k).trim() !== '-') return String(k).trim(); } return '-';
    }
    for (var i = data.length - 1; i >= 1; i--) {
      var rowSub = String(data[i][idxSub]).trim(); var rowKode = getDeepestRekDB(data[i]);
      var rawReal = String(data[i][idxReal] || "0").replace(/[^0-9,-]+/g, ""); var rowReal = Number(rawReal) || 0; var rowTgl = String(data[i][idxTgl]).trim();
      if (rowSub === String(sub).trim() && rowKode === String(kode).trim() && rowReal === Number(realisasi) && rowTgl === String(tgl).trim()) { 
        sheet.deleteRow(i + 1); SpreadsheetApp.flush(); recalculasiSisaSheet(); return { success: true, message: "Data Transaksi berhasil dihapus secara permanen." }; 
      }
    }
    return { success: false, message: "Data transaksi tidak ditemukan di database." };
  } catch (e) { return { success: false, message: e.message }; }
}

function hapusDataRekeningSpesifik(namaSubKegiatan, kodeRekening) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("Sheet_Keuangan");
    var data = sheet.getDataRange().getValues();
    var headers = data[0]; var idxSub = headers.indexOf('Sub_Kegiatan');
    function getDeepestRekDB(row) {
        for (let j = 6; j >= 1; j--) { let kode = row[headers.indexOf('Kode_Lvl_' + j)]; if (kode && String(kode).trim() !== '' && String(kode).trim() !== '-') return String(kode).trim(); } return '-';
    }
    var isDeleted = false;
    for (var i = data.length - 1; i >= 1; i--) {
        if (String(data[i][idxSub]).trim() === String(namaSubKegiatan).trim()) {
            if (getDeepestRekDB(data[i]) === String(kodeRekening).trim()) { sheet.deleteRow(i + 1); isDeleted = true; }
        }
    }
    if (isDeleted) { SpreadsheetApp.flush(); return { success: true, message: "Rekening " + kodeRekening + " dan transaksinya berhasil dihapus!" }; }
    return { success: false, message: "Data rekening tidak ditemukan." };
  } catch(e) { return { success: false, message: "Error Server: " + e.message }; }
}

function hapusSubKegiatanDPA(namaSubKegiatan) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("Sheet_Keuangan");
    var data = sheet.getDataRange().getValues();
    var idIndex = data[0].indexOf('Sub_Kegiatan');
    var adaYangDihapus = false;
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][idIndex]).trim() === String(namaSubKegiatan).trim()) { sheet.deleteRow(i + 1); adaYangDihapus = true; }
    }
    if (adaYangDihapus) { SpreadsheetApp.flush(); recalculasiSisaSheet(); return { success: true, message: "Seluruh data pada Sub Kegiatan tersebut berhasil dihapus." }; }
    return { success: false, message: "Data Sub Kegiatan tidak ditemukan." };
  } catch (e) { return { success: false, message: e.message }; }
}

function editSubKegiatanDPA(namaLama, payload) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("Sheet_Keuangan");
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idxDPA = headers.indexOf('ID_Transaksi'); var idxWaktu = headers.indexOf('Waktu_Pelaksanaan'); var idxProg = headers.indexOf('Program'); var idxKeg = headers.indexOf('Kegiatan'); var idxSub = headers.indexOf('Sub_Kegiatan'); var idxPagu = headers.indexOf('Jumlah_Anggaran'); var idxHasil = headers.indexOf('Target_Hasil'); var idxManfaat = headers.indexOf('Target_Manfaat'); var idxDampak = headers.indexOf('Target_Dampak'); var idxSifat = headers.indexOf('Sifat_Kegiatan');
    var isUpdated = false;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idxSub]).trim() === String(namaLama).trim()) {
        if(idxDPA > -1) sheet.getRange(i + 1, idxDPA + 1).setValue(payload.ID_Transaksi);
        if(idxWaktu > -1) sheet.getRange(i + 1, idxWaktu + 1).setValue(payload.Waktu_Pelaksanaan);
        if(idxProg > -1) sheet.getRange(i + 1, idxProg + 1).setValue(payload.Program);
        if(idxKeg > -1) sheet.getRange(i + 1, idxKeg + 1).setValue(payload.Kegiatan);
        if(idxSub > -1) sheet.getRange(i + 1, idxSub + 1).setValue(payload.Sub_Kegiatan);
        if(idxPagu > -1) sheet.getRange(i + 1, idxPagu + 1).setValue(Number(payload.Jumlah_Anggaran));
        if(idxHasil > -1) sheet.getRange(i + 1, idxHasil + 1).setValue(payload.Target_Hasil);
        if(idxManfaat > -1) sheet.getRange(i + 1, idxManfaat + 1).setValue(payload.Target_Manfaat);
        if(idxDampak > -1) sheet.getRange(i + 1, idxDampak + 1).setValue(payload.Target_Dampak);
        if(idxSifat > -1) sheet.getRange(i + 1, idxSifat + 1).setValue(payload.Sifat_Kegiatan);
        isUpdated = true;
      }
    }
    if (isUpdated) { SpreadsheetApp.flush(); recalculasiSisaSheet(); return { success: true, message: "Seluruh Data Umum Sub Kegiatan berhasil diperbarui." }; }
    return { success: false, message: "Data tidak ditemukan di database." };
  } catch (e) { return { success: false, message: e.message }; }
}

function editPaguDanRealisasiRekening(namaSubKegiatan, kodeRekening, paguBaru, realisasiBaru, keteranganBaru, payloadRekBaru) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("Sheet_Keuangan");
    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    var idxSub = headers.indexOf('Sub_Kegiatan'); var idxPaguDana = headers.indexOf('Pagu_Dana'); var idxRealisasi = headers.indexOf('Realisasi_SPJ_Rp'); var idxBulan = headers.indexOf('Bulan'); var idxKet = headers.indexOf('Keterangan');
    var idxKode1 = headers.indexOf('Kode_Lvl_1'); var idxUraian1 = headers.indexOf('Uraian_Lvl_1'); var idxKode2 = headers.indexOf('Kode_Lvl_2'); var idxUraian2 = headers.indexOf('Uraian_Lvl_2'); var idxKode3 = headers.indexOf('Kode_Lvl_3'); var idxUraian3 = headers.indexOf('Uraian_Lvl_3'); var idxKode4 = headers.indexOf('Kode_Lvl_4'); var idxUraian4 = headers.indexOf('Uraian_Lvl_4'); var idxKode5 = headers.indexOf('Kode_Lvl_5'); var idxUraian5 = headers.indexOf('Uraian_Lvl_5'); var idxKode6 = headers.indexOf('Kode_Lvl_6'); var idxUraian6 = headers.indexOf('Uraian_Lvl_6');

    var isUpdated = false; var realisasiSudahDiupdate = false;

    function getDeepestRekDB(row) {
        for (let j = 6; j >= 1; j--) { let kode = row[headers.indexOf('Kode_Lvl_' + j)]; let uraian = row[headers.indexOf('Uraian_Lvl_' + j)]; if (kode && String(kode).trim() !== '' && String(kode).trim() !== '-') return { kode: String(kode).trim(), uraian: String(uraian).trim() }; } return { kode: '-', uraian: '-' };
    }

    for (var i = data.length - 1; i >= 1; i--) {
        if (String(data[i][idxSub]).trim() === String(namaSubKegiatan).trim()) {
            var deepRek = getDeepestRekDB(data[i]);
            if (deepRek.kode === kodeRekening) {
                if (data[i][idxBulan] === '-') {
                    sheet.getRange(i + 1, idxPaguDana + 1).setValue(Number(paguBaru));
                    if(idxKet > -1) sheet.getRange(i + 1, idxKet + 1).setValue(keteranganBaru);
                    if (!realisasiSudahDiupdate && Number(realisasiBaru) > 0) { sheet.getRange(i + 1, idxRealisasi + 1).setValue(Number(realisasiBaru)); realisasiSudahDiupdate = true; }
                    if (payloadRekBaru) {
                        if(idxKode1 > -1) sheet.getRange(i+1, idxKode1+1).setValue(payloadRekBaru.Kode_Lvl_1); if(idxUraian1 > -1) sheet.getRange(i+1, idxUraian1+1).setValue(payloadRekBaru.Uraian_Lvl_1);
                        if(idxKode2 > -1) sheet.getRange(i+1, idxKode2+1).setValue(payloadRekBaru.Kode_Lvl_2); if(idxUraian2 > -1) sheet.getRange(i+1, idxUraian2+1).setValue(payloadRekBaru.Uraian_Lvl_2);
                        if(idxKode3 > -1) sheet.getRange(i+1, idxKode3+1).setValue(payloadRekBaru.Kode_Lvl_3); if(idxUraian3 > -1) sheet.getRange(i+1, idxUraian3+1).setValue(payloadRekBaru.Uraian_Lvl_3);
                        if(idxKode4 > -1) sheet.getRange(i+1, idxKode4+1).setValue(payloadRekBaru.Kode_Lvl_4); if(idxUraian4 > -1) sheet.getRange(i+1, idxUraian4+1).setValue(payloadRekBaru.Uraian_Lvl_4);
                        if(idxKode5 > -1) sheet.getRange(i+1, idxKode5+1).setValue(payloadRekBaru.Kode_Lvl_5); if(idxUraian5 > -1) sheet.getRange(i+1, idxUraian5+1).setValue(payloadRekBaru.Uraian_Lvl_5);
                        if(idxKode6 > -1) sheet.getRange(i+1, idxKode6+1).setValue(payloadRekBaru.Kode_Lvl_6); if(idxUraian6 > -1) sheet.getRange(i+1, idxUraian6+1).setValue(payloadRekBaru.Uraian_Lvl_6);
                    }
                    isUpdated = true;
                } else if (data[i][idxBulan] !== '-') {
                    if (!realisasiSudahDiupdate && Number(realisasiBaru) > 0) { sheet.getRange(i + 1, idxRealisasi + 1).setValue(Number(realisasiBaru)); realisasiSudahDiupdate = true; }
                    if (payloadRekBaru) {
                        if(idxKode1 > -1) sheet.getRange(i+1, idxKode1+1).setValue(payloadRekBaru.Kode_Lvl_1); if(idxUraian1 > -1) sheet.getRange(i+1, idxUraian1+1).setValue(payloadRekBaru.Uraian_Lvl_1);
                        if(idxKode2 > -1) sheet.getRange(i+1, idxKode2+1).setValue(payloadRekBaru.Kode_Lvl_2); if(idxUraian2 > -1) sheet.getRange(i+1, idxUraian2+1).setValue(payloadRekBaru.Uraian_Lvl_2);
                        if(idxKode3 > -1) sheet.getRange(i+1, idxKode3+1).setValue(payloadRekBaru.Kode_Lvl_3); if(idxUraian3 > -1) sheet.getRange(i+1, idxUraian3+1).setValue(payloadRekBaru.Uraian_Lvl_3);
                        if(idxKode4 > -1) sheet.getRange(i+1, idxKode4+1).setValue(payloadRekBaru.Kode_Lvl_4); if(idxUraian4 > -1) sheet.getRange(i+1, idxUraian4+1).setValue(payloadRekBaru.Uraian_Lvl_4);
                        if(idxKode5 > -1) sheet.getRange(i+1, idxKode5+1).setValue(payloadRekBaru.Kode_Lvl_5); if(idxUraian5 > -1) sheet.getRange(i+1, idxUraian5+1).setValue(payloadRekBaru.Uraian_Lvl_5);
                        if(idxKode6 > -1) sheet.getRange(i+1, idxKode6+1).setValue(payloadRekBaru.Kode_Lvl_6); if(idxUraian6 > -1) sheet.getRange(i+1, idxUraian6+1).setValue(payloadRekBaru.Uraian_Lvl_6);
                    }
                    isUpdated = true;
                }
            }
        }
    }

    if (isUpdated) { SpreadsheetApp.flush(); recalculasiSisaSheet(); return { success: true, message: "Pagu, Realisasi, dan Rincian berhasil diperbarui." }; }
    return { success: false, message: "Data Rekening tidak ditemukan di Database." };
  } catch(e) { return { success: false, message: e.message }; }
}

function recalculasiSisaSheet() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Sheet_Keuangan');
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    var headers = data[0];
    var rows = data.slice(1);
    var urutanBulan = { "JANUARI":1,"FEBRUARI":2,"MARET":3,"APRIL":4,"MEI":5,"JUNI":6,"JULI":7,"AGUSTUS":8,"SEPTEMBER":9,"OKTOBER":10,"NOVEMBER":11,"DESEMBER":12 };

    rows.sort((a, b) => {
      let blnA = urutanBulan[String(a[headers.indexOf("Bulan")]).toUpperCase()] || 99;
      let blnB = urutanBulan[String(b[headers.indexOf("Bulan")]).toUpperCase()] || 99;
      return blnA - blnB;
    });

    let akumulasiRealisasi = {}; 
    let dataTerhitung = rows.map(row => {
      let uraian = row[headers.indexOf("Uraian_Lvl_6")] || row[headers.indexOf("Uraian_Lvl_1")];
      let pagu = Number(row[headers.indexOf("Pagu_Dana")]) || 0;
      let realisasi = Number(row[headers.indexOf("Realisasi_SPJ_Rp")]) || 0;
      
      if (!akumulasiRealisasi[uraian]) akumulasiRealisasi[uraian] = 0;
      akumulasiRealisasi[uraian] += realisasi;
      row[headers.indexOf("Sisa_Anggaran")] = pagu - akumulasiRealisasi[uraian];
      return row;
    });

    sheet.getRange(2, 1, dataTerhitung.length, dataTerhitung[0].length).setValues(dataTerhitung);
    SpreadsheetApp.flush();
  } catch (e) {}
}

function getMergedPdfBase64(bulan, namaUser, nipUser) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheetCetak = ss.getSheetByName('Cetak_FormA');
    var sheetCover = ss.getSheetByName('Cover_FormA');
    if (!sheetCetak || !sheetCover) return { error: true, message: "Tab Cetak atau Cover tidak ditemukan!" };
    if (bulan) { sheetCetak.getRange("R18").setValue(bulan); SpreadsheetApp.flush(); Utilities.sleep(1500); }

    var token = ScriptApp.getOAuthToken();
    var baseUrl = "[https://docs.google.com/spreadsheets/d/](https://docs.google.com/spreadsheets/d/)" + ss.getId() + "/export?exportFormat=pdf&format=pdf";
    var urlCover = baseUrl + "&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=false&gid=" + sheetCover.getSheetId();
    var blobCover = UrlFetchApp.fetch(urlCover, { headers: { 'Authorization': 'Bearer ' + token } }).getBlob();
    var urlLaporan = baseUrl + "&size=A4&portrait=false&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=false&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5&gid=" + sheetCetak.getSheetId();
    var blobLaporan = UrlFetchApp.fetch(urlLaporan, { headers: { 'Authorization': 'Bearer ' + token } }).getBlob();

    return { error: false, coverBase64: Utilities.base64Encode(blobCover.getBytes()), laporanBase64: Utilities.base64Encode(blobLaporan.getBytes()), filename: "Laporan_Keuangan_FormA_" + bulan + ".pdf" };
  } catch (error) { return { error: true, message: error.message }; }
}

function listKegiatan(bulan) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('DB_Reakeu');
    if (!sheet) return [];
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return [];
    
    var map = {};
    for (var i = 1; i < data.length; i++) {
      var id = data[i][0];
      if (!id || (bulan && data[i][1] !== bulan)) continue;
      if (!map[id]) { map[id] = { idKegiatan: id, bulan: data[i][1], program: data[i][2], kegiatan: data[i][3], subKegiatan: data[i][4], jumlahAnggaran: data[i][5] }; }
    }
    return Object.values(map);
  } catch (e) { return []; }
}

function saveKegiatan(payload) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('DB_Reakeu');
    if (!sheet) {
      sheet = ss.insertSheet('DB_Reakeu');
      sheet.appendRow(['ID_Kegiatan','Bulan','Program','Kegiatan','Sub_Kegiatan','Jumlah_Anggaran','Tolak_Ukur','Hasil','Manfaat','Dampak','Sifat_Kegiatan','Waktu_Pelaksanaan','Username','Uraian_Item','Pagu_Item','Volume_Item','Satuan_Item','Target_SP2D','Realisasi_SPJ','Fisik_Volume','Nama_Perusahaan','Nilai_Kontrak','Tgl_Mulai','Masa_Kontrak','No_Kontrak','Keterangan']);
    }
    
    var idKegiatan = 'KEG-' + new Date().getTime();
    var rowsToAppend = [];
    for (var i = 0; i < payload.items.length; i++) {
      var item = payload.items[i];
      rowsToAppend.push([ idKegiatan, payload.bulan, payload.program, payload.kegiatan, payload.subKegiatan, payload.jumlahAnggaran, payload.tolakUkur, payload.hasil, payload.manfaat, payload.dampak, payload.sifat, payload.waktuPelaksanaan, payload.username, item.uraian, item.pagu, item.volume, item.satuan, item.targetSp2d, item.realisasiSpj, item.realisasiFisikVolume, item.perusahaan, item.nilaiKontrak, item.tanggalMulaiKerja, item.masaKontrak, item.noKontrak, item.keterangan ]);
    }
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    SpreadsheetApp.flush();
    return { success: true, message: "Kegiatan berhasil disimpan.", idKegiatan: idKegiatan };
  } catch (e) { return { success: false, message: e.message }; }
}

function getKegiatan(idKegiatan) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('DB_Reakeu');
    if (!sheet) return null;
    var data = sheet.getDataRange().getDisplayValues();
    var result = null;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === idKegiatan) {
        if (!result) { result = { idKegiatan: data[i][0], bulan: data[i][1], program: data[i][2], kegiatan: data[i][3], subKegiatan: data[i][4], tolakUkur: data[i][6], items: [] }; }
        result.items.push({ uraian: data[i][13], pagu: data[i][14], volume: data[i][15], satuan: data[i][16], targetSp2d: data[i][17], realisasiSpj: data[i][18], realisasiFisikVolume: data[i][19] });
      }
    }
    return result;
  } catch (e) { return null; }
}

function deleteKegiatan(idKegiatan) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('DB_Reakeu');
    if (!sheet) return { success: false, message: 'Sheet tidak ditemukan.' };
    var data = sheet.getDataRange().getValues();
    var rowsDeleted = 0;
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === idKegiatan) { sheet.deleteRow(i + 1); rowsDeleted++; }
    }
    SpreadsheetApp.flush();
    return { success: true, message: rowsDeleted + ' baris item kegiatan berhasil dihapus.' };
  } catch (e) { return { success: false, message: e.message }; }
}

// =======================================================================
// 6. MODUL BMD (ASET BARANG) & KIR
// =======================================================================
function setupHeaderBMD() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_BMD_ID);
    var sheet = ss.getSheetByName("DB_BMD");
    if (!sheet) return;
    var headers = ["No", "Nama_Barang", "Merk_Model", "No_Seri", "Ukuran", "Bahan", "Tahun_Perolehan", "Kode_Barang", "Jumlah", "Harga_Perolehan", "Kondisi_B", "Kondisi_KB", "Kondisi_RB", "Keterangan", "Foto_Aset"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("left");
    SpreadsheetApp.flush();
  } catch (e) {}
}

function simpanDataBMDDenganFoto(payload) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_BMD_ID);
    var sheet = ss.getSheetByName("DB_BMD");
    if (!sheet) return { success: false, message: "Tab 'DB_BMD' tidak ditemukan di database!" };
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idxFoto = headers.indexOf("Foto_Aset");
    if (idxFoto === -1) { idxFoto = headers.length; sheet.getRange(1, idxFoto + 1).setValue("Foto_Aset"); }

    var nomorUrut = 1; 
    if (data.length > 1) { var noTerakhir = data[data.length - 1][0]; if (!isNaN(noTerakhir) && noTerakhir !== "") nomorUrut = Number(noTerakhir) + 1; }

    var urlFotoBaru = "";
    if (payload.base64) {
      var folderId = "1c-Pj3ybM_-4-rV11-xjiWHmiCWv5UcoR"; 
      var folder = DriveApp.getFolderById(folderId);
      var blob = Utilities.newBlob(Utilities.base64Decode(payload.base64), payload.mimeType, payload.fileName);
      var file = folder.createFile(blob); file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); urlFotoBaru = file.getUrl();
    }

    var barisDataBaru = []; for(var k=0; k<=idxFoto; k++) barisDataBaru.push("");
    barisDataBaru[0] = nomorUrut; barisDataBaru[1] = payload.Nama_Barang; barisDataBaru[2] = payload.Merk_Model; barisDataBaru[3] = payload.No_Seri; barisDataBaru[4] = payload.Ukuran; barisDataBaru[5] = payload.Bahan; barisDataBaru[6] = payload.Tahun_Perolehan; barisDataBaru[7] = payload.Kode_Barang; barisDataBaru[8] = payload.Jumlah; barisDataBaru[9] = payload.Harga_Perolehan; barisDataBaru[10] = payload.Kondisi_B; barisDataBaru[11] = payload.Kondisi_KB; barisDataBaru[12] = payload.Kondisi_RB; barisDataBaru[13] = payload.Keterangan; barisDataBaru[idxFoto] = urlFotoBaru;

    sheet.appendRow(barisDataBaru); SpreadsheetApp.flush();
    return { success: true, message: "Data Inventaris BMD berhasil disimpan!" };
  } catch (error) { return { success: false, message: error.message }; }
}

function editDataBMDDenganFoto(payload) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_BMD_ID);
    var sheet = ss.getSheetByName("DB_BMD");
    if (!sheet) return { success: false, message: "Tab DB_BMD tidak ditemukan!" };
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idxFoto = headers.indexOf("Foto_Aset");
    if (idxFoto === -1) { idxFoto = headers.length; sheet.getRange(1, idxFoto + 1).setValue("Foto_Aset"); }

    var urlFotoBaru = "";
    if (payload.base64) {
      var folderId = "1CSuHamLm9b2YodgAU0hCJpRqN1VCtLCQ"; 
      var folder = DriveApp.getFolderById(folderId);
      var blob = Utilities.newBlob(Utilities.base64Decode(payload.base64), payload.mimeType, payload.fileName);
      var file = folder.createFile(blob); file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); urlFotoBaru = file.getUrl();
    }

    var isUpdated = false;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][7]) === String(payload.Kode_Barang)) { 
        sheet.getRange(i + 1, 2).setValue(payload.Nama_Barang); sheet.getRange(i + 1, 3).setValue(payload.Merk_Model); sheet.getRange(i + 1, 4).setValue(payload.No_Seri); sheet.getRange(i + 1, 7).setValue(payload.Tahun_Perolehan); sheet.getRange(i + 1, 9).setValue(payload.Jumlah); sheet.getRange(i + 1, 10).setValue(payload.Harga_Perolehan); sheet.getRange(i + 1, 11).setValue(payload.Kondisi_B); sheet.getRange(i + 1, 12).setValue(payload.Kondisi_KB); sheet.getRange(i + 1, 13).setValue(payload.Kondisi_RB); 
        if (urlFotoBaru !== "") sheet.getRange(i + 1, idxFoto + 1).setValue(urlFotoBaru);
        isUpdated = true; break;
      }
    }
    if (isUpdated) { SpreadsheetApp.flush(); return { success: true, message: "Aset BMD berhasil diperbarui." }; } 
    else { return { success: false, message: "Kode Aset tidak ditemukan di database." }; }
  } catch (e) { return { success: false, message: e.message }; }
}

function getTableDataBMD() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_BMD_ID);
    var sheet = ss.getSheetByName("DB_BMD");
    if (!sheet) return { error: true, message: "Tab DB_BMD tidak ditemukan." };
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return []; 
    var headers = data[0]; var result = [];
    for (var i = 1; i < data.length; i++) {
      var obj = {}; for (var j = 0; j < headers.length; j++) { obj[headers[j]] = data[i][j]; } result.push(obj);
    }
    return result;
  } catch (e) { return { error: true, message: e.message }; }
}

function deleteDataBMD(kodeBarang, nip, token) {
  let isAuthorized = validasiOtoritas(nip, token, ["SUPER ADMIN", "ADMIN"]);
  if (!isAuthorized) return { success: false, message: "Akses Ditolak! Sesi Anda tidak valid atau Anda tidak memiliki wewenang." };
  try {
    var ss = SpreadsheetApp.openById(SHEET_BMD_ID);
    var sheet = ss.getSheetByName("DB_BMD");
    if (!sheet) return { success: false, message: "Tab DB_BMD tidak ditemukan!" };
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][7] == kodeBarang) { 
        sheet.deleteRow(i + 1); SpreadsheetApp.flush(); return { success: true, message: "Data Aset berhasil dihapus." };
      }
    }
    return { success: false, message: "Data tidak ditemukan." };
  } catch (error) { return { success: false, message: "Gagal menghapus: " + error.message }; }
}

function getMergedKIRPdfBase64(bulanPilihan) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_BMD_ID);
    var sheetCover = ss.getSheetByName('Cover_KIR');
    var sheetKIR = ss.getSheetByName('Cetak_KIR'); 
    var sheetBMD = ss.getSheetByName('DB_BMD');

    if (!sheetCover || !sheetKIR || !sheetBMD) return { error: true, message: "Data sheet untuk KIR tidak lengkap!" };

    if (bulanPilihan) {
      sheetCover.getRange("A30:I30").setValue("PERIODE " + bulanPilihan.toUpperCase());
      sheetKIR.getRange("A4").setValue("PER " + bulanPilihan.toUpperCase()); 
    }

    var lastRowKIR = sheetKIR.getLastRow();
    if (lastRowKIR >= 18) sheetKIR.getRange(18, 1, lastRowKIR - 17, 14).clearContent(); 

    var lastRowBMD = sheetBMD.getLastRow();
    if (lastRowBMD > 1) {
      var dataBMD = sheetBMD.getRange(2, 1, lastRowBMD - 1, 14).getValues();
      sheetKIR.getRange(18, 1, dataBMD.length, dataBMD[0].length).setValues(dataBMD);
    }

    lastRowKIR = sheetKIR.getLastRow();
    if (lastRowKIR >= 18) {
      var dataRange = sheetKIR.getRange("A18:N" + lastRowKIR);
      dataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      dataRange.setVerticalAlignment("top");
    }

    SpreadsheetApp.flush(); Utilities.sleep(2500); 

    var token = ScriptApp.getOAuthToken();
    var baseUrl = "[https://docs.google.com/spreadsheets/d/](https://docs.google.com/spreadsheets/d/)" + ss.getId() + "/export?exportFormat=pdf&format=pdf";
    var urlCover = baseUrl + "&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=false&gid=" + sheetCover.getSheetId();
    var blobCover = UrlFetchApp.fetch(urlCover, { headers: { 'Authorization': 'Bearer ' + token } }).getBlob();
    var urlLaporan = baseUrl + "&size=A4&portrait=false&fitw=true&gridlines=true&printtitle=false&sheetnames=false&pagenum=false&top_margin=0.4&bottom_margin=0.4&left_margin=0.4&right_margin=0.4&gid=" + sheetKIR.getSheetId();
    var blobLaporan = UrlFetchApp.fetch(urlLaporan, { headers: { 'Authorization': 'Bearer ' + token } }).getBlob();

    return { error: false, coverBase64: Utilities.base64Encode(blobCover.getBytes()), laporanBase64: Utilities.base64Encode(blobLaporan.getBytes()), filename: "KIR_BMD_" + bulanPilihan + ".pdf" };
  } catch (error) { return { error: true, message: error.message }; }
}

// =======================================================================
// 7. MODUL SURAT TUGAS (SPT)
// =======================================================================
function simpanDataSPT(payload) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_SPT");
    if (!sheet) {
      sheet = ss.insertSheet("DB_SPT");
      var headers = ["Tanggal_Input", "No_SPT", "Tgl_SPT", "Petugas", "Lokasi", "Maksud", "Lama_Hari"];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    sheet.insertRowAfter(1);
    sheet.getRange(2, 1, 1, 7).setValues([[ payload.Tanggal_Input, payload.No_SPT, payload.Tgl_SPT, payload.Petugas, payload.Lokasi, payload.Maksud, payload.Lama_Hari ]]);
    SpreadsheetApp.flush();
    return { success: true, message: "Surat Tugas Berhasil Tersimpan." };
  } catch (e) { return { success: false, message: "Gagal menyimpan: " + e.message }; }
}

function getTableDataSPT() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_SPT");
    if (!sheet) return [];
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return [];
    var result = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] !== "") { 
        result.push({ Tanggal_Input: data[i][0], No_SPT: data[i][1], Tgl_SPT: data[i][2], Petugas: data[i][3], Lokasi: data[i][4], Maksud: data[i][5], Lama_Hari: data[i][6] });
      }
    }
    return result;
  } catch (e) { return []; }
}

function deleteDataSPT(id) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_SPT");
    if (!sheet) return { success: false, message: "Data DB_SPT tidak ditemukan." };
    var data = sheet.getDataRange().getDisplayValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).trim() === String(id).trim()) { 
        sheet.deleteRow(i + 1); SpreadsheetApp.flush(); return { success: true, message: "Surat Tugas berhasil dihapus." }; 
      }
    }
    return { success: false, message: "Data tidak ditemukan untuk dihapus." };
  } catch (e) { return { success: false, message: e.message }; }
}

// =========================================================================
// 8. MODUL PENGADUAN KETENAGAKERJAAN
// =========================================================================
function simpanDataPengaduan(payload) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_Pengaduan");
    var rowData = [ payload.Tanggal_Input, payload.Nama_Pelapor, payload.Telp_Pelapor, payload.Pelapor_Ikutan, payload.Alamat_Pelapor, payload.Status_Pekerja, payload.Perusahaan, payload.Alamat_Perusahaan, payload.Pokok_Pengaduan, payload.Kronologi, payload.Proses_Penanganan, payload.Tingkat_Urgensi, payload.Verifikasi, payload.Disposisi, payload.Status ];
    sheet.appendRow(rowData);
    return { success: true, message: "Data pengaduan tersimpan rapi ke 15 Kolom." };
  } catch (error) { return { success: false, message: "Gagal menyimpan: " + error.message }; }
}

function getTableDataPengaduan() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_Pengaduan");
    var data = sheet.getDataRange().getDisplayValues();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      result.push({ Tanggal_Input: data[i][0], Nama_Pelapor: data[i][1], Telp_Pelapor: data[i][2], Pelapor_Ikutan: data[i][3], Alamat_Pelapor: data[i][4], Status_Pekerja: data[i][5], Perusahaan: data[i][6], Alamat_Perusahaan: data[i][7], Pokok_Pengaduan: data[i][8], Kronologi: data[i][9], Proses_Penanganan: data[i][10], Tingkat_Urgensi: data[i][11], Verifikasi: data[i][12], Disposisi: data[i][13], Status: data[i][14] });
    }
    return result;
  } catch (error) { return []; }
}

function deleteDataPengaduan(idTanggal) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_Pengaduan");
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(idTanggal)) { sheet.deleteRow(i + 1); return { success: true }; }
    }
    return { success: false, message: "ID Data tidak ditemukan" };
  } catch (error) { return { success: false, message: error.message }; }
}

function updateProgresPengaduan(payload) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_Pengaduan"); 
    if (!sheet) return { success: false, message: "Tab DB_Pengaduan tidak ditemukan!" };
    var data = sheet.getDataRange().getValues();
    var isUpdated = false;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(payload.id)) { 
        sheet.getRange(i + 1, 8).setValue(payload.progres); sheet.getRange(i + 1, 10).setValue(payload.telaah); isUpdated = true; break;
      }
    }
    if (isUpdated) { SpreadsheetApp.flush(); return { success: true, message: "Status penanganan berhasil di-update." }; }
    return { success: false, message: "Data referensi pengaduan tidak ditemukan." };
  } catch (e) { return { success: false, message: "Error Server: " + e.message }; }
}

// =======================================================================
// 9. MODUL ABSENSI BULANAN
// =======================================================================
function importDataAbsensi(data2D) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_Absensi");
    if (!sheet) return { success: false, message: "Tab DB_Absensi tidak ditemukan!" };

    var dataUntukDisimpan = data2D.slice(1);
    if (dataUntukDisimpan.length === 0) return { success: false, message: "File kosong." };

    var existingData = sheet.getDataRange().getValues();
    var mapExisting = {};
    var numCols = Math.max(existingData[0].length, data2D[0].length);

    function buatKunciNama(namaRaw) {
        if (!namaRaw) return "";
        var namaBersih = String(namaRaw).toLowerCase().split(',')[0];
        namaBersih = namaBersih.replace(/[^a-z0-9]/g, '');
        namaBersih = namaBersih.replace(/ff/g, 'f');
        return namaBersih;
    }

    function ekstrakTahun(tglRaw) {
        var str = String(tglRaw).trim();
        var match = str.match(/\d{4}/);
        return match ? match[0] : "";
    }

    if (existingData.length > 1) {
      for (var i = 1; i < existingData.length; i++) {
        var row = existingData[i];
        var nama = buatKunciNama(row[1]);
        var periode = String(row[2] || "").trim().toLowerCase();
        
        var tahunDari = ekstrakTahun(row[3]); 
        var tahunSampai = ekstrakTahun(row[4]);
        var tahunFix = tahunDari || tahunSampai || new Date().getFullYear().toString();
        
        var key = nama + "|" + periode + "|" + tahunFix;
        mapExisting[key] = i; 
      }
    }

    var newRows = [];
    var updatedCount = 0;

    for (var j = 0; j < dataUntukDisimpan.length; j++) {
      var incomingRow = dataUntukDisimpan[j];
      while (incomingRow.length < numCols) { incomingRow.push(""); }
      
      var incNama = buatKunciNama(incomingRow[1]);
      var incPeriode = String(incomingRow[2] || "").trim().toLowerCase();
      var incTahunDari = ekstrakTahun(incomingRow[3]);
      var incTahunSampai = ekstrakTahun(incomingRow[4]);
      var incTahunFix = incTahunDari || incTahunSampai || new Date().getFullYear().toString();
      
      var incKey = incNama + "|" + incPeriode + "|" + incTahunFix;

      if (mapExisting.hasOwnProperty(incKey)) {
        var existingIndex = mapExisting[incKey];
        if (existingIndex < existingData.length) {
            for (var col = 0; col < incomingRow.length; col++) { 
                existingData[existingIndex][col] = incomingRow[col]; 
            }
            updatedCount++;
        } else {
            var newRowIndex = existingIndex - existingData.length;
            newRows[newRowIndex] = incomingRow;
        }
      } else {
        newRows.push(incomingRow);
        mapExisting[incKey] = existingData.length + (newRows.length - 1);
      }
    }

    if (existingData.length > 1 && updatedCount > 0) {
      existingData.forEach(r => { while(r.length < numCols) r.push(""); });
      sheet.getRange(1, 1, existingData.length, numCols).setValues(existingData);
    }

    if (newRows.length > 0) {
      var startRow = existingData.length + 1;
      sheet.getRange(startRow, 1, newRows.length, numCols).setValues(newRows);
    }

    SpreadsheetApp.flush();
    return { success: true, message: "Sinkronisasi selesai! " + updatedCount + " data diperbarui, " + newRows.length + " data baru ditambahkan." };
  } catch (e) { return { success: false, message: "Gagal import: " + e.message }; }
}

function getAbsensiData() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_Absensi");
    return sheet ? sheet.getDataRange().getDisplayValues() : [];
  } catch (e) { return []; }
}

function hapusDataAbsensi(bulan, tahun) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_Absensi");
    if (!sheet) return { success: false, message: "Tab DB_Absensi tidak ditemukan!" };

    var data = sheet.getDataRange().getDisplayValues();
    var rowsDeleted = 0;
    for (var i = data.length - 1; i >= 1; i--) { 
      if ((bulan === "" || String(data[i][2]).toLowerCase().indexOf(bulan.toLowerCase()) > -1) && 
          (tahun === "" || String(data[i][3]).indexOf(tahun) > -1 || String(data[i][4]).indexOf(tahun) > -1)) {
        sheet.deleteRow(i + 1); rowsDeleted++;
      }
    }
    SpreadsheetApp.flush();
    return rowsDeleted > 0 ? { success: true, message: "Berhasil menghapus " + rowsDeleted + " baris data." } : { success: false, message: "Data tidak cocok." };
  } catch (e) { return { success: false, message: e.message }; }
}

function getTableDataApel() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_Apel") || ss.getSheetByName("Sheet_Apel");
    if (!sheet) return [];
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return [];
    var headers = data[0];
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var obj = {};
      for (var j = 0; j < headers.length; j++) { obj[headers[j]] = data[i][j]; }
      result.push(obj);
    }
    return result;
  } catch (e) { return []; }
}

function simpanDataApelSenin(payload) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_Apel") || ss.getSheetByName("Sheet_Apel");
    if (!sheet) {
      sheet = ss.insertSheet("DB_Apel");
      sheet.appendRow(["Tanggal_Input", "Tanggal_Apel", "Foto_URL", "Daftar_Hadir_Pegawai", "Keterangan_Arahan"]);
    }

    var fileUrl = "Tanpa Foto";
    if (payload.File_Base64) {
      var blob = Utilities.newBlob(Utilities.base64Decode(payload.File_Base64), payload.File_MimeType, payload.File_Name);
      var folder;
      try { folder = DriveApp.getFolderById("1CSuHamLm9b2YodgAU0hCJpRqN1VCtLCQ"); } catch(e) { folder = DriveApp.createFolder("Dokumentasi Apel Senin"); }
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = folder.createFile(blob).getUrl();
    }

    var tglInput = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
    sheet.appendRow([tglInput, payload.Tanggal_Apel, fileUrl, payload.Daftar_Hadir, payload.Keterangan]);
    return { success: true, message: "Laporan Apel Senin berhasil disimpan!" };
  } catch (e) { return { success: false, message: "Gagal: " + e.message }; }
}

function deleteDataApel(idData) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_Apel") || ss.getSheetByName("Sheet_Apel");
    if (!sheet) return { success: false, message: "Tab tidak ditemukan." };
    
    var data = sheet.getDataRange().getDisplayValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).trim() === String(idData).trim()) {
        sheet.deleteRow(i + 1); SpreadsheetApp.flush();
        return { success: true, message: "Arsip apel dihapus." };
      }
    }
    return { success: false, message: "Data tidak ditemukan." };
  } catch(e) { return { success: false, message: e.message }; }
}

function getBase64FotosApel(urls) {
  var results = [];
  for (var i = 0; i < urls.length; i++) {
    var url = urls[i].url;
    var tgl = urls[i].tgl;
    var b64 = "";
    if (url && url.indexOf("drive.google.com") !== -1) {
       var match = url.match(/[-\w]{25,}/);
       if (match && match[0]) {
          try {
            var blob = DriveApp.getFileById(match[0]).getBlob();
            b64 = "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
          } catch(e) {}
       }
    }
    results.push({ tgl: tgl, base64: b64 });
  }
  return results;
}

// =======================================================================
// 10. MODUL K3 & KOMPETENSI (KERTAS KERJA & KKPAK)
// =======================================================================
function getDatabaseKompetensi() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_Kompetensi") || ss.getSheetByName("Sheet_Kompetensi");
    if (!sheet) {
      sheet = ss.insertSheet("DB_Kompetensi");
      sheet.appendRow(["Nama Pegawai", "NIP", "Nama Diklat / Kegiatan", "Tanggal Pelaksanaan", "Jam Pelajaran (JP)", "URL Sertifikat"]);
    }
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return []; 
    
    var hasil = [];
    for (var i = 1; i < data.length; i++) {
      hasil.push({
        nama: data[i][0], nip: data[i][1], diklat: data[i][2],
        tgl: data[i][3] ? Utilities.formatDate(new Date(data[i][3]), "GMT+7", "dd MMMM yyyy") : "-",
        jp: Number(data[i][4]) || 0, file: data[i][5] || "#"
      });
    }
    return hasil;
  } catch (error) { return []; }
}

function simpanKertasKerjaK3(payload) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_Pengawasan_K3");
    
    var lastCol = sheet.getLastColumn();
    var headers = [];
    if (lastCol > 0) {
      headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    } else {
      headers = ["ID", "Tanggal_Input", "Nama_Perusahaan", "Sektor_Usaha", "Skor_Akhir", "Status", "Data_JSON"];
      sheet.appendRow(headers);
    }

    var isNew = false;
    var idData = payload.ID_Data;
    if (!idData || idData === "") {
      idData = "K3-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmmss");
      isNew = true;
    }

    var flatData = {};
    flatData["ID"] = idData;
    flatData["Tanggal_Input"] = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
    flatData["Nama_Perusahaan"] = payload.Nama_Perusahaan || "";
    flatData["Sektor_Usaha"] = payload.Sektor_Usaha || "";
    flatData["Skor_Akhir"] = payload.Skor_Akhir || "0%";
    flatData["Status"] = payload.Status || "";

    var checklist = payload.Checklist_Data || {};
    for (var key in checklist) {
        var item = checklist[key];
        if (typeof item === 'object' && item !== null) {
            flatData[key + "_Status"] = item.Status || "";
            flatData[key + "_Keterangan"] = item.Keterangan || "";
        } else {
            flatData[key] = item;
        }
    }

    var newHeaders = [];
    for (var key in flatData) {
        if (headers.indexOf(key) === -1) {
            headers.push(key);
            newHeaders.push(key);
        }
    }
    if (newHeaders.length > 0) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    var rowDataArray = [];
    for (var i = 0; i < headers.length; i++) {
        var colName = headers[i];
        rowDataArray.push(flatData[colName] !== undefined ? flatData[colName] : "");
    }

    if (isNew) {
        sheet.appendRow(rowDataArray);
    } else {
        var data = sheet.getDataRange().getValues();
        var rowIndex = -1;
        for (var i = 1; i < data.length; i++) {
            if (String(data[i][0]).trim() === String(idData).trim()) {
                rowIndex = i + 1; break;
            }
        }
        if (rowIndex > -1) {
            var oldRow = data[rowIndex - 1];
            for (var i = 0; i < headers.length; i++) {
                if (flatData[headers[i]] === undefined && oldRow[i] !== undefined) {
                    rowDataArray[i] = oldRow[i]; 
                }
            }
            sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowDataArray]);
        } else {
            sheet.appendRow(rowDataArray);
        }
    }

    SpreadsheetApp.flush();
    return { success: true, message: "Kertas Kerja berhasil disimpan dan dipisah per kolom." };
  } catch (e) {
    return { success: false, message: "Error Server: " + e.message };
  }
}

function simpanBulkKertasKerjaK3(dataArray) {
   try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_Pengawasan_K3");
    if (!sheet) {
      sheet = ss.insertSheet("DB_Pengawasan_K3");
      sheet.appendRow(["ID", "Tanggal_Input", "Nama_Perusahaan", "Sektor_Usaha", "Skor_Akhir", "Status"]);
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var rowsToAppend = [];
    
    for (var d = 0; d < dataArray.length; d++) {
        var payload = dataArray[d];
        var idData = "K3-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmmss") + "-" + d;
        
        var flatData = {};
        flatData["ID"] = idData;
        flatData["Tanggal_Input"] = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
        flatData["Nama_Perusahaan"] = payload.Nama_Perusahaan || "";
        flatData["Sektor_Usaha"] = payload.Sektor_Usaha || "";
        flatData["Skor_Akhir"] = payload.Skor_Akhir || "0%";
        flatData["Status"] = payload.Status || "MERAH";

        if (payload.Poin_Sesuai) {
             payload.Poin_Sesuai.forEach(function(kode) {
                 flatData["cek_" + kode + "_Status"] = "Ada / Sesuai";
             });
        } else {
             var checklist = payload.Checklist_Data || {};
             for (var key in checklist) {
                var item = checklist[key];
                if (typeof item === 'object' && item !== null) {
                    flatData[key + "_Status"] = item.Status || "";
                    flatData[key + "_Keterangan"] = item.Keterangan || "";
                } else {
                    flatData[key] = item;
                }
            }
        }

        for (var key in flatData) {
            if (headers.indexOf(key) === -1) { headers.push(key); }
        }
        rowsToAppend.push(flatData);
    }

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    var finalData2D = [];
    for (var i = 0; i < rowsToAppend.length; i++) {
         var rData = [];
         for (var j = 0; j < headers.length; j++) {
             rData.push(rowsToAppend[i][headers[j]] !== undefined ? rowsToAppend[i][headers[j]] : "");
         }
         finalData2D.push(rData);
    }

    if (finalData2D.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, finalData2D.length, headers.length).setValues(finalData2D);
    }
    return { success: true, message: finalData2D.length + " data berhasil diekstrak ke kolom terpisah." };
   } catch(e) { return { success: false, message: "Error Server: " + e.message }; }
}

function getTableDataK3() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID); 
    var sheet = ss.getSheetByName("DB_Pengawasan_K3");
    if (!sheet) return [];
    var data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return [];

    var result = [];
    for (var i = 1; i < data.length; i++) {
      result.push({ ID: data[i][0] || '-', Tanggal_Input: data[i][1] || '-', Nama_Perusahaan: data[i][2] || '-', Sektor_Usaha: data[i][3] || '-', Skor_Akhir: data[i][4] || '0%', Status: data[i][5] || 'MERAH', Data_JSON: data[i][6] || '{}' });
    }
    return result;
  } catch (e) { return { error: true, message: e.message }; }
}

function deleteDataK3(idData) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_Pengawasan_K3");
    if (!sheet) return { success: false, message: "Sheet tidak ditemukan" };
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(idData)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "Kertas Kerja K3 berhasil dihapus." };
      }
    }
    return { success: false, message: "Data tidak ditemukan." };
  } catch (e) { return { success: false, message: e.message }; }
}

function getKKPAKData() {
  try {
      var ss = SpreadsheetApp.openById(SHEET_ID);
      var sheet = ss.getSheetByName("DB_KKPAK");
      if (!sheet) return [];
      return sheet.getDataRange().getDisplayValues();
  } catch (e) { return []; }
}

function simpanDataKKPAK(payload) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_KKPAK");
    if (!sheet) return { success: false, message: "Sheet DB_KKPAK tidak ditemukan!" };
    
    var timestamp = new Date().toLocaleString('id-ID');
    sheet.appendRow([ payload.tglLapor, payload.tglKecelakaan, payload.perusahaan, payload.kbli, payload.pekerja, payload.tempat, payload.uraian, payload.tipe, payload.sumber, payload.lapor1, payload.lapor2, timestamp ]);
    return { success: true, message: "Data KK & PAK berhasil disimpan." };
  } catch (error) { return { success: false, message: error.message }; }
}

function hapusDataKKPAK(timestamp) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_KKPAK");
    if (!sheet) return { success: false, message: "Sheet DB_KKPAK tidak ditemukan!" };
    
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][11] == timestamp) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "Data berhasil dihapus." };
      }
    }
    return { success: false, message: "Data tidak ditemukan untuk dihapus." };
  } catch (error) { return { success: false, message: error.message }; }
}

function cekDanBuatHeaderSheet(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["ID", "NAMA", "NIP", "DIKLAT", "TANGGAL", "JP", "FILE"]);
    sheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#e5e7eb");
  }
}


// =====================================================================
// FUNGSI SIMPAN DATA SATUAN (MANUAL) - DILENGKAPI PELACAK ERROR
// =====================================================================
function simpanDBDataGeneric(sheetName, data) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error("Sheet '" + sheetName + "' tidak ditemukan di Spreadsheet!");
    
    // PERBAIKAN: Tambahkan "'" + data.nip agar disimpan secara presisi sebagai Teks
    if (sheetName === "DB_Kompetensi" || sheetName === "DB_AntiKorupsi") {
      sheet.appendRow([data.id, data.nama, "'" + data.nip, data.diklat, data.penyelenggara, data.tgl, data.jp, data.file]);
    } else {
      sheet.appendRow([data.id, data.nama, "'" + data.nip, data.diklat, data.tgl, data.jp, data.file]);
    }
    
    // PERBAIKAN: Paksa Google Sheets menyimpan riwayat antrean penulisan saat itu juga
    SpreadsheetApp.flush(); 
    return true;
  } catch(e) {
    throw new Error(e.message); 
  }
}

// =====================================================================
// FUNGSI SIMPAN DATA BANYAK (DARI AI EKSTRAK)
// =====================================================================
function simpanBulkDataGeneric(sheetName, dataArray) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error("Sheet '" + sheetName + "' tidak ditemukan!");
    
    var rows = [];
    for (var i = 0; i < dataArray.length; i++) {
      var data = dataArray[i];
      // PERBAIKAN: Tambahkan "'" + data.nip 
      if (sheetName === "DB_Kompetensi" || sheetName === "DB_AntiKorupsi") {
        rows.push([data.id, data.nama, "'" + data.nip, data.diklat, data.penyelenggara, data.tgl, data.jp, data.file]);
      } else {
        rows.push([data.id, data.nama, "'" + data.nip, data.diklat, data.tgl, data.jp, data.file]);
      }
    }
    
    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      // PERBAIKAN: Paksa Google Sheets menyimpan saat itu juga
      SpreadsheetApp.flush();
    }
    return true;
  } catch(e) {
    throw new Error(e.message);
  }
}

// =====================================================================
// FUNGSI TARIK DATA
// =====================================================================
function getDBDataGeneric(sheetName) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
  // PERBAIKAN: Gunakan getDisplayValues() alih-alih getValues()
  var data = sheet.getDataRange().getDisplayValues();
  
  if (data.length <= 1) return [];
  
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (sheetName === "DB_Kompetensi" || sheetName === "DB_AntiKorupsi") {
      result.push({
        id: data[i][0], nama: data[i][1], nip: data[i][2], diklat: data[i][3],
        penyelenggara: data[i][4], tgl: data[i][5], jp: data[i][6], file: data[i][7]
      });
    } else {
      result.push({
        id: data[i][0], nama: data[i][1], nip: data[i][2], diklat: data[i][3],
        tgl: data[i][4], jp: data[i][5], file: data[i][6]
      });
    }
  }
  return result;
}

function hapusDBDataGeneric(sheetName, id) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if(!sheet) return false;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1); break;
    }
  }
  return true;
}

// =====================================================================
// 2. FUNGSI UPLOAD FILE KE GOOGLE DRIVE
// =====================================================================
function uploadKeDrive(base64Data, fileName) {
  try {
    var folderId = PropertiesService.getScriptProperties().getProperty("FOLDER_SERTIFIKAT_ID");
    if (!folderId) throw new Error("ID Folder tidak ditemukan! Pastikan Anda sudah menjalankan fungsi setupSistemKompetensi().");
    
    var folder = DriveApp.getFolderById(folderId);
    var splitBase = base64Data.split(',');
    var type = splitBase[0].split(';')[0].replace('data:', '');
    var byteCharacters = Utilities.base64Decode(splitBase[1]);
    var blob = Utilities.newBlob(byteCharacters, type, fileName);
    
    var file = folder.createFile(blob);
    return file.getUrl();
  } catch (e) {
    throw new Error("Gagal mengunggah file: " + e.message);
  }
}

function getRekapHierarkiBangkom(userRole) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheetUser = ss.getSheetByName("DATA_USER");
    var sheetKomp = ss.getSheetByName("DB_Kompetensi");
    if (!sheetUser) return { success: false, message: "Sheet DATA_USER tidak ditemukan!" };
    
    var dataUser = sheetUser.getDataRange().getDisplayValues();
    var dataKomp = sheetKomp ? sheetKomp.getDataRange().getValues() : [];

    var mapJP = {};
    if (dataKomp.length > 1) {
      for (var i = 1; i < dataKomp.length; i++) {
        var rawNip = String(dataKomp[i][2]).replace(/[^0-9]/g, ''); 
        var jp = Number(dataKomp[i][5]) || 0;
        if (!mapJP[rawNip]) { mapJP[rawNip] = 0; }
        mapJP[rawNip] += jp;
      }
    }

    var result = [];
    var roleTarget = "";
    var roleAkses = String(userRole).toUpperCase();

    if (roleAkses.indexOf("PIMPINAN") !== -1) { roleTarget = "SUPERVISOR_TU"; } 
    else if (roleAkses.indexOf("SUPERVISOR") !== -1 || roleAkses.indexOf("TATA USAHA") !== -1) { roleTarget = "STAF"; } 
    else { return { success: true, data: [], targetRole: "TIDAK ADA AKSES" }; }

    for (var u = 1; u < dataUser.length; u++) {
        var usernameNIP = String(dataUser[u][1]);
        var cleanNIP = usernameNIP.replace(/[^0-9]/g, '');
        var roleDB = String(dataUser[u][4]).toUpperCase();
        var namaLengkap = String(dataUser[u][3]);
        var isMatch = false;
        
        if (roleTarget === "SUPERVISOR_TU" && (roleDB.indexOf("SUPERVISOR") !== -1 || roleDB.indexOf("TATA USAHA") !== -1)) { isMatch = true; } 
        else if (roleTarget === "STAF" && roleDB.indexOf("STAF") !== -1) { isMatch = true; }

        if (isMatch) {
            var totalJP = mapJP[cleanNIP] || 0;
            var targetJP = 40; 
            result.push({ nama: namaLengkap, nip: usernameNIP, role: roleDB, realisasiJP: totalJP, targetJP: targetJP, status: totalJP >= targetJP ? "MEMENUHI" : "BELUM MEMENUHI" });
        }
    }
    return { success: true, data: result, targetRole: roleTarget };
  } catch(e) { return { success: false, message: "Error Server: " + e.message }; }
}

// =======================================================================
// 11. MODUL AI GEMINI (REKAP, K3, RENCANA, SERTIFIKAT, SURAT)
// =======================================================================
function ekstrakRincianKehadiranDariGambar(base64String) {
  try {
    var apiKey = getGeminiAPIKey(); 
    apiKey = apiKey.trim();
    var url = "[https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=](https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=)" + apiKey;
    var base64Data = base64String.split(',')[1];
    var mimeType = base64String.substring(5, base64String.indexOf(';'));
    
    var promptTeks = `Ekstrak tabel rincian absensi dari gambar screenshot ini.
    PENTING: Coba cari Nama Pegawai dan NIP yang tertera di bagian atas/profil gambar.
    Kembalikan HANYA array objek JSON murni tanpa markdown/teks lain.
    Struktur key WAJIB persis seperti ini:
    [
      {
        "Nama_Pegawai": "Nama dari profil",
        "NIP": "NIP dari profil",
        "Tanggal": "Tgl dari baris tabel",
        "Hari": "Hari",
        "Kehadiran_Masuk": "Jam Hadir Masuk",
        "Kehadiran_Pulang": "Jam Hadir Pulang",
        "TLT": "Keterlambatan",
        "POL_TLT": "Pola TLT",
        "PSW": "Pulang Sebelum Waktu",
        "POL_PSW": "Pola PSW",
        "TLT_PSW": "Total TLT+PSW",
        "Alasan": "Keterangan/Alasan"
      }
    ]
    Jika ada kolom yang kosong di tabel, isi dengan "0" (untuk angka) atau "-" (untuk teks).`;

    var payload = {
      "contents": [{ "parts": [ { "text": promptTeks }, { "inlineData": { "mimeType": mimeType, "data": base64Data } } ] }],
      "generationConfig": { "temperature": 0.1 }
    };

    var options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());

    if (json.error) { return { success: false, message: "Error API: " + json.error.message }; }
    if (!json.candidates || !json.candidates[0].content || !json.candidates[0].content.parts[0].text) { 
        return { success: false, message: "AI gagal membaca gambar." }; 
    }

    var extractedText = json.candidates[0].content.parts[0].text;
    var match = extractedText.match(/\[[\s\S]*\]/);
    var cleanJsonStr = match ? match[0] : "[]";
    
    var finalData;
    try { finalData = JSON.parse(cleanJsonStr); } 
    catch(parseError) { return { success: false, message: "AI gagal menyusun JSON." }; }
    
    return { success: true, data: finalData };
  } catch (e) {
    return { success: false, message: "Sistem Backend Gagal: " + e.message };
  }
}

function ekstrakDataAbsensiDariGambar(base64String) {
  try {
    var apiKey = getGeminiAPIKey(); 
    apiKey = apiKey.trim();
    var url = "[https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=](https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=)" + apiKey;
    var base64Data = base64String.split(',')[1];
    var mimeType = base64String.substring(5, base64String.indexOf(';'));
    
    var promptTeks = `Tugas Anda adalah mengekstrak tabel rekapitulasi absensi pada gambar ini dengan SANGAT TELITI.
    PENTING: JANGAN PERNAH MEMOTONG ATAU MENYINGKAT BARIS!
    
    Aturan Ekstraksi:
    1. Kembalikan HANYA dalam bentuk JSON Array of Array (2D Array), TANPA markdown.
    2. Format Kolom WAJIB:
    ["No", "Nama", "Periode", "Dari", "Sampai", "Hadir", "Tidak Hadir", "TLT+PSW", "BTA", "CAP", "CB", "CH", "CTLN", "CM", "CS", "CT", "D", "GS", "H", "IDT", "IPA", "MD", "P1", "S", "TA", "TB", "TMK"]
    3. Baris pertama wajib header. Jika ada cell kosong, isi dengan "0" (angka) atau "-" (teks).
    `;

    var payload = { "contents": [{ "parts": [ { "text": promptTeks }, { "inlineData": { "mimeType": mimeType, "data": base64Data } } ] }], "generationConfig": { "temperature": 0.1 } };
    var options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());

    if (json.error) { return { success: false, message: "Error API: " + json.error.message }; }
    if (!json.candidates || !json.candidates[0].content || !json.candidates[0].content.parts[0].text) { return { success: false, message: "AI gagal membaca gambar." }; }

    var extractedText = json.candidates[0].content.parts[0].text;
    var match = extractedText.match(/\[[\s\S]*\]/);
    var cleanJsonStr = match ? match[0] : "[]";

    var finalData;
    try { finalData = JSON.parse(cleanJsonStr); } catch(parseError) { return { success: false, message: "AI gagal menyusun format JSON." }; }
    return { success: true, data: finalData };
  } catch (e) { return { success: false, message: "Backend Error: " + e.message }; }
}

function ekstrakDataRencanaDariGambar(base64String) {
  try {
    var apiKey = getGeminiAPIKey(); apiKey = apiKey.trim();
    var url = "[https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=](https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=)" + apiKey;
    var base64Data = base64String.split(',')[1]; var mimeType = base64String.substring(5, base64String.indexOf(';'));

    var prompt = `Anda adalah Asisten Administrasi. Analisis gambar jadwal pengawasan yang dilampirkan. Ekstrak data dari gambar tersebut dan kembalikan HANYA dalam format array of JSON yang valid, tanpa teks markdown lainnya.
      Struktur JSON yang WAJIB digunakan:
      [ { "Tanggal_Rencana": "Isi dengan tanggal (DD-MM-YYYY)", "Nama_Perusahaan": "Isi dengan nama perusahaan/target", "Alamat_Perusahaan": "Isi dengan alamat perusahaan", "Petugas": "Isi dengan nama petugas yang ditugaskan", "Keterangan": "Isi dengan tujuan/keterangan lain jika ada" } ]`;

    var payload = { "contents": [{ "parts": [ {"text": prompt}, {"inlineData": { "mimeType": mimeType, "data": base64Data }} ] }] };
    var options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
    var response = UrlFetchApp.fetch(url, options); var json = JSON.parse(response.getContentText());
    if (json.error) return { success: false, message: "Error dari Google: " + json.error.message };
    if (!json.candidates || json.candidates.length === 0) return { success: false, message: "AI gagal membaca gambar." };

    var aiText = json.candidates[0].content.parts[0].text; 
    var match = aiText.match(/\[[\s\S]*\]/); var cleanJsonStr = match ? match[0] : "[]";
    return { success: true, data: JSON.parse(cleanJsonStr) };
  } catch (e) { return { success: false, message: "Terjadi kesalahan di sistem: " + e.message }; }
}

function ekstrakDataSertifikatDariGambar(base64String) {
  try {
    var apiKey = getGeminiAPIKey(); apiKey = apiKey.trim();
    var url = "[https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=](https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=)" + apiKey;
    var base64Data = base64String.split(',')[1]; var mimeType = base64String.substring(5, base64String.indexOf(';'));
    var prompt = `Anda adalah Asisten Administrasi Kepegawaian yang sangat teliti. Analisis gambar sertifikat / piagam yang dilampirkan. Ekstrak data dari gambar tersebut dan kembalikan HANYA dalam format array of JSON tunggal yang valid.
      Struktur JSON yang WAJIB digunakan:
      [ { "Nama_Diklat": "Isi dengan nama kegiatan/pelatihan/diklat", "Tanggal_Pelaksanaan": "Isi dengan tanggal pelaksanaan (wajib format DD-MM-YYYY)", "Total_JP": "Isi HANYA dengan angka Jumlah Jam Pelajaran (JP)." } ]`;

    var payload = { "contents": [{ "parts": [ {"text": prompt}, {"inlineData": { "mimeType": mimeType, "data": base64Data }} ] }] };
    var options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
    var response = UrlFetchApp.fetch(url, options); var json = JSON.parse(response.getContentText());
    if (json.error) return { success: false, message: "Error dari Google: " + json.error.message };
    if (!json.candidates || json.candidates.length === 0) return { success: false, message: "AI gagal membaca sertifikat." };

    var aiText = json.candidates[0].content.parts[0].text; 
    var match = aiText.match(/\[[\s\S]*\]/); var cleanJsonStr = match ? match[0] : "[]";
    return { success: true, data: JSON.parse(cleanJsonStr) };
  } catch (e) { return { success: false, message: "Terjadi kesalahan di sistem AI: " + e.message }; }
}

function ekstrakDataSuratDariGambar(base64String) {
  try {
    var apiKey = getGeminiAPIKey(); apiKey = apiKey.trim(); 
    var url = "[https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=](https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=)" + apiKey;
    var base64Data = base64String.split(',')[1]; var mimeType = base64String.substring(5, base64String.indexOf(';'));
    var prompt = `Anda adalah Asisten Administrasi Arsip Surat. Analisis dokumen / surat yang dilampirkan. Ekstrak data penting dari surat tersebut dan kembalikan HANYA dalam format array of JSON tunggal yang valid.
      Struktur JSON yang WAJIB digunakan:
      [ { "No_Surat": "Isi dengan Nomor Surat", "Tgl_Surat": "Isi dengan Tanggal Surat (format YYYY-MM-DD)", "Pihak_Terkait": "Isi dengan Pengirim atau Tujuan", "Perihal": "Isi dengan Hal/Perihal surat" } ]`;

    var payload = { "contents": [{ "parts": [ {"text": prompt}, {"inlineData": { "mimeType": mimeType, "data": base64Data }} ] }] };
    var options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
    var response = UrlFetchApp.fetch(url, options); var json = JSON.parse(response.getContentText());
    if (json.error) return { success: false, message: "Error dari Google API: " + json.error.message };
    if (!json.candidates || json.candidates.length === 0) return { success: false, message: "AI gagal membaca dokumen surat." };
    
    var aiText = json.candidates[0].content.parts[0].text; 
    var match = aiText.match(/\[[\s\S]*\]/); var cleanJsonStr = match ? match[0] : "[]";
    return { success: true, data: JSON.parse(cleanJsonStr) };
  } catch (e) { return { success: false, message: "Terjadi kesalahan di sistem AI: " + e.message }; }
}

function generateAnalisisNormaK3AI(judulNorma, normaId, dataChecklist) {
  try {
    var apiKey = getGeminiAPIKey(); apiKey = apiKey.trim();
    if (!apiKey || apiKey === "") return { success: false, message: "API Key Gemini belum dimasukkan." };
    var prompt = "Anda adalah Inspektur Pengawas Ketenagakerjaan dan K3 Senior di Indonesia yang tegas, profesional, dan sangat hafal hukum ketenagakerjaan Republik Indonesia.\n\n Tugas Anda: Buatkan SATU PARAGRAF analisis temuan untuk bagian: [" + judulNorma + "]. PASTIKAN KALIMAT UTUH DAN SELESAI SAMPAI TITIK TERAKHIR, JANGAN TERPOTONG.\n Berikut adalah hasil pemeriksaan di lapangan (Checklist) yang harus Anda jadikan dasar:\n" + JSON.stringify(dataChecklist, null, 2) + "\n\n Instruksi Wajib:\n 1. Jika semua status indikator adalah 'Ada / Sesuai', nyatakan bahwa perusahaan telah patuh terhadap regulasi, dan sebutkan secara singkat rujukan UU/Permenaker yang relevan.\n 2. Jika terdapat indikator bersatus 'Tidak Ada / Belum Sesuai', tegaskan letak pelanggarannya secara lugas. Sebutkan peraturan spesifik yang dilanggar, dan berikan saran perbaikan wajib.\n 3. Format output: MURNI 1 PARAGRAF TEKS SAJA YANG SELESAI SEMPURNA.\n 4. Jangan gunakan kata pengantar. Langsung masuk ke kalimat utama.\n 5. Apabila ada Keterangan khusus dari pengawas, masukkan sebagai bahan pertimbangan.";
    var payload = { "contents": [{"parts": [{"text": prompt}]}], "generationConfig": { "temperature": 0.2, "maxOutputTokens": 2048 } };
    var options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
    var url = "[https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=](https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=)" + apiKey; 
    
    var maxPercobaan = 3; var attempt = 0; var berhasil = false; var json = null;
    while (attempt < maxPercobaan && !berhasil) {
        var response = UrlFetchApp.fetch(url, options); var responseText = response.getContentText(); json = JSON.parse(responseText);
        if (json.error) {
            var pesanError = json.error.message.toLowerCase();
            if (pesanError.includes("high demand") || pesanError.includes("quota") || pesanError.includes("429") || pesanError.includes("overloaded")) { attempt++; if (attempt < maxPercobaan) { Utilities.sleep(3000 * attempt); continue; } else return { success: false, message: "Server AI sedang penuh (High Demand). Silakan coba lagi." }; } else return { success: false, message: "Ditolak Google AI: " + json.error.message };
        } else if (!json.candidates || json.candidates.length === 0) return { success: false, message: "AI gagal menyusun paragraf analisis." };
        else berhasil = true;
    }
    var resultText = json.candidates[0].content.parts[0].text.trim(); return { success: true, data: resultText };
  } catch (e) { return { success: false, message: "Error Eksekusi Script: " + e.toString() }; }
}

function ekstrakDataK3DariGambar(base64String) {
  try { 
    var apiKey = getGeminiAPIKey(); apiKey = apiKey.trim();
    if (!apiKey || apiKey === "") return { success: false, message: "API Key Gemini belum dimasukkan!" };
    var splitData = base64String.split(','); if (splitData.length < 2) return { success: false, message: "Format gambar rusak." };
    var base64Data = splitData[1]; var mimeType = base64String.substring(5, base64String.indexOf(';'));
    if(mimeType !== 'application/pdf' && mimeType !== 'image/jpeg' && mimeType !== 'image/png') return { success: false, message: "Format ditolak! AI hanya membaca JPG, PNG, atau PDF." };
    var url = "[https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-8b:generateContent?key=](https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-8b:generateContent?key=)" + apiKey;
    var promptText = `Anda adalah Asisten Pengawas Ketenagakerjaan K3 yang teliti. Tugas Anda: Baca dokumen terlampir (PDF/Gambar) dan ekstrak informasi ke dalam format JSON.\n\n ATURAN PENTING:\n - Berikan respons HANYA dalam format JSON MURNI (Array of Objects).\n\n STRUKTUR JSON YANG WAJIB DIHASILKAN:\n [\n {\n "Nama_Perusahaan": "[Ekstrak Nama Perusahaan]",\n "Alamat_Perusahaan": "[Ekstrak Alamat Perusahaan]",\n "Jenis_Usaha": "[Ekstrak Jenis/Sektor Usaha]",\n "Poin_Sesuai": [\n // Masukkan kode indikator HANYA JIKA terlihat dicentang "Ada/Sesuai":\n "NORMA_A1", "NORMA_B1", "NORMA_C1", "NORMA_D1", "NORMA_E1", "NORMA_F1", "NORMA_G1", "NORMA_H1", "NORMA_I1", "NORMA_J1", "NORMA_K1", "NORMA_L1", "NORMA_M1", "NORMA_N1", "NORMA_O1", "NORMA_P1", "NORMA_Q1", "NORMA_R1", "NORMA_S1", "NORMA_T1"\n ]\n }\n ]`;
    var payload = { "contents": [{ "parts": [ {"text": promptText}, { "inlineData": { "mimeType": mimeType, "data": base64Data } } ] }], "generationConfig": { "responseMimeType": "application/json" } };
    var options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
    var maxPercobaan = 3; var attempt = 0; var berhasil = false; var json = null;
    while (attempt < maxPercobaan && !berhasil) {
        var response = UrlFetchApp.fetch(url, options); json = JSON.parse(response.getContentText());
        if (json.error) {
            var pesanError = json.error.message.toLowerCase();
            if (pesanError.includes("high demand") || pesanError.includes("quota") || pesanError.includes("429") || pesanError.includes("overloaded")) { attempt++; if (attempt < maxPercobaan) { Utilities.sleep(3000 * attempt); continue; } else return { success: false, message: "Server AI penuh. Silakan coba lagi." }; } else return { success: false, message: "Ditolak Google AI: " + json.error.message };
        } else if (!json.candidates || json.candidates.length === 0) return { success: false, message: "AI gagal membaca gambar." };
        else berhasil = true;
    }
    var aiResultText = json.candidates[0].content.parts[0].text;
    var match = aiResultText.match(/\[[\s\S]*\]/); var cleanJsonStr = match ? match[0] : "[]";
    var parsedData = JSON.parse(cleanJsonStr); var dataArray = Array.isArray(parsedData) ? parsedData : [parsedData];
    return { success: true, data: dataArray };
  } catch (e) { return { success: false, message: "Sistem Error: " + e.message }; }
}
// =======================================================================
// 13. FUNGSI DATABASE ABSENSI HARIAN YANG TERHAPUS
// =======================================================================
function getTableDataKinerjaHarian() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("ABSENSI_HARIAN") || ss.getSheetByName("DB_Kinerja_Harian");
    if (!sheet) return [];
    return sheet.getDataRange().getDisplayValues();
  } catch(e) { return []; }
}

function simpanBulkKinerjaHarian(dataArray){
  try{
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("ABSENSI_HARIAN") || ss.getSheetByName("DB_Kinerja_Harian");
    if(!sheet) throw new Error("Sheet ABSENSI_HARIAN tidak ditemukan");

    const lastRow = sheet.getLastRow();
    const oldData = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues() : [];

    let update = 0, tambah = 0;

    dataArray.forEach(item => {
      let nama = String(item.Nama_Pegawai || item["Nama Pegawai"] || "").trim().toUpperCase();
      let tanggal = String(item.Tanggal || "").trim(); 
      let posisi = -1;

      for(let i = 0; i < oldData.length; i++){
        let namaLama = String(oldData[i][1] || "").trim().toUpperCase();
        
        let rawTgl = oldData[i][3];
        let tanggalLama = "";
        if (rawTgl instanceof Date) {
          tanggalLama = rawTgl.getFullYear() + "-" + String(rawTgl.getMonth() + 1).padStart(2, '0') + "-" + String(rawTgl.getDate()).padStart(2, '0');
        } else {
          tanggalLama = String(rawTgl || "").trim();
        }

        if(nama === namaLama && tanggal === tanggalLama){ posisi = i + 2; break; }
      }

      let baris = [
        posisi > 0 ? oldData[posisi - 2][0] : "ABS-" + Date.now() + "-" + Math.floor(Math.random() * 1000), 
        item.Nama_Pegawai || item["Nama Pegawai"],
        item.NIP || "-",
        item.Tanggal,
        // KOREKSI: Terima format 'Hari' maupun 'Hari' dari Excel
        item.Hari || item["Hari"] || "",
        // KOREKSI: Terima format 'Tipe_Hari' manual maupun 'Tipe Hari' Excel
        item.Tipe_Hari || item["Tipe Hari"] || "Kerja", 
        item.Kehadiran_Masuk || item["Kehadiran Masuk"] || "-",
        item.Kehadiran_Pulang || item["Kehadiran Pulang"] || "-",
        item.TLT || "0",
        item["POL.TLT"] || "0",
        item.PSW || "0",
        item["POL.PSW"] || "0",
        item["TLT+PSW"] || "0",
        item.Alasan || ""
      ];

      if(posisi > 0){
        sheet.getRange(posisi, 1, 1, baris.length).setValues([baris]);
        update++;
      }else{
        sheet.appendRow(baris);
        tambah++;
      }
    });

    // Memicu rekap bulanan otomatis
    if (typeof rekapAbsensiBulananOtomatis === "function") {
        rekapAbsensiBulananOtomatis();
    }

    return {
      success: true,
      message: "Berhasil. Update: " + update + " | Tambah baru: " + tambah
    };

  } catch(e) {
    return { success: false, message: e.message };
  }
}

function updateDataKinerjaHarian(rowIndex, payload) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("ABSENSI_HARIAN") || ss.getSheetByName("DB_Kinerja_Harian");
    sheet.getRange(rowIndex, 2).setValue(payload.Nama_Pegawai); sheet.getRange(rowIndex, 4).setValue(payload.Tanggal);
    sheet.getRange(rowIndex, 5).setValue(payload.Hari); sheet.getRange(rowIndex, 6).setValue(payload.Tipe_Hari);
    sheet.getRange(rowIndex, 7).setValue(payload.Kehadiran_Masuk); sheet.getRange(rowIndex, 8).setValue(payload.Kehadiran_Pulang);
    sheet.getRange(rowIndex, 14).setValue(payload.Alasan); 
    SpreadsheetApp.flush();
    
    // --- PEMICU REKAP OTOMATIS ---
    rekapAbsensiBulananOtomatis();
    // -----------------------------

    return { success: true, message: "Data berhasil diperbarui." };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteDataKinerjaHarian(idData) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("ABSENSI_HARIAN") || ss.getSheetByName("DB_Kinerja_Harian");
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(idData)) { 
        sheet.deleteRow(i + 1); 
        SpreadsheetApp.flush(); 
        
        // --- PEMICU REKAP OTOMATIS ---
        rekapAbsensiBulananOtomatis();
        // -----------------------------
        
        return { success: true, message: "Data absensi dihapus." }; 
      }
    }
    return { success: false, message: "Data tidak ditemukan." };
  } catch(e) { return { success: false, message: e.message }; }
}

function rekapAbsensiBulananOtomatis() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheetHarian = ss.getSheetByName("ABSENSI_HARIAN") || ss.getSheetByName("DB_Kinerja_Harian");
    if (!sheetHarian) return { success: false, message: "Sheet Harian tidak ditemukan." };
    
    var sheetBulanan = ss.getSheetByName("DB_Absensi");
    if (!sheetBulanan) sheetBulanan = ss.insertSheet("DB_Absensi");

    var dataHarian = sheetHarian.getDataRange().getValues();
    if (dataHarian.length <= 1) return { success: true, message: "Data harian kosong." };

    // Header sheet Rekap Bulanan
    var headersBulanan = ["No", "Nama", "Periode", "Dari", "Sampai", "Hadir", "Tidak Hadir", "TLT+PSW", "BTA", "CAP", "CB", "CH", "CTLN", "CM", "CS", "CT", "D", "GS", "H", "IDT", "IPA", "MD", "P1", "S", "TA", "TB", "TMK"];
    
    // ===============================================================
    // KAMUS KODE ABSEN (Definitif sesuai standar)
    // ===============================================================
    var KODE_ABSEN = ["BTA", "CAP", "CB", "CH", "CTLN", "CM", "CS", "CT", "D", "GS", "H", "IDT", "IPA", "MD", "P1", "S", "TA", "TB", "TMK"];
    
    var mapRekap = {};
    var bulanNames = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];

    for (var i = 1; i < dataHarian.length; i++) {
      var row = dataHarian[i];
      var nama = String(row[1] || "").trim();
      var rawTgl = row[3];
      if (!nama || nama === "-" || !rawTgl) continue;

      var d = new Date(rawTgl);
      if (isNaN(d.getTime())) continue;

      var date = d.getDate();
      var month = d.getMonth();
      var year = d.getFullYear();
      
      var pMonth = month;
      var pYear = year;
      var startDate, endDate;
      
      // LOGIKA CUT-OFF BULANAN (Tgl 21 - 20)
      if (date >= 21) {
        pMonth = month + 1;
        if (pMonth > 11) { pMonth = 0; pYear++; }
        startDate = new Date(year, month, 21);
        endDate = new Date(pYear, pMonth, 20);
      } else {
        var prevMonth = month - 1;
        var prevYear = year;
        if (prevMonth < 0) { prevMonth = 11; prevYear--; }
        startDate = new Date(prevYear, prevMonth, 21);
        endDate = new Date(year, month, 20);
      }

      var namaPeriode = bulanNames[pMonth];
      var key = nama + "|" + namaPeriode + "|" + pYear;

      if (!mapRekap[key]) {
        var iso = function(dt) { return dt.getFullYear() + '-' + String(dt.getMonth() + 1).padStart(2, '0') + '-' + String(dt.getDate()).padStart(2, '0'); };
        mapRekap[key] = {
          nama: nama, periode: namaPeriode, dari: iso(startDate), sampai: iso(endDate),
          hadir: 0, absen: 0, tltpsw: 0, details: {}
        };
        for(var c = 8; c < headersBulanan.length; c++) mapRekap[key].details[headersBulanan[c]] = 0;
      }

      var tipe = String(row[5] || "").trim().toUpperCase();
      var hari = String(row[4] || "").trim().toUpperCase();
      // Abaikan hari libur dari perhitungan
      if (tipe.includes('LIBUR') || hari === 'MINGGU' || hari === 'SABTU') continue;

      var masuk = String(row[6] || "").trim();
      var pulang = String(row[7] || "").trim();
      var tpsw = parseInt(String(row[12] || "0").replace(/[^0-9]/g, '')) || 0;
      
      var alasan = String(row[13] || "").trim().toUpperCase();
      if (alasan === '-') alasan = ""; 

      var st = 'Hadir';
      
      // ===============================================================
      // PENENTUAN STATUS
      // ===============================================================
      // 1. Cek langsung ke Kamus Kode
      if (KODE_ABSEN.includes(alasan)) {
        st = alasan;
      } 
      // 2. Fallback: jika operator mengetik kata panjang alih-alih kode
      else if (alasan.includes('SAKIT')) st = 'S';
      else if (alasan.includes('IZIN')) st = 'IPA';
      else if (alasan.includes('CUTI')) st = 'CT';
      else if (alasan.includes('DINAS')) st = 'D';
      else if (alasan.includes('TANPA')) st = 'TA';
      else if (alasan.includes('TUGAS BELAJAR')) st = 'TB';
      // 3. Fallback: Jam masuk/pulang kosong dan tidak ada alasan sah = TMK
      else if (masuk === '-' || pulang === '-' || !masuk || !pulang) st = 'TMK';
      else st = 'Hadir';

      // ===============================================================
      // DISTRIBUSI KE KOLOM DB_ABSENSI
      // ===============================================================
      if (st === 'Hadir') {
        mapRekap[key].hadir++;
        mapRekap[key].tltpsw += tpsw;
      } else {
        // HANYA TMK yang dihitung masuk ke angka "Tidak Hadir" utama
        if (st === 'TMK') {
          mapRekap[key].absen++;
          mapRekap[key].details["TMK"]++;
        } 
        // Kode lainnya (GS, H, S, D, dll) masuk ke rincian spesifik tanpa menambah "Tidak Hadir"
        else if (mapRekap[key].details[st] !== undefined) {
          mapRekap[key].details[st]++;
        }
      }
    }

    // Susun data array untuk dicetak ke DB_Absensi
    var finalData = [headersBulanan];
    var no = 1;
    for (var k in mapRekap) {
      var r = mapRekap[k];
      var rowArr = [no++, r.nama, r.periode, r.dari, r.sampai, r.hadir, r.absen, r.tltpsw];
      for(var c = 8; c < headersBulanan.length; c++) {
        rowArr.push(r.details[headersBulanan[c]] || 0);
      }
      finalData.push(rowArr);
    }

    sheetBulanan.clear();
    sheetBulanan.getRange(1, 1, finalData.length, headersBulanan.length).setValues(finalData);
    sheetBulanan.getRange(1, 1, 1, headersBulanan.length).setFontWeight("bold").setBackground("#0f172a").setFontColor("#00f3ff").setHorizontalAlignment("center");
    SpreadsheetApp.flush();

    return { success: true, message: "Sinkronisasi selesai! (" + (finalData.length - 1) + " pegawai tersimpan)." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
// =======================================================================
// FUNGSI AI GENERATOR ANALISIS ABSENSI (MENGGUNAKAN GROQ - LLAMA 3.1)
// =======================================================================
// =======================================================================
// FUNGSI AI GENERATOR ANALISIS ABSENSI (GROQ - QWEN 27B)
// =======================================================================
function generateAnalisisAbsensiAI(payloadStr) {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY');
    if (!apiKey) return "Catatan: GROQ_API_KEY belum dipasang di Script Properties.";

    var data = JSON.parse(payloadStr);
    var prompt = "Anda adalah Analis Kepegawaian Eksekutif ahli PP 94 Tahun 2021 tentang Disiplin PNS. Buat SATU paragraf padat (maksimal 4 kalimat) analisis evaluasi kedisiplinan dan rekomendasi penegakan aturan berdasarkan ringkasan data absensi ini:\n\n" +
                 "- Periode: " + data.periode + "\n" +
                 "- Total Data: " + data.totalPegawai + "\n" +
                 "- Hadir Tepat Waktu: " + data.tepatWaktu + "\n" +
                 "- Terlambat / Pulang Cepat (PSW): " + data.telat + "\n" +
                 "- Mangkir / Tanpa Keterangan (TMK/TK): " + data.mangkir + "\n" +
                 "- Alasan Sah (Izin/Sakit/Cuti/Dinas Luar): " + data.izin + "\n\n" +
                 "Fokuskan paragraf pada kesimpulan tingkat kepatuhan dan saran perbaikan kedisiplinan yang berwibawa. Jangan gunakan list/poin-poin, murni 1 paragraf teks saja.";

    var url = "https://api.groq.com/openai/v1/chat/completions";
    
    var payload = {
      "model": "qwen/qwen3.8-27b", // Model chat/teks yang tersedia dan aktif di akun Anda
      "messages": [
        { "role": "user", "content": prompt }
      ],
      "temperature": 0.2,
      "max_tokens": 300
    };
    
    var options = {
      "method": "post",
      "headers": {
        "Authorization": "Bearer " + apiKey.trim(),
        "Content-Type": "application/json"
      },
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());

    if (json.error) return "Gagal menganalisis: " + json.error.message;
    if (!json.choices || json.choices.length === 0) return "AI gagal memberikan respon.";

    return json.choices[0].message.content.trim().replace(/\n/g, " ");
  } catch (e) {
    return "Error Mesin AI: " + e.message;
  }
}
// =======================================================================
// FUNGSI SETUP DATABASE KEPEGAWAIAN
// =======================================================================
function setupKepegawaian() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var createdSheets = [];

    // 1. Setup Sheet ABSENSI_HARIAN
    var shAbsenHarian = ss.getSheetByName("ABSENSI_HARIAN");
    if (!shAbsenHarian) {
      shAbsenHarian = ss.insertSheet("ABSENSI_HARIAN");
      shAbsenHarian.appendRow(["ID", "Nama Pegawai", "NIP", "Tanggal", "Hari", "Tipe Hari", "Kehadiran Masuk", "Kehadiran Pulang", "TLT", "POL.TLT", "PSW", "POL.PSW", "TLT+PSW", "Alasan"]);
      shAbsenHarian.getRange(1, 1, 1, 14).setFontWeight("bold").setBackground("#1e293b").setFontColor("white");
      createdSheets.push("ABSENSI_HARIAN");
    }

    // 2. Setup Sheet DB_Absensi (Untuk Rekap Bulanan)
    var shAbsenBulanan = ss.getSheetByName("DB_Absensi");
    if (!shAbsenBulanan) {
      shAbsenBulanan = ss.insertSheet("DB_Absensi");
      shAbsenBulanan.appendRow(["No", "Nama", "Periode", "Dari", "Sampai", "Hadir", "Tidak Hadir", "TLT+PSW", "BTA", "CAP", "CB", "CH", "CTLN", "CM", "CS", "CT", "D", "GS", "H", "IDT", "IPA", "MD", "P1", "S", "TA", "TB", "TMK"]);
      shAbsenBulanan.getRange(1, 1, 1, 27).setFontWeight("bold").setBackground("#1e293b").setFontColor("white");
      createdSheets.push("DB_Absensi");
    }

    // 3. Setup Sheet DB_Apel
    var shApel = ss.getSheetByName("DB_Apel");
    if (!shApel) {
      shApel = ss.insertSheet("DB_Apel");
      shApel.appendRow(["Tanggal_Input", "Tanggal_Apel", "Foto_URL", "Daftar_Hadir_Pegawai", "Keterangan_Arahan"]);
      shApel.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#1e293b").setFontColor("white");
      createdSheets.push("DB_Apel");
    }

    SpreadsheetApp.flush();

    if (createdSheets.length === 0) {
      return { success: true, sheets: ["Semua tabel sudah tersedia (Tidak ada yang baru dibuat)"] };
    }

    return { success: true, sheets: createdSheets };
  } catch (e) {
    throw new Error("Gagal membuat database: " + e.message);
  }
}
function setupSistemKompetensi() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // A. Buat Sheet Pengembangan Diri (DB_Kompetensi)
  var sheetKompetensi = ss.getSheetByName("DB_Kompetensi");
  if (!sheetKompetensi) {
    sheetKompetensi = ss.insertSheet("DB_Kompetensi");
    // HEADER BARU DENGAN KOLOM PENYELENGGARA
    sheetKompetensi.appendRow(["id", "nama", "nip", "diklat", "penyelenggara", "tgl", "jp", "file"]);
    sheetKompetensi.getRange("A1:H1").setFontWeight("bold").setBackground("#d9ead3").setHorizontalAlignment("center");
    sheetKompetensi.setFrozenRows(1);
    sheetKompetensi.setColumnWidths(1, 8, 150);
  }
  
  // B. Buat Sheet Anti Korupsi (DB_AntiKorupsi)
  var sheetAntiKorupsi = ss.getSheetByName("DB_AntiKorupsi");
  if (!sheetAntiKorupsi) {
    sheetAntiKorupsi = ss.insertSheet("DB_AntiKorupsi");
    // HEADER BARU DENGAN KOLOM PENYELENGGARA
    sheetAntiKorupsi.appendRow(["id", "nama", "nip", "diklat", "penyelenggara", "tgl", "jp", "file"]);
    sheetAntiKorupsi.getRange("A1:H1").setFontWeight("bold").setBackground("#fce5cd").setHorizontalAlignment("center");
    sheetAntiKorupsi.setFrozenRows(1);
    sheetAntiKorupsi.setColumnWidths(1, 8, 150);
  }
  
  // C. Buat Folder Arsip Sertifikat di Google Drive
  var folderName = "Arsip_Sertifikat_Kompetensi";
  var folders = DriveApp.getFoldersByName(folderName);
  var folderId = "";
  
  if (folders.hasNext()) {
    folderId = folders.next().getId();
  } else {
    var newFolder = DriveApp.createFolder(folderName);
    newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    folderId = newFolder.getId();
  }
  
  PropertiesService.getScriptProperties().setProperty("FOLDER_SERTIFIKAT_ID", folderId);
  return "Setup Berhasil! Folder dan Sheet sudah siap digunakan.";
}
// =======================================================================
// 14. MODUL STRUKTUR ATASAN & BAWAHAN
// =======================================================================
function getHierarkiPegawai() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_Hierarki");
    if (!sheet) {
      sheet = ss.insertSheet("DB_Hierarki");
      sheet.appendRow(["NIP_Bawahan", "Nama_Bawahan", "NIP_Atasan", "Nama_Atasan"]);
      sheet.getRange("A1:D1").setFontWeight("bold").setBackground("#1e293b").setFontColor("white");
      return [];
    }
    var data = sheet.getDataRange().getDisplayValues();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      result.push({
        nip_bawahan: String(data[i][0]).replace(/[^0-9]/g, ''),
        nama_bawahan: data[i][1],
        nip_atasan: String(data[i][2]).replace(/[^0-9]/g, ''),
        nama_atasan: data[i][3]
      });
    }
    return result;
  } catch(e) { return []; }
}

function simpanHierarkiBawahan(nipAtasan, namaAtasan, listBawahan) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName("DB_Hierarki");
    if (!sheet) {
      sheet = ss.insertSheet("DB_Hierarki");
      sheet.appendRow(["NIP_Bawahan", "Nama_Bawahan", "NIP_Atasan", "Nama_Atasan"]);
    }
    
    var existingData = sheet.getDataRange().getValues();
    
    // Ekstrak NIP bawahan yang sedang diproses
    var nipsBawahanTerpilih = listBawahan.map(function(b) { return String(b.nip).replace(/[^0-9]/g, ''); });
    
    // Hapus relasi lama dari bawahan yang diubah agar tidak ada duplikasi atasan (1 bawahan = 1 atasan langsung)
    for (var i = existingData.length - 1; i >= 1; i--) {
      var nipB = String(existingData[i][0]).replace(/[^0-9]/g, '');
      if (nipsBawahanTerpilih.indexOf(nipB) > -1) {
        sheet.deleteRow(i + 1);
      }
    }
    
    // Siapkan baris baru untuk ditambahkan
    var rowsToAppend = [];
    for(var j = 0; j < listBawahan.length; j++) {
        rowsToAppend.push([ "'" + listBawahan[j].nip, listBawahan[j].nama, "'" + nipAtasan, namaAtasan ]);
    }
    
    if (rowsToAppend.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 4).setValues(rowsToAppend);
    }
    
    SpreadsheetApp.flush();
    return { success: true, message: "Struktur hierarki bawahan berhasil diperbarui!" };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// =======================================================================
// 15. MODUL EVALUASI BANGKOM (ATASAN -> BAWAHAN)
// =======================================================================
function getDataEvaluasiBawahan(nipAtasan) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheetHierarki = ss.getSheetByName("DB_Hierarki");
    var sheetKomp = ss.getSheetByName("DB_Kompetensi");
    
    if (!sheetHierarki) return { success: false, message: "Database Struktur Hierarki belum dikonfigurasi." };
    
    // PERBAIKAN: getDisplayValues agar NIP 18 digit terbaca utuh
    var dataHierarki = sheetHierarki.getDataRange().getDisplayValues();
    var targetNipAtasan = String(nipAtasan).replace(/[^0-9]/g, '');
    var listBawahan = {}; 
    
    for (var i = 1; i < dataHierarki.length; i++) {
      var nipA = String(dataHierarki[i][2]).replace(/[^0-9]/g, '');
      if (nipA === targetNipAtasan) {
        var nipB = String(dataHierarki[i][0]).replace(/[^0-9]/g, '');
        var namaB = dataHierarki[i][1];
        listBawahan[nipB] = { nip: nipB, nama: namaB, jp_kompetensi: 0, rincian: [] };
      }
    }
    
    if (Object.keys(listBawahan).length === 0) {
      return { success: true, is_atasan: false, data: [] };
    }
    
    if (sheetKomp) {
      // PERBAIKAN: getDisplayValues untuk DB Kompetensi
      var dataKomp = sheetKomp.getDataRange().getDisplayValues();
      for (var j = 1; j < dataKomp.length; j++) {
        var nipPegawai = String(dataKomp[j][2]).replace(/[^0-9]/g, '');
        if (listBawahan[nipPegawai]) {
          var jp = Number(dataKomp[j][6]) || 0; 
          listBawahan[nipPegawai].jp_kompetensi += jp;
          listBawahan[nipPegawai].rincian.push({ diklat: dataKomp[j][3], tgl: dataKomp[j][5], jp: jp });
        }
      }
    }
    
    var resultArr = [];
    for (var key in listBawahan) {
      var b = listBawahan[key];
      resultArr.push({
        nip: b.nip, nama: b.nama, realisasi_jp: b.jp_kompetensi,
        target_jp: 40, status: b.jp_kompetensi >= 40 ? "MEMENUHI" : "BELUM MEMENUHI", rincian: b.rincian
      });
    }
    resultArr.sort(function(a, b) { return a.nama.localeCompare(b.nama); });
    
    return { success: true, is_atasan: true, data: resultArr };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
// =======================================================================
// FUNGSI AI GENERATOR ANALISIS EVALUASI BANGKOM BULANAN (VIA GROQ)
// =======================================================================
function generateAnalisisBangkomAIBulanan(payloadStr) {
  try {
    // Memanggil API Key Groq sesuai dengan standar laporan kepegawaian Anda
    var apiKey = PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY');
    if (!apiKey) return { success: false, message: "Catatan: GROQ_API_KEY belum dipasang di Script Properties." };

    var data = JSON.parse(payloadStr);
    
    // Prompt yang diarahkan untuk Groq
    var prompt = "Anda adalah Analis Kepegawaian Eksekutif ahli manajemen SDM pemerintahan. Buat SATU paragraf padat (maksimal 4 kalimat) berisi analisis evaluasi dan kesimpulan tingkat partisipasi Pengembangan Kompetensi (Bangkom) pegawai untuk periode " + data.periode + ".\n\n" +
                 "Data Laporan:\n" +
                 "- Total Pegawai Dievaluasi: " + data.totalPegawai + " orang\n" +
                 "- Memenuhi Standar Minimal (>= 40 JP): " + data.memenuhi + " orang\n" +
                 "- Belum Memenuhi Standar: " + data.belumMemenuhi + " orang\n" +
                 "- Rata-rata Capaian: " + data.rataJP + " JP\n\n" +
                 "Fokuskan paragraf pada kesimpulan kepatuhan dan rekomendasi pembinaan yang tegas dan berwibawa. Jangan gunakan list atau poin-poin, murni 1 paragraf teks narasi saja.";

    var url = "https://api.groq.com/openai/v1/chat/completions";
    
    var payload = {
      "model": "llama-3.1-8b-instant", // Anda bisa mengubahnya ke "qwen-2.5-32b" atau model Groq lain yang biasa Anda pakai
      "messages": [
        { "role": "user", "content": prompt }
      ],
      "temperature": 0.2,
      "max_tokens": 300
    };
    
    var options = {
      "method": "post",
      "headers": {
        "Authorization": "Bearer " + apiKey.trim(),
        "Content-Type": "application/json"
      },
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());

    if (json.error) return { success: false, message: "Gagal menganalisis: " + json.error.message };
    if (!json.choices || json.choices.length === 0) return { success: false, message: "AI gagal memberikan respon." };

    // Mengembalikan teks hasil analisis
    return { success: true, data: json.choices[0].message.content.trim().replace(/\n/g, " ") };
  } catch (e) {
    return { success: false, message: "Error Mesin AI Groq: " + e.message };
  }
}
// =======================================================================
// FUNGSI AI GENERATOR ANALISIS PENGEMBANGAN DIRI (VIA GROQ)
// =======================================================================
function generateAnalisisPengembanganDiriAI(payloadStr) {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY');
    if (!apiKey) return { success: false, message: "Catatan: GROQ_API_KEY belum dipasang di Script Properties." };

    var data = JSON.parse(payloadStr);
    
    var prompt = "Anda adalah Analis Kepegawaian Eksekutif. Buat SATU paragraf padat (maksimal 4 kalimat) berisi analisis dan kesimpulan atas laporan kegiatan Pengembangan Diri / Kompetensi ASN untuk periode " + data.periode + ".\n\n" +
                 "Data Laporan:\n" +
                 "- Total Kegiatan Diklat/Bimtek: " + data.totalKegiatan + " Kegiatan\n" +
                 "- Total Pegawai Terlibat: " + data.totalPegawai + " Orang\n" +
                 "- Total Jam Pelajaran (JP) Diperoleh: " + data.totalJP + " JP\n\n" +
                 "Fokuskan paragraf pada apresiasi terhadap peningkatan kompetensi pegawai dan berikan harapan/rekomendasi agar ilmu tersebut diimplementasikan untuk mendorong kinerja instansi. Jangan gunakan list/poin-poin, murni 1 paragraf teks narasi saja.";

    var url = "https://api.groq.com/openai/v1/chat/completions";
    var payload = {
      "model": "llama-3.1-8b-instant",
      "messages": [ { "role": "user", "content": prompt } ],
      "temperature": 0.2,
      "max_tokens": 300
    };
    
    var options = {
      "method": "post",
      "headers": {
        "Authorization": "Bearer " + apiKey.trim(),
        "Content-Type": "application/json"
      },
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());

    if (json.error) return { success: false, message: "Gagal menganalisis: " + json.error.message };
    if (!json.choices || json.choices.length === 0) return { success: false, message: "AI gagal memberikan respon." };

    return { success: true, data: json.choices[0].message.content.trim().replace(/\n/g, " ") };
  } catch (e) {
    return { success: false, message: "Error Mesin AI Groq: " + e.message };
  }
}
function pancingIzinDrive() {
  DriveApp.getFiles();
}
function perbaikiIzinDrive() {
  // Memancing Google agar meminta izin akses Drive
  DriveApp.getFolderById("1CSuHamLm9b2YodgAU0hCJpRqN1VCtLCQ");
}