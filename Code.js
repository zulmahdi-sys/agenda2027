/**
 * System Agenda Kegiatan, Announcement & Booking Ruangan (Islamic Luxe 3D)
 * Integration: Google Sheets & Google Drive
 */

const SPREADSHEET_ID = "1gDH6zcX02eaMZwp8O_-MXZvQSqG6MF_GIPTSXU-jrFU";
const DRIVE_FOLDER_ID = "1tdT0Tlq60Z0qzWyM38vNSrJQQ2ImQk8G";

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Agenda & Booking Ruangan System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Combined Data Retrieval for Agenda & Announcements
 */
function getInitialData() {
  try {
    const agenda = getAgendaData();
    const announcements = getAnnouncementData();
    return { agenda: agenda, announcements: announcements };
  } catch (e) {
    throw new Error("Gagal mengambil data sistem: " + e.message);
  }
}

/**
 * Mengambil data agenda dari sheet "Agenda"
 */
function getAgendaData() {
  try {
    const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Agenda");
    
    if (!sheet) {
      sheet = ss.insertSheet("Agenda");
      sheet.appendRow([
        "ID", 
        "Nama Kegiatan", 
        "Penyelenggara / Unit Kerja", 
        "Fakultas / Unit",
        "Ruangan / Aula", 
        "Tanggal Mulai", 
        "Tanggal Selesai", 
        "Waktu Mulai", 
        "Waktu Selesai", 
        "Kategori", 
        "Status", 
        "URL Berkas / Lampiran", 
        "Keterangan", 
        "Waktu Input"
      ]);
      sheet.getRange("A1:N1").setFontWeight("bold").setBackground("#064E3B").setFontColor("#F59E0B");
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const rows = data.slice(1);
    return rows.map((row, index) => {
      return {
        rowId: index + 2,
        id: row[0],
        namaKegiatan: row[1],
        penyelenggara: row[2],
        fakultas: row[3] || "Umum",
        ruangan: row[4],
        tglMulai: row[5] instanceof Date ? formatDate(row[5]) : row[5],
        tglSelesai: row[6] instanceof Date ? formatDate(row[6]) : row[6],
        waktuMulai: row[7],
        waktuSelesai: row[8],
        kategori: row[9],
        status: row[10],
        urlBerkas: row[11],
        keterangan: row[12],
        created: row[13]
      };
    });
  } catch (error) {
    throw new Error("Gagal mengambil agenda: " + error.message);
  }
}

/**
 * Simpan Agenda Baru (Khusus Admin Input)
 */
function simpanAgendaAdmin(formObject, fileData) {
  try {
    const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Agenda");

    let fileUrl = "";
    if (fileData && fileData.data) {
      const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      const decodedData = Utilities.base64Decode(fileData.data.split(',')[1]);
      const blob = Utilities.newBlob(decodedData, fileData.mimeType, fileData.fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
    }

    const newId = "AGD-" + Date.now();
    const timestamp = new Date();

    sheet.appendRow([
      newId,
      formObject.namaKegiatan,
      formObject.penyelenggara,
      formObject.fakultas || "Umum",
      formObject.ruangan,
      formObject.tglMulai,
      formObject.tglSelesai || formObject.tglMulai,
      formObject.waktuMulai,
      formObject.waktuSelesai,
      formObject.kategori || "Umum",
      "Approved", // Direct Admin Input is automatically approved
      fileUrl,
      formObject.keterangan || "",
      timestamp
    ]);

    return { success: true, message: "Agenda berhasil ditambahkan & disetujui!" };
  } catch (error) {
    return { success: false, message: "Terjadi kesalahan: " + error.message };
  }
}

/**
 * Kelola Data Pengumuman Beranda
 */
function getAnnouncementData() {
  try {
    const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Pengumuman");
    if (!sheet) {
      sheet = ss.insertSheet("Pengumuman");
      sheet.appendRow(["ID", "Judul Pengumuman", "Isi Pengumuman", "Tanggal", "Status"]);
      sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#064E3B").setFontColor("#F59E0B");
    }
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    return data.slice(1).map(row => ({
      id: row[0],
      judul: row[1],
      isi: row[2],
      tanggal: row[3] instanceof Date ? formatDate(row[3]) : row[3],
      status: row[4]
    }));
  } catch (e) {
    return [];
  }
}

function simpanPengumuman(judul, isi) {
  try {
    const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Pengumuman");
    if (!sheet) {
      sheet = ss.insertSheet("Pengumuman");
      sheet.appendRow(["ID", "Judul Pengumuman", "Isi Pengumuman", "Tanggal", "Status"]);
    }
    const newId = "ANN-" + Date.now();
    const dateStr = formatDate(new Date());
    sheet.appendRow([newId, judul, isi, dateStr, "Aktif"]);
    return { success: true, message: "Pengumuman berhasil diterbitkan!" };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function deleteAgenda(id) {
  try {
    const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Agenda");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "Agenda berhasil dihapus!" };
      }
    }
    return { success: false, message: "Data tidak ditemukan." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function formatDate(date) {
  const d = new Date(date);
  let month = '' + (d.getMonth() + 1);
  let day = '' + d.getDate();
  const year = d.getFullYear();
  if (month.length < 2) month = '0' + month;
  if (day.length < 2) day = '0' + day;
  return [year, month, day].join('-');
}