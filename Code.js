/**
 * System Agenda Kegiatan, Announcement & Booking Ruangan (Islamic Luxe 3D)
 * Integration: Google Sheets & Google Drive
 */

const SPREADSHEET_ID = "1gDH6zcX02eaMZwp8O_-MXZvQSqG6MF_GIPTSXU-jrFU";
const DRIVE_FOLDER_ID = "1tdT0Tlq60Z0qzWyM38vNSrJQQ2ImQk8G";

function doGet(e) {
  if (e && e.parameter && e.parameter.action === "getInitialData") {
    try {
      return jsonResponse_(getInitialData());
    } catch (error) {
      return jsonResponse_({ success: false, message: error.message });
    }
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Agenda & Booking Ruangan System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  const params = e && e.parameter ? e.parameter : {};
  try {
    if (params.action === "verifyAdmin") {
      verifyAdminPassword_(params.adminPassword);
      return jsonResponse_({ success: true });
    }

    verifyAdminPassword_(params.adminPassword);
    if (params.action === "simpanAgendaAdmin") {
      if (!params.data) throw new Error("Data agenda tidak ditemukan.");
      return jsonResponse_(simpanAgendaAdmin(JSON.parse(params.data)));
    }
    if (params.action === "publishAnnouncement") {
      const text = String(params.text || "").trim();
      if (!text) throw new Error("Teks pengumuman tidak boleh kosong.");
      if (text.length > 5000) throw new Error("Teks pengumuman maksimal 5.000 karakter.");
      return jsonResponse_(saveAnnouncement_(text, true));
    }
    if (params.action === "unpublishAnnouncement") {
      return jsonResponse_(saveAnnouncement_(null, false));
    }
    if (params.action === "addCalendarEvent") {
      if (!params.data) throw new Error("Data hari khusus tidak ditemukan.");
      return jsonResponse_(addCalendarEvent_(JSON.parse(params.data)));
    }
    if (params.action === "deleteCalendarEvent") {
      return jsonResponse_(deleteCalendarEvent_(params.id));
    }
    throw new Error("Aksi tidak dikenali.");
  } catch (error) {
    return jsonResponse_({ success: false, message: error.message });
  }
}

function jsonResponse_(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function verifyAdminPassword_(password) {
  // Configure ADMIN_PASSWORD under Apps Script Project Settings > Script Properties.
  const configuredPassword = PropertiesService.getScriptProperties().getProperty("ADMIN_PASSWORD");
  if (!configuredPassword) {
    throw new Error("Kata sandi admin belum dikonfigurasi. Atur Script Property ADMIN_PASSWORD di Apps Script.");
  }
  if (!password || password !== configuredPassword) {
    throw new Error("Kata sandi admin salah.");
  }
}

/**
 * Combined Data Retrieval for Agenda & Announcements
 */
function getInitialData() {
  try {
    const agenda = getAgendaData();
    const pengumuman = getAnnouncementData();
    const ruangan = getRuanganData();
    return {
      agenda: agenda,
      pengumuman: pengumuman.text,
      pengumumanPublished: pengumuman.isPublished,
      hariKhusus: getCalendarEvents_(),
      ruangan: ruangan
    };
  } catch (e) {
    throw new Error("Gagal mengambil data sistem: " + e.message);
  }
}

function getCalendarEventsSheet_() {
  const ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("HariKhusus");
  if (!sheet) {
    sheet = ss.insertSheet("HariKhusus");
    sheet.appendRow(["ID", "Tanggal", "Nama", "Jenis", "Waktu Pembaruan"]);
    sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#064E3B").setFontColor("#F59E0B");
  }
  return sheet;
}

function getCalendarEvents_() {
  const sheet = getCalendarEventsSheet_();
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues()
    .filter(row => row[0] && row[1] && row[2])
    .map(row => ({
      id: String(row[0]),
      tanggal: row[1] instanceof Date
        ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), "yyyy-MM-dd")
        : String(row[1]),
      nama: String(row[2]),
      jenis: String(row[3])
    }));
}

function addCalendarEvent_(event) {
  const tanggal = String(event.tanggal || "");
  const nama = String(event.nama || "").trim();
  const jenis = String(event.jenis || "");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(tanggal) || isNaN(Date.parse(tanggal + "T00:00:00Z"))) {
    throw new Error("Tanggal harus diisi dengan format yang valid.");
  }
  if (!nama || nama.length > 150) throw new Error("Nama hari khusus wajib diisi (maksimal 150 karakter).");
  if (!["Libur nasional", "Cuti bersama", "Hari besar Islam"].includes(jenis)) {
    throw new Error("Jenis hari khusus tidak valid.");
  }

  const sheet = getCalendarEventsSheet_();
  const existing = getCalendarEvents_();
  if (existing.some(item => item.tanggal === tanggal && item.nama.toLowerCase() === nama.toLowerCase())) {
    throw new Error("Tanggal dan nama hari tersebut sudah terdaftar.");
  }
  const id = "HK-" + Date.now();
  sheet.appendRow([id, tanggal, nama, jenis, new Date()]);
  return {
    success: true,
    item: { id: id, tanggal: tanggal, nama: nama, jenis: jenis },
    message: "Hari khusus berhasil disimpan ke spreadsheet."
  };
}

function deleteCalendarEvent_(id) {
  if (!id) throw new Error("ID hari khusus tidak ditemukan.");
  const sheet = getCalendarEventsSheet_();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "Hari khusus berhasil dihapus." };
    }
  }
  throw new Error("Hari khusus tidak ditemukan.");
}

function getAnnouncementSheet_() {
  const ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Pengumuman");
  if (!sheet) {
    sheet = ss.insertSheet("Pengumuman");
    sheet.appendRow(["Teks Pengumuman", "Status Terbit", "Waktu Pembaruan"]);
    sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#064E3B").setFontColor("#F59E0B");
  }
  return sheet;
}

function getAnnouncementData() {
  const sheet = getAnnouncementSheet_();
  if (sheet.getLastRow() < 2) return { text: "", isPublished: false };
  const row = sheet.getRange(2, 1, 1, 3).getValues()[0];
  return {
    text: row[0] ? String(row[0]) : "",
    isPublished: row[1] === true || String(row[1]).toLowerCase() === "true"
  };
}

function saveAnnouncement_(text, isPublished) {
  const sheet = getAnnouncementSheet_();
  const existing = getAnnouncementData();
  const savedText = text === null ? existing.text : text;
  const values = [[savedText, isPublished, new Date()]];
  sheet.getRange(2, 1, 1, 3).setValues(values);
  return {
    success: true,
    pengumuman: savedText,
    pengumumanPublished: isPublished,
    message: isPublished ? "Pengumuman berhasil diterbitkan." : "Pengumuman berhasil dicabut."
  };
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
 * Simpan Agenda Baru (Khusus Admin Input) dengan Validasi Bentrok Tanggal & Ruangan
 */
function simpanAgendaAdmin(formObject, fileData) {
  try {
    const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Agenda");

    // ================= VALIDASI PENCEGAHAN BENTROK =================
    const data = sheet.getDataRange().getValues();
    if (data.length > 1) {
      const targetTgl = formObject.tglMulai;
      const targetRuangan = formObject.ruangan;

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const tglExist = row[5] instanceof Date ? formatDate(row[5]) : row[5];
        const ruanganExist = row[4];
        const statusExist = row[10];

        // Jika tanggal sama, ruangan sama, dan status Approved -> TOLAK!
        if (tglExist === targetTgl && ruanganExist === targetRuangan && statusExist === "Approved") {
          return { 
            success: false, 
            message: `[BENTROK JADWAL] Ruangan "${targetRuangan}" sudah di-booking untuk kegiatan lain pada tanggal ${targetTgl}. Permohonan ditolak otomatis!` 
          };
        }
      }
    }
    // ==============================================================

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
      "Approved", // Direct Admin Input
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
 * Kelola Data Ruangan / Aula (Two-Way Synchronization Support)
 */
function getRuanganData() {
  try {
    const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Ruangan");
    if (!sheet) {
      sheet = ss.insertSheet("Ruangan");
      sheet.appendRow(["ID", "Nama Ruangan", "Keterangan"]);
      sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#064E3B").setFontColor("#F59E0B");
      // Default initial rooms
      sheet.appendRow(["RNG-1", "Aula Utama", "Gedung Utama Gold Dome"]);
      sheet.appendRow(["RNG-2", "Ruang Rapat 1", "Ruang Rapat Lantai 1"]);
      sheet.appendRow(["RNG-3", "Ruang Rapat 2", "Ruang Rapat Lantai 2"]);
      sheet.appendRow(["RNG-4", "Auditorium", "Auditorium Spatial"]);
    }
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    // Mengambil seluruh baris data dari Google Sheets (Baris 2 ke bawah)
    return data.slice(1).filter(row => row[1] && row[1].toString().trim() !== "").map((row, idx) => ({
      id: row[0] ? row[0].toString() : ("RNG-AUTO-" + (idx + 1)),
      nama: row[1].toString().trim(),
      keterangan: row[2] ? row[2].toString().trim() : ""
    }));
  } catch (e) {
    return [];
  }
}

function simpanRuangan(id, nama, keterangan) {
  try {
    const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Ruangan");
    if (!sheet) {
      sheet = ss.insertSheet("Ruangan");
      sheet.appendRow(["ID", "Nama Ruangan", "Keterangan"]);
    }
    const data = sheet.getDataRange().getValues();

    if (id) {
      // Edit Mode: Update baris yang cocok di Spreadsheet berdasarkan ID atau Nama
      for (let i = 1; i < data.length; i++) {
        if (data[i][0].toString() === id.toString() || data[i][1].toString().trim().toLowerCase() === nama.trim().toLowerCase()) {
          sheet.getRange(i + 1, 1).setValue(id);
          sheet.getRange(i + 1, 2).setValue(nama);
          sheet.getRange(i + 1, 3).setValue(keterangan || "");
          return { success: true, message: `Ruangan "${nama}" berhasil diperbarui di Google Sheets!` };
        }
      }
    }
    
    // Add Mode: Baris Baru ke Spreadsheet
    const newId = id || ("RNG-" + Date.now());
    sheet.appendRow([newId, nama, keterangan || ""]);
    return { success: true, message: `Ruangan baru "${nama}" berhasil tersimpan ke Google Sheets!` };
  } catch (e) {
    return { success: false, message: "Gagal memperbarui Google Sheets: " + e.message };
  }
}

function deleteRuangan(id) {
  try {
    const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Ruangan");
    if (!sheet) return { success: false, message: "Sheet Ruangan tidak ditemukan." };
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === id.toString()) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "Ruangan berhasil dihapus dari Google Sheets!" };
      }
    }
    return { success: false, message: "Ruangan tidak ditemukan di Google Sheets." };
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