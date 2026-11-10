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
  if (e && e.parameter && e.parameter.action === "getAutomaticCalendar") {
    try {
      const year = Number(e.parameter.year);
      if (!Number.isInteger(year) || year < 2000 || year > 2100) {
        throw new Error("Tahun kalender tidak valid.");
      }
      return jsonResponse_({
        success: true,
        year: year,
        events: getAutomaticCalendarEvents_(year)
      });
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
    if (params.action === "updateAgendaAdmin") {
      if (!params.id || !params.data) throw new Error("ID dan data agenda wajib diisi.");
      return jsonResponse_(updateAgendaAdmin_(params.id, JSON.parse(params.data)));
    }
    if (params.action === "deleteAgendaAdmin") {
      return jsonResponse_(deleteAgendaAdmin_(params.id));
    }
    if (params.action === "setAgendaPublishStatus") {
      return jsonResponse_(setAgendaPublishStatus_(params.id, params.published === "true"));
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

function getAutomaticCalendarEvents_(year) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "automatic-calendar-" + year;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const events = year === 2026
    ? getOfficialCalendar2026_()
    : year === 2027
      ? getOfficialCalendar2027_()
      : getExternalCalendarEvents_(year);
  cache.put(cacheKey, JSON.stringify(events), 21600);
  return events;
}

function getOfficialCalendar2026_() {
  const events = [
    ["2026-01-01", "Tahun Baru 2026 Masehi", "Libur nasional"],
    ["2026-01-16", "Isra Mikraj Nabi Muhammad SAW", "Libur nasional / Hari besar Islam"],
    ["2026-02-16", "Cuti Bersama Tahun Baru Imlek 2577 Kongzili", "Cuti bersama"],
    ["2026-02-17", "Tahun Baru Imlek 2577 Kongzili", "Libur nasional"],
    ["2026-03-18", "Cuti Bersama Hari Suci Nyepi", "Cuti bersama"],
    ["2026-03-19", "Hari Suci Nyepi (Tahun Baru Saka 1948)", "Libur nasional"],
    ["2026-03-20", "Cuti Bersama Hari Raya Idulfitri 1447 H", "Cuti bersama"],
    ["2026-03-21", "Hari Raya Idulfitri 1447 H", "Libur nasional / Hari besar Islam"],
    ["2026-03-22", "Hari Raya Idulfitri 1447 H", "Libur nasional / Hari besar Islam"],
    ["2026-03-23", "Cuti Bersama Hari Raya Idulfitri 1447 H", "Cuti bersama"],
    ["2026-03-24", "Cuti Bersama Hari Raya Idulfitri 1447 H", "Cuti bersama"],
    ["2026-04-03", "Wafat Yesus Kristus", "Libur nasional"],
    ["2026-04-05", "Kebangkitan Yesus Kristus (Paskah)", "Libur nasional"],
    ["2026-05-01", "Hari Buruh Internasional", "Libur nasional"],
    ["2026-05-14", "Kenaikan Yesus Kristus", "Libur nasional"],
    ["2026-05-15", "Cuti Bersama Kenaikan Yesus Kristus", "Cuti bersama"],
    ["2026-05-27", "Hari Raya Iduladha 1447 H", "Libur nasional / Hari besar Islam"],
    ["2026-05-28", "Cuti Bersama Hari Raya Iduladha 1447 H", "Cuti bersama"],
    ["2026-05-31", "Hari Raya Waisak 2570 BE", "Libur nasional"],
    ["2026-06-01", "Hari Lahir Pancasila", "Libur nasional"],
    ["2026-06-16", "Tahun Baru Islam 1448 H", "Libur nasional / Hari besar Islam"],
    ["2026-08-17", "Proklamasi Kemerdekaan Republik Indonesia", "Libur nasional"],
    ["2026-08-25", "Maulid Nabi Muhammad SAW", "Libur nasional / Hari besar Islam"],
    ["2026-12-24", "Cuti Bersama Kelahiran Yesus Kristus (Natal)", "Cuti bersama"],
    ["2026-12-25", "Kelahiran Yesus Kristus (Natal)", "Libur nasional"]
  ];
  return events.map(function(event, index) {
    return {
      id: "AUTO-2026-" + (index + 1),
      tanggal: event[0],
      nama: event[1],
      jenis: event[2],
      sumber: "SKB 3 Menteri 2026",
      otomatis: true
    };
  });
}

function getOfficialCalendar2027_() {
  const events = [
    ["2027-01-01", "Tahun Baru Masehi", "Libur nasional"],
    ["2027-01-05", "Isra Mikraj Nabi Muhammad SAW 1448 H", "Libur nasional / Hari besar Islam"],
    ["2027-02-05", "Cuti Bersama Tahun Baru Imlek", "Cuti bersama"],
    ["2027-02-06", "Tahun Baru Imlek 2578 Kongzili", "Libur nasional"],
    ["2027-03-08", "Hari Raya Nyepi Tahun Baru Saka 1949", "Libur nasional"],
    ["2027-03-09", "Cuti Bersama Hari Raya Idul Fitri 1448 H", "Cuti bersama"],
    ["2027-03-10", "Hari Raya Idul Fitri 1448 H", "Libur nasional / Hari besar Islam"],
    ["2027-03-11", "Hari Raya Idul Fitri 1448 H", "Libur nasional / Hari besar Islam"],
    ["2027-03-12", "Cuti Bersama Hari Raya Idul Fitri 1448 H", "Cuti bersama"],
    ["2027-03-15", "Cuti Bersama Hari Raya Idul Fitri 1448 H", "Cuti bersama"],
    ["2027-03-25", "Cuti Bersama Wafat Isa Almasih", "Cuti bersama"],
    ["2027-03-26", "Wafat Isa Almasih", "Libur nasional"],
    ["2027-03-28", "Kebangkitan Yesus Kristus (Paskah)", "Libur nasional"],
    ["2027-05-01", "Hari Buruh Internasional", "Libur nasional"],
    ["2027-05-06", "Kenaikan Isa Almasih", "Libur nasional"],
    ["2027-05-17", "Hari Raya Idul Adha 1448 H", "Libur nasional / Hari besar Islam"],
    ["2027-05-18", "Cuti Bersama Hari Raya Idul Adha 1448 H", "Cuti bersama"],
    ["2027-05-19", "Cuti Bersama Hari Raya Waisak 2571", "Cuti bersama"],
    ["2027-05-20", "Hari Raya Waisak 2571", "Libur nasional"],
    ["2027-06-01", "Hari Lahir Pancasila", "Libur nasional"],
    ["2027-06-06", "Tahun Baru Islam 1449 H", "Libur nasional / Hari besar Islam"],
    ["2027-08-15", "Maulid Nabi Muhammad SAW 1449 H", "Libur nasional / Hari besar Islam"],
    ["2027-08-17", "Hari Kemerdekaan Republik Indonesia", "Libur nasional"],
    ["2027-12-24", "Cuti Bersama Hari Raya Natal", "Cuti bersama"],
    ["2027-12-25", "Hari Raya Natal", "Libur nasional"],
    ["2027-12-26", "Isra Mikraj Nabi Muhammad SAW 1449 H", "Libur nasional / Hari besar Islam"]
  ];
  return events.map(function(event, index) {
    return {
      id: "AUTO-2027-" + (index + 1),
      tanggal: event[0],
      nama: event[1],
      jenis: event[2],
      sumber: "SKB 3 Menteri 2027",
      otomatis: true
    };
  });
}

function getExternalCalendarEvents_(year) {
  const requests = [{
    url: "https://date.nager.at/api/v3/PublicHolidays/" + year + "/ID",
    muteHttpExceptions: true
  }];
  for (let month = 1; month <= 12; month++) {
    requests.push({
      url: "https://api.aladhan.com/v1/gToHCalendar/" + month + "/" + year + "?method=20",
      muteHttpExceptions: true
    });
  }

  const responses = UrlFetchApp.fetchAll(requests);
  const publicHolidays = parseCalendarApiResponse_(responses[0], "kalender libur nasional");
  if (!Array.isArray(publicHolidays)) {
    throw new Error("Format kalender libur nasional tidak sesuai.");
  }

  const events = publicHolidays
    .filter(function(item) {
      return item && /^\d{4}-\d{2}-\d{2}$/.test(String(item.date || "")) && item.localName;
    })
    .map(function(item, index) {
      return {
        id: "AUTO-" + year + "-N-" + (index + 1),
        tanggal: item.date,
        nama: String(item.localName),
        jenis: "Libur nasional",
        sumber: "Kalender publik Nager.Date (referensi)",
        perkiraan: true,
        otomatis: true
      };
    });

  const islamicEvents = {
    "lailat-ul-miraj": "Isra Mikraj Nabi Muhammad SAW",
    "lailat-ul-bara'at": "Nisfu Syaban",
    "eid-ul-fitr": "Hari Raya Idul Fitri",
    "eid-ul-adha": "Hari Raya Idul Adha",
    "ashura": "Hari Asyura",
    "mawlid (birth) al-nabi": "Maulid Nabi Muhammad SAW",
    "lailat-ul-qadr": "Lailatul Qadar (perkiraan)"
  };
  const seenIslamicDates = {};

  responses.slice(1).forEach(function(response, monthIndex) {
    const calendar = parseCalendarApiResponse_(response, "kalender Hijriah bulan " + (monthIndex + 1));
    if (!Array.isArray(calendar)) {
      throw new Error("Format kalender Hijriah bulan " + (monthIndex + 1) + " tidak sesuai.");
    }
    calendar.forEach(function(day) {
      const hijri = day && day.hijri;
      const gregorianDate = day && day.gregorian && String(day.gregorian.date || "");
      const dateMatch = gregorianDate.match(/^(\d{2})-(\d{2})-(\d{4})$/);
      if (!hijri || !dateMatch) return;
      const date = dateMatch[3] + "-" + dateMatch[2] + "-" + dateMatch[1];
      const dayNumber = String(hijri.day || "").replace(/^0+/, "");
      const holidays = Array.isArray(hijri.holidays) ? hijri.holidays : [];
      if (Number(hijri.month && hijri.month.number) === 1 && dayNumber === "1") {
        addIslamicCalendarEvent_(events, seenIslamicDates, date, "Tahun Baru Islam " + hijri.year + " H");
      }
      holidays.forEach(function(holiday) {
        const normalizedHoliday = String(holiday).toLowerCase();
        const nameKey = Object.keys(islamicEvents).find(function(key) {
          return normalizedHoliday.indexOf(key) !== -1;
        });
        if (nameKey) {
          addIslamicCalendarEvent_(events, seenIslamicDates, date, islamicEvents[nameKey] + " " + hijri.year + " H");
        }
      });
    });
  });

  return events.sort(function(a, b) {
    return a.tanggal.localeCompare(b.tanggal) || a.nama.localeCompare(b.nama);
  });
}

function parseCalendarApiResponse_(response, sourceName) {
  if (response.getResponseCode() !== 200) {
    throw new Error("Gagal mengambil " + sourceName + " (HTTP " + response.getResponseCode() + ").");
  }
  let result;
  try {
    result = JSON.parse(response.getContentText());
  } catch (error) {
    throw new Error("Respons " + sourceName + " bukan JSON yang valid.");
  }
  if (result && result.code !== undefined && Number(result.code) !== 200) {
    throw new Error("Sumber " + sourceName + " mengembalikan status " + result.status + ".");
  }
  return result && result.data !== undefined ? result.data : result;
}

function addIslamicCalendarEvent_(events, seenDates, date, name) {
  const eventKey = date + "|" + name;
  if (seenDates[eventKey]) return;
  seenDates[eventKey] = true;
  events.push({
    id: "AUTO-" + date.replace(/-/g, "") + "-I-" + Object.keys(seenDates).length,
    tanggal: date,
    nama: name,
    jenis: "Hari besar Islam",
    sumber: "Kalender Hijriah AlAdhan (perkiraan)",
    perkiraan: true,
    otomatis: true
  });
}

function addCalendarEvent_(event) {
  const tanggal = String(event.tanggal || "");
  const nama = String(event.nama || "").trim();
  const jenis = String(event.jenis || "");
  const parsedDate = /^\d{4}-\d{2}-\d{2}$/.test(tanggal)
    ? new Date(tanggal + "T00:00:00Z")
    : null;
  if (!parsedDate || isNaN(parsedDate.getTime()) || parsedDate.toISOString().slice(0, 10) !== tanggal) {
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

    const operatorColumn = getAgendaOperatorColumn_(sheet);
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
        created: row[13],
        operatorAula: row[operatorColumn - 1] || ""
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
    if (!sheet) throw new Error("Sheet Agenda tidak ditemukan.");
    validateAgendaForm_(formObject);
    assertNoAgendaConflict_(sheet, formObject, "");

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
      formObject.namaKegiatan || formObject.nama,
      formObject.penyelenggara,
      formObject.fakultas || "Umum",
      formObject.ruangan,
      formObject.tglMulai,
      formObject.tglSelesai || formObject.tglMulai,
      formObject.waktuMulai || formObject.jamMulai,
      formObject.waktuSelesai || formObject.jamSelesai,
      formObject.kategori || "Umum",
      "Approved", // Direct Admin Input
      fileUrl,
      formObject.keterangan || "",
      timestamp
    ]);
    sheet.getRange(sheet.getLastRow(), getAgendaOperatorColumn_(sheet))
      .setValue(String(formObject.operatorAula || "").trim());

    return { success: true, id: newId, message: "Agenda berhasil ditambahkan & disetujui!" };
  } catch (error) {
    return { success: false, message: "Terjadi kesalahan: " + error.message };
  }
}

function validateAgendaForm_(formObject) {
  const name = String(formObject.namaKegiatan || formObject.nama || "").trim();
  const organizer = String(formObject.penyelenggara || "").trim();
  const room = String(formObject.ruangan || "").trim();
  const faculty = String(formObject.fakultas || "").trim();
  const startDate = String(formObject.tglMulai || "");
  const endDate = String(formObject.tglSelesai || startDate);
  const startTime = String(formObject.waktuMulai || formObject.jamMulai || "");
  const endTime = String(formObject.waktuSelesai || formObject.jamSelesai || "");
  if (!name || !organizer || !room || !faculty) throw new Error("Nama kegiatan, penyelenggara, fakultas/unit, dan ruangan wajib diisi.");
  if (!isValidAgendaDate_(startDate) || !isValidAgendaDate_(endDate) || startDate > endDate) {
    throw new Error("Rentang tanggal kegiatan tidak valid.");
  }
  if (agendaTimeMinutes_(startTime) === null || agendaTimeMinutes_(endTime) === null ||
      agendaTimeMinutes_(startTime) >= agendaTimeMinutes_(endTime)) {
    throw new Error("Rentang waktu kegiatan tidak valid.");
  }
}

function isValidAgendaDate_(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  const date = new Date(value + "T00:00:00Z");
  return !isNaN(date.getTime()) && date.toISOString().slice(0, 10) === value;
}

function agendaTimeMinutes_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value.getHours() * 60 + value.getMinutes();
  }
  const match = String(value || "").match(/^(\d{1,2}):(\d{2})/);
  if (!match) return null;
  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  return hours <= 23 && minutes <= 59 ? hours * 60 + minutes : null;
}

function assertNoAgendaConflict_(sheet, formObject, excludedId) {
  const targetStart = String(formObject.tglMulai);
  const targetEnd = String(formObject.tglSelesai || targetStart);
  const targetStartMinutes = agendaTimeMinutes_(formObject.waktuMulai || formObject.jamMulai);
  const targetEndMinutes = agendaTimeMinutes_(formObject.waktuSelesai || formObject.jamSelesai);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (excludedId && String(row[0]) === String(excludedId)) continue;
    if (String(row[10] || "").toLowerCase() !== "approved") continue;
    const existingRoom = String(row[4] || "").trim().toLowerCase();
    if (existingRoom !== String(formObject.ruangan || "").trim().toLowerCase()) continue;
    const existingStart = row[5] instanceof Date ? formatDate(row[5]) : String(row[5] || "");
    const existingEnd = row[6] instanceof Date ? formatDate(row[6]) : String(row[6] || existingStart);
    const datesOverlap = targetStart <= existingEnd && targetEnd >= existingStart;
    const existingStartMinutes = agendaTimeMinutes_(row[7]);
    const existingEndMinutes = agendaTimeMinutes_(row[8]);
    const timesOverlap = existingStartMinutes === null || existingEndMinutes === null ||
      (targetStartMinutes < existingEndMinutes && targetEndMinutes > existingStartMinutes);
    if (datesOverlap && timesOverlap) {
      throw new Error(`[BENTROK JADWAL] Ruangan "${formObject.ruangan}" sudah dibooking pada tanggal ${existingStart} sampai ${existingEnd}.`);
    }
  }
}

function findAgendaRow_(sheet, id) {
  if (!id) throw new Error("ID booking tidak ditemukan.");
  const values = sheet.getDataRange().getValues();
  for (let index = 1; index < values.length; index++) {
    if (String(values[index][0]) === String(id)) return { rowNumber: index + 1, values: values[index] };
  }
  throw new Error("Booking tidak ditemukan di spreadsheet.");
}

function getAgendaOperatorColumn_(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), 14);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const headerName = "Petugas / Operator Aula/Gedung";
  let column = headers.indexOf(headerName) + 1;
  if (!column) {
    column = lastColumn + 1;
    sheet.getRange(1, column).setValue(headerName);
    sheet.getRange(1, column).setFontWeight("bold").setBackground("#064E3B").setFontColor("#F59E0B");
  }
  return column;
}

function updateAgendaAdmin_(id, formObject) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Agenda");
  if (!sheet) throw new Error("Sheet Agenda tidak ditemukan.");
  const existing = findAgendaRow_(sheet, id);
  validateAgendaForm_(formObject);
  assertNoAgendaConflict_(sheet, formObject, id);
  sheet.getRange(existing.rowNumber, 2, 1, 8).setValues([[
    formObject.namaKegiatan || formObject.nama,
    formObject.penyelenggara,
    formObject.fakultas || "Umum",
    formObject.ruangan,
    formObject.tglMulai,
    formObject.tglSelesai || formObject.tglMulai,
    formObject.waktuMulai || formObject.jamMulai,
    formObject.waktuSelesai || formObject.jamSelesai
  ]]);
  sheet.getRange(existing.rowNumber, getAgendaOperatorColumn_(sheet))
    .setValue(String(formObject.operatorAula || "").trim());
  return { success: true, message: "Booking berhasil diperbarui." };
}

function deleteAgendaAdmin_(id) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Agenda");
  if (!sheet) throw new Error("Sheet Agenda tidak ditemukan.");
  const existing = findAgendaRow_(sheet, id);
  sheet.deleteRow(existing.rowNumber);
  return { success: true, message: "Booking berhasil dihapus." };
}

function setAgendaPublishStatus_(id, published) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Agenda");
  if (!sheet) throw new Error("Sheet Agenda tidak ditemukan.");
  const existing = findAgendaRow_(sheet, id);
  if (published) {
    const formObject = {
      namaKegiatan: existing.values[1],
      penyelenggara: existing.values[2],
      fakultas: existing.values[3] || "Umum",
      ruangan: existing.values[4],
      tglMulai: existing.values[5] instanceof Date ? formatDate(existing.values[5]) : String(existing.values[5] || ""),
      tglSelesai: existing.values[6] instanceof Date ? formatDate(existing.values[6]) : String(existing.values[6] || ""),
      waktuMulai: existing.values[7],
      waktuSelesai: existing.values[8]
    };
    validateAgendaForm_(formObject);
    assertNoAgendaConflict_(sheet, formObject, id);
  }
  sheet.getRange(existing.rowNumber, 11).setValue(published ? "Approved" : "Draft");
  return {
    success: true,
    status: published ? "Approved" : "Draft",
    message: published ? "Booking berhasil diterbitkan kembali." : "Booking dipindahkan ke draf."
  };
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