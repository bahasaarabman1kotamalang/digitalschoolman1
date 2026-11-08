/**
 * DIGITAL SCHOOL MAN 1 KOTA MALANG
 * Backend Google Apps Script + Google Sheets
 *
 * PENTING:
 * - Database sudah pernah dibuat dan DATABASE_ID tersimpan.
 * - Jangan menjalankan setupDatabase() lagi jika database sudah ada.
 * - Buat file HTML bernama Index dan tempel file Index.html yang disediakan.
 */

const NAMA_DATABASE = "Database Digital School MAN 1 Kota Malang";

/**
 * KONFIGURASI AWAL PRESENSI GPS
 * Nilai ini hanya menjadi default di kode.
 * Tidak ada penulisan/perubahan otomatis ke sheet Settings.
 * Zona waktu dan waktu validasi ditentukan di sisi server.
 */
const KONFIGURASI_PRESENSI_AWAL = {
  active: true,
  schoolName: "MAN 1 Kota Malang",
  latitude: -7.932858303917564,
  longitude: 112.60621037589507,
  radiusMeters: 190,
  startTime: "05:30",
  onTimeUntil: "06:45",
  endTime: "08:00",
  maxAccuracyMeters: 75,
  timeZone: "Asia/Jakarta"
};


const STRUKTUR = {
  Users: ["ID", "Username", "Password", "Nama", "Role", "Status", "PasswordHash", "Salt", "ReferenceID", "MustChangePassword", "CreatedAt", "UpdatedAt", "LastLoginAt", "Permissions"],
  Students: ["ID", "NIS", "Nama", "Kelas", "Jenis Kelamin", "Telepon"],
  Teachers: ["ID", "NIP", "Nama", "Jabatan", "Mata Pelajaran", "Kategori", "Unit Kerja", "Telepon", "Status"],
  Attendance: ["ID", "Tanggal", "ID Siswa", "Nama", "Kelas", "Status", "Latitude", "Longitude"],
  Exams: ["ID", "Judul Ujian", "Mata Pelajaran", "Kelas", "Tanggal"],
  Questions: ["ID", "ID Ujian", "Pertanyaan", "Opsi A", "Opsi B", "Opsi C", "Opsi D", "Opsi E", "Kunci"],
  Grades: ["ID", "NIS", "Nama", "Kelas", "Mata Pelajaran", "Nilai"],
  Finance: ["ID", "Tanggal", "Jenis", "Nominal", "Keterangan"],
  News: ["ID", "Tanggal", "Judul", "Isi"],
  Settings: ["Pengaturan", "Nilai"],
  AuditLog: ["ID", "Waktu", "Pengguna", "Aktivitas"]
};

/**
 * Membuat database hanya jika belum ada.
 * Jika DATABASE_ID sudah tersimpan, fungsi tidak membuat spreadsheet kedua.
 */
function setupDatabase_() {
  const properties = PropertiesService.getScriptProperties();
  const existingId = properties.getProperty("DATABASE_ID");

  if (existingId) {
    const existing = SpreadsheetApp.openById(existingId);
    Logger.log("Database sudah tersedia: " + existing.getName());
    Logger.log("BUKA DATABASE: " + existing.getUrl());
    return { id: existingId, url: existing.getUrl(), existing: true };
  }

  const ss = SpreadsheetApp.create(NAMA_DATABASE);

  Object.keys(STRUKTUR).forEach(nama => {
    const sheet = ss.insertSheet(nama);
    const header = STRUKTUR[nama];
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.getRange(1, 1, 1, header.length)
      .setFontWeight("bold")
      .setBackground("#123653")
      .setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, header.length);
  });

  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet) ss.deleteSheet(defaultSheet);

  const settings = ss.getSheetByName("Settings");
  settings.appendRow(["Nama Madrasah", "MAN 1 Kota Malang"]);
  settings.appendRow(["Latitude", "-7.932858303917564"]);
  settings.appendRow(["Longitude", "112.60621037589507"]);
  settings.appendRow(["Radius Presensi", "200 meter"]);
  settings.appendRow(["Jam Presensi Mulai", "06:30"]);
  settings.appendRow(["Jam Presensi Selesai", "07:00"]);

  properties.setProperty("DATABASE_ID", ss.getId());

  Logger.log("DATABASE BERHASIL DIBUAT");
  Logger.log("BUKA DATABASE: " + ss.getUrl());
  return { id: ss.getId(), url: ss.getUrl(), existing: false };
}

/** Menayangkan antarmuka HTML melalui Apps Script Web App. */
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('MAN 1 Kota Malang - Digital School');
}

const TABEL = {
  students: {
    sheet: "Students",
    fields: ["ID", "NIS", "Nama", "Kelas", "Jenis Kelamin", "Telepon"]
  },
  teachers: {
    sheet: "Teachers",
    fields: ["ID", "NIP", "Nama", "Jabatan", "Mata Pelajaran", "Kategori", "Unit Kerja", "Telepon", "Status"]
  },
  attendance: {
    sheet: "Attendance",
    fields: ["ID", "Tanggal", "ID Siswa", "Nama", "Kelas", "Status", "Latitude", "Longitude"]
  },
  exams: {
    sheet: "Exams",
    fields: ["ID", "Judul Ujian", "Mata Pelajaran", "Kelas", "Tanggal"]
  },
  grades: {
    sheet: "Grades",
    fields: ["ID", "Tanggal", "NIS", "Nama", "Kelas", "Mata Pelajaran", "Nilai"]
  },
  finance: {
    sheet: "Finance",
    fields: ["ID", "Tanggal", "Jenis", "Nominal", "Keterangan"]
  },
  news: {
    sheet: "News",
    fields: ["ID", "Tanggal", "Judul", "Isi"]
  }
};

function getDatabase_() {
  const id = PropertiesService.getScriptProperties().getProperty("DATABASE_ID");
  if (!id) {
    throw new Error("DATABASE_ID belum ditemukan. Database sebelumnya belum terhubung.");
  }
  return SpreadsheetApp.openById(id);
}

function buatId_() {
  return Utilities.getUuid();
}

function bacaTabel_(nama) {
  const konfigurasi = TABEL[nama];
  if (!konfigurasi) throw new Error("Tabel tidak dikenal: " + nama);

  const sheet = getDatabase_().getSheetByName(konfigurasi.sheet);
  if (!sheet) throw new Error("Sheet tidak ditemukan: " + konfigurasi.sheet);

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0];
  return values.slice(1)
    .filter(row => row.some(cell => cell !== ""))
    .map(row => {
      const obj = {};
      headers.forEach((header, i) => obj[header] = row[i]);
      return obj;
    });
}

/** Mengambil data siswa untuk modul Data Siswa. */

/**
 * Mengambil konfigurasi presensi default.
 * Fungsi ini hanya membaca konstanta kode dan tidak menulis ke Google Sheets.
 */
function apiAmbilKonfigurasiPresensi(sessionToken) {
  const actor = requireSession_(sessionToken);
  requirePermission_(actor, "attendance.view");
  return Object.assign({}, KONFIGURASI_PRESENSI_AWAL);
}

/**
 * Uji kelayakan lokasi GPS dari perangkat.
 * PENTING: fungsi ini hanya memvalidasi dan mengembalikan hasil simulasi.
 * Fungsi ini TIDAK menyimpan presensi ke sheet Attendance.
 * Penyimpanan presensi produksi harus menunggu autentikasi siswa yang aman.
 */
function apiUjiLokasiPresensi(payload, sessionToken) {
  const actor = requireSession_(sessionToken);
  requirePermission_(actor, "attendance.test");
  if (!payload) throw new Error("Data lokasi tidak diterima.");

  const cfg = KONFIGURASI_PRESENSI_AWAL;
  const latitude = Number(payload.latitude);
  const longitude = Number(payload.longitude);
  const accuracy = Number(payload.accuracyMeters);

  const now = new Date();
  const tanggalServer = Utilities.formatDate(now, cfg.timeZone, "yyyy-MM-dd");
  const waktuServer = Utilities.formatDate(now, cfg.timeZone, "HH:mm");
  const waktuLengkap = Utilities.formatDate(now, cfg.timeZone, "yyyy-MM-dd HH:mm:ss");

  const hasil = {
    eligible: false,
    previewOnly: true,
    saved: false,
    date: tanggalServer,
    serverTime: waktuLengkap,
    time: waktuServer,
    distanceMeters: null,
    radiusMeters: cfg.radiusMeters,
    accuracyMeters: isFinite(accuracy) ? Math.round(accuracy) : null,
    maxAccuracyMeters: cfg.maxAccuracyMeters,
    attendanceStatus: "",
    message: ""
  };

  if (!cfg.active) {
    hasil.message = "Presensi sedang dinonaktifkan oleh administrator.";
    return hasil;
  }

  if (!isFinite(latitude) || latitude < -90 || latitude > 90 ||
      !isFinite(longitude) || longitude < -180 || longitude > 180) {
    hasil.message = "Koordinat GPS tidak valid. Aktifkan lokasi perangkat dan coba kembali.";
    return hasil;
  }

  if (!isFinite(accuracy) || accuracy <= 0) {
    hasil.message = "Akurasi GPS tidak terbaca. Coba aktifkan GPS dengan mode akurasi tinggi.";
    return hasil;
  }

  hasil.distanceMeters = Math.round(
    hitungJarakMeter_(
      latitude,
      longitude,
      Number(cfg.latitude),
      Number(cfg.longitude)
    )
  );

  if (accuracy > Number(cfg.maxAccuracyMeters)) {
    hasil.message = "Akurasi GPS belum memenuhi syarat. Maksimal ketidakpastian yang diterima adalah " +
      cfg.maxAccuracyMeters + " meter.";
    return hasil;
  }

  if (hasil.distanceMeters > Number(cfg.radiusMeters)) {
    hasil.message = "Lokasi perangkat berada di luar radius presensi madrasah.";
    return hasil;
  }

  if (waktuServer < cfg.startTime || waktuServer > cfg.endTime) {
    hasil.message = "Di luar jadwal presensi. Jadwal yang berlaku " +
      cfg.startTime + "–" + cfg.endTime + " WIB.";
    return hasil;
  }

  hasil.eligible = true;
  hasil.attendanceStatus = waktuServer <= cfg.onTimeUntil ? "Hadir" : "Terlambat";
  hasil.message = hasil.attendanceStatus === "Hadir"
    ? "Lokasi dan waktu memenuhi ketentuan presensi tepat waktu."
    : "Lokasi memenuhi ketentuan. Jika presensi resmi dibuka, status waktu adalah terlambat.";

  return hasil;
}

/** Menghitung jarak dua koordinat GPS menggunakan rumus Haversine. */
function hitungJarakMeter_(lat1, lon1, lat2, lon2) {
  const radiusBumi = 6371000;
  const toRad = function(derajat) { return derajat * Math.PI / 180; };

  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
    Math.sin(dLon / 2) * Math.sin(dLon / 2);

  return 2 * radiusBumi * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function apiAmbilSiswa(sessionToken) {
  const actor = requireSession_(sessionToken);
  requirePermission_(actor, "students.view");
  const records = bacaTabel_("students").map(x => ({
    recordId: String(x.ID || ""),
    id: String(x.NIS || ""),
    name: String(x.Nama || ""),
    kelas: String(x.Kelas || ""),
    gender: String(x["Jenis Kelamin"] || ""),
    phone: String(x.Telepon || "")
  }));
  if (actor.role === "SISWA") {
    return records.filter(x => x.id === actor.referenceId);
  }
  return records;
}

/**
 * Menambah atau memperbarui siswa.
 * recordId adalah ID internal database; id adalah NIS siswa.
 */
function apiSimpanSiswa(data, sessionToken) {
  const actor = requireSession_(sessionToken);
  requirePermission_(actor, "students.manage");
  if (!data) throw new Error("Data siswa tidak diterima.");

  const nis = String(data.id || "").trim();
  const nama = String(data.name || "").trim();
  const kelas = String(data.kelas || "").trim();
  const gender = String(data.gender || "").trim();
  const telepon = String(data.phone || "").trim();
  const recordId = String(data.recordId || "").trim();

  if (!nis || !nama || !kelas) {
    throw new Error("NIS, nama, dan kelas wajib diisi.");
  }
  if (gender && ["L", "P"].indexOf(gender) === -1) {
    throw new Error("Jenis kelamin harus L atau P.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getDatabase_().getSheetByName("Students");
    if (!sheet) throw new Error("Sheet Students tidak ditemukan.");

    const lastRow = sheet.getLastRow();
    const values = lastRow > 1
      ? sheet.getRange(2, 1, lastRow - 1, 6).getValues()
      : [];

    const duplicate = values.findIndex(row =>
      String(row[1] || "").trim() === nis &&
      String(row[0] || "").trim() !== recordId
    );

    if (duplicate !== -1) {
      throw new Error("NIS " + nis + " sudah terdaftar. Periksa kembali data siswa.");
    }

    if (recordId) {
      const index = values.findIndex(row => String(row[0]) === recordId);
      if (index === -1) throw new Error("Data siswa yang akan diperbarui tidak ditemukan.");

      const rowNumber = index + 2;
      sheet.getRange(rowNumber, 1, 1, 6)
        .setValues([[recordId, nis, nama, kelas, gender, telepon]]);

      return { success: true, recordId: recordId, action: "updated" };
    }

    const newRecordId = buatId_();
    sheet.appendRow([newRecordId, nis, nama, kelas, gender, telepon]);

    return { success: true, recordId: newRecordId, action: "created" };
  } finally {
    lock.releaseLock();
  }
}

/** Menghapus satu siswa berdasarkan ID internal, bukan berdasarkan nomor baris. */
function apiHapusSiswa(recordId, sessionToken) {
  const actor = requireSession_(sessionToken);
  requirePermission_(actor, "students.manage");
  const id = String(recordId || "").trim();
  if (!id) throw new Error("ID internal siswa tidak valid.");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getDatabase_().getSheetByName("Students");
    if (!sheet) throw new Error("Sheet Students tidak ditemukan.");

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("Data siswa belum tersedia.");

    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const index = ids.findIndex(row => String(row[0]) === id);
    if (index === -1) throw new Error("Data siswa tidak ditemukan.");

    sheet.deleteRow(index + 2);
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}


/**
 * MIGRASI RINGAN SHEET TEACHERS
 * Mempertahankan kolom dan baris lama, hanya menambahkan header yang belum ada.
 * Tidak membuat spreadsheet baru dan tidak menghapus/mengubah data personel lama.
 */
function siapkanSheetGuruTendik_() {
  const db = getDatabase_();
  let sheet = db.getSheetByName("Teachers");
  if (!sheet) sheet = db.insertSheet("Teachers");

  const required = [
    "ID", "NIP", "Nama", "Jabatan", "Mata Pelajaran",
    "Kategori", "Unit Kerja", "Telepon", "Status"
  ];

  const lastCol = sheet.getLastColumn();
  let headers = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map((x, i) => {
        const name = String(x || "").trim();
        return name || (i < 5 ? "Kolom Lama " + (i + 1) : "Kolom Tambahan " + (i + 1));
      })
    : [];

  if (!headers.length || headers.every(h => !h)) headers = [];
  required.forEach(name => {
    if (headers.indexOf(name) === -1) headers.push(name);
  });

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#123653")
    .setFontColor("#ffffff");
  sheet.setFrozenRows(1);

  // Beri ID internal pada baris lama yang belum memilikinya agar dapat diedit/dihapus dengan aman.
  const idCol = headers.indexOf("ID");
  const dataLastRow = sheet.getLastRow();
  if (idCol >= 0 && dataLastRow > 1) {
    const idRange = sheet.getRange(2, idCol + 1, dataLastRow - 1, 1);
    const ids = idRange.getValues();
    let changed = false;
    ids.forEach(row => {
      if (!String(row[0] || "").trim()) {
        row[0] = buatId_();
        changed = true;
      }
    });
    if (changed) idRange.setValues(ids);
  }

  return { sheet: sheet, headers: headers };
}

function nilaiKolomGuru_(row, headers, name) {
  const index = headers.indexOf(name);
  return index < 0 ? "" : row[index];
}

function teksAmanSheet_(value) {
  const text = String(value == null ? "" : value).trim();
  return /^[=+@-]/.test(text) ? "'" + text : text;
}

/** Mengambil seluruh data Guru dan Tendik dari sheet Teachers. */
function apiAmbilGuruPegawai(sessionToken) {
  const actor = requireSession_(sessionToken);
  requirePermission_(actor, "teachers.view");
  const setup = siapkanSheetGuruTendik_();
  const sheet = setup.sheet;
  const headers = setup.headers;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const rows = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const records = rows.filter(row => row.some(cell => cell !== "" && cell !== null)).map(row => {
    const position = String(nilaiKolomGuru_(row, headers, "Jabatan") || "").trim();
    const subject = String(nilaiKolomGuru_(row, headers, "Mata Pelajaran") || "").trim();
    let category = String(nilaiKolomGuru_(row, headers, "Kategori") || "").trim();

    // Data lama tidak diubah. Kategori hanya diinferensikan untuk tampilan jika kolom baru kosong.
    if (!category) {
      const p = position.toLowerCase();
      if (/guru|kepala madrasah|wakil kepala|waka/.test(p) || subject) category = "Guru";
      else if (/pegawai|tenaga kependidikan|tata usaha|operator|pustakawan|laboran|satpam|pramubakti/.test(p)) category = "Tenaga Kependidikan";
      else category = "Belum dikategorikan";
    }

    return {
      recordId: String(nilaiKolomGuru_(row, headers, "ID") || ""),
      id: String(nilaiKolomGuru_(row, headers, "NIP") || ""),
      name: String(nilaiKolomGuru_(row, headers, "Nama") || ""),
      category: category,
      position: position,
      subject: subject,
      unit: String(nilaiKolomGuru_(row, headers, "Unit Kerja") || "").trim(),
      phone: String(nilaiKolomGuru_(row, headers, "Telepon") || "").trim(),
      status: String(nilaiKolomGuru_(row, headers, "Status") || "Aktif").trim() || "Aktif"
    };
  });
  // Guru dan Tendik hanya menerima profilnya sendiri; Admin dapat melihat seluruh personel.
  if (actor.role === "ADMIN" || actor.permissions.indexOf("teachers.manage") >= 0) return records;
  return records.filter(item => item.id === actor.referenceId);
}

/** Menambah atau memperbarui data Guru/Tendik dengan validasi NIP/ID unik. */
function apiSimpanGuruPegawai(data, sessionToken) {
  const actor = requireSession_(sessionToken);
  requirePermission_(actor, "teachers.manage");
  if (!data) throw new Error("Data personel tidak diterima.");

  const id = String(data.id || "").trim();
  const name = String(data.name || "").trim();
  const category = String(data.category || "").trim();
  const position = String(data.position || "").trim();
  const subject = String(data.subject || "").trim();
  const unit = String(data.unit || "").trim();
  const phone = String(data.phone || "").trim();
  const status = String(data.status || "Aktif").trim();
  const recordId = String(data.recordId || "").trim();

  if (!id || !name) throw new Error("NIP/ID dan nama lengkap wajib diisi.");
  if (["Guru", "Tenaga Kependidikan"].indexOf(category) === -1) {
    throw new Error("Kategori harus Guru atau Tenaga Kependidikan.");
  }
  if (["Aktif", "Nonaktif"].indexOf(status) === -1) {
    throw new Error("Status harus Aktif atau Nonaktif.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const setup = siapkanSheetGuruTendik_();
    const sheet = setup.sheet;
    const headers = setup.headers;
    const lastRow = sheet.getLastRow();
    const rows = lastRow > 1
      ? sheet.getRange(2, 1, lastRow - 1, headers.length).getValues()
      : [];
    const idCol = headers.indexOf("ID");
    const nipCol = headers.indexOf("NIP");

    const duplicate = rows.findIndex(row =>
      String(row[nipCol] || "").trim() === id &&
      String(row[idCol] || "").trim() !== recordId
    );
    if (duplicate !== -1) throw new Error("NIP/ID tersebut sudah terdaftar.");

    let targetRow = -1;
    if (recordId) {
      targetRow = rows.findIndex(row => String(row[idCol] || "").trim() === recordId);
      if (targetRow === -1) throw new Error("Data personel yang akan diperbarui tidak ditemukan. Silakan sinkronkan ulang.");
    }

    const internalId = recordId || buatId_();
    const output = new Array(headers.length).fill("");
    const values = {
      "ID": internalId,
      "NIP": teksAmanSheet_(id),
      "Nama": teksAmanSheet_(name),
      "Jabatan": teksAmanSheet_(position),
      "Mata Pelajaran": teksAmanSheet_(subject),
      "Kategori": category,
      "Unit Kerja": teksAmanSheet_(unit),
      "Telepon": teksAmanSheet_(phone),
      "Status": status
    };
    headers.forEach((header, index) => {
      if (Object.prototype.hasOwnProperty.call(values, header)) output[index] = values[header];
      else if (targetRow >= 0) output[index] = rows[targetRow][index]; // pertahankan kolom tambahan yang tidak dikenal
    });

    if (targetRow >= 0) {
      sheet.getRange(targetRow + 2, 1, 1, headers.length).setValues([output]);
      return { success: true, recordId: internalId, action: "updated" };
    }

    sheet.appendRow(output);
    return { success: true, recordId: internalId, action: "created" };
  } finally {
    lock.releaseLock();
  }
}

/** Menghapus Guru/Tendik berdasarkan ID internal, bukan NIP atau nomor baris. */
function apiHapusGuruPegawai(recordId, sessionToken) {
  const actor = requireSession_(sessionToken);
  requirePermission_(actor, "teachers.manage");
  const id = String(recordId || "").trim();
  if (!id) throw new Error("ID internal personel tidak valid.");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const setup = siapkanSheetGuruTendik_();
    const sheet = setup.sheet;
    const idCol = setup.headers.indexOf("ID");
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("Belum ada data Guru/Tendik.");

    const ids = sheet.getRange(2, idCol + 1, lastRow - 1, 1).getDisplayValues();
    const index = ids.findIndex(row => String(row[0] || "").trim() === id);
    if (index === -1) throw new Error("Data personel tidak ditemukan.");

    sheet.deleteRow(index + 2);
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

/** Mengambil data semua modul; modul selain Data Siswa belum diaktifkan di frontend. */
function apiAmbilSemuaData(sessionToken) {
  requireAdmin_(sessionToken);
  const siswa = apiAmbilSiswa(sessionToken);
  const guru = apiAmbilGuruPegawai(sessionToken);
  const presensi = bacaTabel_("attendance");
  const ujian = bacaTabel_("exams");
  const nilai = bacaTabel_("grades");
  const keuangan = bacaTabel_("finance");
  const berita = bacaTabel_("news");

  return {
    students: siswa,
    teachers: guru,
    attendance: presensi.map(x => ({
      id: x.ID,
      date: tanggalTeks_(x.Tanggal),
      sid: String(x["ID Siswa"] || ""),
      name: x.Nama,
      kelas: x.Kelas,
      status: x.Status
    })),
    exams: ujian.map(x => ({
      id: x.ID,
      title: x["Judul Ujian"],
      subject: x["Mata Pelajaran"],
      kelas: x.Kelas,
      date: tanggalTeks_(x.Tanggal),
      questions: bacaSoal_(x.ID)
    })),
    grades: nilai.map(x => ({
      id: x.ID,
      date: tanggalTeks_(x.Tanggal),
      sid: String(x.NIS || ""),
      name: x.Nama,
      kelas: x.Kelas,
      subject: x["Mata Pelajaran"],
      score: x.Nilai
    })),
    finance: keuangan.map(x => ({
      id: x.ID,
      date: tanggalTeks_(x.Tanggal),
      type: x.Jenis,
      amount: x.Nominal,
      note: x.Keterangan
    })),
    news: berita.map(x => ({
      id: x.ID,
      date: tanggalTeks_(x.Tanggal),
      title: x.Judul,
      body: x.Isi
    }))
  };
}

function tanggalTeks_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );
  }
  return value ? String(value) : "";
}

function bacaSoal_(idUjian) {
  const sheet = getDatabase_().getSheetByName("Questions");
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  return values.slice(1)
    .filter(row => String(row[1]) === String(idUjian))
    .map(row => ({
      q: row[2],
      opts: [row[3], row[4], row[5], row[6], row[7]],
      a: ["A", "B", "C", "D", "E"].indexOf(String(row[8]).toUpperCase())
    }));
}

/**
 * Fungsi simpan generik untuk pengembangan modul berikutnya.
 * Untuk saat ini frontend hanya memakai apiSimpanSiswa().
 */
function apiSimpanData(kategori, data, sessionToken) {
  requireAdmin_(sessionToken);
  if (kategori === "students") return apiSimpanSiswa(data, sessionToken);
  if (kategori === "teachers") return apiSimpanGuruPegawai(data, sessionToken);

  const cfg = TABEL[kategori];
  if (!cfg) throw new Error("Kategori tidak dikenal.");

  const sheet = getDatabase_().getSheetByName(cfg.sheet);
  if (!sheet) throw new Error("Sheet tidak ditemukan: " + cfg.sheet);

  const id = buatId_();
  let row = [];

  if (kategori === "teachers") {
    row = [id, data.id, data.name, data.role, data.subject];
  } else if (kategori === "attendance") {
    row = [id, data.date, data.sid, data.name, data.kelas, data.status, data.latitude || "", data.longitude || ""];
  } else if (kategori === "exams") {
    row = [id, data.title, data.subject, data.kelas, data.date];
  } else if (kategori === "grades") {
    row = [id, data.date, data.sid, data.name, data.kelas, data.subject, data.score];
  } else if (kategori === "finance") {
    row = [id, data.date, data.type, data.amount, data.note];
  } else if (kategori === "news") {
    row = [id, data.date, data.title, data.body];
  } else {
    throw new Error("Kategori belum didukung.");
  }

  sheet.appendRow(row);

  if (kategori === "exams" && Array.isArray(data.questions)) {
    simpanSoal_(id, data.questions);
  }

  return { success: true, id: id };
}

function simpanSoal_(idUjian, questions) {
  const sheet = getDatabase_().getSheetByName("Questions");
  if (!sheet) throw new Error("Sheet Questions tidak ditemukan.");

  questions.forEach(q => {
    const opsi = q.opts || [];
    const kunci = ["A", "B", "C", "D", "E"][Number(q.a)] || "";
    sheet.appendRow([
      buatId_(), idUjian, q.q,
      opsi[0] || "", opsi[1] || "",
      opsi[2] || "", opsi[3] || "",
      opsi[4] || "", kunci
    ]);
  });
}

/** Tes koneksi database yang sudah dibuat sebelumnya. */
function testKoneksiDatabase_() {
  const db = getDatabase_();
  const hasil = {
    namaDatabase: db.getName(),
    jumlahSiswa: bacaTabel_("students").length,
    jumlahGuru: bacaTabel_("teachers").length,
    jumlahPresensi: bacaTabel_("attendance").length,
    status: "Koneksi berhasil"
  };
  Logger.log(JSON.stringify(hasil, null, 2));
  return hasil;
}

function diagnostikDatabaseAktif_() {
  const id = PropertiesService
    .getScriptProperties()
    .getProperty("DATABASE_ID");

  if (!id) {
    throw new Error("DATABASE_ID tidak ditemukan.");
  }

  const ss = SpreadsheetApp.openById(id);
  const sheet = ss.getSheetByName("Teachers");

  if (!sheet) {
    throw new Error("Sheet Teachers tidak ditemukan.");
  }

  const values = sheet.getDataRange().getDisplayValues();

  const jumlahData = values.slice(1).filter(row =>
    row.some(cell => String(cell).trim() !== "")
  ).length;

  const hasil = {
    namaDatabase: ss.getName(),
    idDatabase: id,
    namaSheet: sheet.getName(),
    jumlahKolom: sheet.getLastColumn(),
    jumlahBarisData: jumlahData,
    header: values[0]
  };

  Logger.log(JSON.stringify(hasil, null, 2));
  return hasil;
}

function tampilkanLinkDatabaseAktif_() {
  const id = PropertiesService
    .getScriptProperties()
    .getProperty("DATABASE_ID");

  if (!id) {
    throw new Error("DATABASE_ID tidak ditemukan.");
  }

  const db = SpreadsheetApp.openById(id);

  Logger.log("NAMA DATABASE: " + db.getName());
  Logger.log("LINK DATABASE AKTIF: " + db.getUrl());
}


/* ============================================================================
 * AUTENTIKASI DAN MANAJEMEN AKUN
 * - Empat role: ADMIN, GURU, TENDIK, SISWA.
 * - Password di-hash menggunakan PBKDF2-HMAC-SHA256 + pepper di ScriptProperties.
 * - Token sesi disimpan di CacheService dan divalidasi pada setiap API terlindungi.
 * - Tidak ada pendaftaran mandiri. Akun dibuat oleh Admin.
 * ========================================================================== */

const AUTH_SESSION_TTL_SECONDS = 21600; // 6 jam
const AUTH_HASH_ITERATIONS = 12000;

const AUTH_PERMISSION_KEYS = [
  "dashboard",
  "students.view", "students.manage",
  "teachers.view", "teachers.manage",
  "attendance.view", "attendance.test",
  "exams.view", "grades.view",
  "finance.view", "news.view", "news.manage",
  "settings.manage", "accounts.manage"
];

const AUTH_ROLE_DEFAULTS = {
  ADMIN: AUTH_PERMISSION_KEYS.slice(),
  GURU: [
    "dashboard", "students.view", "teachers.view",
    "attendance.view", "attendance.test"
  ],
  TENDIK: [
    "dashboard", "teachers.view", "attendance.view"
  ],
  SISWA: [
    "dashboard", "attendance.view", "attendance.test"
  ]
};

const AUTH_USER_HEADERS = [
  "ID", "Username", "Password", "Nama", "Role", "Status",
  "PasswordHash", "Salt", "ReferenceID", "MustChangePassword",
  "CreatedAt", "UpdatedAt", "LastLoginAt", "Permissions"
];

/**
 * Jalankan fungsi privat ini satu kali dari editor Apps Script untuk membuat
 * akun admin pertama. Password sementara akan ditampilkan di Execution log.
 * Pengguna wajib menggantinya pada login pertama.
 */
function buatAdminPertama_() {
  const setup = siapkanSheetUsers_();
  const sheet = setup.sheet;
  const headers = setup.headers;
  const lastRow = sheet.getLastRow();
  const existing = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, headers.length).getValues()
        .filter(row => String(row[headers.indexOf("Username")] || "").trim() !== "")
    : [];

  if (existing.length) {
    throw new Error("Sheet Users sudah berisi akun. Pembuatan admin pertama dibatalkan agar akun yang ada tidak tertimpa.");
  }

  const temporaryPassword = buatPasswordSementara_();
  const salt = buatSalt_();
  const now = new Date();
  const values = {
    ID: buatId_(),
    Username: "admin",
    Password: "",
    Nama: "Administrator MAN 1 Kota Malang",
    Role: "ADMIN",
    Status: "AKTIF",
    PasswordHash: hashPassword_(temporaryPassword, salt),
    Salt: salt,
    ReferenceID: "",
    MustChangePassword: true,
    CreatedAt: now,
    UpdatedAt: now,
    LastLoginAt: "",
    Permissions: JSON.stringify(AUTH_ROLE_DEFAULTS.ADMIN)
  };

  const row = headers.map(header =>
    Object.prototype.hasOwnProperty.call(values, header) ? values[header] : ""
  );
  sheet.appendRow(row);

  Logger.log("AKUN ADMIN PERTAMA BERHASIL DIBUAT.");
  Logger.log("Username: admin");
  Logger.log("Password sementara: " + temporaryPassword);
  Logger.log("Segera login dan ganti password. Jangan bagikan log ini.");
  catatAudit_("SYSTEM", "Membuat akun admin pertama");
  return { success: true, username: "admin", temporaryPassword: temporaryPassword };
}

/** Endpoint login publik. Tidak mengembalikan hash atau salt kepada browser. */
function apiLogin(username, password) {
  const uname = normalisasiUsername_(username);
  const pass = String(password || "");
  if (!uname || !pass) throw new Error("Username dan password wajib diisi.");

  const failKey = "auth_fail_" + hashTeks_(uname).substring(0, 32);
  const cache = CacheService.getScriptCache();
  const currentFailures = Number(cache.get(failKey) || 0);
  if (currentFailures >= 5) {
    throw new Error("Terlalu banyak percobaan login. Coba kembali setelah 15 menit.");
  }

  const setup = siapkanSheetUsers_();
  const sheet = setup.sheet;
  const headers = setup.headers;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error("Akun belum disiapkan. Jalankan fungsi buatAdminPertama_ dari editor Apps Script terlebih dahulu.");
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const usernameCol = headers.indexOf("Username");
  const rowIndex = rows.findIndex(row =>
    normalisasiUsername_(row[usernameCol]) === uname
  );

  if (rowIndex < 0) {
    // Perhitungan dummy membantu mengurangi perbedaan waktu untuk username yang tidak ditemukan.
    hashPassword_(pass, "0".repeat(64));
    cache.put(failKey, String(currentFailures + 1), 900);
    throw new Error("Username atau password salah.");
  }

  const row = rows[rowIndex];
  const account = objekAkunDariBaris_(row, headers);
  if (account.Status !== "AKTIF") {
    cache.put(failKey, String(currentFailures + 1), 900);
    throw new Error("Username atau password salah, atau akun tidak aktif.");
  }

  if (!account.PasswordHash || !account.Salt) {
    cache.put(failKey, String(currentFailures + 1), 900);
    throw new Error("Akun belum memiliki password yang valid. Hubungi Admin.");
  }
  const calculated = hashPassword_(pass, String(account.Salt || ""));
  if (!bandingKonstan_(calculated, String(account.PasswordHash || ""))) {
    cache.put(failKey, String(currentFailures + 1), 900);
    throw new Error("Username atau password salah.");
  }

  cache.remove(failKey);

  const token = Utilities.getUuid().replace(/-/g, "") +
    Utilities.getUuid().replace(/-/g, "");
  const session = { userId: String(account.ID), issuedAt: new Date().toISOString() };
  cache.put(kunciSesi_(token), JSON.stringify(session), AUTH_SESSION_TTL_SECONDS);

  const lastLoginCol = headers.indexOf("LastLoginAt");
  const updatedRow = row.slice();
  if (lastLoginCol >= 0) updatedRow[lastLoginCol] = new Date();
  sheet.getRange(rowIndex + 2, 1, 1, headers.length).setValues([updatedRow]);

  catatAudit_(account.Username, "Login berhasil");
  return {
    success: true,
    sessionToken: token,
    user: profilPublikAkun_(account)
  };
}

/** Memulihkan profil sesi saat halaman dimuat ulang. */
function apiSessionMe(sessionToken) {
  const actor = requireSession_(sessionToken, true);
  return profilPublikAkun_(actor);
}

/** Mengakhiri sesi pada browser yang sedang digunakan. */
function apiLogout(sessionToken) {
  const token = String(sessionToken || "");
  if (token) CacheService.getScriptCache().remove(kunciSesi_(token));
  return { success: true };
}

/** Pengguna mengganti password sendiri; diwajibkan pada login pertama/reset. */
function apiGantiPassword(passwordLama, passwordBaru, sessionToken) {
  const actor = requireSession_(sessionToken, true);
  const oldPass = String(passwordLama || "");
  const newPass = String(passwordBaru || "");

  if (newPass.length < 10 || !/[A-Za-z]/.test(newPass) || !/[0-9]/.test(newPass)) {
    throw new Error("Password baru minimal 10 karakter dan harus mengandung huruf serta angka.");
  }
  if (newPass === oldPass) throw new Error("Password baru harus berbeda dari password lama.");

  const found = cariAkunById_(actor.id);
  if (!found) throw new Error("Akun tidak ditemukan.");
  const oldHash = hashPassword_(oldPass, String(found.account.Salt || ""));
  if (!bandingKonstan_(oldHash, String(found.account.PasswordHash || ""))) {
    throw new Error("Password lama tidak sesuai.");
  }

  const salt = buatSalt_();
  const row = found.row.slice();
  setKolom_(row, found.headers, "Password", "");
  setKolom_(row, found.headers, "PasswordHash", hashPassword_(newPass, salt));
  setKolom_(row, found.headers, "Salt", salt);
  setKolom_(row, found.headers, "MustChangePassword", false);
  setKolom_(row, found.headers, "UpdatedAt", new Date());
  found.sheet.getRange(found.rowNumber, 1, 1, found.headers.length).setValues([row]);

  catatAudit_(actor.username, "Mengganti password");
  const refreshed = cariAkunById_(actor.id);
  return { success: true, user: profilPublikAkun_(refreshed.account) };
}

/** Daftar akun hanya dapat dibaca oleh Admin. */
function apiDaftarAkun(sessionToken) {
  requireAdmin_(sessionToken);
  const setup = siapkanSheetUsers_();
  const sheet = setup.sheet;
  const headers = setup.headers;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, headers.length).getValues()
    .filter(row => String(row[headers.indexOf("Username")] || "").trim() !== "")
    .map(row => profilPublikAkun_(objekAkunDariBaris_(row, headers)));
}

/** Admin membuat akun baru; password sementara dibuat otomatis oleh server. */
function apiBuatAkun(data, sessionToken) {
  const actor = requireAdmin_(sessionToken);
  if (!data) throw new Error("Data akun tidak diterima.");

  const role = normalisasiRole_(data.role);
  const username = normalisasiUsername_(data.username);
  if (!/^[a-z0-9._-]{3,50}$/.test(username)) {
    throw new Error("Username harus 3–50 karakter dan hanya boleh berisi huruf kecil, angka, titik, garis bawah, atau tanda hubung.");
  }

  const referenceId = String(data.referenceId || "").trim();
  const identity = resolveIdentitasAkun_(role, referenceId, data.name);
  const setup = siapkanSheetUsers_();
  const sheet = setup.sheet;
  const headers = setup.headers;
  const lastRow = sheet.getLastRow();
  const rows = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, headers.length).getValues()
    : [];

  const usernameCol = headers.indexOf("Username");
  const roleCol = headers.indexOf("Role");
  const refCol = headers.indexOf("ReferenceID");
  if (rows.some(row => normalisasiUsername_(row[usernameCol]) === username)) {
    throw new Error("Username sudah digunakan.");
  }
  if (role !== "ADMIN" && rows.some(row =>
    String(row[roleCol] || "") === role &&
    String(row[refCol] || "") === referenceId
  )) {
    throw new Error("Identitas tersebut sudah memiliki akun dengan kategori yang sama.");
  }

  const permissions = sanitasiPermissions_(role, data.permissions);
  const temporaryPassword = buatPasswordSementara_();
  const salt = buatSalt_();
  const now = new Date();
  const values = {
    ID: buatId_(),
    Username: username,
    Password: "",
    Nama: identity.name,
    Role: role,
    Status: "AKTIF",
    PasswordHash: hashPassword_(temporaryPassword, salt),
    Salt: salt,
    ReferenceID: identity.referenceId,
    MustChangePassword: true,
    CreatedAt: now,
    UpdatedAt: now,
    LastLoginAt: "",
    Permissions: JSON.stringify(permissions)
  };

  sheet.appendRow(headers.map(header =>
    Object.prototype.hasOwnProperty.call(values, header) ? values[header] : ""
  ));

  catatAudit_(actor.username, "Membuat akun " + role + " untuk " + identity.name);
  return {
    success: true,
    account: profilPublikAkun_(values),
    temporaryPassword: temporaryPassword
  };
}

/** Admin mengubah identitas login, role, status, dan hak akses akun. */
function apiUbahAkun(data, sessionToken) {
  const actor = requireAdmin_(sessionToken);
  if (!data || !data.id) throw new Error("ID akun tidak valid.");

  const found = cariAkunById_(String(data.id));
  if (!found) throw new Error("Akun tidak ditemukan.");

  const oldAccount = found.account;
  const role = normalisasiRole_(data.role);
  const username = normalisasiUsername_(data.username);
  const status = String(data.status || "").toUpperCase();
  if (!/^[a-z0-9._-]{3,50}$/.test(username)) throw new Error("Format username tidak valid.");
  if (["AKTIF", "NONAKTIF"].indexOf(status) === -1) throw new Error("Status akun tidak valid.");

  if (actor.id === oldAccount.ID && (role !== "ADMIN" || status !== "AKTIF")) {
    throw new Error("Admin yang sedang login tidak boleh menonaktifkan atau menurunkan role akunnya sendiri.");
  }

  if (oldAccount.Role === "ADMIN" && oldAccount.Status === "AKTIF" &&
      (role !== "ADMIN" || status !== "AKTIF") && hitungAdminAktif_() <= 1) {
    throw new Error("Tidak dapat menonaktifkan Admin terakhir yang masih aktif.");
  }

  const usernameCol = found.headers.indexOf("Username");
  const rows = found.sheet.getLastRow() > 1
    ? found.sheet.getRange(2, 1, found.sheet.getLastRow() - 1, found.headers.length).getValues()
    : [];
  if (rows.some((row, i) =>
    i + 2 !== found.rowNumber &&
    normalisasiUsername_(row[usernameCol]) === username
  )) {
    throw new Error("Username sudah digunakan akun lain.");
  }

  const referenceId = String(data.referenceId || "").trim();
  const identity = resolveIdentitasAkun_(role, referenceId, data.name);

  const roleCol = found.headers.indexOf("Role");
  const refCol = found.headers.indexOf("ReferenceID");
  if (role !== "ADMIN" && rows.some((row, i) =>
    i + 2 !== found.rowNumber &&
    String(row[roleCol] || "") === role &&
    String(row[refCol] || "") === identity.referenceId
  )) {
    throw new Error("Identitas tersebut sudah memiliki akun dengan kategori yang sama.");
  }

  const row = found.row.slice();
  setKolom_(row, found.headers, "Username", username);
  setKolom_(row, found.headers, "Nama", identity.name);
  setKolom_(row, found.headers, "Role", role);
  setKolom_(row, found.headers, "Status", status);
  setKolom_(row, found.headers, "ReferenceID", identity.referenceId);
  setKolom_(row, found.headers, "Permissions", JSON.stringify(sanitasiPermissions_(role, data.permissions)));
  setKolom_(row, found.headers, "UpdatedAt", new Date());
  found.sheet.getRange(found.rowNumber, 1, 1, found.headers.length).setValues([row]);

  catatAudit_(actor.username, "Memperbarui akun " + username);
  return { success: true };
}

/** Admin mereset password; password sementara hanya dikembalikan satu kali. */
function apiResetPasswordAkun(userId, sessionToken) {
  const actor = requireAdmin_(sessionToken);
  const found = cariAkunById_(String(userId || ""));
  if (!found) throw new Error("Akun tidak ditemukan.");

  const temporaryPassword = buatPasswordSementara_();
  const salt = buatSalt_();
  const row = found.row.slice();
  setKolom_(row, found.headers, "Password", "");
  setKolom_(row, found.headers, "PasswordHash", hashPassword_(temporaryPassword, salt));
  setKolom_(row, found.headers, "Salt", salt);
  setKolom_(row, found.headers, "MustChangePassword", true);
  setKolom_(row, found.headers, "UpdatedAt", new Date());
  found.sheet.getRange(found.rowNumber, 1, 1, found.headers.length).setValues([row]);

  catatAudit_(actor.username, "Reset password akun " + found.account.Username);
  return { success: true, temporaryPassword: temporaryPassword };
}

/* --------------------------- helper autentikasi --------------------------- */

function siapkanSheetUsers_() {
  const db = getDatabase_();
  let sheet = db.getSheetByName("Users");
  if (!sheet) sheet = db.insertSheet("Users");

  const lastCol = sheet.getLastColumn();
  let headers = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0]
        .map(value => String(value || "").trim())
        .filter((value, index, all) => value || index < all.length)
    : [];
  if (!headers.length || headers.every(value => !value)) headers = [];

  AUTH_USER_HEADERS.forEach(name => {
    if (headers.indexOf(name) === -1) headers.push(name);
  });

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#123653")
    .setFontColor("#ffffff");
  sheet.setFrozenRows(1);
  getAuthPepper_();

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const range = sheet.getRange(2, 1, lastRow - 1, headers.length);
    const rows = range.getValues();
    let changed = false;
    const idx = name => headers.indexOf(name);

    rows.forEach(row => {
      const username = String(row[idx("Username")] || "").trim();
      if (!username) return;

      const legacyPassword = String(row[idx("Password")] || "");
      const existingHash = String(row[idx("PasswordHash")] || "");
      if (legacyPassword && !existingHash) {
        const salt = buatSalt_();
        row[idx("Salt")] = salt;
        row[idx("PasswordHash")] = hashPassword_(legacyPassword, salt);
        row[idx("Password")] = "";
        row[idx("MustChangePassword")] = true;
        row[idx("UpdatedAt")] = new Date();
        changed = true;
      }
      if (!String(row[idx("Status")] || "").trim()) {
        row[idx("Status")] = "AKTIF";
        changed = true;
      }
      const rawRoleValue = String(row[idx("Role")] || "").trim();
      if (!rawRoleValue) {
        row[idx("Role")] = "SISWA";
        changed = true;
      } else {
        try {
          const normalizedRole = normalisasiRole_(rawRoleValue);
          if (normalizedRole !== rawRoleValue) {
            row[idx("Role")] = normalizedRole;
            changed = true;
          }
        } catch (roleError) {
          row[idx("Role")] = "SISWA";
          changed = true;
        }
      }
      if (!String(row[idx("Permissions")] || "").trim()) {
        const rawRole = String(row[idx("Role")] || "").toUpperCase();
        const role = ["ADMIN", "GURU", "TENDIK", "SISWA"].indexOf(rawRole) >= 0
          ? rawRole : "SISWA";
        row[idx("Permissions")] = JSON.stringify(AUTH_ROLE_DEFAULTS[role] || []);
        changed = true;
      }
    });

    if (changed) range.setValues(rows);
  }

  return { sheet: sheet, headers: headers };
}

function getAuthPepper_() {
  const props = PropertiesService.getScriptProperties();
  let pepper = props.getProperty("DIGITAL_SCHOOL_AUTH_PEPPER");
  if (!pepper) {
    pepper = Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid();
    props.setProperty("DIGITAL_SCHOOL_AUTH_PEPPER", pepper);
  }
  return pepper;
}

function buatSalt_() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, "");
}

function buatPasswordSementara_() {
  const raw = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, "");
  return raw.substring(0, 6) + "A9!" + raw.substring(6, 16);
}

function hashPassword_(password, saltHex) {
  const passwordBytes = Utilities.newBlob(String(password) + "\u0000" + getAuthPepper_()).getBytes();
  const saltBytes = hexKeBytes_(String(saltHex || ""));
  const block = saltBytes.concat([0, 0, 0, 1]);
  let u = Utilities.computeHmacSha256Signature(block, passwordBytes);
  let result = u.slice();

  for (let i = 1; i < AUTH_HASH_ITERATIONS; i++) {
    u = Utilities.computeHmacSha256Signature(u, passwordBytes);
    for (let j = 0; j < result.length; j++) {
      result[j] = result[j] ^ u[j];
    }
  }
  return bytesKeHex_(result);
}

function hexKeBytes_(hex) {
  if (!/^[0-9a-f]+$/i.test(hex) || hex.length % 2 !== 0) {
    throw new Error("Salt akun tidak valid.");
  }
  const out = [];
  for (let i = 0; i < hex.length; i += 2) {
    const value = parseInt(hex.substr(i, 2), 16);
    out.push(value > 127 ? value - 256 : value);
  }
  return out;
}

function bytesKeHex_(bytes) {
  return bytes.map(value => {
    const unsigned = value & 255;
    return ("0" + unsigned.toString(16)).slice(-2);
  }).join("");
}

function hashTeks_(value) {
  return bytesKeHex_(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value),
    Utilities.Charset.UTF_8
  ));
}

function bandingKonstan_(a, b) {
  a = String(a || "");
  b = String(b || "");
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function normalisasiUsername_(value) {
  return String(value || "").trim().toLowerCase();
}

function normalisasiRole_(value) {
  const role = String(value || "").trim().toUpperCase();
  if (role === "ADMIN" || role === "ADMINISTRATOR") return "ADMIN";
  if (role === "GURU") return "GURU";
  if (role === "TENDIK" || role === "TENAGA KEPENDIDIKAN") return "TENDIK";
  if (role === "SISWA") return "SISWA";
  throw new Error("Kategori akun harus Admin, Guru, Tenaga Kependidikan, atau Siswa.");
}

function sanitasiPermissions_(role, permissions) {
  if (role === "ADMIN") return AUTH_ROLE_DEFAULTS.ADMIN.slice();
  const requested = Array.isArray(permissions) ? permissions : AUTH_ROLE_DEFAULTS[role];
  return requested.filter((key, index, all) =>
    AUTH_PERMISSION_KEYS.indexOf(key) >= 0 &&
    key !== "accounts.manage" &&
    all.indexOf(key) === index
  );
}

function permissionsAkun_(account) {
  if (account.Role === "ADMIN") return AUTH_ROLE_DEFAULTS.ADMIN.slice();
  try {
    const parsed = JSON.parse(String(account.Permissions || ""));
    if (Array.isArray(parsed)) return sanitasiPermissions_(account.Role, parsed);
  } catch (err) {}
  return (AUTH_ROLE_DEFAULTS[account.Role] || []).slice();
}

function objekAkunDariBaris_(row, headers) {
  const obj = {};
  headers.forEach((header, index) => obj[header] = row[index]);
  obj.ID = String(obj.ID || "");
  obj.Username = normalisasiUsername_(obj.Username);
  obj.Nama = String(obj.Nama || "");
  obj.Role = normalisasiRole_(obj.Role);
  obj.Status = String(obj.Status || "AKTIF").toUpperCase();
  obj.ReferenceID = String(obj.ReferenceID || "");
  obj.MustChangePassword = obj.MustChangePassword === true ||
    String(obj.MustChangePassword).toLowerCase() === "true";
  obj.Permissions = permissionsAkun_(obj);
  return obj;
}

function profilPublikAkun_(account) {
  const permissions = Array.isArray(account.Permissions)
    ? account.Permissions
    : (Array.isArray(account.permissions) ? account.permissions : permissionsAkun_(account));
  return {
    id: String(account.ID || account.id || ""),
    username: String(account.Username || account.username || ""),
    name: String(account.Nama || account.name || ""),
    role: String(account.Role || account.role || ""),
    referenceId: String(account.ReferenceID || account.referenceId || ""),
    status: String(account.Status || account.status || ""),
    mustChangePassword: !!(account.MustChangePassword || account.mustChangePassword),
    permissions: permissions,
    createdAt: account.CreatedAt instanceof Date
  ? account.CreatedAt.toISOString()
  : String(account.CreatedAt || account.createdAt || ""),
lastLoginAt: account.LastLoginAt instanceof Date
  ? account.LastLoginAt.toISOString()
  : String(account.LastLoginAt || account.lastLoginAt || "")
  };
}

function kunciSesi_(token) {
  return "auth_session_" + hashTeks_(token);
}

function requireSession_(sessionToken, allowMustChangePassword) {
  const token = String(sessionToken || "");
  if (!token) throw new Error("Sesi tidak ditemukan. Silakan login kembali.");

  const cache = CacheService.getScriptCache();
  const raw = cache.get(kunciSesi_(token));
  if (!raw) throw new Error("Sesi login telah berakhir. Silakan login kembali.");

  let session;
  try { session = JSON.parse(raw); }
  catch (err) { throw new Error("Sesi tidak valid. Silakan login kembali."); }

  const found = cariAkunById_(String(session.userId || ""));
  if (!found) throw new Error("Akun tidak ditemukan. Silakan login kembali.");
  const account = found.account;
  if (account.Status !== "AKTIF") throw new Error("Akun tidak aktif. Hubungi Admin.");
  if (account.MustChangePassword && !allowMustChangePassword) {
    throw new Error("Anda wajib mengganti password sementara sebelum menggunakan aplikasi.");
  }

  return {
    id: account.ID,
    username: account.Username,
    name: account.Nama,
    role: account.Role,
    referenceId: account.ReferenceID,
    status: account.Status,
    mustChangePassword: account.MustChangePassword,
    permissions: account.Permissions
  };
}

function requireAdmin_(sessionToken) {
  const actor = requireSession_(sessionToken);
  if (actor.role !== "ADMIN") throw new Error("Akses ditolak. Fitur ini hanya untuk Admin.");
  return actor;
}

function requirePermission_(actor, permission) {
  if (!actor || (actor.role !== "ADMIN" &&
      (!Array.isArray(actor.permissions) || actor.permissions.indexOf(permission) === -1))) {
    throw new Error("Akses ditolak. Anda tidak memiliki hak untuk melakukan tindakan ini.");
  }
}

function cariAkunById_(id) {
  const setup = siapkanSheetUsers_();
  const sheet = setup.sheet;
  const headers = setup.headers;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const rows = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const idCol = headers.indexOf("ID");
  const index = rows.findIndex(row => String(row[idCol] || "") === String(id));
  if (index < 0) return null;

  return {
    sheet: sheet,
    headers: headers,
    rowNumber: index + 2,
    row: rows[index],
    account: objekAkunDariBaris_(rows[index], headers)
  };
}

function setKolom_(row, headers, name, value) {
  const index = headers.indexOf(name);
  if (index >= 0) row[index] = value;
}

function hitungAdminAktif_() {
  const setup = siapkanSheetUsers_();
  const headers = setup.headers;
  const lastRow = setup.sheet.getLastRow();
  if (lastRow < 2) return 0;
  const rows = setup.sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const roleCol = headers.indexOf("Role");
  const statusCol = headers.indexOf("Status");
  const usernameCol = headers.indexOf("Username");
  return rows.filter(row =>
    normalisasiUsername_(row[usernameCol]) &&
    String(row[roleCol] || "").toUpperCase() === "ADMIN" &&
    String(row[statusCol] || "").toUpperCase() === "AKTIF"
  ).length;
}

function resolveIdentitasAkun_(role, referenceId, suppliedName) {
  if (role === "ADMIN") {
    const name = String(suppliedName || "").trim();
    if (!name) throw new Error("Nama Admin wajib diisi.");
    return { name: name, referenceId: "" };
  }

  if (!referenceId) throw new Error("NIP/ID atau NIS wajib diisi untuk akun Guru, Tendik, dan Siswa.");

  if (role === "SISWA") {
    const sheet = getDatabase_().getSheetByName("Students");
    if (!sheet || sheet.getLastRow() < 2) throw new Error("Data siswa belum tersedia.");
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(6, sheet.getLastColumn())).getDisplayValues();
    const found = rows.find(row => String(row[1] || "").trim() === referenceId);
    if (!found) throw new Error("NIS tidak ditemukan pada sheet Students.");
    return { name: String(found[2] || "").trim(), referenceId: referenceId };
  }

  const setup = siapkanSheetGuruTendik_();
  const headers = setup.headers;
  const rows = setup.sheet.getLastRow() > 1
    ? setup.sheet.getRange(2, 1, setup.sheet.getLastRow() - 1, headers.length).getValues()
    : [];
  const nipCol = headers.indexOf("NIP");
  const found = rows.find(row => String(row[nipCol] || "").trim() === referenceId);
  if (!found) throw new Error("NIP/ID tidak ditemukan pada sheet Teachers.");

  const position = String(nilaiKolomGuru_(found, headers, "Jabatan") || "").toLowerCase();
  const subject = String(nilaiKolomGuru_(found, headers, "Mata Pelajaran") || "").trim();
  let category = String(nilaiKolomGuru_(found, headers, "Kategori") || "").trim();
  if (!category) {
    if (/guru|kepala madrasah|wakil kepala|waka/.test(position) || subject) category = "Guru";
    else if (/pegawai|tenaga kependidikan|tata usaha|operator|pustakawan|laboran|satpam|pramubakti/.test(position)) category = "Tenaga Kependidikan";
  }
  const expected = role === "GURU" ? "Guru" : "Tenaga Kependidikan";
  if (category !== expected) throw new Error("Kategori NIP/ID tidak sesuai dengan role akun yang dipilih.");
  const status = String(nilaiKolomGuru_(found, headers, "Status") || "Aktif").trim();
  if (status.toLowerCase() === "nonaktif") throw new Error("Personel berstatus Nonaktif tidak dapat diberi akun aktif.");
  return {
    name: String(nilaiKolomGuru_(found, headers, "Nama") || "").trim(),
    referenceId: referenceId
  };
}

function catatAudit_(user, activity) {
  try {
    const sheet = getDatabase_().getSheetByName("AuditLog");
    if (!sheet) return;
    sheet.appendRow([buatId_(), new Date(), String(user || "SYSTEM"), String(activity || "")]);
  } catch (err) {
    // Audit tidak boleh menggagalkan operasi utama.
  }
}

