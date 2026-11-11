/**
 * Bansos Distribution Reminder: Google Sheets -> Email (prototype)
 * Author: Ageng Windarto Jati
 *
 * Bound Apps Script for a Google Sheet with three tabs:
 *   SEMBAKO           : distribution timeline (data starts at row 4)
 *   REMINDER_CONFIG   : one row per notification rule
 *   NOTIFICATION_LOG  : every send attempt (also used to prevent duplicates)
 *
 * Main entry points:
 *   pasangTriggerEdit()           installs the on-edit trigger -> handleSheetEdit()
 *   pasangTriggerPengingatBank()  installs a daily 10:00 trigger -> prosesPengingatBank()
 *   validasiKonfigurasi()         checks REMINDER_CONFIG before going live
 *   simulasiPengingatBank()       dry run: lists H-1/H candidates without sending
 */

// Email address used by the test functions (ujiPencatatan, ujiKirimEmail,
// ujiKirimPengingatBankSatuBaris). Replace it with your own before testing.
const TEST_EMAIL = "your-test-email@example.com";

function cekStrukturSpreadsheet() {
  // Mengambil spreadsheet yang terhubung dengan Apps Script
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Mengambil sheet SEMBAKO
  const sheet = ss.getSheetByName("SEMBAKO");

  // Memastikan sheet tersedia
  if (!sheet) {
    throw new Error("Sheet SEMBAKO tidak ditemukan!");
  }

  // Mengambil daftar nama seluruh sheet
  const daftarSheet = ss.getSheets()
    .map(sheet => sheet.getName());

  // Menampilkan informasi di execution log
  Logger.log("Nama Spreadsheet: " + ss.getName());
  Logger.log("Daftar Sheet: " + daftarSheet.join(", "));
  Logger.log("Jumlah Baris SEMBAKO: " + sheet.getLastRow());
  Logger.log("Jumlah Kolom SEMBAKO: " + sheet.getLastColumn());
}

function bacaKonfigurasi() {
  // Mengambil spreadsheet aktif
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Mengambil sheet konfigurasi
  const configSheet = ss.getSheetByName("REMINDER_CONFIG");

  // Memastikan sheet tersedia
  if (!configSheet) {
    throw new Error("Sheet REMINDER_CONFIG tidak ditemukan!");
  }

  // Mengambil seluruh isi konfigurasi
  const data = configSheet.getDataRange().getValues();

  // Memastikan terdapat header
  if (data.length < 2) {
    throw new Error("REMINDER_CONFIG belum memiliki data konfigurasi!");
  }

  // Mengambil header pada baris pertama
  const headers = data[0];

  // Mengubah setiap baris menjadi objek konfigurasi
  const konfigurasi = data.slice(1)
    .filter(row => row.some(value => value !== ""))
    .map(row => {
      const item = {};

      headers.forEach((header, index) => {
        item[header] = row[index];
      });

      return item;
    });

  // Menampilkan konfigurasi yang berhasil dibaca
  Logger.log(JSON.stringify(konfigurasi, null, 2));

  return konfigurasi;
}

function validasiKonfigurasi() {
  const konfigurasi = bacaKonfigurasi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let jumlahValid = 0;
  let jumlahTidakValid = 0;

  konfigurasi.forEach((config, index) => {
    const nomorBaris = index + 2;
    const kesalahan = [];

    // Memeriksa apakah konfigurasi aktif
    if (config.ACTIVE !== true) {
      Logger.log(`Baris ${nomorBaris}: Konfigurasi tidak aktif.`);
      return;
    }

    // Memeriksa keberadaan sheet
    if (!ss.getSheetByName(config.SHEET_NAME)) {
      kesalahan.push("Sheet tidak ditemukan");
    }

    // Memeriksa alamat email
    if (!config.RECIPIENT_EMAIL ||
        !String(config.RECIPIENT_EMAIL).includes("@")) {
      kesalahan.push("Email penerima tidak valid");
    }

    // Memeriksa aturan berdasarkan jenis trigger
    if (config.TRIGGER_TYPE === "DATE") {
      if (!config.DATE_COLUMN) {
        kesalahan.push("DATE_COLUMN belum diisi");
      }

      if (!config.REMINDER_DAYS) {
        kesalahan.push("REMINDER_DAYS belum diisi");
      }

    } else if (config.TRIGGER_TYPE === "STATUS_CHANGE") {
      if (!config.STATUS_COLUMN) {
        kesalahan.push("STATUS_COLUMN belum diisi");
      }

      if (!config.STATUS_VALUE) {
        kesalahan.push("STATUS_VALUE belum diisi");
      }

    } else {
      kesalahan.push("Jenis trigger tidak dikenal");
    }

    // Menampilkan hasil validasi
    if (kesalahan.length === 0) {
      jumlahValid++;
      Logger.log(`Baris ${nomorBaris}: VALID`);
    } else {
      jumlahTidakValid++;
      Logger.log(
        `Baris ${nomorBaris}: TIDAK VALID - ${kesalahan.join(", ")}`
      );
    }
  });

  Logger.log("--------------------------");
  Logger.log("Jumlah konfigurasi valid: " + jumlahValid);
  Logger.log("Jumlah konfigurasi tidak valid: " + jumlahTidakValid);
}

function sudahPernahDikirim(dedupeKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("NOTIFICATION_LOG");

  if (!logSheet) {
    throw new Error("Sheet NOTIFICATION_LOG tidak ditemukan!");
  }

  const data = logSheet.getDataRange().getValues();

  // Jika belum ada catatan, berarti belum pernah dikirim
  if (data.length < 2) {
    return false;
  }

  const headers = data[0];

  const indexStatus = headers.indexOf("SEND_STATUS");
  const indexKey = headers.indexOf("DEDUPE_KEY");

  if (indexStatus === -1 || indexKey === -1) {
    throw new Error("Kolom SEND_STATUS atau DEDUPE_KEY tidak ditemukan!");
  }

  // Periksa seluruh riwayat pengiriman
  for (let i = 1; i < data.length; i++) {
    const status = data[i][indexStatus];
    const key = data[i][indexKey];

    if (key === dedupeKey && status === "SENT") {
      return true;
    }
  }

  return false;
}

function catatNotifikasi(logData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("NOTIFICATION_LOG");

  if (!logSheet) {
    throw new Error("Sheet NOTIFICATION_LOG tidak ditemukan!");
  }

  // Membaca nama kolom pada baris pertama
  const headers = logSheet
    .getRange(1, 1, 1, logSheet.getLastColumn())
    .getValues()[0];

  // Menyusun data berdasarkan nama header
  const barisBaru = headers.map(header => {
    return logData[header] !== undefined
      ? logData[header]
      : "";
  });

  // Menambahkan catatan pada baris terakhir
  logSheet.appendRow(barisBaru);
}

function ujiPencatatan() {
  const dedupeKey = "TEST_SEMBAKO_BANK_001";

  // Memeriksa apakah catatan pengujian sudah ada
  const sudahAda = sudahPernahDikirim(dedupeKey);

  Logger.log("Sudah pernah dikirim: " + sudahAda);

  // Jika belum ada, buat catatan simulasi
  if (!sudahAda) {
    catatNotifikasi({
      TIMESTAMP: new Date(),
      RUN_ID: "TEST_001",
      SHEET_NAME: "SEMBAKO",
      SOURCE_ROW: 4,
      REMINDER_ID: "BANK_FREEZE",
      TRIGGER_TYPE: "DATE",
      ACTIVITY_NAME: "Akhir Pengembalian",
      RECIPIENT_ROLE: "BANK",
      RECIPIENT_EMAIL: TEST_EMAIL,
      SEND_STATUS: "SENT",
      DEDUPE_KEY: dedupeKey
    });

    Logger.log("Catatan simulasi berhasil ditambahkan.");
  } else {
    Logger.log("Catatan sudah ada. Tidak menambahkan duplikat.");
  }
}


function kirimEmailNotifikasi(emailData) {
  const lock = LockService.getScriptLock();

  // Menunggu maksimal 30 detik untuk mendapatkan kunci
  if (!lock.tryLock(30000)) {
    throw new Error(
      "Sistem sedang memproses notifikasi lain. Silakan coba kembali."
    );
  }

  try {
    const dedupeKey = emailData.DEDUPE_KEY;

    if (!dedupeKey) {
      throw new Error("DEDUPE_KEY belum tersedia.");
    }

    // Memeriksa apakah notifikasi sudah pernah sukses
    if (sudahPernahDikirim(dedupeKey)) {
      Logger.log("Notifikasi sudah pernah dikirim: " + dedupeKey);
      return "DUPLICATE";
    }

    let status = "FAILED";
    let pesanError = "";

    try {
      // Mengirim email
      MailApp.sendEmail({
        to: emailData.RECIPIENT_EMAIL,
        subject: emailData.SUBJECT,
        body: emailData.BODY
      });

      status = "SENT";

      Logger.log("Email berhasil diserahkan ke layanan pengiriman.");

    } catch (error) {
      pesanError = error.message;
      Logger.log("Pengiriman email gagal: " + pesanError);
    }

    // Mencatat hasil pengiriman beserta identitas data bantuan
    catatNotifikasi({
      TIMESTAMP: new Date(),
      RUN_ID: Utilities.getUuid(),
      SHEET_NAME: emailData.SHEET_NAME,
      SOURCE_ROW: emailData.SOURCE_ROW,

      // Identitas data bantuan dari kolom A:D
      PROGRAM: emailData.PROGRAM || "",
      TAHAP: emailData.TAHAP || "",
      BATCH: emailData.BATCH || "",
      BANK: emailData.BANK || "",

      REMINDER_ID: emailData.REMINDER_ID,
      TRIGGER_TYPE: emailData.TRIGGER_TYPE,
      ACTIVITY_NAME: emailData.ACTIVITY_NAME,
      DEADLINE_DATE: emailData.DEADLINE_DATE || "",
      REMINDER_LABEL: emailData.REMINDER_LABEL || "",
      RECIPIENT_ROLE: emailData.RECIPIENT_ROLE,
      RECIPIENT_EMAIL: emailData.RECIPIENT_EMAIL,
      STATUS_BANTUAN: emailData.STATUS_BANTUAN || "",
      STATUS_INPUT: emailData.STATUS_INPUT || "",
      SEND_STATUS: status,
      DEDUPE_KEY: dedupeKey,
      ERROR_MESSAGE: pesanError
    });

    return status;

  } finally {
    // Melepaskan kunci setelah proses selesai
    lock.releaseLock();
  }
}

function ujiKirimEmail() {
  const hasil = kirimEmailNotifikasi({
    SHEET_NAME: "SEMBAKO",
    SOURCE_ROW: 4,

    // Identitas data bantuan untuk pengujian
    PROGRAM: "SEMBAKO",
    TAHAP: "JAN-MAR",
    BATCH: "1162",
    BANK: "MANDIRI",

    REMINDER_ID: "TEST_EMAIL",
    TRIGGER_TYPE: "TEST",
    ACTIVITY_NAME: "Pengujian Email",
    RECIPIENT_ROLE: "TESTER",
    RECIPIENT_EMAIL: TEST_EMAIL,
    SUBJECT: "Uji Coba Sistem Notifikasi SEMBAKO",
    BODY:
      "Halo,\n\n" +
      "Ini adalah email pengujian Google Apps Script.\n\n" +
      "Jika kamu menerima email ini, fungsi pengiriman email sudah berjalan.\n\n" +
      "Salam,\nSistem Notifikasi SEMBAKO",

    DEDUPE_KEY: "TEST_EMAIL_SEMBAKO_002"
  });

  Logger.log("Hasil pengiriman: " + hasil);
}

function buatIdentitasRecord(sheet, row, reminderId, tambahan = "") {
  // Data bantuan dimulai dari baris 4
  if (row < 4) {
    throw new Error("Baris tersebut bukan baris data bantuan.");
  }

  // Mengambil Program, Tahap, Batch, dan Bank dari kolom A:D
  const dataIdentitas = sheet.getRange(row, 1, 1, 4).getValues()[0];

  // Memastikan seluruh identitas terisi
  if (dataIdentitas.some(value => value === "" || value === null)) {
    throw new Error(
      "Identitas record pada kolom A:D belum lengkap di baris " + row
    );
  }

  // Menormalkan nilai agar konsisten
  const identitas = dataIdentitas.map(value =>
    String(value).trim().toUpperCase()
  );

  // Membentuk kunci unik berdasarkan isi kolom A:D
  const dedupeKey = [
    "SEMBAKO",
    reminderId,
    JSON.stringify(identitas),
    tambahan
  ].join("|");

  return dedupeKey;
}

function ujiIdentitasRecord() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SEMBAKO");

  // Menguji identitas pada baris 4
  const row = 4;

  const key = buatIdentitasRecord(
    sheet,
    row,
    "DAERAH_PEMBUKUAN"
  );

  Logger.log("Nomor baris: " + row);
  Logger.log("Identitas notifikasi: " + key);
}

function pasangTriggerEdit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const triggers = ScriptApp.getProjectTriggers();

  // Memeriksa apakah trigger sudah pernah dipasang
  const sudahAda = triggers.some(trigger =>
    trigger.getHandlerFunction() === "handleSheetEdit" &&
    trigger.getEventType() === ScriptApp.EventType.ON_EDIT
  );

  if (sudahAda) {
    Logger.log("Trigger handleSheetEdit sudah terpasang.");
    return;
  }

  // Memasang trigger on-edit
  ScriptApp.newTrigger("handleSheetEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  Logger.log("Trigger handleSheetEdit berhasil dipasang.");
}

function handleSheetEdit(e) {
  // Memastikan fungsi dipanggil melalui event edit
  if (!e || !e.range) {
    Logger.log("Event edit tidak ditemukan.");
    return;
  }

  const range = e.range;
  const sheet = range.getSheet();

  // Hanya memproses sheet SEMBAKO
  if (sheet.getName() !== "SEMBAKO") return;

  // Hanya memproses satu sel pada kolom G
  if (
    range.getColumn() !== 7 ||
    range.getNumRows() !== 1 ||
    range.getNumColumns() !== 1
  ) return;

  const row = range.getRow();

  // Data bantuan dimulai dari baris 4
  if (row < 4) return;

  Logger.log("Perubahan terdeteksi pada baris: " + row);

  // Membaca konfigurasi notifikasi
  const konfigurasi = bacaKonfigurasi();

  // Mencari konfigurasi khusus notifikasi Pokja
  const config = konfigurasi.find(item =>
    item.ACTIVE === true &&
    item.SHEET_NAME === sheet.getName() &&
    item.REMINDER_ID === "DAERAH_PEMBUKUAN" &&
    item.TRIGGER_TYPE === "STATUS_CHANGE"
  );

  if (!config) {
    Logger.log("Konfigurasi DAERAH_PEMBUKUAN tidak ditemukan.");
    return;
  }

  // Membaca nilai status terbaru pada kolom G
  const nilaiBaru = range.getValue();

  Logger.log("Status terbaru: " + nilaiBaru);
  Logger.log("Status pemicu: " + config.STATUS_VALUE);

  // Hanya mengirim notifikasi jika status berubah menjadi Selesai
  if (
    String(nilaiBaru).trim() !==
    String(config.STATUS_VALUE).trim()
  ) {
    Logger.log("Status belum memenuhi kondisi notifikasi.");
    return;
  }

  // Membaca identitas data dari kolom A:D
  const identitas = sheet.getRange(row, 1, 1, 4).getValues()[0];

  const program = identitas[0];
  const tahap = identitas[1];
  const batch = identitas[2];
  const bank = identitas[3];

  // Membaca status bantuan dari kolom T
  const statusBantuan = sheet.getRange(row, 20).getValue();

  // Membuat identitas unik untuk mencegah duplikasi
  const dedupeKey = buatIdentitasRecord(
    sheet,
    row,
    config.REMINDER_ID
  );

  // Menyiapkan dan mengirim notifikasi
  const hasil = kirimEmailNotifikasi({
    SHEET_NAME: config.SHEET_NAME,
    SOURCE_ROW: row,

    // Identitas data bantuan
    PROGRAM: program,
    TAHAP: tahap,
    BATCH: batch,
    BANK: bank,

    REMINDER_ID: config.REMINDER_ID,
    TRIGGER_TYPE: config.TRIGGER_TYPE,
    ACTIVITY_NAME: config.ACTIVITY_NAME,
    RECIPIENT_ROLE: config.RECIPIENT_ROLE,
    RECIPIENT_EMAIL: config.RECIPIENT_EMAIL,

    SUBJECT:
      config.SUBJECT_PREFIX +
      " - Laporan Pemindahbukuan",

    BODY:
      "Yth. Tim Pokja Daerah,\n\n" +
      "Laporan Pemindahbukuan untuk data bantuan berikut telah selesai diinput:\n\n" +
      "Program: " + program + "\n" +
      "Tahap: " + tahap + "\n" +
      "Batch: " + batch + "\n" +
      "Bank: " + bank + "\n" +
      "Status Bantuan: Telah ter top-up " + "\n\n" +
      "Mohon melakukan pemeriksaan lebih lanjut.\n\n" +
      "Salam,\n" +
      "Sistem Notifikasi SEMBAKO",

    STATUS_BANTUAN: statusBantuan,
    STATUS_INPUT: nilaiBaru,
    DEDUPE_KEY: dedupeKey
  });

  Logger.log("Hasil notifikasi pokja: " + hasil);
}

function bacaReminderDays(nilai) {
  // Memastikan konfigurasi tidak kosong
  if (nilai === "" || nilai === null || nilai === undefined) {
    throw new Error("REMINDER_DAYS belum diisi.");
  }

  // Mengubah nilai konfigurasi menjadi teks
  const teks = String(nilai).trim();

  // Memisahkan nilai berdasarkan tanda koma
  const daftarHari = teks.split(",");

  // Mengubah setiap nilai menjadi angka
  const hasil = daftarHari.map(item => {
    const angka = Number(item.trim());

    // Memastikan nilai berupa bilangan bulat nonnegatif
    if (!Number.isInteger(angka) || angka < 0) {
      throw new Error(
        "REMINDER_DAYS tidak valid: " + item
      );
    }

    return angka;
  });

  // Menghindari jadwal pengingat yang sama
  if (new Set(hasil).size !== hasil.length) {
    throw new Error("Terdapat REMINDER_DAYS yang duplikat.");
  }

  return hasil;
}

function ujiBacaReminderDays() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("REMINDER_CONFIG");

  // Mengambil nilai REMINDER_DAYS dari baris konfigurasi BANK_FREEZE
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const indexId = headers.indexOf("REMINDER_ID");
  const indexDays = headers.indexOf("REMINDER_DAYS");

  const configBank = data.find((row, index) =>
    index > 0 && row[indexId] === "BANK_FREEZE"
  );

  if (!configBank) {
    throw new Error("Konfigurasi BANK_FREEZE tidak ditemukan.");
  }

  const nilai = configBank[indexDays];
  const hasil = bacaReminderDays(nilai);

  Logger.log("Nilai konfigurasi: " + nilai);
  Logger.log("Hasil pembacaan: " + JSON.stringify(hasil));
  Logger.log("Jumlah jadwal pengingat: " + hasil.length);
}

function hitungSelisihHari(tanggalDeadline) {
  // Memastikan tanggal deadline tersedia
  if (
    !tanggalDeadline ||
    !(tanggalDeadline instanceof Date) ||
    isNaN(tanggalDeadline.getTime())
  ) {
    throw new Error("Tanggal deadline tidak valid.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Menggunakan zona waktu spreadsheet
  const zonaWaktu = ss.getSpreadsheetTimeZone();

  // Mengambil tanggal hari ini
  const tanggalSekarang = new Date();

  // Mengubah tanggal menjadi format YYYY-MM-DD
  const deadlineText = Utilities.formatDate(
    tanggalDeadline,
    zonaWaktu,
    "yyyy-MM-dd"
  );

  const sekarangText = Utilities.formatDate(
    tanggalSekarang,
    zonaWaktu,
    "yyyy-MM-dd"
  );

  // Memisahkan tahun, bulan, dan hari
  const deadlineParts = deadlineText.split("-").map(Number);
  const sekarangParts = sekarangText.split("-").map(Number);

  // Mengubah tanggal menjadi nilai milidetik UTC
  const deadlineUTC = Date.UTC(
    deadlineParts[0],
    deadlineParts[1] - 1,
    deadlineParts[2]
  );

  const sekarangUTC = Date.UTC(
    sekarangParts[0],
    sekarangParts[1] - 1,
    sekarangParts[2]
  );

  // Menghitung selisih hari
  const selisihHari = Math.round(
    (deadlineUTC - sekarangUTC) / (1000 * 60 * 60 * 24)
  );

  return selisihHari;
}

function ujiHitungSelisihHari() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SEMBAKO");

  // Mengambil tanggal deadline dari kolom R baris 4
  const tanggalDeadline = sheet.getRange("R4").getValue();

  // Menghitung selisih hari
  const hasil = hitungSelisihHari(tanggalDeadline);

  Logger.log("Tanggal hari ini: " + new Date());
  Logger.log("Tanggal deadline: " + tanggalDeadline);
  Logger.log("Selisih hari: " + hasil);

  if (hasil === 1) {
    Logger.log("Keterangan: H-1");
  } else if (hasil === 0) {
    Logger.log("Keterangan: H");
  } else {
    Logger.log("Keterangan: Bukan jadwal pengingat H-1 atau H.");
  }
}

function simulasiPengingatBank() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SEMBAKO");

  if (!sheet) {
    throw new Error("Sheet SEMBAKO tidak ditemukan!");
  }

  // Ambil konfigurasi pengingat bank
  const konfigurasi = bacaKonfigurasi();

  const config = konfigurasi.find(item =>
    item.ACTIVE === true &&
    item.SHEET_NAME === "SEMBAKO" &&
    item.REMINDER_ID === "BANK_FREEZE" &&
    item.TRIGGER_TYPE === "DATE"
  );

  if (!config) {
    throw new Error("Konfigurasi pengingat bank tidak ditemukan!");
  }

  // Baca daftar hari pengingat, misalnya [1, 0]
  const reminderDays = bacaReminderDays(config.REMINDER_DAYS);

  const barisTerakhir = sheet.getLastRow();

  let jumlahH1 = 0;
  let jumlahH = 0;
  let jumlahTerlewat = 0;

  Logger.log("=== SIMULASI PENGINGAT BANK ===");
  Logger.log("Hari pengingat: " + reminderDays);

  // Data dimulai dari baris 4
  for (let row = 4; row <= barisTerakhir; row++) {

    // Ambil tanggal deadline dari kolom R
    const deadline = sheet.getRange(row, 18).getValue();

    // Lewati jika tanggal tidak valid
    if (!(deadline instanceof Date) || isNaN(deadline.getTime())) {
      continue;
    }

    const selisihHari = hitungSelisihHari(deadline);

    // Periksa apakah termasuk jadwal pengingat
    if (!reminderDays.includes(selisihHari)) {
      continue;
    }

    // Ambil identitas data dari kolom A:D
    const identitas = sheet.getRange(row, 1, 1, 4).getValues()[0];

    const program = identitas[0];
    const tahap = identitas[1];
    const batch = identitas[2];
    const bank = identitas[3];

    // Tentukan label pengingat
    const label = selisihHari === 1 ? "H-1" : "H";

    if (label === "H-1") {
      jumlahH1++;
    } else {
      jumlahH++;
    }

    Logger.log(
      "Baris: " + row +
      " | Program: " + program +
      " | Tahap: " + tahap +
      " | Batch: " + batch +
      " | Bank: " + bank +
      " | Deadline: " + deadline +
      " | Pengingat: " + label
    );
  }

  Logger.log("=== RINGKASAN SIMULASI ===");
  Logger.log("Total kandidat H-1: " + jumlahH1);
  Logger.log("Total kandidat H: " + jumlahH);
  Logger.log("Total kandidat pengingat: " + (jumlahH1 + jumlahH));
  Logger.log("Simulasi selesai. Tidak ada email yang dikirim.");
}

function ujiKirimPengingatBankSatuBaris() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SEMBAKO");

  if (!sheet) {
    throw new Error("Sheet SEMBAKO tidak ditemukan!");
  }

  // Target pengujian: baris 42
  const row = 42;

  // Baca konfigurasi pengingat bank
  const konfigurasi = bacaKonfigurasi();

  const config = konfigurasi.find(item =>
    item.ACTIVE === true &&
    item.SHEET_NAME === "SEMBAKO" &&
    item.REMINDER_ID === "BANK_FREEZE" &&
    item.TRIGGER_TYPE === "DATE"
  );

  if (!config) {
    throw new Error("Konfigurasi pengingat bank tidak ditemukan!");
  }

  // Pengaman: pengujian hanya boleh menggunakan email percobaan
  const emailTes = TEST_EMAIL;

  if (String(config.RECIPIENT_EMAIL).trim() !== emailTes) {
    throw new Error(
      "Pengujian dibatalkan. Email penerima bukan alamat email tes."
    );
  }

  // Ambil tanggal deadline dari kolom R
  const deadline = sheet.getRange(row, 18).getValue();

  if (!(deadline instanceof Date) || isNaN(deadline.getTime())) {
    throw new Error("Tanggal deadline pada baris " + row + " tidak valid.");
  }

  // Hitung selisih hari
  const selisihHari = hitungSelisihHari(deadline);

  const reminderDays = bacaReminderDays(config.REMINDER_DAYS);

  if (!reminderDays.includes(selisihHari)) {
    Logger.log("Tidak ada pengingat untuk baris " + row);
    Logger.log("Selisih hari: " + selisihHari);
    return;
  }

  // Tentukan label pengingat
  const label = selisihHari === 1 ? "H-1" : "H";

  // Ambil identitas data A:D
  const identitas = sheet.getRange(row, 1, 1, 4).getValues()[0];

  const program = identitas[0];
  const tahap = identitas[1];
  const batch = identitas[2];
  const bank = identitas[3];

  // Ambil status bantuan untuk pencatatan saja
  const statusBantuan = sheet.getRange(row, 20).getValue();

  // Format tanggal deadline
  const zonaWaktu = ss.getSpreadsheetTimeZone();
  const deadlineText = Utilities.formatDate(
    deadline,
    zonaWaktu,
    "dd MMMM yyyy"
  );

  // Buat kunci unik berdasarkan identitas, label, dan deadline
  const deadlineKey = Utilities.formatDate(
    deadline,
    zonaWaktu,
    "yyyy-MM-dd"
  );

  const dedupeKey = buatIdentitasRecord(
    sheet,
    row,
    config.REMINDER_ID,
    label + "|" + deadlineKey
  );

  Logger.log("Mengirim pengingat untuk baris: " + row);
  Logger.log("Label pengingat: " + label);
  Logger.log("Deadline: " + deadlineText);

  // Kirim email dan catat hasilnya
  const hasil = kirimEmailNotifikasi({
    SHEET_NAME: config.SHEET_NAME,
    SOURCE_ROW: row,
    PROGRAM: program,
    TAHAP: tahap,
    BATCH: batch,
    BANK: bank,
    REMINDER_ID: config.REMINDER_ID,
    TRIGGER_TYPE: config.TRIGGER_TYPE,
    ACTIVITY_NAME: config.ACTIVITY_NAME,
    DEADLINE_DATE: deadline,
    REMINDER_LABEL: label,
    RECIPIENT_ROLE: config.RECIPIENT_ROLE,
    RECIPIENT_EMAIL: config.RECIPIENT_EMAIL,
    SUBJECT: config.SUBJECT_PREFIX + " - " + label +
      " " + config.ACTIVITY_NAME,
    BODY:
      "Yth. Tim Bank,\n\n" +
      "Kami ingin mengingatkan bahwa batas waktu " +
      config.ACTIVITY_NAME + " akan segera tiba.\n\n" +
      "Program: " + program + "\n" +
      "Tahap: " + tahap + "\n" +
      "Batch: " + batch + "\n" +
      "Bank: " + bank + "\n" +
      "Kegiatan: " + config.ACTIVITY_NAME + "\n" +
      "Tanggal Deadline: " + deadlineText + "\n" +
      "Pengingat: " + label + "\n\n" +
      "Mohon memperhatikan batas waktu tersebut.\n\n" +
      "Salam,\nSistem Notifikasi SEMBAKO",
    STATUS_BANTUAN: statusBantuan,
    STATUS_INPUT: "",
    DEDUPE_KEY: dedupeKey
  });

  Logger.log("Hasil pengiriman: " + hasil);
}

function prosesPengingatBank() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SEMBAKO");

  if (!sheet) {
    throw new Error("Sheet SEMBAKO tidak ditemukan!");
  }

  // Ambil konfigurasi pengingat bank
  const konfigurasi = bacaKonfigurasi();

  const config = konfigurasi.find(item =>
    item.ACTIVE === true &&
    item.SHEET_NAME === "SEMBAKO" &&
    item.REMINDER_ID === "BANK_FREEZE" &&
    item.TRIGGER_TYPE === "DATE"
  );

  if (!config) {
    throw new Error("Konfigurasi pengingat bank tidak ditemukan!");
  }

  // Baca aturan pengingat, misalnya [1, 0]
  const reminderDays = bacaReminderDays(config.REMINDER_DAYS);

  const barisTerakhir = sheet.getLastRow();

  if (barisTerakhir < 4) {
    Logger.log("Tidak ada data untuk diperiksa.");
    return;
  }

  // Ambil seluruh data A:T sekaligus
  const data = sheet.getRange(
    4, 1, barisTerakhir - 3, 20
  ).getValues();

  const zonaWaktu = ss.getSpreadsheetTimeZone();

  // Variabel ringkasan
  let jumlahDiperiksa = 0;
  let jumlahKandidat = 0;
  let jumlahTerkirim = 0;
  let jumlahDuplikat = 0;
  let jumlahGagal = 0;

  Logger.log("=== PROSES PENGINGAT BANK ===");

  for (let i = 0; i < data.length; i++) {

    const row = i + 4;
    const nilaiBaris = data[i];

    // Kolom R = indeks 17
    const deadline = nilaiBaris[17];

    // Lewati tanggal yang kosong atau tidak valid
    if (
      !(deadline instanceof Date) ||
      isNaN(deadline.getTime())
    ) {
      continue;
    }

    jumlahDiperiksa++;

    try {
      // Hitung selisih hari
      const selisihHari = hitungSelisihHari(deadline);

      // Periksa apakah termasuk H-1 atau H
      if (!reminderDays.includes(selisihHari)) {
        continue;
      }

      jumlahKandidat++;

      const label = selisihHari === 1 ? "H-1" : "H";

      // Identitas dari kolom A:D
      const program = nilaiBaris[0];
      const tahap = nilaiBaris[1];
      const batch = nilaiBaris[2];
      const bank = nilaiBaris[3];

      // Lewati jika identitas tidak lengkap
      if ([program, tahap, batch, bank].some(
        nilai => nilai === "" || nilai === null
      )) {
        Logger.log(
          "Baris " + row + " dilewati: identitas A:D tidak lengkap."
        );
        continue;
      }

      // Status bantuan hanya untuk pencatatan, bukan filter
      const statusBantuan = nilaiBaris[19];

      // Format tanggal deadline
      const deadlineText = Utilities.formatDate(
        deadline,
        zonaWaktu,
        "dd MMMM yyyy"
      );

      const deadlineKey = Utilities.formatDate(
        deadline,
        zonaWaktu,
        "yyyy-MM-dd"
      );

      // Buat DEDUPE_KEY
      const dedupeKey = buatIdentitasRecord(
        sheet,
        row,
        config.REMINDER_ID,
        label + "|" + deadlineKey
      );

      Logger.log(
        "Memproses baris " + row +
        " | " + program +
        " | " + tahap +
        " | " + batch +
        " | " + bank +
        " | " + label
      );

      // Kirim email dan catat hasilnya
      const hasil = kirimEmailNotifikasi({
        SHEET_NAME: config.SHEET_NAME,
        SOURCE_ROW: row,
        PROGRAM: program,
        TAHAP: tahap,
        BATCH: batch,
        BANK: bank,
        REMINDER_ID: config.REMINDER_ID,
        TRIGGER_TYPE: config.TRIGGER_TYPE,
        ACTIVITY_NAME: config.ACTIVITY_NAME,
        DEADLINE_DATE: deadline,
        REMINDER_LABEL: label,
        RECIPIENT_ROLE: config.RECIPIENT_ROLE,
        RECIPIENT_EMAIL: config.RECIPIENT_EMAIL,
        SUBJECT: config.SUBJECT_PREFIX +
          " - " + label + " " + config.ACTIVITY_NAME,
        BODY:
          "Yth. Tim Bank,\n\n" +
          "Kami ingin mengingatkan bahwa batas waktu " +
          config.ACTIVITY_NAME + " akan segera tiba.\n\n" +
          "Program: " + program + "\n" +
          "Tahap: " + tahap + "\n" +
          "Batch: " + batch + "\n" +
          "Bank: " + bank + "\n" +
          "Kegiatan: " + config.ACTIVITY_NAME + "\n" +
          "Tanggal Deadline: " + deadlineText + "\n" +
          "Pengingat: " + label + "\n\n" +
          "Mohon memperhatikan batas waktu tersebut.\n\n" +
          "Salam,\nSistem Notifikasi SEMBAKO",
        STATUS_BANTUAN: statusBantuan,
        STATUS_INPUT: "",
        DEDUPE_KEY: dedupeKey
      });

      // Klasifikasi hasil pengiriman
      if (hasil === "SENT") {
        jumlahTerkirim++;
      } else if (hasil === "FAILED") {
        jumlahGagal++;
      } else {
        jumlahDuplikat++;
      }

      Logger.log(
        "Hasil baris " + row + ": " + hasil
      );

    } catch (error) {
      jumlahGagal++;

      Logger.log(
        "Gagal memproses baris " + row +
        ": " + error.message
      );
    }
  }

  Logger.log("=== RINGKASAN PROSES BANK ===");
  Logger.log("Tanggal valid diperiksa: " + jumlahDiperiksa);
  Logger.log("Kandidat H-1/H: " + jumlahKandidat);
  Logger.log("Email berhasil dikirim: " + jumlahTerkirim);
  Logger.log("Duplikat dilewati: " + jumlahDuplikat);
  Logger.log("Gagal diproses: " + jumlahGagal);
  Logger.log("=== PROSES SELESAI ===");
}

function pasangTriggerPengingatBank() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const zonaWaktu = ss.getSpreadsheetTimeZone();

  // Periksa apakah trigger pengingat bank sudah ada
  const triggers = ScriptApp.getProjectTriggers();

  const triggerSudahAda = triggers.some(trigger =>
    trigger.getHandlerFunction() === "prosesPengingatBank"
  );

  if (triggerSudahAda) {
    Logger.log("Trigger pengingat bank sudah terpasang.");
    return;
  }

  // Buat trigger harian pukul 10.00
  ScriptApp.newTrigger("prosesPengingatBank")
    .timeBased()
    .atHour(10)
    .everyDays(1)
    .inTimezone(zonaWaktu)
    .create();

  Logger.log("Trigger pengingat bank berhasil dipasang.");
  Logger.log("Jadwal: Setiap hari sekitar pukul 10.00");
  Logger.log("Zona waktu: " + zonaWaktu);
}
