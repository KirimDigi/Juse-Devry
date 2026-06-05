// ==========================================
// GOOGLE APPS SCRIPT FOR RSVP & WISHES
// ==========================================
// Cara Pemasangan:
// 1. Buka Google Spreadsheet Anda.
// 2. Klik menu 'Ekstensi' (Extensions) -> 'Apps Script'.
// 3. Hapus kode bawaan, lalu paste kode di bawah ini.
// 4. Pastikan baris pertama spreadsheet sudah memiliki header berikut (jika kosong, script akan membuatnya secara otomatis):
//    A1: timestamp
//    B1: nama tamu
//    C1: ucapan
//    D1: konfirmasi kehadiran
//    E1: jumlah tamu
// 5. Simpan proyek dengan klik ikon disket.
// 6. Klik tombol 'Terapkan' (Deploy) -> 'Penerapan baru' (New deployment).
// 7. Pilih jenis penerapan: 'Aplikasi Web' (Web App).
// 8. Konfigurasikan:
//    - Jalankan sebagai: 'Saya' (Me)
//    - Siapa yang memiliki akses: 'Siapa saja' (Anyone)
// 9. Klik 'Terapkan' (Deploy). Anda mungkin perlu memberikan izin akses ke akun Google Anda.
// 10. Salin 'URL Aplikasi Web' yang diberikan. URL ini akan dimasukkan ke dalam kode website Anda.

function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var dataRange = sheet.getDataRange();
    var rows = dataRange.getValues();
    
    if (rows.length <= 1) {
      return ContentService.createTextOutput(JSON.stringify({ status: "success", data: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    var headers = rows[0].map(function(h) { return h.toString().toLowerCase().trim(); });
    
    // Temukan indeks kolom berdasarkan header
    var timestampIdx = headers.indexOf("timestamp");
    var namaIdx = headers.indexOf("nama tamu");
    var ucapanIdx = headers.indexOf("ucapan");
    var konfirmasiIdx = headers.indexOf("konfirmasi kehadiran");
    var jumlahIdx = headers.indexOf("jumlah tamu");
    
    // Fallback jika tidak cocok sempurna
    if (timestampIdx === -1) timestampIdx = 0;
    if (namaIdx === -1) namaIdx = 1;
    if (ucapanIdx === -1) ucapanIdx = 2;
    if (konfirmasiIdx === -1) konfirmasiIdx = 3;
    if (jumlahIdx === -1) jumlahIdx = 4;
    
    var wishes = [];
    for (var i = 1; i < rows.length; i++) {
      var row = rows[i];
      var ts = row[timestampIdx];
      var nama = row[namaIdx];
      var ucapan = row[ucapanIdx];
      var konfirmasi = row[konfirmasiIdx];
      var jumlah = row[jumlahIdx];
      
      if (ts instanceof Date) {
        ts = Utilities.formatDate(ts, "GMT+7", "yyyy-MM-dd HH:mm");
      } else {
        ts = ts ? ts.toString() : "";
      }
      
      if (nama && ucapan) {
        wishes.push({
          timestamp: ts,
          nama: nama,
          ucapan: ucapan,
          konfirmasi: konfirmasi,
          jumlah: jumlah
        });
      }
    }
    
    // Tampilkan dari yang paling baru di atas
    wishes.reverse();
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success", data: wishes }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var params = e.parameter;
    
    // Jika data dikirim berupa JSON
    if (!params.nama && e.postData && e.postData.contents) {
      try {
        params = JSON.parse(e.postData.contents);
      } catch (jsonErr) {}
    }
    
    var nama = params.nama || params.author || "";
    var ucapan = params.ucapan || params.comment || "";
    var konfirmasi = params.konfirmasi || params.attendance || "";
    var jumlah = params.jumlah || params.guest || "";
    
    // Normalisasi nilai kehadiran
    if (konfirmasi === "present" || konfirmasi === "Hadir") {
      konfirmasi = "Hadir";
    } else if (konfirmasi === "notpresent" || konfirmasi === "Tidak Hadir") {
      konfirmasi = "Tidak Hadir";
      jumlah = "0"; // Jika tidak hadir, jumlah tamu diatur ke 0
    }
    
    if (!nama || !ucapan) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Nama dan ucapan wajib diisi" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Jika sheet masih kosong, buat baris header
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["timestamp", "nama tamu", "ucapan", "konfirmasi kehadiran", "jumlah tamu"]);
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#dcdcdc");
    }
    
    var timestamp = new Date();
    sheet.appendRow([timestamp, nama, ucapan, konfirmasi, jumlah]);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Berhasil mengirim ucapan!" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
