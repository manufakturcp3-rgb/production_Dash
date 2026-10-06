# Sistem Sinkronisasi & Scraping Data Produksi (Pengganti IMPORTRANGE)

Proyek ini dibuat untuk menggantikan rumus bawaan `=IMPORTRANGE` yang lambat dan sering error dengan script otomatis Google Apps Script yang mengambil data secara berkala, menampilkannya di spreadsheet target, dan menyinkronkannya ke Firebase.

---

### Rumus yang Digantikan:
1. `=IMPORTRANGE("1P34UU9Q1pN4afM9L8Mn23LD90QdnbQ7Y", "WebData_Produk!A1:Z100")`
2. `=IMPORTRANGE("1OrmtFMggqx0j5uW_X5Nxfo23ty61F7m6ppKWEPvPu0s", "DASHBOARD KHUSUS!A1:Z100")`

### Target Tampilan:
* **Spreadsheet ID**: `146f5qPWBsDEyIn1e6WpjN-bfkGZXwGQfdqlybhsXdS4`
* **Target Sheet (GID)**: `1192757151`
* **URL**: [Buka Spreadsheet Target](https://docs.google.com/spreadsheets/d/146f5qPWBsDEyIn1e6WpjN-bfkGZXwGQfdqlybhsXdS4/edit?gid=1192757151#gid=1192757151)
* **Firebase Target**: Realtime Database (`produksi/webdata_produk` dan `produksi/dashboard_khusus`)

---

## 🚀 Keunggulan Dibandingkan IMPORTRANGE
1. **Tidak Ada Batas Kuota IMPORTRANGE**: Tidak akan membebani spreadsheet target dengan formula berat.
2. **Dual-Strategy Scraping (Tahan Masalah Akun Lain)**:
   - Script mencoba membuka via native `SpreadsheetApp.openById()`.
   - Jika akses akun terbatas, script otomatis beralih ke engine fallback **Google Visualization (GViz) CSV Exporter** yang bekerja langsung dengan link berstatus *"Anyone with the link"*.
3. **Pembersihan Otomatis**: Baris/kolom kosong di luar data aktif langsung dipotong rapi (*clean array*).
4. **Terintegrasi Langsung ke Firebase**: Mengubah data tabel baris-kolom menjadi JSON Objek dengan nama header yang aman (karakter khusus disanitasi).
5. **Otomatis Berjadwal (Trigger)**: Dapat disetel sync otomatis tiap 15 menit atau 1 jam tanpa perlu membuka file.

---

## 🛠️ Cara Pasang & Menjalankan

### Cara 1: Langsung di Google Sheets (Paling Praktis)
1. Buka spreadsheet target:
   `https://docs.google.com/spreadsheets/d/146f5qPWBsDEyIn1e6WpjN-bfkGZXwGQfdqlybhsXdS4/edit?gid=1192757151#gid=1192757151`
2. Klik menu **Extensions (Ekstensi)** > **Apps Script**.
3. Hapus kode bawaan, lalu salin seluruh isi file [`Kode.js`](Kode.js) ke dalam editor script.
4. (Opsional) Buka `Project Settings (Ikon Gerigi)` > centang *"Show 'appsscript.json' manifest file in editor"*, lalu salin isi [`appsscript.json`](appsscript.json).
5. Pada pilihan fungsi di atas, pilih `syncAllData` lalu klik **Run (Jalankan)**.
6. Berikan izin akses (*Review Permissions*) pada popup akun Google Anda.
7. Setelah selesai, refresh spreadsheet Anda. Anda akan melihat menu baru: **🔄 Sync Produksi**.

---

### Cara 2: Pasang Menggunakan clasp (CLI)
Jika ingin mengunggah via terminal menggunakan akun `itdt.west@gmail.com`:
1. Pastikan fitur **Google Apps Script API** sudah diaktifkan di akun Anda dengan mengunjungi:
   👉 **https://script.google.com/home/usersettings** (Ubah sakelar menjadi **ON**).
2. Di terminal, masuk ke folder `produksi`:
   ```bash
   cd /home/renagge/Dokumen/appscript/produksi
   clasp create --title "Produksi Sync" --type standalone
   clasp push
   ```

---

## ⚙️ Konfigurasi Firebase
Jika ingin data otomatis dikirim ke Firebase Realtime Database:
1. Buka file `Kode.js` bagian `CONFIG.FIREBASE`:
   ```javascript
   FIREBASE: {
     DATABASE_URL: 'https://nama-project-anda-default-rtdb.firebaseio.com',
     AUTH_SECRET: 'Kunci-Rahasia-Database-Jika-Ada', // biarkan kosong jika test mode
     ENABLED: true // ubah ke true
   }
   ```
2. Struktur data yang akan tersimpan di Firebase:
   * `/produksi/webdata_produk`: Array objek berisi record produk.
   * `/produksi/dashboard_khusus`: Array objek data dashboard.
   * `/produksi/last_sync`: Metadata timestamp & baris data terakhir.

---

## ⏰ Mengaktifkan Auto-Sync (Trigger Latar Belakang)
Anda bisa mengaktifkan jadwal otomatis langsung dari menu spreadsheet:
* Klik menu **🔄 Sync Produksi** > **⏰ Pasang Auto-Sync (Tiap 15 Menit)** atau **(Tiap 1 Jam)**.
* Atau jalankan fungsi `setupTrigger15Min()` sekali di editor Apps Script.
