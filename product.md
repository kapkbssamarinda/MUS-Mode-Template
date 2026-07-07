# AuditWorkpaper Pro - MUS Template Mode

## 🎯 Penjelasan Produk
**AuditWorkpaper Pro** adalah sebuah aplikasi web (Client-Side) yang dirancang untuk mengotomatisasi proses pembuatan Kertas Kerja Audit menggunakan metode **Monetary Unit Sampling (MUS)**. Aplikasi ini memungkinkan auditor untuk menarik sampel transaksi dari buku besar (*General Ledger*) secara objektif berdasarkan batas materialitas, dan menghasilkan format Excel (atau PDF) siap pakai untuk pengujian detail (*Test of Details*). 

Kelebihan utama aplikasi ini adalah beroperasi sepenuhnya di sisi klien (browser), sehingga memastikan keamanan dan kerahasiaan data tanpa adanya pertukaran data ke server.

---

## 🛠️ Detail Fungsi yang Digunakan

Berikut adalah penjelasan mengenai fungsi-fungsi utama yang digunakan di dalam *source code* (JavaScript):

### 1. Fungsi Interaksi UI (User Interface)
* **`handleFiles(files)`**: Fungsi ini menangani *file* yang dimasukkan oleh pengguna, baik melalui klik tombol *upload* maupun melalui metode *drag-and-drop*. Fungsi ini juga memperbarui tampilan untuk menunjukkan bahwa file berhasil diunggah.

### 2. Fungsi Validasi
* **`validateForm()`**: Memeriksa kelengkapan seluruh *field input* yang wajib diisi pada form penugasan (seperti Nama Klien, Periode, Batas Materialitas, Opsi Pengurutan, dan File Excel). Jika ada yang kosong, fungsi ini akan memunculkan *alert* peringatan interaktif menggunakan **SweetAlert2**.

### 3. Fungsi Bantuan (Helpers)
* **`getNilaiMaterialitas()`**: Mengambil *input* angka materialitas dari form, menghilangkan format titik (ribuan), dan mengubahnya menjadi tipe data numerik (integer).
* **`getSortingOption()`**: Mengambil pilihan opsi pengurutan yang dipilih oleh pengguna (berdasarkan tanggal `date` atau nominal `nominal`).
* **`parseToTimestamp(val)`**: Mengonversi berbagai format tanggal (dari format teks, angka seri Excel, atau objek Date) menjadi *timestamp* seragam agar memudahkan proses penyortiran/pengurutan (sorting) tanggal.
* **`formatTanggal(item)`**: Memformat objek atau *timestamp* tanggal kembali menjadi *string* berformat standar `DD/MM/YYYY` untuk ditampilkan di Excel atau PDF.
* **`formatRupiahGL(num)`**: Memformat nilai numerik ke dalam bentuk mata uang (contoh: pemisah ribuan dengan titik, desimal dengan koma) khusus untuk output PDF.

### 4. Fungsi Logika Inti (Sampling)
* **`getSampledItems(inputSheet, batasMaterialitas, sortingOption)`**: Ini adalah fungsi **Logika Sampling MUS (Monetary Unit Sampling)** utama. 
  * Pertama, sistem membaca data tiap baris Excel.
  * Kedua, menyaring transaksi dengan nominal (absolut) **di atas batas materialitas**.
  * Ketiga, mengurutkan data dari nilai tertinggi ke terendah dan mengambil **15 transaksi teratas (Top 15)**.
  * Keempat, mengacak sisa data dari hasil saringan dan mengambil **15 transaksi acak (Random 15)**.
  * Terakhir, menggabungkan total 30 sampel dan mengurutkan ulang hasil akhirnya berdasarkan opsi yang dipilih (Tanggal atau Nominal).

### 5. Fungsi Logika Utama (Proses Export)
* **`processAuditWorkpaper(mode)`**: Fungsi utama (driver) yang dieksekusi saat tombol *Proses* diklik.
  * Fungsi ini pertama kali memanggil `validateForm()`.
  * Menggunakan **ExcelJS** untuk membaca file Excel *input* yang diunggah.
  * Jika parameter mode adalah `'pdf'`, ia memanggil fungsi `generatePDF()`.
  * Jika mode adalah `'excel'`, ia memuat file `Template_Output.xlsx`, lalu membuat sheet baru berdasarkan tiap-tiap akun, dan menggunakan `copyRows()` untuk menduplikasi *header/footer* dari template.
  * Kemudian ia memasukkan 30 hasil sampel dari `getSampledItems()` ke dalam baris yang telah disediakan beserta format garis, *font*, dan *Data Validation* *dropdown* ("V", "X").
  * File hasil akhir kemudian diunduh langsung menggunakan *FileSaver.js*.

### 6. Fungsi Render PDF
* **`generatePDF(allSheetsData, clientName, period)`**: Menggunakan *library* **jsPDF** (dan *autoTable*) untuk mencetak hasil penarikan sampel ke dalam format tabel PDF yang rapi, lengkap dengan header identitas klien dan periode per akunnya.

### 7. Fungsi Manipulasi Template Excel
* **`copyRows(srcSheet, destSheet, srcStartRow, srcEndRow, destStartRow, sheetNameForReplace)`**: Fungsi bantuan khusus untuk menduplikasi sel dari sheet *template* ke sheet hasil. Fungsi ini menyalin *value*, mengubah teks `<<NamaAkun>>` secara dinamis, serta menyalin seluruh *style* dan validasi (*data validation*).
* **`parseRangeString(rangeStr)`**: Menerjemahkan *string* referensi (*merged cells*) seperti `A1:C3` menjadi bentuk indeks baris dan kolom yang dapat dibaca dan disalin ulang oleh ExcelJS.
