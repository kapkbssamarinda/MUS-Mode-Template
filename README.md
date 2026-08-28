# 📄 AuditWorkpaper Pro - MUS Template Mode (Monetary Unit Sampling)

## 📖 Pengertian
**AuditWorkpaper Pro** adalah sebuah aplikasi web pintar (Client-side Web-App) yang dirancang untuk mengotomatisasi proses pembuatan Kertas Kerja Audit menggunakan metode **Monetary Unit Sampling (MUS)**. 

Aplikasi ini sangat memudahkan auditor dalam menarik sampel transaksi secara objektif dari populasi data buku besar (General Ledger), berdasarkan **batas nilai materialitas** yang ditentukan, dan secara otomatis menyusunnya ke dalam format Kertas Kerja Pengujian Detail (*Test of Details*) Excel yang siap pakai.

Aplikasi ini beroperasi sepenuhnya di sisi klien (browser) menggunakan HTML, CSS, JavaScript murni, dan *library* ExcelJS, sehingga menjamin keamanan dan kerahasiaan data klien karena data tidak pernah dikirim ke server mana pun.

---

## ✨ Fitur Utama

1. 🚀 **Pemrosesan Multi-Akun**: Mendukung penarikan sampel untuk banyak akun sekaligus. Anda hanya perlu memisahkan masing-masing akun dalam *Sheet* yang berbeda pada satu file Excel input.
2. 🧮 **Filter Materialitas Otomatis**: Secara otomatis memisahkan dan mengabaikan nilai transaksi yang berada di bawah ambang batas materialitas audit.
3. 🎯 **Smart Sampling Logic**: Mengambil sampel secara terstruktur dengan total maksimal 30 sampel:
   * **15 Sampel Teratas** (Top 15 Nominal terbesar di atas materialitas)
   * **15 Sampel Acak** (Random 15 dari sisa data di atas materialitas)
4. 📑 **Template Kertas Kerja Lengkap**: Menghasilkan output Excel (*Test of Detail*) otomatis yang sudah berisi header penugasan, daftar sampel, kolom pengujian asersi dengan *dropdown validation*, dan rumus/formatting yang rapi.
5. 🔄 **Opsi Pengurutan (Sorting)**: Memungkinkan auditor untuk menyajikan sampel di kertas kerja dengan urutan **Berdasarkan Tanggal** atau **Berdasarkan Nominal**.
6. 📱 **PWA Ready (Progressive Web App)**: Dapat diinstal langsung sebagai aplikasi mandiri di Desktop maupun Smartphone Anda.

---

## 🛠️ Teknologi yang Digunakan
* HTML5, CSS3 (Bootstrap 5), Vanilla JavaScript
* **ExcelJS** & **FileSaver.js** (Untuk memproses, membaca, dan membuat file Excel secara lokal)
* **SweetAlert2** (Untuk notifikasi UI yang interaktif)
* **Service Workers** & Web Manifest (Untuk dukungan Offline & PWA)

---

## 📚 Cara Penggunaan

### 1. Persiapan Data (Input)
1. Buka aplikasi di browser Anda.
2. Pada panel **Langkah Kerja** di aplikasi, klik tombol **"Unduh Template Input"** (`Template_Input.xlsx`).
3. Buka file tersebut dan isi dengan data populasi transaksi (Format: `Tanggal`, `Voucher`, `Keterangan`, `Nominal`). 
4. *Penting:* Jika Anda ingin menguji beberapa akun sekaligus, cukup **buat Sheet baru** di dalam file Excel tersebut untuk tiap-tiap akun (1 Sheet = 1 Akun).

### 2. Lengkapi Informasi Penugasan
Pada panel kiri aplikasi, isi secara lengkap detail penugasan audit Anda:
* **Nama Klien / Entitas**
* **Periode Audit**
* **Batas Nilai Materialitas** *(Sistem akan mengabaikan nilai di bawah nominal ini)*
* **Dibuat Oleh & Tanggal** (Preparer)
* **Direview Oleh & Tanggal** (Reviewer)
* **Opsi Pengurutan Data** (Pilih "Urut Per Tanggal" atau "Urut Per Nominal")

### 3. Eksekusi Sampling (Upload & Proses)
1. Pergi ke panel sebelah kanan (**Eksekusi Sampling**).
2. *Drag and Drop* atau klik untuk mengunggah file `Template_Input.xlsx` yang sudah Anda isi dengan populasi data.
3. Klik tombol **"Proses & Download Kertas Kerja"**.
4. Sistem akan memproses data dan secara otomatis mengunduh file hasil bernama `Kertas_Kerja_[Nama_Klien].xlsx`. File ini sudah berbentuk Kertas Kerja Uji Detail (ToD) yang rapi dan siap Anda lengkapi.

---

## 🧠 Memahami Logika Sampling Aplikasi
Agar kertas kerja dapat diandalkan, penting bagi auditor untuk mengetahui cara sistem ini memilih sampel. Berikut adalah urutan logikanya:
1. **Pengumpulan Data:** Sistem membaca setiap baris transaksi dari Excel Input.
2. **Penyaringan (Filtering):** Sistem membuang semua transaksi yang nilai nominal absolutnya **kurang dari Batas Nilai Materialitas** yang diinputkan di form.
3. **Pengurutan Awal:** Sisa data diurutkan dari nominal terbesar hingga terkecil.
4. **Ekstraksi Sampel (Max 30):**
   * Mengambil **15 transaksi teratas** (Top 15 Nominal).
   * Sisa dari data tersebut kemudian diacak secara sistematis.
   * Mengambil **15 transaksi acak** (Random 15).
5. **Finalisasi Urutan:** Ke-30 sampel tersebut kemudian diurutkan ulang berdasarkan pilihan *Opsi Pengurutan* user (Tanggal atau Nominal) untuk dicetak ke dalam template Kertas Kerja (Excel Output).

---

## 💻 Cara Menjalankan Secara Lokal (Local Setup)

Aplikasi ini dapat dijalankan dengan sangat mudah menggunakan **Python**:

### Menggunakan Python (Direkomendasikan)
1. Buka terminal / command prompt di folder proyek.
2. Jalankan perintah:
   ```bash
   python server.py
   ```
3. Peramban web Anda akan otomatis terbuka di `http://127.0.0.1:8000/`.

> **Pengguna Windows**: Anda juga cukup melakukan **klik ganda (double-click)** pada file `start_server.bat`.

### Opsi Parameter Tambahan
* Menggunakan port tertentu: `python server.py --port 8080`
* Menjalankan tanpa membuka browser otomatis: `python server.py --no-browser`

---

**Dikembangkan oleh:** Viany Ramadhany | Project MUS