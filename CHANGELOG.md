# Changelog / Update Log
Catatan riwayat pembaruan untuk proyek **AuditWorkpaper Pro - MUS Template Mode**.

## [1.1.0] - 2026-07-07
### Ditambahkan (Added)
- **Fitur Kustomisasi Jumlah Sampel**: Menambahkan dua input opsional baru pada antarmuka pengguna (`index.html`) untuk menentukan jumlah spesifik untuk "Sampel Tertinggi" (Top Samples) dan "Sampel Acak" (Random Samples).
- **Penyesuaian Teks UI**: Mengubah teks pada panel "Metode Sampling" agar mencerminkan sifat dinamis dari jumlah pengambilan sampel (secara *default* tetap maksimal 30: 15 Teratas dan 15 Acak).

### Diubah (Changed)
- **Modularisasi Logika JS**: Memecah fungsi `getSampledItems()` di `script.js` menjadi fungsi-fungsi modular yang lebih kecil untuk mempermudah perawatan dan pengembangan ke depannya:
  - `determineSampleCounts()`: Menentukan berapa jumlah sampel yang harus ditarik berdasarkan input dari pengguna (atau menggunakan nilai default).
  - `extractTopSamples()`: Mengurutkan dan mengekstrak jumlah sampel teratas.
  - `extractRandomSamples()`: Mengacak dan mengekstrak jumlah sampel acak.
- **Dukungan Parameter Konduktor**: Fungsi ekspor (PDF & Excel) sekarang meneruskan variabel `sampleCounts` yang bersifat dinamis ke dalam parameter fungsi `getSampledItems()`.

### Info Penting / Isu yang Diketahui (Known Issues)
- **Failed to Fetch pada Excel Export**: Jika aplikasi dijalankan langsung melalui protokol `file:///` tanpa menggunakan *Local Web Server* (misalnya Live Server), ekspor ke Excel akan mengalami *error* "Failed to fetch" akibat kebijakan CORS dari browser saat membaca file template statis (`assets/Template_Output.xlsx`). Ekspor PDF tidak terdampak masalah ini.
