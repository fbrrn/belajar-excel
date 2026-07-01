# 📊 Dokumentasi Belajar Microsoft Excel

Selamat datang di repository belajar Excel saya! Repository ini dibuat sebagai wadah untuk mendokumentasikan proses belajar, latihan formula, hingga pengelolaan data menggunakan Microsoft Excel secara umum.

Tujuan dari repository ini adalah untuk mengasah kemampuan teknis (*hard skills*) dalam mengotomatisasi pekerjaan, merapikan dokumen, dan mengelola data secara efektif dan efisien.

---

## 📁 Proyek 1: Otomatisasi Tabel Pemesanan (VLOOKUP & HLOOKUP)

Di proyek pertama ini, saya mempelajari bagaimana menghubungkan baris data secara otomatis agar proses penginputan data menjadi lebih cepat dan terhindar dari kesalahan manual (*human error*). 

Berdasarkan studi kasus pada file latihan, terdapat tabel data pemesanan yang perlu diisi secara otomatis menggunakan dua metode pencarian berdasarkan format tabel referensinya:

### 🛠️ Rumus & Fungsi yang Dipelajari

1. **`VLOOKUP` (Vertical Lookup)**
   * **Kegunaan:** Mengambil data secara vertikal (tegak lurus). Digunakan untuk mengisi kolom `Nama Customer` dan `Domisili` dari tabel referensi utama berdasarkan kata kunci `No. Pemesanan`.

2. **`HLOOKUP` (Horizontal Lookup)**
   * **Kegunaan:** Mengambil data secara horizontal (mendatar). Digunakan untuk menentukan `Lama Pengiriman` secara otomatis berdasarkan data `Domisili` customer, dengan merujuk pada tabel referensi kota yang disusun menyamping.

3. **Absolute References (`$`)**
   * Mempelajari pentingnya mengunci cell/koordinat tabel (menggunakan tanda `$`) agar rumus dapat disalin ke baris lain di bawahnya dengan rapi tanpa merusak area referensi data.

---

## 📁 Proyek 2: Pengolahan Data Lanjutan (Pivot Table, Logika IF Bersyarat, & Barcode)

Di proyek kedua ini, saya mempelajari cara meringkas data dalam jumlah besar secara cepat, menerapkan logika bersyarat untuk perhitungan otomatis, serta membuat barcode langsung di dalam Excel untuk kebutuhan operasional/inventaris.

### 🛠️ Rumus & Fitur yang Dipelajari
1. **Pivot Table:** Merangkum, mengelompokkan, dan melihat total atau rata-rata data secara instan tanpa perlu menulis banyak rumus.
2. **Eksplorasi Fungsi Logika `IF` (Berbagai Variasi)**
   * **IF Single:** Menentukan keputusan berdasarkan satu kondisi/syarat dasar, yaitu menentukan status Untung atau Rugi, kondisi stok (Aman atau harus Restock), serta klasifikasi harga (Mahal atau Murah).
   * **IF Bertingkat (Nested IF):** Menguji beberapa kondisi sekaligus secara berurutan, yaitu mengklasifikasikan status customer (Silver, Gold, atau Platinum) dan menentukan lama pengiriman ke alamat customer (1, 2, atau 3 hari).
   * **IF Kombinasi `LEFT`, `MID`, `RIGHT`:** Mengambil keputusan berdasarkan kode atau karakter spesifik yang diekstrak dari sebagian teks, yaitu menentukan divisi kerja, jenis benefit, dan gender berdasarkan potongan huruf pada ID Karyawan.
   * **IF Kombinasi `AND` / `OR`:** Menghitung performa kerja, penentuan gaji, dan komisi berdasarkan beberapa syarat yang harus terpenuhi sekaligus (`AND`), atau salah satu syarat saja yang terpenuhi (`OR`).
3. **Pembuatan Barcode:** Mengubah teks/kode produk menjadi format barcode siap cetak menggunakan bantuan *font* khusus.

---

### ⚠️ Tantangan & Solusi (Lessons Learned)

Selama proses belajar, berikut adalah beberapa kendala teknis yang saya temui dan berhasil saya selesaikan:

1. **Error `#VALUE!` saat Menghitung Komisi pada Materi `IF(AND)` / `IF(OR)`**
   * **Masalah:** Rumus menghasilkan error `#VALUE!`. Setelah dianalisis, hal ini terjadi karena adanya kesalahan *human error* dalam meletakkan tanda kurung tutup `)` pada fungsi `SUM` yang bersarang di dalam rumus logika tersebut.
   * **Solusi:** Melakukan *double-check* susunan rumus dan merapikan kembali penempatan tanda kurung tutup `)` agar fungsi `SUM` menyelesaikan perhitungannya terlebih dahulu.

2. **Font "Libre Barcode 39" Tidak Tersedia di Excel**
   * **Masalah:** Teks kode tidak bisa berubah menjadi bentuk barcode karena font bawaan Excel belum mendukung.
   * **Solusi:** Melakukan instalasi font eksternal secara manual. Saya mengunduh font *Libre Barcode 39* (misalnya dari Google Fonts), menginstallnya ke sistem operasi komputer, lalu melakukan *restart* pada aplikasi Excel agar font baru tersebut dapat terbaca dan digunakan.

---

## 📁 Proyek 3: Sistem Absensi & Checklist Otomatis Interaktif

Proyek ketiga ini berfokus pada pembuatan sistem pencatatan kehadiran (absensi) dan tracker progres kerja yang interaktif menggunakan fitur kotak centang (*checkbox*), perhitungan persentase dinamis, serta visualisasi performa langsung di dalam sel.

### 🛠️ Rumus & Fitur yang Dipelajari

1. **Integrasi Checkbox & Logika Data**
   * Memanfaatkan fitur kotak centang interaktif yang menghasilkan nilai logika `TRUE` saat dicentang dan `FALSE` saat kosong untuk melacak kehadiran.
2. **Perhitungan Persentase Kehadiran Dinamis**
   * Menggunakan kombinasi rumus `=COUNTIF(range; TRUE)/COUNTA(range)` untuk menghitung rasio kehadiran secara akurat berdasarkan total hari yang aktif.
3. **Visualisasi Data dengan `SPARKLINE`**
   * Membaca tren progres secara visual langsung di dalam sel menggunakan grafik batang mini (*mini bar chart*) lewat fungsi `=SPARKLINE()`.

---

## ⚠️ Tantangan & Solusi (Lessons Learned)

Selama proses pengerjaan proyek absensi ini, saya menghadapi beberapa tantangan teknis akibat perbedaan ekosistem software dan ketidaksesuaian tutorial, namun berhasil diatasi dengan solusi alternatif:

1. **Keterbatasan Fitur Checkbox bawaan pada Excel 2013**
   * **Masalah:** Fitur menyisipkan checkbox hanya didukung pada Google Spreadsheet atau Excel Office 365, sementara perangkat lokal saya masih menggunakan Excel 2013.
   * **Solusi:** Saya beradaptasi dengan berpindah menggunakan platform cloud **Google Sheets** untuk mempraktikkan materi ini agar fitur checkbox interaktif dan pintasan sistem dapat berjalan dengan semestinya.

2. **Kesalahan Output Persentase pada Rumus Tutorial**
   * **Masalah:** Jika mengikuti tutorial dasar yang hanya menggunakan `=COUNTIF(B19:H19; TRUE)`, hasil yang keluar berupa angka mutlak (1, 2, 3) yang jika diformat ke persen menjadi tidak akurat (100%, 200%).
   * **Solusi:** Saya memodifikasi rumusnya menjadi `=COUNTIF(B3:H3; TRUE)/COUNTA(B3:H3)`. Dengan membagi total tanda centang dengan total cell data (`COUNTA`), hasil kalkulasi persentase menjadi akurat (misal: 1 centang dari 7 hari menghasilkan angka proporsional ~14%).

3. **Kendala Ketiadaan Font Khusus untuk Fungsi `=REPT` pada Progress Bar**
   * **Masalah:** Tutorial menyarankan pembuatan progress bar menggunakan rumus pengulangan teks `=REPT("|"; total)` lalu mengubah jenis fontnya. Namun, saya kesulitan menemukan font yang pas di sistem untuk menyatukan karakter tersebut menjadi balok solid.
   * **Solusi:** Saya mencari alternatif fungsi yang lebih modern dan rapi, yaitu menggunakan rumus `=SPARKLINE(cell; {"charttype"\"bar"; "max"\7; "color"\"blue"})`. Solusi ini menghasilkan indikator visual berupa grafik batang biru yang jauh lebih bersih, solid, dan responsif tanpa bergantung pada ketersediaan font eksternal.
  
---

*Catatan: Repository ini akan terus diperbarui seiring dengan berjalannya proses belajar Excel saya.*
