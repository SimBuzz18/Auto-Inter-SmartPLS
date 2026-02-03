# Inter-SmartPLS

**Inter-SmartPLS** adalah aplikasi otomatisasi untuk menginterpretasikan hasil output SmartPLS (Excel) ke dalam laporan naratif ilmiah (Word/DOCX) yang siap pakai. Aplikasi ini dirancang untuk mempermudah peneliti dan mahasiswa dalam menyusun Bab 4 (Hasil dan Pembahasan) dengan standar akademik yang ketat.

![Icon](icon.ico)

## Fitur Utama

### 1. Interpretasi Ilmiah Otomatis
Aplikasi tidak hanya menyalin angka, tetapi memberikan **narasi interpretatif** berdasarkan standar literatur statistik (Hair et al., Tenenhaus et al., dll.).
*   **Path Coefficients**: Menentukan signifikansi jalur (H1 diterima/ditolak).
*   **Specific Indirect Effects**: Analisis variabel mediasi.
*   **Goodness of Fit (GoF)**: Perhitungan otomatis nilai GoF dan klasifikasi (Small/Medium/Large) sesuai kriteria Tenenhaus (2004) & Akter (2011).
*   **VAF (Variance Accounted For)**: Penentuan tipe mediasi (Tidak Memediasi, Parsial, atau Penuh) berdasarkan Hair et al. (2014).
*   **Analisis Lain**: Outer Loadings, Construct Reliability (CR), AVE, Discriminant Validity (Fornell-Larcker, HTMT), R-Square, f-Square, Q-Square.

### 2. Format Laporan Standar Tesis/Disertasi
Output dokumen diformat secara profesional agar meminimalkan pengeditan manual:
*   **Font**: Times New Roman, Size 12 (Isi), Size 14 Bold (Judul).
*   **Daftar Isi Otomatis (ToC)**: Menggunakan fitur native Word Table of Contents.
*   **Bullet Points Native**: Simbol list dikenali dan diformat sebagai daftar poin asli Word.
*   **Struktur Rapih**: Judul tabel, tabel data, narasi definisi, dan interpretasi data tersusun sistematis.

### 3. Smart Filtering
*   Mendeteksi dan memperbaiki header tabel secara cerdas (misal: "Specific Indirect Effects" -> "Jalur | Nilai Efek").
*   Menyembunyikan output yang tidak relevan (seperti "Total Effect Histogram") untuk menjaga laporan tetap ringkas.

### 4. Build & Portable
*   Dilengkapi dengan **Icon Aplikasi** profesional.
*   Siap dibuild menjadi `.exe` standalone menggunakan PyInstaller (`Inter-SmartPLS.spec`).

## Cara Penggunaan
1.  **Export Hasil SmartPLS**: Export hasil analisis SmartPLS Anda ke Excel.
2.  **Jalankan Aplikasi**: Buka `Inter-SmartPLS`.
3.  **Pilih Folder**: Arahkan ke folder yang berisi file Excel hasil export.
4.  **Proses**: Klik tombol "Proses Interpretasi".
5.  **Hasil**: File `Interpretasi SmartPLS.docx` akan dibuat di folder yang sama.

## Struktur & Dependensi
*   `app.py`: GUI utama menggunakan `customtkinter`.
*   `smartpls_logic.py`: Logika ekstraksi data dan pemrosesan Excel.
*   `interpretation.py`: Engine generasi teks naratif dan pembuatan dokumen Word.
*   `Inter-SmartPLS.spec`: Konfigurasi build PyInstaller.
*   **Libraries**: `pandas`, `openpyxl`, `python-docx`, `customtkinter`, `pillow`.

## Build Executable
Untuk membuat file `.exe`:
```bash
pyinstaller Inter-SmartPLS.spec
```
Hasil build akan muncul di folder `dist/`.

---
*Dikembangkan untuk mempermudah analisis statistik akademik.*
