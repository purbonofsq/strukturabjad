# Dokumentasi Program Analisis Struktur Al-Quran (JBAH & Tabel AH)

## Gambaran Umum

Kumpulan skrip Python ini dirancang untuk menganalisis struktur Al-Quran berdasarkan jumlah ayat dalam setiap Juz menggunakan suatu metode yang disebut JBAH. Tujuan utamanya adalah untuk mengambil data ayat-ayat mentah, memprosesnya melalui beberapa tahapan transformasi (chunking, reordering, flattening), dan menghasilkan output dalam bentuk tabel-tabel terstruktur, termasuk Tabel Asmaul Husna (AH), yang diduga memiliki pola atau makna khusus berdasarkan struktur numerik ini.

## Komponen Program

Program ini terdiri dari beberapa file modul utama dan skrip eksekusi:

### 1. `backend1.py` - Inti Logika Pemrosesan

File ini mendefinisikan kelas `PengolahDataJBAH`, yang berisi seluruh logika inti untuk memuat, memfilter, dan memproses data Al-Quran.

*   **`__init__(self, file_path, ...)`**: Memuat data mentah dari file Excel (`Struktur Quran18.xlsx` atau sumber lainnya) ke dalam `self.raw_data`.
*   **`filter_data(self, filter_col, filter_value)`**: Memfilter `raw_data` untuk Juz tertentu berdasarkan `aq_juz_id`. Ini menjadi titik awal pemrosesan per Juz.
*   **`set_metode(self, metode)`**: Menentukan metode "Poros Siklik" (saat ini `T111`, `T222`, `T333`). Metode ini menggeser urutan ayat dalam Juz yang difilter agar titik referensi tertentu (misalnya, ayat tengah untuk `T111`) jatuh di posisi tengah array data.
*   **`proses_1_dan_2(self)`**:
    *   **Proses 1**: Melakukan chunking data yang telah diatur porosnya menjadi blok-blok simetris (A, a, B, b, C1, c1, ..., F, ..., g) berdasarkan lebar sayap tetap. Hasilnya disimpan dalam `df_p1`.
    *   **Proses 2**: Mengambil 15 blok utama dari Proses 1 dan mengurutkannya ulang sesuai template tertentu. Hasilnya disimpan dalam `df_p2`.
*   **`proses_3(self)`**: Melakukan "flattening" terhadap `df_p2`, mengubah struktur tabel menjadi satu kolom ayat dengan ID (`p3_id`) dan Label blok asalnya (`df_p3`). Ini adalah dataset fundamental untuk pembuatan Tabel AH.
*   **`generate_tabel_jb_init(self)`**: Membuat ringkasan informasi dasar (ayat awal, akhir, tengah, blok F, dll.) berdasarkan metode yang aktif dan Juz yang diproses.
*   **`generate_tabel_ah_final(self)`**: Menggunakan `df_p3` untuk mengisi Tabel Asmaul Husna. Tabel ini mengaitkan huruf Latin/Hari, huruf Arab, dan nomor urut Asmaul Husna dengan ayat-ayat tertentu yang dipilih dari `df_p3` berdasarkan indeks yang ditentukan (misalnya indeks ke-1, 2, 3, 4, dan 101-minus-indeks).
*   **`export_excel(self, output_name, ...)`**: Menyediakan antarmuka untuk mengekspor hasil-hasil DataFrame ke file Excel melalui modul `outputproc`.

### 2. `outputproc.py` - Modul Ekspor dan Styling Excel

File ini bertanggung jawab untuk menulis DataFrame ke file Excel dan menerapkan gaya/formatting.

*   **`write_excel_output(processor, ...)`**: Menyimpan DataFrame (`df_init`, `df_p1`, `df_p2`, `df_p3`, `df_final`) ke sheet-sheet terpisah dalam file Excel.
*   **`export_and_format_output(...)`**: Fungsi utama yang dipanggil dari `backend1`. Memanggil `write_excel_output` terlebih dahulu, lalu menerapkan formatting menggunakan kelas `FormatterTabelFinal`.
*   **`FormatterTabelFinal`**:
    *   **`_apply_standard_style(ws)`**: Menerapkan gaya dasar (font, alignment, border, warna header) untuk sheet umum seperti Init, Proses 1, 2, 3.
    *   **`_apply_final_style(ws)`**: Menerapkan gaya khusus dan kompleks untuk sheet "Final" (Tabel AH), termasuk merging cells untuk header, styling kolom berbeda, dan footer berisi signature dan timestamp.
    *   **`apply_formatting(self)`**: Metode utama yang memuat workbook, mendeteksi sheet target, dan menerapkan style yang sesuai.

### 3. `dataextraction.py` - Utilitas Ekstraksi Data Mentah

File sederhana yang berisi fungsi utilitas untuk menyimpan DataFrame hasil filter (sebelum pemrosesan JBAH) ke file Excel terpisah.

*   **`save_filtered_data_to_excel(...)`**: Membuat folder `rawprocess`, menghapus kolom `TextAyat` jika ada, dan menyimpan DataFrame ke file `FilteredData_Juz<X>.xlsx`.

### 4. `rawdata_ah_alljuz.py` - Skrip Generasi Data Mentah Gabungan (Versi 1)

Skrip utama yang menjalankan pipeline JBAH untuk semua 30 Juz dan menggabungkan hasil `df_p3` dari tiap Juz.

*   **`generate_raw_data_all_juz()`**:
    1.  Memuat data dari `Struktur Quran18.xlsx`.
    2.  Melakukan iterasi dari Juz 1 hingga 30.
    3.  Untuk setiap Juz: filter, set metode `T111`, jalankan `proses_1_dan_2`, `proses_3`.
    4.  Mengambil `df_p3` dan menambahkan kolom `ah_juz_id`.
    5.  Menggabungkan semua `df_p3` yang telah diberi label Juz ke dalam satu DataFrame (`final_df`).
    6.  Menyimpan `final_df` ke file `rawdatatabel_ah_alljuz.xlsx`.

### 5. `rawdata_ah_hrfain_alljuz.py` - Skrip Generasi Data Mentah Gabungan (Versi 2)

Identik dalam fungsi dengan `rawdata_ah_alljuz.py`, namun menggunakan file sumber data yang berbeda (`rawahaq_expanded.xlsx`) dan kelas `PengolahDataJBAH` dari `backend2.py` (yang tidak disertakan tetapi diasumsikan memiliki struktur dan API yang identik dengan `backend1.py`).

*   **`generate_raw_data_all_juz()`**: Fungsi yang sama dengan versi 1, tetapi menggunakan file sumber dan backend yang berbeda untuk menghasilkan `rawdatatabel_ah_hrfain_alljuz.xlsx`.

### 6. `userinput.py` - Skrip Interaktif untuk Eksekusi Manual

File ini menyediakan antarmuka berbasis teks yang memungkinkan pengguna menentukan parameter eksekusi secara manual.

*   **`sanitize_filename(name)`**: Membersihkan input pengguna agar aman digunakan sebagai nama file/folder (mengganti karakter ilegal dengan `_`).
*   **`parse_juz_input(input_str)`**: Menguraikan input pengguna untuk memilih Juz tunggal (misalnya `1`) atau rentang Juz (misalnya `1-5`).
*   **`list_available_datasource(script_dir)`**: Mencari file-file Excel yang tersedia di folder `datasource` untuk dipilih sebagai sumber data.
*   **`main()`**:
    1.  Menampilkan daftar file sumber data yang tersedia di folder `datasource` dan meminta pengguna memilih satu.
    2.  Meminta pengguna memasukkan nomor Juz atau rentang Juz yang ingin diproses.
    3.  Meminta pengguna menentukan nama folder output (default: `dataoutput`).
    4.  Meminta pengguna memasukkan template nama file output (dengan placeholder `{juz}` dan `{metode}` untuk fleksibilitas). Jika placeholder tidak dimasukkan, skrip akan menambahkannya secara otomatis.
    5.  Meminta label konsep dasar dan tanda tangan digital untuk output.
    6.  Menginisialisasi `PengolahDataJBAH`.
    7.  Untuk setiap Juz dalam daftar yang dipilih:
        *   Melakukan filter data.
        *   Menyimpan data hasil filter ke folder `rawprocess` menggunakan `dataextraction.py`.
        *   Untuk setiap metode (`T111`, `T222`, `T333`):
            *   Menjalankan pipeline pemrosesan (`set_metode`, `proses_1_dan_2`, `proses_3`, `generate_tabel_jb_init`, `generate_tabel_ah_final`).
            *   Mengekspor hasil ke file Excel sesuai template nama dan menerapkan formatting menggunakan `export_and_format_output`.

## Alur Pemrosesan Umum

1.  **Muat Data**: Skrip utama (`rawdata_*.py` atau `userinput.py`) membuat instance `PengolahDataJBAH` dan memuat file Excel sumber.
2.  **Iterasi Juz**: Loop melalui Juz yang ditentukan (semua 30 untuk `rawdata_*.py`, atau pilihan pengguna untuk `userinput.py`).
3.  **Filter**: Filter data untuk Juz `X`.
4.  **Simpan Filter (Opsional)**: Untuk `userinput.py`, data hasil filter disimpan.
5.  **Iterasi Metode**: (Untuk `userinput.py` dan simulasi di `backend1.py`) atau gunakan `T111` sebagai standar (untuk `rawdata_*.py`).
6.  **Atur Metode**: Gunakan metode (`T111`, `T222`, `T333`) untuk menggeser poros data.
7.  **Proses**: Jalankan `proses_1_dan_2` dan `proses_3` untuk mendapatkan `df_p3` dan komponen lainnya.
8.  **Generate Output**: Buat `df_init` dan `df_final` (Tabel AH).
9.  **Ekspor & Format**: Simpan DataFrame ke file Excel dan terapkan styling menggunakan `outputproc.py`.
10. **Kumpulkan (untuk `rawdata_*.py`)**: Tambahkan `df_p3` dari Juz `X` ke dalam daftar.
11. **Gabung & Ekspor (untuk `rawdata_*.py`)**: Setelah semua Juz diproses, gabungkan semua `df_p3` dan simpan ke satu file Excel.

## Catatan Penting

*   **Ketergantungan**: Program memerlukan library `pandas`, `numpy`, dan `openpyxl`.
*   **Struktur Data Input**: Program mengharapkan file Excel sumber memiliki kolom-kolom tertentu seperti `aq_juz_id`, `aq_surat_id`, `surat_ayat_no`.
*   **Logika Tabel AH**: Pembuatan Tabel AH bergantung pada indeks tetap tertentu dari `df_p3`. Perubahan dalam jumlah total ayat per Juz atau logika `proses_1_dan_2`/`proses_3` dapat mempengaruhi validitas Tabel AH yang dihasilkan.
*   **Keberadaan `backend2.py`**: File `rawdata_ah_hrfain_alljuz.py` merujuk pada `backend2.py`, yang harus ada dan berfungsi seperti `backend1.py` agar skrip tersebut dapat berjalan.
*   **Input Pengguna (`userinput.py`)**: Memberikan fleksibilitas tinggi dalam menjalankan proses untuk Juz dan metode tertentu dengan konfigurasi output yang dapat disesuaikan.
```