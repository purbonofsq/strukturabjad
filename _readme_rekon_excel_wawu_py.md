_readme_rekon_excel_wawu.py

# Rekon Excel: Ekspansi Data per Huruf Wau

Program Python untuk memperluas data Excel yang tiap barisnya mewakili **satu ayat** menjadi data yang tiap barisnya mewakili **satu huruf wau**. Jumlah baris hasil untuk setiap ayat ditentukan oleh kolom `Bobot`, yaitu jumlah huruf wau pada ayat tersebut.

## Fitur

- Memilih file sumber dari daftar bernomor di folder `datasource`, tanpa mengetik nama file sehingga tidak ada salah tulis.
- Menggandakan setiap baris sesuai nilai `Bobot`.
- Menambahkan ID unik untuk setiap huruf wau, ditambah kolom bantu untuk menelusuri data ke baris aslinya.
- Memvalidasi nilai `Bobot`: nilai kosong, bukan angka, atau negatif dilewati disertai peringatan.
- Menyimpan hasil ke folder `output`, terpisah dari data sumber.

## Kebutuhan

- Python 3.8 atau lebih baru
- Library `pandas` dan `openpyxl` (untuk membaca dan menulis `.xlsx`)
- Library `xlrd` (hanya jika memakai file `.xls`)

```bash
pip install pandas openpyxl
```

## Struktur Folder

Letakkan `rekon_excel.py` dan folder `datasource` di lokasi yang sama:

```
proyek/
├── rekon_excel.py
├── datasource/          <- taruh file Excel sumber di sini
│   ├── rawahaq.xlsx
│   └── data_lain.xlsx
└── output/              <- dibuat otomatis, berisi hasil
    └── rawahaq_expanded.xlsx
```

Folder `datasource` harus dibuat sendiri. Folder `output` dibuat otomatis oleh program.

## Format File Sumber

- Baris pertama berisi **header** (nama kolom).
- Wajib ada kolom bernama **`Bobot`**, berisi bilangan bulat: jumlah huruf wau pada ayat di baris tersebut.
- Kolom lain bebas, dan semuanya ikut tersalin ke hasil.

Contoh:

| Surah | Ayat | Teks | Bobot |
|---|---|---|---|
| Al-Fatihah | 1 | ... | 3 |
| Al-Fatihah | 2 | ... | 2 |

## Cara Penggunaan

1. Salin file Excel sumber ke folder `datasource`.
2. Tutup file tersebut di Excel jika sedang terbuka.
3. Jalankan program:

   ```bash
   python rekon_excel.py
   ```

4. Pilih file dengan mengetik nomornya, lalu tekan Enter:

   ```
   File yang tersedia di folder 'datasource':
     1. data_lain.xlsx
     2. rawahaq.xlsx
     0. Batal
   Pilih nomor file [1-2]: 2
   ```

5. Tentukan nama file output. Tekan Enter untuk memakai nama otomatis, atau ketik nama sendiri (ekstensi `.xlsx` ditambahkan otomatis bila belum ada):

   ```
   Nama file output [default: rawahaq_expanded.xlsx]:
   ```

6. Setelah selesai, program menampilkan ringkasan:

   ```
   Baris awal  : 6236
   Baris hasil : 2814
   Data berhasil disimpan ke '.../output/rawahaq_expanded.xlsx'
   ```

   Angka di atas hanya contoh tampilan.

## Format Hasil

Semua kolom asli dipertahankan, lalu ditambahkan tiga kolom di bagian akhir:

| Kolom | Isi |
|---|---|
| `id_huruf_wau` | ID unik tiap huruf wau, berurutan dari 1 di seluruh data |
| `id_baris_asal` | Nomor baris ayat di file sumber, sesuai nomor baris Excel (header = baris 1) |
| `urutan_wau` | Urutan wau di dalam satu ayat (1, 2, 3, ...), dimulai ulang tiap ayat |

Contoh: ayat di baris 2 dengan `Bobot` = 3 dan ayat di baris 3 dengan `Bobot` = 2 menghasilkan:

| ... | id_huruf_wau | id_baris_asal | urutan_wau |
|---|---|---|---|
| ... | 1 | 2 | 1 |
| ... | 2 | 2 | 2 |
| ... | 3 | 2 | 3 |
| ... | 4 | 3 | 1 |
| ... | 5 | 3 | 2 |

Ayat dengan `Bobot` = 0 tidak menghasilkan baris.

## Konfigurasi

Pengaturan ada di bagian atas `rekon_excel.py`:

| Konstanta | Fungsi | Nilai awal |
|---|---|---|
| `FOLDER_SUMBER` | Nama folder file sumber | `datasource` |
| `FOLDER_OUTPUT` | Nama folder hasil | `output` |
| `KOLOM_BOBOT` | Nama kolom bobot di file sumber | `Bobot` |
| `KOLOM_ID` | Nama kolom ID unik | `id_huruf_wau` |
| `KOLOM_BARIS_ASAL` | Nama kolom nomor baris asal | `id_baris_asal` |
| `KOLOM_URUTAN` | Nama kolom urutan wau dalam ayat | `urutan_wau` |
| `TAMBAH_KOLOM_BANTU` | `False` jika hanya ingin kolom `id_huruf_wau` | `True` |

## Penanganan Masalah

| Pesan / Gejala | Penyebab dan solusi |
|---|---|
| `Tidak ada file yang dipilih / tidak ditemukan file Excel...` | Folder `datasource` belum ada atau kosong, atau Anda memilih `0`. Buat folder dan isi dengan file `.xlsx`. |
| `Kolom 'Bobot' tidak ditemukan...` | Nama kolom tidak persis `Bobot` (huruf besar/kecil dan spasi berpengaruh). Perbaiki di Excel atau ubah `KOLOM_BOBOT`. |
| `Peringatan: Bobot tidak valid pada baris Excel [...]` | Baris tersebut kosong, bukan angka, atau negatif, sehingga dilewati. Perbaiki nilainya jika baris itu seharusnya ikut. |
| `File output '...' sedang dibuka` | File hasil dengan nama yang sama masih terbuka di Excel. Tutup lalu jalankan ulang. |
| `ModuleNotFoundError` | Library belum terpasang. Jalankan `pip install pandas openpyxl`. |

## Catatan

- Jika file output dengan nama yang sama sudah ada di folder `output`, file tersebut akan **ditimpa**.
- File sementara Excel (nama diawali `~$`) tidak ditampilkan di daftar pilihan.
- Bobot berbentuk desimal (misalnya 2,7) dibulatkan ke bawah menjadi bilangan bulat.