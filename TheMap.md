Optimized tool selectionSaya akan meninjau file-file utama di workspace dan merangkum hubungan antar programnya agar peta fungsinya jelas.

Read [](file:///c%3A/Users/purbo/OneDrive/Documents/GitHub/purbonofsq/strukturabjad)

Read [](file:///c%3A/Users/purbo/OneDrive/Documents/GitHub/purbonofsq/strukturabjad/strukturabjad/README.md)

Read [](file:///c%3A/Users/purbo/OneDrive/Documents/GitHub/purbonofsq/strukturabjad/README.md)

Saya sedang mengidentifikasi file-file inti dan alur pemrosesan mereka untuk membuat peta program yang ringkas dan mudah dipahami.

Read [](file:///c%3A/Users/purbo/OneDrive/Documents/GitHub/purbonofsq/strukturabjad/backend_jbah.py#1-1), lines 1 to 250

Read [](file:///c%3A/Users/purbo/OneDrive/Documents/GitHub/purbonofsq/strukturabjad/rekon_excel.py)

Read [](file:///c%3A/Users/purbo/OneDrive/Documents/GitHub/purbonofsq/strukturabjad/strukturabjad/backend1.py#1-1), lines 1 to 260

Read [](file:///c%3A/Users/purbo/OneDrive/Documents/GitHub/purbonofsq/strukturabjad/strukturabjad/backend2.py#1-1), lines 1 to 260

Read [](file:///c%3A/Users/purbo/OneDrive/Documents/GitHub/purbonofsq/strukturabjad/strukturabjad/dataextraction.py)

## Peta mapping program di workspace

Berikut pemetaan singkat dari program-program yang ada:

| File | Peran utama | Hubungan dengan program lain |
|---|---|---|
| backend_jbah.py | Versi awal/standalone untuk memproses data JBAH. | Ini adalah fondasi logika pemrosesan sebelum ada wrapper interaktif. |
| rekon_excel.py | Utility untuk memperluas data berdasarkan kolom “Bobot” ke baris-baris yang lebih detail. | Biasanya dipakai sebagai pra-pemrosesan data sebelum masuk ke pipeline utama. |
| backend1.py | Engine utama pemrosesan data. Berisi kelas `PengolahDataJBAH` dan formatter Excel. | Ini adalah inti dari alur kerja: filter data, mengatur metode T111/T222/T333, menghasilkan sheet Init/Proses/Final, lalu ekspor Excel. |
| backend2.py | Versi/backup logic pemrosesan yang mirip dengan backend1. | Fungsinya hampir sama, tetapi sebagai alternatif atau cadangan. |
| dataextraction.py | Menyimpan data hasil filter ke file Excel di folder rawprocess. | Dipakai setelah data difilter sebelum diproses lebih lanjut. |
| userinput.py | Entry point interaktif. | File ini menerima input pengguna, memilih file sumber, juz, template output, lalu memanggil proses utama. |

## Alur kerja yang paling umum

1. userinput.py menerima input pengguna.
2. Data sumber dari folder datasource dibaca.
3. Data difilter berdasarkan Juz lewat dataextraction.py.
4. Proses utama dijalankan oleh backend1.py atau backend_jbah.py.
5. Hasil akhir diekspor ke Excel dalam bentuk file output.

## Inti singkat

- userinput.py = “tampilan depan”
- backend1.py = “otak pemrosesan”
- dataextraction.py = “penyimpan data hasil filter”
- rekon_excel.py = “alat bantu persiapan data”

Kalau Anda mau, saya bisa lanjut buatkan “diagram alur” versi visual atau “mapping per folder” supaya lebih mudah dipahami.