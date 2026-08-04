# Rencana kerja berikutnya — pin-point

### 1) Selesaikan fitur yang sedang ditambah
- fokus pada fitur yang masih dikerjakan
- jangan ubah struktur besar dulu
- pastikan fungsi lama tetap jalan

### 2) Stabilkan alur utama
- validasi jalur input data
- validasi filter dan metode
- validasi output Excel
- pastikan hasil untuk kasus normal masih benar

### 3) Setelah fitur stabil, baru cleanup kecil
- pisahkan kode yang masih campur di satu file
- kecilkan fungsi yang terlalu panjang
- hilangkan duplikasi logika
- jaga perubahan tetap lokal

### 4) Perbaiki kontrak data antar modul
- input/output tiap fungsi harus jelas
- hindari passing objek besar yang berisi terlalu banyak state
- pastikan tipe data lebih konsisten

### 5) Lakukan refactor bersih hanya saat fitur berhenti berubah
- setelah semua fitur baru selesai
- ketika alur utama sudah stabil
- baru struktur modul dibenahi secara menyeluruh

### 6) Prioritas yang benar
- fitur baru dulu
- stabilisasi alur
- lalu cleanup arsitektur
- terakhir: optimasi dan produksi-ready

> Jadi, pendekatannya: “fitur-terlebih, refactor-belakangan” dengan tetap menjaga kode agar tidak terlalu berantakan.

Kalau mau, saya bisa lanjutkan dengan versi yang lebih formal seperti:
- “Checklist prioritas sprint”
- “Rencana 3 tahap”
- atau “Roadmap pengembangan sampai siap produksi”

## Checklist prioritas sprint

### Sprint prioritas saat ini
- [ ] Selesaikan fitur yang sedang ditambahkan
- [ ] Pastikan alur utama tetap bekerja
- [ ] Validasi input data dan filter
- [ ] Validasi output Excel untuk kasus normal
- [ ] Uji kasus batas dan error ringan
- [ ] Setelah stabil, lakukan cleanup kecil pada area terkait
- [ ] Jangan refactor besar sebelum fitur utama selesai

### Fokus teknis
- [ ] Cek apakah modul input, pipeline, dan output masih sinkron
- [ ] Pastikan data tidak berubah bentuk tiap tahap
- [ ] Pastikan metode dan filter tetap konsisten
- [ ] Pastikan eksport Excel tidak rusak akibat perubahan baru

---

## Rencana 3 tahap

### Tahap 1: Stabilisasi fitur
- fokus pada penambahan fitur aktif
- pastikan semua fungsi lama masih berjalan
- validasi data utama dan output dasar

### Tahap 2: Cleanup terarah
- pisahkan logika yang masih tercampur
- perbaiki nama variabel dan fungsi
- kurangi duplikasi
- perbaiki kontrak antar modul

### Tahap 3: Persiapan produksi
- uji alur utama secara berulang
- buat dokumentasi penggunaan
- batasi manual step
- siap untuk validasi lebih formal sebelum deployment/produksi

---

## Roadmap pengembangan sampai siap produksi

### Fase 1 — Pengembangan aktif
- fitur baru ditambah
- alur utama masih terus diuji
- refactor dibatasi dan terarah

### Fase 2 — Stabilisasi sistem
- validasi semua jalur utama
- data input dan output konsisten
- cleanup kecil pada modul yang paling sering berubah

### Fase 3 — Standarisasi kode
- modul lebih terpisah dan konsisten
- tipe data jelas
- dokumentasi fungsi dan alur kerja lebih rapi

### Fase 4 — Produksi-ready
- semua proses utama teruji
- output valid
- error ditangani dengan jelas
- dokumentasi operasional siap dipakai

> Ringkasnya: project Anda masih berada di fase pengembangan, dan roadmap yang paling tepat adalah “stabilkan fitur dulu, lalu refactor, lalu siap produksi.”

Kalau Anda mau, saya bisa lanjutkan dengan versi yang lebih formal dan siap dipindahkan ke README atau catatan proyek.
