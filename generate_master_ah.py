"""
generate_master_ah.py
=====================
Skrip utilitas untuk membangun "Master Tabel Asmaul Husna" (flat-table)
dari seluruh 30 Juz × 3 Metode (T111, T222, T333).

Fitur:
  - Pemilihan file sumber data secara interaktif dari folder 'datasource'
  - Pemrosesan 30 Juz × 3 Metode
  - Output: file CSV datar (flat) berisi kolom:
    juz | metode | huruf_latin | huruf_arab | ayat_1 | ayat_2 | ayat_3 | ayat_4

Catatan Penting:
  - Kolom ayat DIPAKSA sebagai STRING untuk mencegah konversi "2.50" -> 2.5
  - File CSV ini menjadi "sumber kebenaran tunggal" untuk query_ah.py

Cara pakai:
  python generate_master_ah.py
"""

import os
import gc
import sys
import pandas as pd
from backend1 import PengolahDataJBAH


# ============================================================
# HELPER: Pastikan nilai ayat adalah string yang aman
# ============================================================
def safe_ayat(val):
    """
    Mengonversi nilai ayat menjadi string bersih.
    Mencegah NaN, None, atau float yang tidak diinginkan.
    """
    if pd.isna(val) or val is None:
        return ""
    return str(val).strip()


# ============================================================
# UTILITAS: List file di folder datasource
# ============================================================
def list_available_datasource(script_dir):
    """
    Scan folder 'datasource' dan kembalikan daftar file Excel yang tersedia.
    Returns: list of tuples (nama_tanpa_ekstensi, nama_lengkap)
    """
    datasource_dir = os.path.join(script_dir, "datasource")
    available = []
    if os.path.isdir(datasource_dir):
        for f in os.listdir(datasource_dir):
            if f.lower().endswith(('.xlsx', '.xls')):
                name_without_ext = os.path.splitext(f)[0]
                available.append((name_without_ext, f))
    return available


# ============================================================
# FUNGSI UTAMA
# ============================================================
def generate_master_ah():
    """
    Memproses seluruh 30 Juz × 3 Metode, mengekstrak data inti dari df_final,
    dan menyimpannya ke dalam satu file CSV master (flat-table).
    """
    print("=" * 60)
    print(" MEMULAI GENERASI MASTER TABEL ASMAUL HUSNA (FLAT-TABLE)")
    print("=" * 60)

    # --- Konfigurasi Dasar ---
    script_dir = os.path.dirname(os.path.abspath(__file__))
    datasource_dir = os.path.join(script_dir, "datasource")
    KOLOM_FILTER = "aq_juz_id"
    DAFTAR_METODE = ["T111", "T222", "T333"]

    # ========================================================
    # TAHAP 1: Pemilihan File Sumber Data (Interaktif)
    # ========================================================
    print("\n--- PILIH FILE SUMBER DATA ---")
    available_files = list_available_datasource(script_dir)

    if not available_files:
        print(f"[ERROR] Tidak ada file Excel (.xlsx/.xls) di folder: {datasource_dir}")
        print(f"        Silakan letakkan file data sumber Anda di folder tersebut.")
        sys.exit(1)

    # Cari file default (prioritas: Struktur Quran18 > rawahaq_expanded > rawahaq)
    default_idx = None
    prioritas_default = ['struktur quran18', 'rawahaq_expanded', 'rawahaq']
    for idx, (name_no_ext, full_name) in enumerate(available_files):
        if name_no_ext.lower() in prioritas_default:
            default_idx = idx + 1  # 1-based
            break

    if default_idx is None:
        default_idx = 1  # fallback ke file pertama

    print(f"File Excel yang tersedia di folder 'datasource':")
    for idx, (name_no_ext, full_name) in enumerate(available_files, 1):
        marker = " [default]" if idx == default_idx else ""
        print(f"  {idx}. {full_name}{marker}")

    while True:
        choice = input(f"\nMasukkan nomor pilihan Anda (default: {default_idx}): ").strip()
        if not choice:
            choice = str(default_idx)

        try:
            choice_num = int(choice)
            if 1 <= choice_num <= len(available_files):
                name_no_ext, full_name = available_files[choice_num - 1]
                FILE_MENTAH = os.path.join(datasource_dir, full_name)
                print(f"\n[INFO] Pilihan #{choice_num}: {full_name}")
                print(f"[INFO] Menggunakan file sumber: {FILE_MENTAH}")
                break
            else:
                print(f"[ERROR] Nomor tidak valid. Masukkan angka 1-{len(available_files)}.")
        except ValueError:
            print(f"[ERROR] Masukkan hanya ANGKA (1-{len(available_files)}).")

    # Nama file output dinamis (mengikuti nama file sumber)
    NAMA_FILE_OUTPUT = os.path.join(script_dir, f"master_tabel_ah_{name_no_ext}.csv")

    # ========================================================
    # TAHAP 2: Inisialisasi Pengolah Data
    # ========================================================
    app = PengolahDataJBAH(FILE_MENTAH)
    if app.raw_data is None:
        print("[FATAL] Gagal memuat data mentah. Proses dihentikan.")
        sys.exit(1)

    # ========================================================
    # TAHAP 3: Iterasi Juz × Metode & Ekstraksi Data
    # ========================================================
    list_row_master = []
    total_berhasil = 0
    total_gagal = 0

    for juz_number in range(1, 31):
        print(f"\n[INFO] === Memproses Juz {juz_number} ===")

        # Filter data untuk Juz saat ini
        if app.filter_data(KOLOM_FILTER, juz_number) is None:
            print(f"[WARN] Tidak ada data untuk Juz {juz_number}. Dilewati.")
            total_gagal += len(DAFTAR_METODE)
            continue

        for metode in DAFTAR_METODE:
            try:
                # Jalankan pipeline pemrosesan JBAH
                app.set_metode(metode)
                app.proses_1_dan_2()
                app.proses_3()
                app.generate_tabel_ah_final()

                # Ekstrak data inti dari df_final
                # --------------------------------------------------
                # Struktur df_final (dari backend1.py):
                #   Baris 0 : Konsep Dasar (header)
                #   Baris 1 : Metode & Juz (header)
                #   Baris 2 : Sub-judul (header)
                #   Baris 3 : Nama kolom (Latin, No AH, Ayat#3, ..., Arab)
                #   Baris 4-36 : Data 33 huruf AH
                #
                # Mapping kolom df_final -> flat-table:
                #   indeks 0 (Latin)   -> huruf_latin
                #   indeks 2 (Ayat #3) -> ayat_3
                #   indeks 4 (Ayat #4) -> ayat_4
                #   indeks 6 (Ayat #2) -> ayat_2
                #   indeks 8 (Ayat #1) -> ayat_1
                #   indeks 9 (Arab)    -> huruf_arab
                #
                # Kolom "No AH" (indeks 1,3,5,7) TIDAK dimasukkan
                # karena tidak dibutuhkan untuk query.
                # --------------------------------------------------
                df_final = app.df_final
                if df_final is None or len(df_final) < 5:
                    print(f"[WARN] df_final Juz {juz_number} {metode} tidak valid. Dilewati.")
                    total_gagal += 1
                    continue

                # Ambil baris data saja (lewati 4 baris header)
                data_rows = df_final.iloc[4:]

                for _, row in data_rows.iterrows():
                    list_row_master.append({
                        'juz': juz_number,
                        'metode': metode,
                        'huruf_latin': safe_ayat(row.iloc[0]),
                        'huruf_arab': safe_ayat(row.iloc[9]),
                        'ayat_1': safe_ayat(row.iloc[8]),  # Ayat #1
                        'ayat_2': safe_ayat(row.iloc[6]),  # Ayat #2
                        'ayat_3': safe_ayat(row.iloc[2]),  # Ayat #3
                        'ayat_4': safe_ayat(row.iloc[4]),  # Ayat #4
                    })

                total_berhasil += 1
                print(f"[OK] Juz {juz_number} - {metode} berhasil diekstrak.")

            except Exception as e:
                print(f"[ERROR] Gagal memproses Juz {juz_number} - {metode}: {e}")
                total_gagal += 1

        # Bersihkan memori antar-Juz (penting untuk 30 iterasi)
        gc.collect()

    # ========================================================
    # TAHAP 4: Kompilasi dan Ekspor ke CSV
    # ========================================================
    if not list_row_master:
        print("\n[END] Tidak ada data yang berhasil diproses. File CSV tidak dibuat.")
        return

    print(f"\n[INFO] Menggabungkan {len(list_row_master)} baris data...")
    df_master = pd.DataFrame(list_row_master)

    # Urutkan agar rapi: juz -> metode -> huruf_latin
    df_master = df_master.sort_values(
        by=['juz', 'metode', 'huruf_latin'],
        ignore_index=True
    )

    print(f"[INFO] Menyimpan hasil ke file CSV: {NAMA_FILE_OUTPUT}")
    try:
        df_master.to_csv(
            NAMA_FILE_OUTPUT,
            index=False,
            encoding='utf-8',
            float_format='%s'  # Cegah konversi float saat menulis
        )
        print(f"\n{'=' * 60}")
        print(f"[SUCCESS] Master Tabel AH berhasil dibuat!")
        print(f"  - Total baris     : {len(df_master)}")
        print(f"  - Total kombinasi : {total_berhasil} berhasil, {total_gagal} gagal")
        print(f"  - Lokasi file     : {NAMA_FILE_OUTPUT}")
        print(f"{'=' * 60}")
    except Exception as e:
        print(f"\n[ERROR] Gagal menyimpan file CSV: {e}")


# ============================================================
# ENTRY POINT
# ============================================================
if __name__ == "__main__":
    generate_master_ah()