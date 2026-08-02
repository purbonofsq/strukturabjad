import os
import pandas as pd
from backend2 import PengolahDataJBAH

def generate_raw_data_all_juz():
    """
    This script processes data for all 30 Juz using the JBAH method,
    extracts the flattened data from 'proses_3' for each Juz,
    and compiles it into a single Excel file.
    """
    print("--- MEMULAI GENERASI RAW DATA TABEL AH (ALL JUZ) ---")

    # --- Konfigurasi Awal ---
    script_dir = os.path.dirname(os.path.abspath(__file__))
    FILE_MENTAH = os.path.join(script_dir, "datasource", "rawahaq_expanded.xlsx")
    KOLOM_FILTER = "aq_juz_id"
    METODE_UTAMA = "T111"  # Menggunakan T111 sebagai metode standar
    NAMA_FILE_OUTPUT = os.path.join(script_dir, "rawdatatabel_ah_hrfain_alljuz.xlsx")

    # --- Inisialisasi Pengolah Data ---
    app = PengolahDataJBAH(FILE_MENTAH)
    if app.raw_data is None:
        print("[FATAL] Gagal memuat data mentah. Proses dihentikan.")
        return

    # --- Proses Iterasi per Juz ---
    list_df_p3_all_juz = []
    for juz_number in range(1, 31):
        print(f"\n[INFO] Memproses Juz {juz_number}...")

        # 1. Filter data untuk Juz saat ini
        if app.filter_data(KOLOM_FILTER, juz_number) is None:
            print(f"[WARN] Tidak ada data untuk Juz {juz_number} atau terjadi error. Melanjutkan ke Juz berikutnya.")
            continue

        # 2. Jalankan pipeline pemrosesan
        app.set_metode(METODE_UTAMA)
        app.proses_1_dan_2()
        app.proses_3()

        # 3. Ambil hasil dari df_p3 dan tambahkan kolom ah_juz_id
        df_p3_current = app.df_p3.copy()
        df_p3_current['ah_juz_id'] = juz_number
        list_df_p3_all_juz.append(df_p3_current)
        print(f"[OK] Juz {juz_number} berhasil diproses.")

    # --- Kompilasi dan Ekspor Hasil ---
    if not list_df_p3_all_juz:
        print("\n[END] Tidak ada data yang berhasil diproses. File Excel tidak dibuat.")
        return

    print("\n[INFO] Menggabungkan hasil dari semua Juz...")
    final_df = pd.concat(list_df_p3_all_juz, ignore_index=True)

    # Reorder kolom agar 'ah_juz_id' di depan
    final_df = final_df[['ah_juz_id', 'p3_id', 'Label', 'Ayat']]

    print(f"[INFO] Menyimpan hasil ke file Excel: {NAMA_FILE_OUTPUT}")
    try:
        with pd.ExcelWriter(NAMA_FILE_OUTPUT, engine='openpyxl') as writer:
            final_df.to_excel(writer, sheet_name='Raw_Data_Proses_3_All_Juz', index=False)
        print(f"\n[SUCCESS] File '{NAMA_FILE_OUTPUT}' berhasil dibuat.")
    except Exception as e:
        print(f"\n[ERROR] Gagal menyimpan file Excel: {e}")

if __name__ == "__main__":
    generate_raw_data_all_juz()