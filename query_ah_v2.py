"""
query_ah.py
===========
Modul query untuk mengambil ayat-ayat dari Master Tabel Asmaul Husna (flat-table).

Fitur utama:
  - Fungsi `ambil_ayat_dari_ah(...)` yang fleksibel
  - Kolom ayat DIPAKSA sebagai STRING saat load CSV (mencegah "2.50" -> 2.5)
"""

import os
import sys
import pandas as pd


# ============================================================
# KONSTANTA
# ============================================================
KOLOM_AYAT_VALID = {'ayat_1', 'ayat_2', 'ayat_3', 'ayat_4'}
METODE_VALID = {'T111', 'T222', 'T333'}


# ============================================================
# UTILITAS: LIST & LOAD MASTER CSV
# ============================================================
def list_available_master(script_dir):
    """
    Scan direktori dan kembalikan daftar file master CSV yang tersedia.
    """
    available = []
    if os.path.isdir(script_dir):
        for f in os.listdir(script_dir):
            if f.lower().startswith('master_tabel_ah_') and f.lower().endswith('.csv'):
                name_without_ext = os.path.splitext(f)[0]
                available.append((name_without_ext, f))
    return available


def load_master_ah(file_path):
    """
    Memuat file master CSV ke dalam DataFrame pandas.
    KOLOM AYAT DIPAKSA SEBAGAI STRING untuk mencegah konversi desimal.
    """
    if not os.path.exists(file_path):
        print(f"[ERROR] File master tidak ditemukan: {file_path}")
        return None

    try:
        # 🔑 KUNCI UTAMA: Baca SEMUA kolom sebagai string terlebih dahulu.
        # Ini mencegah pandas melakukan inferensi tipe data otomatis
        # yang bisa mengubah "2.50" menjadi float 2.5, dan sekaligus 
        # menghilangkan warning Pylance terkait tipe dict pada dtype.
        df = pd.read_csv(file_path, encoding='utf-8', dtype=str)

        # Validasi kolom wajib
        kolom_wajib = {'juz', 'metode', 'huruf_arab', 'ayat_1', 'ayat_2', 'ayat_3', 'ayat_4'}
        if not kolom_wajib.issubset(set(df.columns)):
            missing = kolom_wajib - set(df.columns)
            print(f"[ERROR] File master tidak memiliki kolom wajib: {missing}")
            return None

        # Pastikan tipe data juz konsisten (integer)
        # Karena tadi dibaca sebagai string, kita konversi eksplisit ke int
        df['juz'] = df['juz'].astype(int)
        
        # Bersihkan dan standarisasi kolom metode
        df['metode'] = df['metode'].astype(str).str.upper().str.strip()

        # Bersihkan nilai NaN/None di kolom ayat menjadi string kosong
        for kol in ['ayat_1', 'ayat_2', 'ayat_3', 'ayat_4']:
            df[kol] = df[kol].fillna('').astype(str).str.strip()

        print(f"[OK] Master CSV berhasil dimuat: {len(df)} baris dari {file_path}")
        print(f"     Semua kolom dibaca sebagai STRING (aman dari konversi desimal)")
        return df

    except Exception as e:
        print(f"[ERROR] Gagal memuat master CSV: {e}")
        import traceback
        traceback.print_exc()
        return None


# ============================================================
# FUNGSI CORE: AMBIL AYAT DARI TABEL AH
# ============================================================
def ambil_ayat_dari_ah(juz: int, metode: str, huruf_arab: str,
                       daftar_kolom_ayat: list, df_master: pd.DataFrame) -> list:
    """
    Mengambil ayat-ayat dari Master Tabel AH berdasarkan parameter yang diberikan.
    """
    # 1. Validasi daftar_kolom_ayat
    if not isinstance(daftar_kolom_ayat, list):
        print(f"[ERROR] Parameter 'daftar_kolom_ayat' harus berupa list.")
        return []
    if len(daftar_kolom_ayat) == 0:
        print("[ERROR] Parameter 'daftar_kolom_ayat' tidak boleh kosong.")
        return []
    if len(daftar_kolom_ayat) > 4:
        print(f"[ERROR] Jumlah ayat maksimal 4. Anda meminta {len(daftar_kolom_ayat)} ayat.")
        return []
    for kolom in daftar_kolom_ayat:
        if kolom not in KOLOM_AYAT_VALID:
            print(f"[ERROR] Kolom '{kolom}' tidak dikenali. Gunakan: {sorted(KOLOM_AYAT_VALID)}")
            return []

    # 2. Validasi juz
    if not isinstance(juz, int) or not (1 <= juz <= 30):
        print(f"[ERROR] Parameter 'juz' harus integer antara 1-30.")
        return []

    # 3. Validasi metode
    metode_upper = str(metode).upper().strip()
    if metode_upper not in METODE_VALID:
        print(f"[ERROR] Parameter 'metode' harus salah satu dari {sorted(METODE_VALID)}.")
        return []

    # 4. Validasi huruf_arab
    if not isinstance(huruf_arab, str) or len(huruf_arab.strip()) == 0:
        print(f"[ERROR] Parameter 'huruf_arab' harus string tidak kosong.")
        return []
    huruf_arab = huruf_arab.strip()

    # 5. Validasi DataFrame
    if df_master is None or df_master.empty:
        print("[ERROR] DataFrame master kosong atau None.")
        return []

    # PENCARIAN DATA
    try:
        mask = (
            (df_master['juz'] == juz) &
            (df_master['metode'] == metode_upper) &
            (df_master['huruf_arab'] == huruf_arab)
        )
        hasil_filter = df_master[mask]

        if hasil_filter.empty:
            print(f"[ERROR] Data tidak ditemukan untuk: Juz={juz}, Metode={metode_upper}, Huruf='{huruf_arab}'")
            return []

        baris_data = hasil_filter.iloc[0]

        # EKSTRAK AYAT SESUAI URUTAN YANG DIMINTA
        daftar_ayat = []
        for kolom in daftar_kolom_ayat:
            nilai = baris_data[kolom]
            if pd.isna(nilai) or nilai is None or str(nilai).strip() == '':
                daftar_ayat.append("")
            else:
                daftar_ayat.append(str(nilai).strip())

        return daftar_ayat

    except Exception as e:
        print(f"[ERROR] Terjadi kesalahan saat mengambil data: {e}")
        return []