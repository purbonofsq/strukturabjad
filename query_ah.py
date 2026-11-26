"""
query_ah.py
===========
Modul query untuk mengambil ayat-ayat dari Master Tabel Asmaul Husna (flat-table).

Fitur utama:
  - Fungsi `ambil_ayat_dari_ah(...)` yang fleksibel:
    * User menentukan ayat apa saja yang diambil (ayat_1 s/d ayat_4)
    * User menentukan urutan/susunan ayat sesuai keinginan
  - Validasi ketat: nama kolom, jumlah maksimal 4, keberadaan data
  - Penanganan error yang informatif (cetak pesan, return list kosong)
  - Kolom ayat DIPAKSA sebagai STRING saat load CSV (mencegah "2.50" -> 2.5)

Cara pakai (sebagai library):
  >>> from query_ah import load_master_ah, ambil_ayat_dari_ah
  >>> df = load_master_ah("master_tabel_ah_Struktur_Quran18.csv")
  >>> hasil = ambil_ayat_dari_ah(
  ...     juz=16, metode="T111", huruf_arab="ب",
  ...     daftar_kolom_ayat=['ayat_1', 'ayat_3', 'ayat_4', 'ayat_2'],
  ...     df_master=df
  ... )
  >>> print(hasil)
  ['16.20', '16.18', '16.17', '16.19']

Cara pakai (demo interaktif):
  python query_ah.py
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
    Returns: list of tuples (nama_tanpa_ekstensi, nama_lengkap)
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

    PENTING: Kolom ayat DIPAKSA sebagai STRING (dtype=str) untuk mencegah
    pandas mengubah "2.50" menjadi float 2.5 secara otomatis.

    Args:
        file_path (str): Path lengkap ke file CSV master.

    Returns:
        pd.DataFrame atau None jika gagal.
    """
    if not os.path.exists(file_path):
        print(f"[ERROR] File master tidak ditemukan: {file_path}")
        return None

    try:
        # 🔑 KUNCI UTAMA: Paksa semua kolom sensitif dibaca sebagai string
        # Ini mencegah pandas melakukan inferensi tipe data otomatis
        # yang bisa mengubah "2.50" menjadi float 2.5
        dtype_paksa_string = {
            'huruf_latin': str,
            'huruf_arab': str,
            'metode': str,
            'ayat_1': str,
            'ayat_2': str,
            'ayat_3': str,
            'ayat_4': str,
        }

        df = pd.read_csv(file_path, encoding='utf-8', dtype=dtype_paksa_string)

        # Validasi kolom wajib
        kolom_wajib = {'juz', 'metode', 'huruf_arab', 'ayat_1', 'ayat_2', 'ayat_3', 'ayat_4'}
        if not kolom_wajib.issubset(set(df.columns)):
            missing = kolom_wajib - set(df.columns)
            print(f"[ERROR] File master tidak memiliki kolom wajib: {missing}")
            return None

        # Pastikan tipe data juz konsisten (integer)
        df['juz'] = df['juz'].astype(int)
        df['metode'] = df['metode'].astype(str).str.upper().str.strip()

        # Bersihkan nilai NaN/None di kolom ayat menjadi string kosong
        for kol in ['ayat_1', 'ayat_2', 'ayat_3', 'ayat_4']:
            df[kol] = df[kol].fillna('').astype(str).str.strip()

        print(f"[OK] Master CSV berhasil dimuat: {len(df)} baris dari {file_path}")
        print(f"     Kolom ayat dipaksa sebagai STRING (aman dari konversi desimal)")
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

    Args:
        juz (int): Nomor Juz (1-30).
        metode (str): Metode pemrosesan ('T111', 'T222', atau 'T333').
        huruf_arab (str): Huruf Arab yang dicari (misal: 'ب', 'ت', 'ا').
        daftar_kolom_ayat (list): Daftar nama kolom ayat yang ingin diambil,
                                  beserta urutannya. Contoh:
                                  ['ayat_3'] -> 1 ayat saja
                                  ['ayat_1', 'ayat_2'] -> 2 ayat berurutan
                                  ['ayat_1', 'ayat_3', 'ayat_4', 'ayat_2'] -> 4 ayat kustom
        df_master (pd.DataFrame): DataFrame master yang sudah dimuat via load_master_ah().

    Returns:
        list: Daftar ayat dalam format string (misal: ['16.20', '16.18']).
              List kosong [] jika terjadi error atau data tidak ditemukan.

    Contoh:
        >>> ambil_ayat_dari_ah(16, 'T111', 'ب', ['ayat_1'], df)
        ['16.20']

        >>> ambil_ayat_dari_ah(16, 'T111', 'ب', ['ayat_1', 'ayat_3', 'ayat_4', 'ayat_2'], df)
        ['16.20', '16.18', '16.17', '16.19']
    """

    # --------------------------------------------------------
    # VALIDASI 1: Parameter daftar_kolom_ayat
    # --------------------------------------------------------
    if not isinstance(daftar_kolom_ayat, list):
        print(f"[ERROR] Parameter 'daftar_kolom_ayat' harus berupa list. "
              f"Diterima: {type(daftar_kolom_ayat).__name__}")
        return []

    if len(daftar_kolom_ayat) == 0:
        print("[ERROR] Parameter 'daftar_kolom_ayat' tidak boleh kosong.")
        return []

    if len(daftar_kolom_ayat) > 4:
        print(f"[ERROR] Jumlah ayat maksimal 4. Anda meminta {len(daftar_kolom_ayat)} ayat.")
        return []

    for kolom in daftar_kolom_ayat:
        if kolom not in KOLOM_AYAT_VALID:
            print(f"[ERROR] Kolom '{kolom}' tidak dikenali. "
                  f"Gunakan salah satu dari: {sorted(KOLOM_AYAT_VALID)}")
            return []

    # --------------------------------------------------------
    # VALIDASI 2: Parameter juz
    # --------------------------------------------------------
    if not isinstance(juz, int) or not (1 <= juz <= 30):
        print(f"[ERROR] Parameter 'juz' harus integer antara 1-30. Diterima: {juz}")
        return []

    # --------------------------------------------------------
    # VALIDASI 3: Parameter metode
    # --------------------------------------------------------
    metode_upper = str(metode).upper().strip()
    if metode_upper not in METODE_VALID:
        print(f"[ERROR] Parameter 'metode' harus salah satu dari {sorted(METODE_VALID)}. "
              f"Diterima: '{metode}'")
        return []

    # --------------------------------------------------------
    # VALIDASI 4: Parameter huruf_arab
    # --------------------------------------------------------
    if not isinstance(huruf_arab, str) or len(huruf_arab.strip()) == 0:
        print(f"[ERROR] Parameter 'huruf_arab' harus string tidak kosong. "
              f"Diterima: '{huruf_arab}'")
        return []
    huruf_arab = huruf_arab.strip()

    # --------------------------------------------------------
    # VALIDASI 5: DataFrame master
    # --------------------------------------------------------
    if df_master is None or df_master.empty:
        print("[ERROR] DataFrame master kosong atau None.")
        return []

    # --------------------------------------------------------
    # PENCARIAN DATA
    # --------------------------------------------------------
    try:
        mask = (
            (df_master['juz'] == juz) &
            (df_master['metode'] == metode_upper) &
            (df_master['huruf_arab'] == huruf_arab)
        )
        hasil_filter = df_master[mask]

        if hasil_filter.empty:
            print(f"[ERROR] Data tidak ditemukan untuk kombinasi:")
            print(f"        Juz     = {juz}")
            print(f"        Metode  = {metode_upper}")
            print(f"        Huruf   = '{huruf_arab}'")
            print(f"        Silakan periksa kembali parameter Anda.")
            return []

        # Ambil baris pertama (seharusnya hanya 1 baris per kombinasi unik)
        baris_data = hasil_filter.iloc[0]

        # --------------------------------------------------------
        # EKSTRAK AYAT SESUAI URUTAN YANG DIMINTA USER
        # --------------------------------------------------------
        daftar_ayat = []
        for kolom in daftar_kolom_ayat:
            nilai = baris_data[kolom]
            # Tangani nilai kosong/NaN (seharusnya sudah dibersihkan di load,
            # tapi double-check untuk keamanan)
            if pd.isna(nilai) or nilai is None or str(nilai).strip() == '':
                daftar_ayat.append("")
            else:
                daftar_ayat.append(str(nilai).strip())

        return daftar_ayat

    except Exception as e:
        print(f"[ERROR] Terjadi kesalahan saat mengambil data: {e}")
        import traceback
        traceback.print_exc()
        return []


# ============================================================
# DEMO INTERAKTIF
# ============================================================
def demo():
    """
    Antarmuka interaktif untuk menguji fungsi ambil_ayat_dari_ah().
    """
    print("=" * 60)
    print(" DEMO QUERY TABEL ASMAUL HUSNA")
    print("=" * 60)

    script_dir = os.path.dirname(os.path.abspath(__file__))

    # --- Pilih file master ---
    print("\n--- PILIH FILE MASTER CSV ---")
    available_masters = list_available_master(script_dir)

    if not available_masters:
        print(f"[ERROR] Tidak ada file master CSV (master_tabel_ah_*.csv) "
              f"di folder: {script_dir}")
        print("        Silakan jalankan 'generate_master_ah.py' terlebih dahulu.")
        return

    print("File master CSV yang tersedia:")
    for idx, (name_no_ext, full_name) in enumerate(available_masters, 1):
        print(f"  {idx}. {full_name}")

    while True:
        choice = input(f"\nMasukkan nomor pilihan (1-{len(available_masters)}): ").strip()
        try:
            choice_num = int(choice)
            if 1 <= choice_num <= len(available_masters):
                _, full_name = available_masters[choice_num - 1]
                file_master = os.path.join(script_dir, full_name)
                break
            else:
                print(f"[ERROR] Nomor tidak valid.")
        except ValueError:
            print("[ERROR] Masukkan hanya ANGKA.")

    # --- Load master ---
    df_master = load_master_ah(file_master)
    if df_master is None:
        return

    # --- Sesi query interaktif ---
    print("\n" + "=" * 60)
    print(" SESI QUERY (ketik 'q' untuk keluar)")
    print("=" * 60)

    while True:
        print("\n--- INPUT PARAMETER QUERY ---")

        # Juz
        juz_input = input("Masukkan nomor Juz (1-30) [q=quit]: ").strip()
        if juz_input.lower() == 'q':
            break
        try:
            juz = int(juz_input)
        except ValueError:
            print("[ERROR] Juz harus angka. Coba lagi.")
            continue

        # Metode
        metode_input = input("Masukkan metode (T111/T222/T333) [default: T111]: ").strip()
        metode = metode_input if metode_input else "T111"

        # Huruf Arab
        huruf_arab = input("Masukkan huruf Arab (contoh: ب): ").strip()
        if not huruf_arab:
            print("[ERROR] Huruf Arab tidak boleh kosong. Coba lagi.")
            continue

        # Daftar kolom ayat
        print("\n  Kolom yang tersedia: ayat_1, ayat_2, ayat_3, ayat_4")
        print("  Contoh input: ayat_1,ayat_3,ayat_4,ayat_2")
        kolom_input = input("  Masukkan daftar kolom (pisahkan koma): ").strip()
        if not kolom_input:
            print("[ERROR] Daftar kolom tidak boleh kosong. Coba lagi.")
            continue

        daftar_kolom = [k.strip() for k in kolom_input.split(',')]

        # --- Panggil fungsi core ---
        print("\n" + "-" * 40)
        print(f"Query: Juz={juz}, Metode={metode.upper()}, Huruf='{huruf_arab}'")
        print(f"Kolom diminta: {daftar_kolom}")
        print("-" * 40)

        hasil = ambil_ayat_dari_ah(
            juz=juz,
            metode=metode,
            huruf_arab=huruf_arab,
            daftar_kolom_ayat=daftar_kolom,
            df_master=df_master
        )

        if hasil:
            print(f"\n[HASIL] ({len(hasil)} ayat):")
            for i, ayat in enumerate(hasil, 1):
                label = daftar_kolom[i - 1] if i - 1 < len(daftar_kolom) else "?"
                print(f"  {i}. [{label}] = {ayat}")
        else:
            print("\n[HASIL] Tidak ada ayat yang berhasil diambil.")

    print("\n[INFO] Sesi query selesai. Terima kasih!")


# ============================================================
# ENTRY POINT
# ============================================================
if __name__ == "__main__":
    demo()