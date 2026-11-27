import os
import pandas as pd

FOLDER_SUMBER = 'datasource'
FOLDER_OUTPUT = 'output'
KOLOM_BOBOT = 'Bobot'
KOLOM_ID = 'id_huruf_wau'            # ID unik tiap huruf wau (1, 2, 3, ... di seluruh data)
KOLOM_BARIS_ASAL = 'id_baris_asal'   # nomor baris ayat di file sumber (sesuai nomor baris Excel)
KOLOM_URUTAN = 'urutan_wau'          # wau ke-1, ke-2, dst. di dalam satu ayat
TAMBAH_KOLOM_BANTU = True            # set False jika hanya ingin kolom id_huruf_wau
EKSTENSI = ('.xlsx', '.xls')


def daftar_file_sumber(folder):
    """Mengembalikan daftar file Excel di folder (diurutkan, tanpa file temp ~$)."""
    if not os.path.isdir(folder):
        return []
    return sorted(
        f for f in os.listdir(folder)
        if f.lower().endswith(EKSTENSI) and not f.startswith('~$')
    )


def pilih_file(folder):
    """Menampilkan daftar file dan meminta pengguna memilih berdasarkan nomor."""
    files = daftar_file_sumber(folder)
    if not files:
        return None

    print(f"\nFile yang tersedia di folder '{os.path.basename(folder)}':")
    for i, nama in enumerate(files, start=1):
        print(f"  {i}. {nama}")
    print("  0. Batal")

    while True:
        pilihan = input(f"Pilih nomor file [1-{len(files)}]: ").strip()
        if pilihan == '0':
            return None
        if pilihan.isdigit() and 1 <= int(pilihan) <= len(files):
            return files[int(pilihan) - 1]
        print("Pilihan tidak valid, silakan coba lagi.")


def expand_data(input_path, output_path):
    """
    Membaca file Excel berisi data per ayat, lalu menggandakan setiap baris
    sebanyak nilai kolom 'Bobot' (jumlah huruf wau pada ayat tersebut),
    sehingga tiap huruf wau menjadi satu baris dengan ID unik (id_huruf_wau).
    """
    df = pd.read_excel(input_path, header=0)

    if KOLOM_BOBOT not in df.columns:
        raise ValueError(
            f"Kolom '{KOLOM_BOBOT}' tidak ditemukan. "
            f"Kolom yang ada: {list(df.columns)}"
        )

    # Validasi bobot: non-angka / kosong / negatif dianggap 0 (baris dilewati)
    bobot = pd.to_numeric(df[KOLOM_BOBOT], errors='coerce')
    tidak_valid = bobot.isna() | (bobot < 0)
    if tidak_valid.any():
        baris_excel = (df.index[tidak_valid] + 2).tolist()  # +2: header + index mulai 0
        print(f"Peringatan: Bobot tidak valid pada baris Excel {baris_excel}, dilewati.")
    bobot = bobot.where(~tidak_valid, 0).astype(int)

    # Gandakan baris sesuai bobot (vektorisasi, jauh lebih cepat dari iterrows)
    expanded_df = df.loc[df.index.repeat(bobot)]
    nomor_baris_asal = expanded_df.index + 2  # +2: baris 1 = header, index mulai 0
    expanded_df = expanded_df.reset_index(drop=True)

    # ID unik berurutan mulai dari 1
    expanded_df[KOLOM_ID] = range(1, len(expanded_df) + 1)

    if TAMBAH_KOLOM_BANTU:
        expanded_df[KOLOM_BARIS_ASAL] = nomor_baris_asal
        expanded_df[KOLOM_URUTAN] = expanded_df.groupby(KOLOM_BARIS_ASAL).cumcount() + 1

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    expanded_df.to_excel(output_path, index=False)

    print(f"Baris awal  : {len(df)}")
    print(f"Baris hasil : {len(expanded_df)}")
    print(f"Data berhasil disimpan ke '{output_path}'")


if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    folder_sumber = os.path.join(script_dir, FOLDER_SUMBER)
    folder_output = os.path.join(script_dir, FOLDER_OUTPUT)

    nama_file_sumber = pilih_file(folder_sumber)
    if nama_file_sumber is None:
        print(f"Tidak ada file yang dipilih / tidak ditemukan file Excel di folder '{FOLDER_SUMBER}'.")
        raise SystemExit(1)

    nama_dasar = os.path.splitext(nama_file_sumber)[0]
    output_default = f"{nama_dasar}_expanded.xlsx"
    nama_file_output = input(
        f"Nama file output [default: {output_default}]: "
    ).strip() or output_default
    if not nama_file_output.lower().endswith('.xlsx'):
        nama_file_output += '.xlsx'

    try:
        expand_data(
            os.path.join(folder_sumber, nama_file_sumber),
            os.path.join(folder_output, nama_file_output),
        )
    except PermissionError:
        print(f"Error: File output '{nama_file_output}' sedang dibuka. Tutup dulu lalu coba lagi.")
    except Exception as e:
        print(f"Terjadi kesalahan saat memproses file: {e}")