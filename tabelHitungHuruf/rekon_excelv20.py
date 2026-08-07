import os
import sys
import pandas as pd
import numpy as np


# Folder utama tempat program mencari file Excel.
# Jika ingin mencari file di folder tempat script dijalankan,
# ganti menjadi: BASE_DIR = os.getcwd()
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def resolve_path(filename, base_dir=None):
    """
    Mengubah nama file menjadi path lengkap.
    Jika filename sudah berupa path absolut, gunakan langsung.
    Jika tidak, gabungkan dengan base_dir.
    """
    if base_dir is None:
        base_dir = BASE_DIR

    filename = os.path.expanduser(filename)

    if os.path.isabs(filename):
        return filename

    return os.path.join(base_dir, filename)


def list_excel_files(folder=None):
    """
    Mengambil daftar file Excel .xlsx di folder tertentu.
    File temporary Excel seperti ~$file.xlsx diabaikan.
    """
    if folder is None:
        folder = BASE_DIR

    if not os.path.isdir(folder):
        return []

    files = []

    for name in os.listdir(folder):
        if name.lower().endswith(".xlsx") and not name.startswith("~$"):
            full_path = os.path.join(folder, name)

            if os.path.isfile(full_path):
                files.append(name)

    return sorted(files, key=lambda x: x.lower())


def tampilkan_daftar_file(files, folder, judul):
    """
    Menampilkan daftar file Excel dalam bentuk list bernomor.
    """
    print("\n" + "=" * 70)
    print(judul)
    print(f"Folder: {folder}")
    print("-" * 70)

    if not files:
        print("Tidak ada file .xlsx di folder ini.")
    else:
        for i, name in enumerate(files, start=1):
            print(f"{i:>3}. {name}")

    print("-" * 70)


def pilih_file_excel(judul, default_name, mode, folder=None):
    """
    Menampilkan daftar file Excel dan meminta user memilih.

    mode:
    - 'input'  : untuk file sumber
    - 'output' : untuk file hasil

    User bisa:
    - memilih nomor file
    - mengetik nama file manual
    - mengetik 'r' untuk refresh daftar
    - menekan Enter untuk memakai nama default
    """
    if folder is None:
        folder = BASE_DIR

    files = list_excel_files(folder)

    while True:
        tampilkan_daftar_file(files, folder, judul)

        print("Pilihan: nomor file, nama file manual, 'r' refresh, Enter default.")

        if mode == "output":
            print("Jika memilih file yang sudah ada, file tersebut akan ditimpa.")

        pilihan = input(">> ").strip().strip('"').strip("'")

        # Refresh daftar file
        if pilihan.lower() == "r":
            files = list_excel_files(folder)
            continue

        # Pilih default
        if pilihan == "":
            if mode == "input":
                default_path = resolve_path(default_name, folder)

                if os.path.exists(default_path):
                    return default_name

                print(
                    f"File default '{default_name}' tidak ditemukan. "
                    "Silakan pilih file lain atau ketik nama file secara manual."
                )
                continue

            return default_name

        # Pilih berdasarkan nomor
        if pilihan.isdigit():
            idx = int(pilihan)

            if 1 <= idx <= len(files):
                return files[idx - 1]

            print("Nomor tidak valid. Silakan coba lagi.")
            continue

        # Jika user mengetik nama file manual
        if not pilihan.lower().endswith(".xlsx"):
            pilihan += ".xlsx"

        return pilihan


def expand_data(
    input_filename="rawahaq.xlsx",
    output_filename="rawahaq_expanded.xlsx",
    base_dir=None
):
    """
    Membaca file Excel, memperluas baris berdasarkan kolom 'Bobot',
    dan menyimpan hasilnya ke file Excel baru dengan kolom ID tambahan.
    """

    if base_dir is None:
        base_dir = BASE_DIR

    input_path = resolve_path(input_filename, base_dir)
    output_path = resolve_path(output_filename, base_dir)

    if not os.path.exists(input_path):
        raise FileNotFoundError(f"File input tidak ditemukan: {input_path}")

    # Baca file Excel.
    # Jika header berada di baris kedua Excel, ubah header=0 menjadi header=1.
    df = pd.read_excel(input_path, header=0)

    # Reset index agar aman saat menggunakan repeat
    df = df.reset_index(drop=True)

    # Validasi kolom Bobot
    if "Bobot" not in df.columns:
        raise ValueError("Kolom 'Bobot' tidak ditemukan di file input.")

    # Konversi Bobot menjadi angka
    bobot_numeric = pd.to_numeric(df["Bobot"], errors="coerce")

    # Cek nilai yang tidak bisa dikonversi
    invalid_mask = bobot_numeric.isna()
    if invalid_mask.any():
        invalid_rows = df.index[invalid_mask].tolist()
        raise ValueError(f"Nilai 'Bobot' tidak valid pada baris: {invalid_rows}")

    # Cek nilai tak hingga
    if not np.isfinite(bobot_numeric).all():
        raise ValueError("Kolom 'Bobot' mengandung nilai tidak valid atau tak hingga.")

    # Cek nilai negatif
    negative_mask = bobot_numeric < 0
    if negative_mask.any():
        negative_rows = df.index[negative_mask].tolist()
        raise ValueError(f"'Bobot' tidak boleh negatif pada baris: {negative_rows}")

    # Cek apakah semua nilai merupakan bilangan bulat
    decimal_mask = (bobot_numeric % 1) != 0
    if decimal_mask.any():
        decimal_rows = df.index[decimal_mask].tolist()
        raise ValueError(f"'Bobot' harus bilangan bulat pada baris: {decimal_rows}")

    # Konversi menjadi integer
    bobot = bobot_numeric.astype(int)

    # Perluas baris berdasarkan bobot
    expanded_df = df.loc[df.index.repeat(bobot)].copy()

    # Tambahkan ID unik
    expanded_df["ay_huruf_mimm_id"] = range(1, len(expanded_df) + 1)

    # Letakkan ID di kolom paling kiri
    column_order = ["ay_huruf_mimm_id"] + [
        col for col in expanded_df.columns if col != "ay_huruf_mimm_id"
    ]
    expanded_df = expanded_df[column_order]

    # Reset index akhir
    expanded_df = expanded_df.reset_index(drop=True)

    # Pastikan folder output tersedia
    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    # Simpan ke Excel
    expanded_df.to_excel(output_path, index=False)

    print(f"\nData berhasil diproses dan disimpan ke '{output_path}'")
    return output_path


if __name__ == "__main__":
    folder = BASE_DIR

    try:
        nama_file_sumber = pilih_file_excel(
            judul="Pilih File Sumber Excel",
            default_name="rawahaq.xlsx",
            mode="input",
            folder=folder
        )

        nama_file_output = pilih_file_excel(
            judul="Pilih File Output Excel",
            default_name="rawahaq_expanded.xlsx",
            mode="output",
            folder=folder
        )

        output_path = resolve_path(nama_file_output, folder)

        # Konfirmasi jika file output sudah ada
        if os.path.exists(output_path):
            jawab = input(
                f"\nFile output '{os.path.basename(output_path)}' sudah ada. "
                "Timpa? (y/n): "
            ).strip().lower()

            if jawab not in ("y", "ya", "yes"):
                print("Proses dibatalkan. File tidak ditimpa.")
                sys.exit(0)

        expand_data(
            input_filename=nama_file_sumber,
            output_filename=nama_file_output,
            base_dir=folder
        )

    except KeyboardInterrupt:
        print("\nProses dibatalkan oleh user.")
        sys.exit(0)

    except FileNotFoundError as e:
        print(f"Error: file atau folder tidak ditemukan. Detail: {e}")

    except ValueError as ve:
        print(f"Error validasi data: {ve}")

    except Exception as e:
        print(f"Terjadi kesalahan saat memproses file: {e}")