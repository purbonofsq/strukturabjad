import pandas as pd
import os

def expand_data(input_filename='rawahaq.xlsx', output_filename='rawahaq_expanded.xlsx'):
    """
    Membaca file Excel, memperluas baris berdasarkan kolom 'Bobot',
    dan menyimpan hasilnya ke file Excel baru dengan kolom ID tambahan.

    Args:
        input_filename (str): Nama file Excel input.
        output_filename (str): Nama file Excel output.
    """
    # Dapatkan direktori di mana script ini berada
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # Gabungkan dengan nama file untuk mendapatkan path lengkap
    input_path = os.path.join(script_dir, input_filename)
    output_path = os.path.join(script_dir, output_filename)
    
    # Membaca data dari file Excel (header ada pada baris ke-2 / index 1)
    df = pd.read_excel(input_path, header=0)

    # Inisialisasi list untuk menyimpan baris-baris hasil ekspansi
    expanded_rows = []
    ay_huruf_ain_id_counter = 1 # Mulai ID dari 1

    # Iterasi melalui setiap baris di DataFrame asli
    for index, row in df.iterrows():
        # Dapatkan nilai bobot
        weight = int(row['Bobot']) # Pastikan bobot adalah integer

        # Ulangi sebanyak nilai bobot
        for _ in range(weight):
            # Salin baris saat ini
            new_row = row.copy()
            # Tambahkan ID unik ke salinan baris
            new_row['ay_huruf_ain_id'] = ay_huruf_ain_id_counter
            # Tambahkan salinan baris ke list hasil
            expanded_rows.append(new_row)
            # Naikkan counter ID
            ay_huruf_ain_id_counter += 1

    # Buat DataFrame baru dari list hasil
    expanded_df = pd.DataFrame(expanded_rows)

    # Urutkan berdasarkan ID (opsional, tapi biasanya diharapkan)
    expanded_df.sort_values(by='ay_huruf_ain_id', inplace=True, ignore_index=True)

    # Simpan DataFrame baru ke file Excel
    expanded_df.to_excel(output_path, index=False)
    print(f"Data berhasil diproses dan disimpan ke '{output_path}'")

if __name__ == "__main__":
    # Meminta nama file input dan output dari pengguna
    nama_file_sumber = input("Masukkan nama file sumber (dengan ekstensi .xlsx): [default: rawahaq.xlsx] ") or 'rawahaq.xlsx'
    nama_file_output = input("Masukkan nama file output (dengan ekstensi .xlsx): [default: rawahaq_expanded.xlsx] ") or 'rawahaq_expanded.xlsx'

    try:
        # Panggil fungsi untuk memproses data
        expand_data(nama_file_sumber, nama_file_output)
    except FileNotFoundError:
        print(f"Error: File '{nama_file_sumber}' tidak ditemukan.")
    except Exception as e:
        print(f"Terjadi kesalahan saat memproses file: {e}")
