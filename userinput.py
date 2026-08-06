# userinput.py

import os
import re
import sys
from backend1 import PengolahDataJBAH
from dataextraction import save_filtered_data_to_excel
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

def sanitize_filename(name):
    """Hapus karakter ilegal pada nama file/folder Windows dan bersihkan tanda kutip."""
    # Hapus tanda kutip tunggal/ganda yang mungkin membungkus input
    name = name.strip().strip('"').strip("' ")
    # Daftar karakter ilegal di Windows: \ / : * ? " < > |
    name = re.sub(r'[\/:*?"<>|]', '_', name)
    return name.strip()

def parse_juz_input(input_str):
    """Parse user input for single Juz, range (e.g., "1" or "1-5"), or comma-separated list (e.g., "1,3,5")."""
    input_str = input_str.strip()
    # Coba parsing sebagai daftar yang dipisahkan koma terlebih dahulu
    if "," in input_str:
        try:
            # Split by comma, strip whitespace, convert to int
            juz_numbers = [int(x.strip()) for x in input_str.split(",")]
            # Validasi: pastikan semua angka > 0 dan unik (opsional)
            if any(juz < 1 for juz in juz_numbers):
                 raise ValueError("Nomor Juz harus lebih besar dari 0.")
            # Opsional: hapus duplikat dan urutkan
            # juz_numbers = sorted(set(juz_numbers))
            return juz_numbers
        except ValueError as ve:
            # Jika gagal parsing sebagai koma, coba sebagai range
            if "invalid literal for int()" in str(ve):
                raise ValueError(f"Format daftar Juz tidak valid: {input_str}. Gunakan angka, koma, dan/atau tanda hubung. Contoh: 1,3,5 atau 1-5 atau 7")
            else:
                raise ve # Re-raise jika error lain (misalnya, angka < 1)

    # Jika tidak ada koma, gunakan logika asli untuk single atau range
    if "-" in input_str:
        # Range input
        try:
            start, end = map(int, input_str.split("-"))
            if start > end:
                raise ValueError("Start must be <= end")
            return list(range(start, end + 1))
        except ValueError:
            raise ValueError("Invalid range format. Use e.g., 1-5")
    else:
        # Single Juz
        try:
            return [int(input_str)]
        except ValueError:
            raise ValueError("Invalid number. Use e.g., 1")

def list_available_datasource(script_dir):
    """List file Excel yang tersedia di folder datasource."""
    datasource_dir = os.path.join(script_dir, "datasource")
    available = []
    if os.path.isdir(datasource_dir):
        for f in os.listdir(datasource_dir):
            if f.lower().endswith(('.xlsx', '.xls')):
                name_without_ext = os.path.splitext(f)[0]
                available.append((name_without_ext, f))
    return available

def main():
    print("--- MEMULAI PENGOLAH DATA JBAH (USER INPUT) ---")
    # Get script directory to handle paths correctly
    script_dir = os.path.dirname(os.path.abspath(__file__))
    datasource_dir = os.path.join(script_dir, "datasource")

    ## Default settings
    DEFAULT_SOURCE_NAME = "rawahaq"
    KOLOM_FILTER = "aq_juz_id"
    JENIS_OUTPUT_DEFAULT = "pengguna"  # Ganti default
    FORMAT_OUTPUT_DEFAULT = "excel"    # Default format

    # User input for data source file (pilih nomor)
    print("\n--- PILIH FILE SUMBER DATA ---")
    available_files = list_available_datasource(script_dir)
    if not available_files:
        print(f"[ERROR] Tidak ada file Excel (.xlsx/.xls) di folder: {datasource_dir}")
        print(f"        Silakan letakkan file data sumber Anda di folder tersebut terlebih dahulu.")
        sys.exit(1)

    # Cari index default (DEFAULT_SOURCE_NAME)
    default_idx = None
    for idx, (name_no_ext, full_name) in enumerate(available_files):
        if name_no_ext.lower() == DEFAULT_SOURCE_NAME.lower():
            default_idx = idx + 1  # 1-based
            break
    if default_idx is None:
        default_idx = 1  # fallback ke file pertama jika default tidak ada

    print(f"File Excel yang tersedia di folder 'datasource':")
    for idx, (name_no_ext, full_name) in enumerate(available_files, 1):
        marker = " [default]" if idx == default_idx else ""
        print(f"  {idx}. {full_name}{marker}")

    while True:
        choice = input(f"Masukkan nomor pilihan Anda (default: {default_idx}): ").strip()
        if not choice:
            choice = str(default_idx)
        # Cek apakah input adalah nomor
        try:
            choice_num = int(choice)
            if 1 <= choice_num <= len(available_files):
                name_no_ext, full_name = available_files[choice_num - 1]
                FILE_MENTAH = os.path.join(datasource_dir, full_name)
                print(f"[INFO] Pilihan #{choice_num}: {full_name}")
                print(f"[INFO] Menggunakan file sumber: {FILE_MENTAH}")
                break
            else:
                print(f"[ERROR] Nomor tidak valid. Masukkan angka antara 1 sampai {len(available_files)}.")
        except ValueError:
            print(f"[ERROR] Masukkan hanya ANGKA (1-{len(available_files)}). Nama file manual tidak didukung.")

    # User input for Juz number(s) - PINDAH DEFINISI KE SINI AGAR BISA DI-ACCESS SECARA GLOBAL
    juz_list = None
    while True:
        try:
            user_input = input("Masukkan nomor Juz (contoh: 1 atau 1-5 atau 1,3,5): ")
            juz_list = parse_juz_input(user_input)
            break
        except ValueError as e:
            print(f"[ERROR] {e}. Mohon coba lagi!")

    # User input for output directory
    user_output_dir = sanitize_filename(input("Masukkan nama folder output (default: dataoutput): ").strip() or "dataoutput")

    # User input for output filename template
    default_template = "Output_JBAH_Juz{juz}_{metode}.xlsx"
    print(f"\nFormat penamaan file (placeholder: {{juz}} = nomor juz, {{metode}} = T111/T222/T333)")
    print(f"  * Tips: Anda bisa memasukkan nama file tanpa placeholder, nanti akan ditambahkan otomatis.")
    file_template_raw = input(f"Masukkan template nama file (default: {default_template}): ").strip() or default_template

    # Sanitize template (tapi placeholder {juz} dan {metode} harus tetap dipertahankan)
    placeholder_juz = "__JUZ_PLACEHOLDER__"
    placeholder_metode = "__METODE_PLACEHOLDER__"
    temp_template = file_template_raw.replace("{juz}", placeholder_juz).replace("{metode}", placeholder_metode)
    temp_template = sanitize_filename(temp_template)
    file_template = temp_template.replace(placeholder_juz, "{juz}").replace(placeholder_metode, "{metode}")

    # Logika fleksibel: pastikan template punya placeholder agar file tidak saling menimpa
    has_juz = "{juz}" in file_template
    has_metode = "{metode}" in file_template

    # Pisahkan nama file dan extension
    base_name, ext = os.path.splitext(file_template)
    if not ext:
        ext = ".xlsx"

    if not has_juz and not has_metode:
        # User tidak masukkan placeholder sama sekali → auto-tambah keduanya
        file_template = f"{base_name}_Juz{{juz}}_{{metode}}{ext}"
        print(f"[INFO] Template disesuaikan (auto-tambah placeholder): {file_template}")
    elif has_juz and not has_metode:
        # Hanya ada {juz}, tambahkan {metode} agar 3 metode tidak saling timpa
        file_template = f"{base_name}_{{metode}}{ext}"
        print(f"[INFO] Template disesuaikan (auto-tambah {{metode}}): {file_template}")
    elif not has_juz and has_metode:
        # Hanya ada {metode}, tambahkan {juz} agar multi juz tidak saling timpa
        file_template = f"{base_name}_Juz{{juz}}{ext}"
        print(f"[INFO] Template disesuaikan (auto-tambah {{juz}}): {file_template}")
    # else: keduanya ada, biarkan seperti input user

    # --- Tambahkan bagian ini ---
    while True:
        jenis_pengguna_input = input("\nPilih jenis output - a. Pengguna (Init & Final saja)   b. Praktisi (Semua Tab)\nMasukkan pilihan (a/b) (default: a): ").strip().lower()
        if jenis_pengguna_input == '' or jenis_pengguna_input == 'a':
            JENIS_OUTPUT_FINAL = "pengguna"
            break
        elif jenis_pengguna_input == 'b':
            JENIS_OUTPUT_FINAL = "praktisi"
            break
        else:
            print("[ERROR] Pilihan tidak valid. Harap masukkan 'a' atau 'b'.")

    while True:
        format_output_input = input("\nPilih format output - a. Excel   b. HTML\nMasukkan pilihan (a/b) (default: a): ").strip().lower()
        if format_output_input == '' or format_output_input == 'a':
            FORMAT_OUTPUT_FINAL = "excel"
            ext_default = ".xlsx"
            break
        elif format_output_input == 'b':
            FORMAT_OUTPUT_FINAL = "html"
            ext_default = ".html"
            break
        else:
            print("[ERROR] Pilihan tidak valid. Harap masukkan 'a' atau 'b'.")

    # --- Tambahkan bagian ini untuk pilihan output Excel ---
    if FORMAT_OUTPUT_FINAL == "excel":
        while True:
            excel_output_choice = input("\nPilih mode output Excel:\n  a. Satu file besar untuk semua Juz & Metode (dengan formatting otomatis per sheet)\n  b. Satu file per Juz (semua metode dalam satu file)\nMasukkan pilihan (a/b) (default: a): ").strip().lower()
            if excel_output_choice == '' or excel_output_choice == 'a':
                EXCEL_OUTPUT_MODE = "single_file"
                break
            elif excel_output_choice == 'b':
                EXCEL_OUTPUT_MODE = "per_juz"
                break
            else:
                print("[ERROR] Pilihan tidak valid. Harap masukkan 'a' atau 'b'.")
    else:
        # Jika bukan Excel, set mode ke None atau nilai default
        EXCEL_OUTPUT_MODE = None # atau "single_file" jika HTML selalu single file

    # ----------------------------

    default_concept_label = "Konsep Dasar (metode jumlah ayat dalam juz)"
    concept_label = input(f"Masukkan label konsep dasar (default: {default_concept_label}): ").strip() or default_concept_label

    default_signature = "Digital Signature"
    signature_text = input(f"Masukkan tanda tangan / penandatangan (default: {default_signature}): ").strip() or default_signature

    # Initialize and process
    app = PengolahDataJBAH(FILE_MENTAH, concept_label=concept_label, signature_text=signature_text)
    if app.raw_data is not None:
        # Process all methods for each Juz
        daftar_metode = ["T111", "T222", "T333"]
        # Create output directory if it doesn't exist
        output_dir = os.path.join(script_dir, user_output_dir)
        os.makedirs(output_dir, exist_ok=True)
        print(f"[INFO] Folder output: {output_dir}")
        print(f"[INFO] Template file: {file_template}")

        # --- Logika Output Excel ---
        if FORMAT_OUTPUT_FINAL == "excel" and EXCEL_OUTPUT_MODE == "single_file":
            print("\n[INFO] Mode: Excel - Satu file besar. Mengumpulkan data...")
            # Buat nama file untuk output single file
            single_file_name = sanitize_filename(file_template.format(juz="ALL", metode="ALL")) # Ganti placeholder dengan ALL
            if not single_file_name.lower().endswith(ext_default):
                single_file_name += ext_default
            NAMA_FILE_OUTPUT_SINGLE = os.path.join(output_dir, single_file_name)

            # Inisialisasi writer untuk file besar
            # KUMPULKAN DATA TERLEBIH DAHULU UNTUK MENGECEK APAKAH ADA YANG DIPROSES
            data_to_write = [] # List tuple: (sheet_name, df_to_write)
            for NILAI_FILTER in juz_list: # Gunakan juz_list yang didefinisikan di atas
                print(f"\n=== Memproses Juz {NILAI_FILTER} ===")
                # Filter data
                staging_data = app.filter_data(KOLOM_FILTER, NILAI_FILTER)
                if staging_data is None:
                    print(f"[WARNING] Melewatkan Juz {NILAI_FILTER} karena tidak ada data.")
                    continue

                # Simpan data yang sudah difilter ke folder rawprocess
                save_filtered_data_to_excel(staging_data, NILAI_FILTER)

                for metode in daftar_metode:
                    # Apply method and process
                    app.set_metode(metode)
                    app.proses_1_dan_2()
                    app.proses_3()
                    app.generate_tabel_jb_init()
                    app.generate_tabel_ah_final()

                    # --- Simpan ke list sementara ---
                    sheet_prefix = f"Juz{NILAI_FILTER}_{metode}_"
                    if JENIS_OUTPUT_FINAL == "praktisi":
                        data_to_write.append((sheet_prefix + "Init", app.df_init))
                        data_to_write.append((sheet_prefix + "Proses_1", app.df_p1))
                        data_to_write.append((sheet_prefix + "Proses_2", app.df_p2))
                        data_to_write.append((sheet_prefix + "Proses_3", app.df_p3))
                        data_to_write.append((sheet_prefix + "Final", app.df_final))
                    elif JENIS_OUTPUT_FINAL == "pengguna":
                        data_to_write.append((sheet_prefix + "Init", app.df_init))
                        data_to_write.append((sheet_prefix + "Final", app.df_final))

            # PERIKSA APAKAH ADA DATA YANG DIPROSES
            if not data_to_write:
                 print(f"[WARNING] Tidak ada data yang berhasil diproses dari Juz {juz_list}. File Excel besar tidak akan dibuat.")
            else:
                 print(f"[INFO] Menemukan {len(data_to_write)} sheet untuk ditulis ke file gabungan.")
                 with pd.ExcelWriter(NAMA_FILE_OUTPUT_SINGLE, engine='openpyxl') as writer:
                     for sheet_name, df in data_to_write:
                         df.to_excel(writer, sheet_name=sheet_name, index=False)

                 print(f"[SAVED] File Excel gabungan disimpan sebagai: {NAMA_FILE_OUTPUT_SINGLE}")
                 print("[INFO] File gabungan dibuat. Sekarang menerapkan formatting otomatis ke setiap sheet...")

                 # --- Apply formatting ke setiap sheet dalam file besar ---
                 try:
                     wb = openpyxl.load_workbook(NAMA_FILE_OUTPUT_SINGLE)
                     sheet_names = wb.sheetnames

                     for sheet_name in sheet_names:
                         ws = wb[sheet_name]
                         # Cek apakah sheet_name adalah Final berdasarkan nama (mengandung "_Final")
                         if sheet_name.endswith("_Final"):
                             # Terapkan styling Final
                             apply_final_style(ws, signature_text)
                         else:
                             # Terapkan styling standar
                             apply_standard_style(ws)

                     wb.save(NAMA_FILE_OUTPUT_SINGLE)
                     print(f"[OK] Formatting otomatis berhasil diterapkan ke file: {NAMA_FILE_OUTPUT_SINGLE}")
                 except Exception as e:
                     print(f"[ERROR] Gagal menerapkan formatting ke file gabungan: {e}")
                     import traceback
                     traceback.print_exc()

        elif FORMAT_OUTPUT_FINAL == "excel" and EXCEL_OUTPUT_MODE == "per_juz":
            print("\n[INFO] Mode: Excel - Satu file per Juz.")
            for NILAI_FILTER in juz_list: # Gunakan juz_list yang didefinisikan di atas
                print(f"\n=== Memproses Juz {NILAI_FILTER} ===")
                # Filter data
                staging_data = app.filter_data(KOLOM_FILTER, NILAI_FILTER)
                if staging_data is None:
                    print(f"[WARNING] Melewatkan Juz {NILAI_FILTER} karena tidak ada data.")
                    continue

                # Simpan data yang sudah difilter ke folder rawprocess
                save_filtered_data_to_excel(staging_data, NILAI_FILTER)

                # Buat nama file per Juz
                juz_file_name = sanitize_filename(file_template.format(juz=NILAI_FILTER, metode="ALL")) # Gunakan ALL untuk metode
                if not juz_file_name.lower().endswith(ext_default):
                    juz_file_name += ext_default
                NAMA_FILE_OUTPUT_PER_JUZ = os.path.join(output_dir, juz_file_name)

                # Proses semua metode untuk Juz ini dan kumpulkan DataFrame
                all_data_frames_per_juz = {}
                for sheet_name in ["Init", "Proses 1", "Proses 2", "Proses 3", "Final"]:
                    all_data_frames_per_juz[sheet_name] = []

                for metode in daftar_metode:
                    # Apply method and process
                    app.set_metode(metode)
                    app.proses_1_dan_2()
                    app.proses_3()
                    app.generate_tabel_jb_init()
                    app.generate_tabel_ah_final()

                    # Tambahkan hasil ke list sesuai sheet
                    if JENIS_OUTPUT_FINAL == "praktisi":
                        all_data_frames_per_juz["Init"].append(app.df_init)
                        all_data_frames_per_juz["Proses 1"].append(app.df_p1)
                        all_data_frames_per_juz["Proses 2"].append(app.df_p2)
                        all_data_frames_per_juz["Proses 3"].append(app.df_p3)
                        all_data_frames_per_juz["Final"].append(app.df_final)
                    elif JENIS_OUTPUT_FINAL == "pengguna":
                        all_data_frames_per_juz["Init"].append(app.df_init)
                        all_data_frames_per_juz["Final"].append(app.df_final)

                # Gabungkan DataFrame per sheet untuk Juz ini
                final_data_per_juz = {}
                for sheet_name, df_list in all_data_frames_per_juz.items():
                    if df_list:
                        # Tambahkan kolom identifikasi opsional (metode, juz) jika perlu
                        combined_df = pd.concat(df_list, ignore_index=True)
                        # Misalnya, tambahkan kolom 'Metode' ke setiap baris
                        metode_labels = []
                        for df_orig in df_list:
                             metode_labels.extend([df_orig.iloc[0, 0].split()[2]] * len(df_orig)) # Ambil metode dari kolom pertama baris pertama tiap df
                        combined_df.insert(0, 'Metode', metode_labels)
                        final_data_per_juz[sheet_name] = combined_df

                # Ekspor ke file Excel per Juz
                if final_data_per_juz:
                     with pd.ExcelWriter(NAMA_FILE_OUTPUT_PER_JUZ, engine='openpyxl') as writer:
                         for sheet_name, df in final_data_per_juz.items():
                             df.to_excel(writer, sheet_name=sheet_name, index=False)
                     print(f"[SAVED] File Excel untuk Juz {NILAI_FILTER} disimpan sebagai: {NAMA_FILE_OUTPUT_PER_JUZ}")
                     # Kita tidak bisa langsung memformat file ini karena struktur datanya (banyak metode digabung) berbeda
                     # dengan file single-metode/single-juz biasa.
                     print("[INFO] File per Juz tidak diberi formatting otomatis karena struktur datanya berbeda.")
                else:
                    print(f"[WARNING] Tidak ada data untuk diekspor ke file Juz {NILAI_FILTER}.")

        else: # --- Mode Output Lain (HTML atau Excel Single File Per Metode Lama) ---
            # Mode lama: satu file per metode (seperti sebelum fitur tambahan)
            # Kita bisa mempertahankan logika lama di sini jika format bukan excel atau mode bukan single_file/per_juz
            # Tapi karena kita sudah menangani excel dengan mode tertentu di atas,
            # bagian ini bisa diabaikan untuk excel, dan hanya berlaku untuk HTML.
            if FORMAT_OUTPUT_FINAL == "html": # Pastikan hanya HTML yang masuk sini
                 for NILAI_FILTER in juz_list: # Gunakan juz_list yang didefinisikan di atas
                     print(f"\n=== Memproses Juz {NILAI_FILTER} ===")
                     # Filter data
                     staging_data = app.filter_data(KOLOM_FILTER, NILAI_FILTER)
                     if staging_data is None:
                         print(f"[WARNING] Melewatkan Juz {NILAI_FILTER} karena tidak ada data.")
                         continue

                     # Simpan data yang sudah difilter ke folder rawprocess
                     save_filtered_data_to_excel(staging_data, NILAI_FILTER)

                     for metode in daftar_metode:
                         filename = file_template.format(juz=NILAI_FILTER, metode=metode)
                         # Sanitize filename akhir (jika ada karakter ilegal dari juz/metode - tidak mungkin, tapi antisipasi)
                         filename = sanitize_filename(filename)
                         # Pastikan ekstensi sesuai dengan format output
                         if not filename.lower().endswith(('.xlsx', '.html')):
                             filename += ext_default
                         elif (FORMAT_OUTPUT_FINAL == "excel" and not filename.lower().endswith('.xlsx')) or \
                              (FORMAT_OUTPUT_FINAL == "html" and not filename.lower().endswith('.html')):
                             # Ganti ekstensi jika tidak cocok dengan format yang dipilih
                             name_root = os.path.splitext(filename)[0]
                             filename = name_root + ext_default

                         NAMA_FILE_OUTPUT = os.path.join(output_dir, filename)
                         # Apply method and process
                         app.set_metode(metode)
                         app.proses_1_dan_2()
                         app.proses_3()
                         app.generate_tabel_jb_init()
                         app.generate_tabel_ah_final()
                         # Export file dan terapkan formatting melalui modul output terpisah
                         # Gunakan JENIS_OUTPUT_FINAL dan FORMAT_OUTPUT_FINAL yang baru ditentukan
                         app.export_excel(NAMA_FILE_OUTPUT, jenis=JENIS_OUTPUT_FINAL, format=FORMAT_OUTPUT_FINAL)
            else:
                 print(f"[ERROR] Mode output tidak dikenali untuk format {FORMAT_OUTPUT_FINAL}. Proses dibatalkan.")
                 return # Batalkan jika mode tidak sesuai

    print("\n[OK] SELURUH PROSES SELESAI")

# --- Fungsi Styling untuk File Gabungan ---
def apply_standard_style(ws):
    """Apply basic Excel styling to a worksheet."""
    font_header = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    font_body = Font(name="Arial", size=10)
    font_label = Font(name="Arial", size=10, bold=True)
    font_title = Font(name="Arial", size=12, bold=True)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_alt = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )

    if ws.max_row > 0:
        for row in range(1, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                cell.border = thin_border
                cell.alignment = align_center
                if row == 1:
                    cell.font = font_header
                    cell.fill = fill_header
                else:
                    cell.font = font_body
                    if col == 1:
                        cell.font = font_label
                    if row % 2 == 0:
                        cell.fill = fill_alt
    ws.freeze_panes = "A2"
    for col in range(1, ws.max_column + 1):
        column_letter = get_column_letter(col)
        max_length = 0
        for row in range(1, ws.max_row + 1):
            cell_value = ws.cell(row=row, column=col).value
            if cell_value is None:
                continue
            if isinstance(cell_value, (int, float)):
                cell_length = len(str(cell_value)) + 2
            else:
                cell_length = len(str(cell_value)) + 2
            max_length = max(max_length, cell_length)
        ws.column_dimensions[column_letter].width = min(max_length + 2, 40)

def apply_final_style(ws, signature_text):
    """Apply specific Excel styling for the 'Final' worksheet."""
    font_title = Font(name="Arial", size=14, bold=True, color="000000")
    font_subtitle = Font(name="Arial", size=11, bold=True, italic=True)
    font_header = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    font_body_regular = Font(name="Arial", size=10)
    font_arab = Font(name="Arial", size=12, bold=True)
    font_label = Font(name="Arial", size=10, bold=True, color="000000")
    font_label_italic = Font(name="Arial", size=10, bold=True, italic=True)
    font_label_bold = Font(name="Arial", size=10, bold=True)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_center_top = Alignment(horizontal="center", vertical="top", wrap_text=True)
    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_alt_row = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    thick_border = Border(
        left=Side(style="thick"), right=Side(style="thick"),
        top=Side(style="thick"), bottom=Side(style="thick")
    )

    # Cek jumlah baris sebelum menggabungkan
    max_r = ws.max_row
    if max_r < 6:
        print(f"[WARNING] Sheet '{ws.title}' tidak cukup baris untuk styling Final. Membatalkan styling.")
        return

    ws.merge_cells("A1:J1")
    ws["A1"].font = font_title
    ws["A1"].alignment = align_center
    ws.merge_cells("A2:J2")
    ws["A2"].font = font_subtitle
    ws["A2"].alignment = align_center
    ws.merge_cells("A4:C4")
    ws["A4"].alignment = align_center
    ws["A4"].border = thick_border
    ws["A4"].font = font_label
    ws.merge_cells("D4:G4")
    ws["D4"].alignment = align_center
    ws["D4"].border = thick_border
    ws["D4"].font = font_label
    ws.merge_cells("H4:J4")
    ws["H4"].alignment = align_center
    ws["H4"].border = thick_border
    ws["H4"].font = font_label
    ws["B3"].font = font_label_italic
    ws["I3"].font = font_subtitle

    for col in range(1, 11):
        cell = ws.cell(row=5, column=col)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border

    max_r = ws.max_row
    max_c = min(10, ws.max_column)
    for row in range(6, max_r + 1):
        for col in range(1, max_c + 1):
            cell = ws.cell(row=row, column=col)
            cell.alignment = align_center_top
            cell.border = thin_border
            if col == 10:
                cell.font = font_arab
            elif col == 1:
                cell.font = font_label_italic
            elif col in {2, 4, 6, 8}:
                cell.font = font_label_bold
            else:
                cell.font = font_body_regular
            if row % 2 == 0:
                cell.fill = fill_alt_row

    column_widths = {"A": 5, "B": 7, "C": 10, "D": 4, "E": 10, "F": 4, "G": 10, "H": 7, "I": 15, "J": 5}
    for col_letter, width in column_widths.items():
        if col_letter in ws.column_dimensions:
            ws.column_dimensions[col_letter].width = width

    signature_row = ws.max_row + 2
    signature_title = "__"
    signature_stamp = f": {signature_text}"
    signature_time = f": {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"

    ws.merge_cells(start_row=signature_row, start_column=1, end_row=signature_row, end_column=10)
    ws.cell(signature_row, 1, value=signature_title)
    ws.cell(signature_row, 1).font = Font(name="Arial", size=12, bold=True, color="1F497D")
    ws.cell(signature_row, 1).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(signature_row, 1).fill = PatternFill(start_color="E9F2FF", end_color="E9F2FF", fill_type="solid")
    ws.cell(signature_row, 1).border = thin_border

    ws.merge_cells(start_row=signature_row + 1, start_column=1, end_row=signature_row + 1, end_column=10)
    ws.cell(signature_row + 1, 1, value=signature_stamp)
    ws.cell(signature_row + 1, 1).font = font_label_bold
    ws.cell(signature_row + 1, 1).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(signature_row + 1, 1).border = thin_border

    ws.merge_cells(start_row=signature_row + 2, start_column=1, end_row=signature_row + 2, end_column=10)
    ws.cell(signature_row + 2, 1, value=signature_time)
    ws.cell(signature_row + 2, 1).font = Font(name="Arial", size=9, italic=True)
    ws.cell(signature_row + 2, 1).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(signature_row + 2, 1).border = thin_border

from datetime import datetime # Import datetime untuk signature di final style

if __name__ == "__main__":
    main()
