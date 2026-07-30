import os
import re
import sys
from backend1 import PengolahDataJBAH
from dataextraction import save_filtered_data_to_excel

def sanitize_filename(name):
    """Hapus karakter ilegal pada nama file/folder Windows dan bersihkan tanda kutip."""
    # Hapus tanda kutip tunggal/ganda yang mungkin membungkus input
    name = name.strip().strip('"').strip("'")
    # Daftar karakter ilegal di Windows: \ / : * ? " < > |
    name = re.sub(r'[\\/:*?"<>|]', '_', name)
    return name.strip()

def parse_juz_input(input_str):
    """Parse user input for single Juz or range (e.g., "1" or "1-5")."""
    input_str = input_str.strip()
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
    JENIS_OUTPUT = "advance"
    
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
            # Jika bukan nomor, coba terima sebagai nama file manual (sebagai fallback)
            # Tapi ini kasus langka, berikan pesan jelas
            print(f"[ERROR] Masukkan hanya ANGKA (1-{len(available_files)}). Nama file manual tidak didukung.")
    
    # User input for Juz number(s)
    while True:
        try:
            user_input = input("Masukkan nomor Juz (contoh: 1 atau 1-5): ")
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
        
        for NILAI_FILTER in juz_list:
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
                # Pastikan ekstensi .xlsx
                if not filename.lower().endswith('.xlsx'):
                    filename += '.xlsx'
                # Sanitize filename akhir (jika ada karakter ilegal dari juz/metode - tidak mungkin, tapi antisipasi)
                filename = sanitize_filename(filename)
                if not filename.lower().endswith('.xlsx'):
                    filename += '.xlsx'
                NAMA_FILE_OUTPUT = os.path.join(output_dir, filename)
                
                # Apply method and process
                app.set_metode(metode)
                app.proses_1_dan_2()
                app.proses_3()
                app.generate_tabel_jb_init()
                app.generate_tabel_ah_final()
                
                # Export file dan terapkan formatting melalui modul output terpisah
                app.export_excel(NAMA_FILE_OUTPUT, jenis=JENIS_OUTPUT)
    
    print("\n[OK] SELURUH PROSES SELESAI")

if __name__ == "__main__":
    main()
