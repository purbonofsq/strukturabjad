import os
from backend1 import PengolahDataJBAH
from dataextraction import save_filtered_data_to_excel

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

def main():
    print("--- MEMULAI PENGOLAH DATA JBAH (USER INPUT) ---")
    
    # Get script directory to handle paths correctly
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    ## Default settings
    FILE_MENTAH = os.path.join(script_dir, "datasource", "rawahaq.xlsx")  # rawahaq.xlsx is in datasource!
    KOLOM_FILTER = "aq_juz_id"
    JENIS_OUTPUT = "advance"
    
    # User input for Juz number(s)
    while True:
        try:
            user_input = input("Masukkan nomor Juz (contoh: 1 atau 1-5): ")
            juz_list = parse_juz_input(user_input)
            break
        except ValueError as e:
            print(f"[ERROR] {e}. Mohon coba lagi!")
    
    # Initialize and process
    app = PengolahDataJBAH(FILE_MENTAH)
    
    if app.raw_data is not None:
        # Process all methods for each Juz
        daftar_metode = ["T111", "T222", "T333"]
        
        # Create dataoutput directory if it doesn't exist
        output_dir = os.path.join(script_dir, "jbah_ayat_per_hain")
        os.makedirs(output_dir, exist_ok=True)
        
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
                NAMA_FILE_OUTPUT = os.path.join(output_dir, f"Output_JBAH_Juz{NILAI_FILTER}_{metode}.xlsx")
                
                # Apply method and process
                app.set_metode(metode)
                app.proses_1_dan_2()
                app.proses_3()
                app.generate_tabel_jb_init()
                app.generate_tabel_ah_final()
                
                # Export file
                app.export_excel(NAMA_FILE_OUTPUT, jenis=JENIS_OUTPUT)
    
    print("\n[OK] SELURUH PROSES SELESAI")

if __name__ == "__main__":
    main()
