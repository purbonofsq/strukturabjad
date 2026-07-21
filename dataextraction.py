import os
import pandas as pd

def save_filtered_data_to_excel(df: pd.DataFrame, juz_number: int, output_folder: str = "rawprocess") -> str:
    """
    Simpan DataFrame yang sudah difilter ke file Excel di folder output.
    
    Args:
        df: DataFrame yang berisi data ayat yang sudah difilter
        juz_number: Nomor Juz (NILAI_FILTER)
        output_folder: Nama folder untuk menyimpan file (default: rawprocess)
        
    Returns:
        Path file Excel yang disimpan
    """
    # Dapatkan direktori di mana script ini berada
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(script_dir, output_folder)
    
    # Buat folder jika belum ada
    os.makedirs(output_dir, exist_ok=True)
    
    # Buat nama file
    output_filename = os.path.join(output_dir, f"FilteredData_Juz{juz_number}.xlsx")
    
    # Hapus kolom 'TextAyat' jika ada
    if 'TextAyat' in df.columns:
        df = df.drop(columns=['TextAyat'])
    
    # Simpan ke Excel
    df.to_excel(output_filename, index=False)
    print(f"[OK] Data untuk Juz {juz_number} disimpan ke: {output_filename}")
    
    return output_filename
