import os
import pandas as pd
import numpy as np
from typing import Any, List, Optional

# Tambahan Import untuk Styling Excel
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

class PengolahDataJBAH:
    def __init__(self, file_path):
        """Inisialisasi class dan memuat data mentah Excel."""
        self.file_path = file_path
        self.raw_data: Optional[pd.DataFrame] = None
        try:
            self.raw_data = pd.read_excel(self.file_path)
            print(f"[OK] Data mentah Excel berhasil dimuat dari: {self.file_path}")
        except Exception as e:
            print(f"[ERROR] Error memuat data: {e}")
            self.raw_data = None

    def filter_data(self, filter_col, filter_value):
        """Tahap II.1: Staging Data Utama (Filter dan Inisialisasi Dasar)"""
        if self.raw_data is None:
            print("[ERROR] raw_data is None, cannot filter.")
            return None

        try:
            # Check if filter column exists
            if filter_col not in self.raw_data.columns:
                print(f"[ERROR] Kolom '{filter_col}' tidak ditemukan di data. Kolom yang tersedia: {list(self.raw_data.columns)}")
                return None

            self.NILAI_FILTER = filter_value
            self.staging_data = self.raw_data[self.raw_data[filter_col] == filter_value].copy()
            
            # Check if any rows were found
            if len(self.staging_data) == 0:
                unique_values = self.raw_data[filter_col].unique()
                print(f"[ERROR] Tidak ada data untuk {filter_col} = {filter_value}. Nilai yang tersedia: {sorted(unique_values)}")
                return None

            self.staging_data.reset_index(drop=True, inplace=True)
            
            # Membuat string penggabungan surat.ayat
            self.staging_data['surat_ayat'] = self.staging_data['aq_surat_id'].astype(str) + '.' + self.staging_data['surat_ayat_no'].astype(str)
            self.total_ayat = len(self.staging_data)
            
            print(f"[OK] Filter Data Siap: Total {self.total_ayat} ayat.")
            return self.staging_data
        except Exception as e:
            print(f"[ERROR] Error saat filtering data: {type(e).__name__}: {e}")
            import traceback
            traceback.print_exc()
            return None

    def set_metode(self, metode="T111"):
        """
        Tahap Penentuan Poros (Metode Siklik):
        Menggeser (shifting) array data agar target Blok F jatuh tepat di tengah.
        """
        self.metode_aktif = metode
        N = self.total_ayat
        
        # 1. Tentukan posisi tengah absolut dari array
        if N % 2 == 0:
            center_pos = N // 2 - 1  # Indeks kiri dari 2 ayat tengah
        else:
            center_pos = N // 2
            
        # 2. Tentukan target indeks asli yang ingin dijadikan Blok F
        if metode == "T111":
            target_f_idx = center_pos 
        elif metode == "T222":
            target_f_idx = 0  # Poros di awal
        elif metode == "T333":
            if N % 2 == 0:
                target_f_idx = N - 2  # Poros di akhir (2 ayat)
            else:
                target_f_idx = N - 1  # Poros di akhir (1 ayat)
        else:
            target_f_idx = center_pos # Default
            
        # 3. Hitung offset dan geser data secara sirkular (Modulo)
        # Mencari indeks awal (0) baru agar target_f_idx jatuh tepat di center_pos
        start_idx = (target_f_idx - center_pos) % N
        new_order = [(start_idx + i) % N for i in range(N)]
        
        self.staging_data_met = self.staging_data.iloc[new_order].copy()
        self.staging_data_met.reset_index(drop=True, inplace=True)
        
        # 4. Update Pelaporan Dasar berdasarkan poros baru
        self.ayat_awal = self.staging_data_met.loc[0, 'surat_ayat']
        self.ayat_akhir = self.staging_data_met.loc[N - 1, 'surat_ayat']
        
        print(f"\n[INFO] Set Metode: {metode} | Ayat Awal Baru: {self.ayat_awal} | Ayat Akhir Baru: {self.ayat_akhir}")

    def proses_1_dan_2(self):
        """Tahap II.2.1 & II.2.2: Chunking Baku & Symmetrical Reordering"""
        N = self.total_ayat
        self.p1_dict = {}
        
        # 1. Menentukan Titik Tengah
        if N % 2 == 0:
            mid_left = N // 2 - 1
            mid_right = N // 2
            self.titik_tengah = f"{self.staging_data_met.loc[mid_left, 'surat_ayat']}+{self.staging_data_met.loc[mid_right, 'surat_ayat']}"
            left_idx = mid_left - 1
            right_idx = mid_right + 1
        else:
            mid = N // 2
            self.titik_tengah = self.staging_data_met.loc[mid, 'surat_ayat']
            left_idx = mid - 1
            right_idx = mid + 1
            
        self.p1_dict['F'] = [self.titik_tengah]

        # 2. Template Baku Lebar Sayap (Mekar ke luar)
        met1_widths = [
            ('A', 'a', 5), ('B', 'b', 7), ('C1', 'c1', 7), ('C2', 'c2', 7),
            ('D1', 'd1', 7), ('D2', 'd2', 7), ('D3', 'd3', 9),
            ('E', 'e', 5), ('G', 'g', 1)
        ]

        for left_lbl, right_lbl, width in met1_widths:
            left_chunk = self.staging_data_met.loc[left_idx - width + 1 : left_idx, 'surat_ayat'].tolist()
            right_chunk = self.staging_data_met.loc[right_idx : right_idx + width - 1, 'surat_ayat'].tolist()
            
            self.p1_dict[left_lbl] = left_chunk
            self.p1_dict[right_lbl] = right_chunk
            
            left_idx -= width
            right_idx += width

        # 3. Merakit DataFrame Proses 1
        order_p1 = ['G', 'E', 'D3', 'D2', 'D1', 'C2', 'C1', 'B', 'A', 'F', 'a', 'b', 'c1', 'c2', 'd1', 'd2', 'd3', 'e', 'g']
        df_p1 = pd.DataFrame(dict([(k, pd.Series(self.p1_dict.get(k, []))) for k in order_p1])).T
        df_p1.reset_index(inplace=True)
        df_p1.columns = ['Label'] + [f'Ayat {i+1}' for i in range(df_p1.shape[1]-1)]
        self.df_p1 = df_p1

        # 4. Merakit DataFrame Proses 2 (Reordering Eksklusif 15 Blok)
        order_p2 = ['F', 'A', 'a', 'B', 'b', 'C2', 'C1', 'c1', 'c2', 'D3', 'D2', 'D1', 'd1', 'd2', 'd3']
        df_p2 = pd.DataFrame(dict([(k, pd.Series(self.p1_dict.get(k, []))) for k in order_p2])).T
        df_p2.reset_index(inplace=True)
        df_p2.columns = ['Label'] + [f'Ayat {i+1}' for i in range(df_p2.shape[1]-1)]
        self.df_p2 = df_p2
        
        print(f"[OK] Proses 1 & Proses 2 ({self.metode_aktif}) Selesai.")

    def proses_3(self):
        """Tahap II.2.3: Flattening, Merger & Indexing Mutlak"""
        flat_data = []
        p3_id_counter = 1
        
        for idx, row in self.df_p2.iterrows():
            label = row['Label']
            for col in self.df_p2.columns[1:]:
                ayat = row[col]
                if pd.notna(ayat):
                    flat_data.append([p3_id_counter, label, ayat])
                    p3_id_counter += 1
                    
        self.df_p3 = pd.DataFrame(flat_data, columns=['p3_id', 'Label', 'Ayat'])
        
        try:
            self.at_h = self.df_p3.loc[38, 'Ayat'] # p3_id 39
            self.at_ah = self.df_p3.loc[49, 'Ayat'] # p3_id 50 (Dhzo)
        except KeyError:
            self.at_h = "Index Tidak Ditemukan"
            self.at_ah = "Index Tidak Ditemukan"

    def generate_tabel_jb_init(self):
        """Tahap II.2.4: Dashboard JB (Init)"""
        def get_padded_list(label, length=5):
            lst = self.p1_dict.get(label, [])
            return lst + [None] * (length - len(lst))

        tabel_jb_data = [
            ["Variabel", "Ayat1", "Ayat2", "Ayat3", "Ayat4", "Ayat5"],
            [f"Ayat Awal {self.metode_aktif}", self.ayat_awal, None, None, None, None],
            [f"Ayat Akhir {self.metode_aktif}", self.ayat_akhir, None, None, None, None],
            [f"Ayat Tengah {self.metode_aktif}", self.titik_tengah, None, None, None, None],
            [f"Ayat F {self.metode_aktif}"] + get_padded_list('F'),
            [f"Ayat A {self.metode_aktif}"] + get_padded_list('A'),
            [f"Ayat a {self.metode_aktif}"] + get_padded_list('a'),
            [f"Ayat Dhzo {self.metode_aktif}", self.at_ah, None, None, None, None],
            [f"Ayat E {self.metode_aktif}"] + get_padded_list('E'),
            [f"Ayat G {self.metode_aktif}"] + get_padded_list('G'),
            [f"Ayat e {self.metode_aktif}"] + get_padded_list('e'),
            [f"Ayat g {self.metode_aktif}"] + get_padded_list('g'),
            [f"Ayat AT-H {self.metode_aktif}", self.at_h, None, None, None, None],
            [f"Ayat AT-Juz {self.metode_aktif}", self.titik_tengah, None, None, None, None],
            [f"Ayat AT-AH {self.metode_aktif}", self.at_ah, None, None, None, None],
            ["Total Ayat", f"{self.total_ayat} ayat", None, None, None, None]
        ]
        
        kolom = [f'Tabel JB {self.metode_aktif} Juz {self.NILAI_FILTER}', 'Ayat1', 'Ayat2', 'Ayat3', 'Ayat4', 'Ayat5']
        self.df_init = pd.DataFrame(tabel_jb_data[1:], columns=kolom)

    def generate_tabel_ah_final(self):
        """Tahap II.2.5: Tabel Asmaul Husna"""
        latin_chars = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'I', 'II', 'III', 'Iv', 'V', 'VI', 'VII']
        arab_chars = ['ا', 'ب', 'ت', 'ث', 'ج', 'ح', 'خ', 'د', 'ذ', 'ر', 'ز', 'س', 'ش', 'ص', 'ض', 'ط', 'ظ', 'ع', 'غ', 'ف', 'ق', 'ك', 'ل', 'م', 'ن', 'و', 'ھ', 'لا', 'ء', 'ي', 'ال', 'ة', 'هؤ']
        
        baris_final: List[List[Any]] = []
        baris_final.append(['Konsep Dasar (metode jumlah ayat dalam juz)', None, None, None, None, None, None, None, None, None])
        baris_final.append([None, self.metode_aktif, None, None, None, None, None, None, f'Juz {self.NILAI_FILTER}', None])
        baris_final.append(['Latin & Hari', None, None, 'Ayat Khusus & Titik Kesadaran', None, None, None, 'Arab & Dasar', None, None])
        baris_final.append(['Latin', 'No AH', 'Ayat #3', 'No AH', 'Ayat #4', 'No AH', 'Ayat #2', 'No AH', 'Ayat #1', 'Arab'])
        
        for i in range(1, 34):
            idx_1 = 3 * i - 2
            idx_2 = 3 * i - 1
            idx_3 = 3 * i
            idx_4 = 101 - idx_3
            
            ayat1 = self.df_p3.loc[idx_1 - 1, 'Ayat'] if idx_1 <= len(self.df_p3) else None
            ayat2 = self.df_p3.loc[idx_2 - 1, 'Ayat'] if idx_2 <= len(self.df_p3) else None
            ayat3 = self.df_p3.loc[idx_3 - 1, 'Ayat'] if idx_3 <= len(self.df_p3) else None
            ayat4 = self.df_p3.loc[idx_4 - 1, 'Ayat'] if idx_4 <= len(self.df_p3) else None
            
            huruf_latin = latin_chars[i-1]
            huruf_arab = arab_chars[i-1]
            
            baris = [huruf_latin, idx_3, ayat3, idx_4, ayat4, idx_2, ayat2, idx_1, ayat1, huruf_arab]
            baris_final.append(baris)

        kolom = ['TABEL ASMAUL HUSNA', None, None, None, None, None, None, None, None, None]
        self.df_final = pd.DataFrame(baris_final, columns=kolom)

    def export_excel(self, output_name, jenis="advance"):
        """Ekspor hasil ke file Excel"""
        with pd.ExcelWriter(output_name, engine='openpyxl') as writer:
            if jenis == "advance":
                self.df_init.to_excel(writer, sheet_name='Init', index=False)
                self.df_p1.to_excel(writer, sheet_name='Proses 1', index=False)
                self.df_p2.to_excel(writer, sheet_name='Proses 2', index=False)
                self.df_p3.to_excel(writer, sheet_name='Proses 3', index=False)
                self.df_final.to_excel(writer, sheet_name='Final', index=False)
            else:
                self.df_init.to_excel(writer, sheet_name='Init', index=False)
                self.df_final.to_excel(writer, sheet_name='Final', index=False)
        print(f"[OK] File {self.metode_aktif} disimpan sebagai: {output_name}")

# ==========================================
# --- CLASS BARU UNTUK STYLING TABEL ---
# ==========================================
class FormatterTabelFinal:
    def __init__(self, file_path, sheet_name="Final"):
        self.file_path = file_path
        self.sheet_name = sheet_name

    def apply_formatting(self):
        try:
            wb = openpyxl.load_workbook(self.file_path)
            
            if self.sheet_name not in wb.sheetnames:
                print(f"[WARNING] Sheet '{self.sheet_name}' tidak ditemukan.")
                return
                
            ws = wb[self.sheet_name]

            # Style definitions
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
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
            thick_border = Border(
                left=Side(style='thick'), right=Side(style='thick'),
                top=Side(style='thick'), bottom=Side(style='thick')
            )

            # Format Header (Merge Cells & Styling)
            ws.merge_cells('A1:J1')
            ws['A1'].font = font_title
            ws['A1'].alignment = align_center

            ws.merge_cells('A2:J2')
            ws['A2'].font = font_subtitle
            ws['A2'].alignment = align_center

            ws.merge_cells('A4:C4')
            # ws['A4'].font = font_header
            ws['A4'].alignment = align_center
            ws['A4'].border = thick_border
            ws['A4'].font = font_label

            ws.merge_cells('D4:G4')
            # ws['D4'].font = font_header
            ws['D4'].alignment = align_center
            ws['D4'].border = thick_border
            ws['D4'].font = font_label

            ws.merge_cells('H4:J4')
            # ws['H4'].font = font_header
            ws['H4'].alignment = align_center
            ws['H4'].border = thick_border
            ws['H4'].font = font_label

            ws['B3'].font = font_label_italic
            ws['I3'].font = font_subtitle
                       
            
            # Format Kolom Header Aktual (Baris 5)
            for col in range(1, 11):
                cell = ws.cell(row=5, column=col)
                cell.font = font_header
                cell.fill = fill_header
                cell.alignment = align_center
                cell.border = thin_border

            # Format Body
            max_r = ws.max_row
            max_c = 10 
            
            for row in range(6, max_r + 1):
                for col in range(1, max_c + 1):
                    cell = ws.cell(row=row, column=col)
                    cell.alignment = align_center_top
                    cell.border = thin_border
                    
                    if col == 10:
                        cell.font = font_arab
                    elif col == 1:
                        cell.font = font_label_italic
                    elif col == 2 or col == 4 or col == 6 or col == 8:
                        cell.font = font_label_bold
                    else:
                        cell.font = font_body_regular
                        
                    if row % 2 == 0:
                        cell.fill = fill_alt_row

            # Lebar Kolom
            column_widths = {'A': 5, 'B': 7, 'C': 10, 'D': 4, 'E': 10, 'F': 4, 'G': 10, 'H': 7, 'I': 15, 'J': 5}
            for col_letter, width in column_widths.items():
                ws.column_dimensions[col_letter].width = width

            wb.save(self.file_path)
            print(f"[OK] Format dan Style berhasil diaplikasikan pada sheet '{self.sheet_name}'")

        except Exception as e:
            print(f"[ERROR] Gagal melakukan styling: {e}")

# ==========================================
# --- SIMULASI EKSEKUSI MULTI-METODE ---
# ==========================================
if __name__ == "__main__":
    print("--- MEMULAI PENGOLAH DATA JBAH (ALL METHODS) ---")
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    FILE_MENTAH = os.path.join(script_dir, "datasource", "Struktur Quran18.xlsx")
    KOLOM_FILTER = "aq_juz_id"
    NILAI_FILTER = 1  
    JENIS_OUTPUT = "advance"
    
    app = PengolahDataJBAH(FILE_MENTAH)
    
    if app.raw_data is not None:
        app.filter_data(KOLOM_FILTER, NILAI_FILTER)
        
        daftar_metode = ["T111", "T222", "T333"]
        
        output_dir = os.path.join(script_dir, "dataoutput")
        os.makedirs(output_dir, exist_ok=True)
        
        for metode in daftar_metode:
            NAMA_FILE_OUTPUT = os.path.join(output_dir, f"Output_JBAH_Juz{NILAI_FILTER}_{metode}.xlsx")
            
            app.set_metode(metode)
            app.proses_1_dan_2()
            app.proses_3()
            app.generate_tabel_jb_init()
            app.generate_tabel_ah_final()
            
            # Tahap ekspor file ke raw Excel 
            app.export_excel(NAMA_FILE_OUTPUT, jenis=JENIS_OUTPUT)
            
            # TAHAP BARU: Styling File Excel yg baru saja dibuat
            formatter = FormatterTabelFinal(NAMA_FILE_OUTPUT, sheet_name="Final")
            formatter.apply_formatting()
            
    print("\n[OK] SELURUH PROSES SELESAI")