"""
pola_baptis.py
==============
Program untuk menghasilkan tabel "Pola Baptis (Taubatan Nasuha)" dari data Master Tabel AH.

Fitur:
  - Pemilihan file master CSV secara interaktif
  - Pilihan mode output: per-Juz (30 file) atau all-Juz (1 file)
  - Proses otomatis Juz 1-30 × 3 Metode (T111, T222, T333)
  - Styling Excel mirip file asli "POLA BABTIS.xlsx"
  - 9 huruf Arab statis dengan konstanta angka baku
  - Kolom A & B: 1 konstanta (merge 2 cell), Kolom C: 2 konstanta

Cara pakai:
  python pola_baptis.py
"""

import os
import sys
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Import fungsi dari modul yang sudah ada
from query_ah import load_master_ah, ambil_ayat_dari_ah, list_available_master


# ============================================================
# KONSTANTA
# ============================================================
DAFTAR_HURUF_ARAB = ['ت', 'ج', 'خ', 'ز', 'ش', 'ظ', 'ف', 'ء', 'ة']
DAFTAR_METODE = ['T111', 'T222', 'T333']

# Konstanta angka per huruf (format baru: A, B, C2, C3)
# Kolom A dan B: 1 angka (akan di-merge 2 cell)
# Kolom C: 2 angka (tidak di-merge)
# Format: {huruf: (A, B, C2, C3)}
KONSTANTA_ANGKA = {
    'ت': (6, 13, 1, 16),
    'ج': (5, 11, 2, 15),
    'خ': (2, 12, 3, 14),
    'ز': (4, 10, 4, 13),
    'ش': (3, 9, 5, 12),
    'ظ': (1, 7, 6, 11),
    'ف': (15, 8, 7, 10),
    'ء': (14, 16, 8, 9),
    'ة': (None, None, None, None),  # Huruf ة tidak punya konstanta
}


# ============================================================
# FUNGSI STYLING
# ============================================================
def apply_pola_baptis_style(ws, juz_number, metode, total_huruf_ain):
    """
    Terapkan styling Excel untuk sheet Pola Baptis.
    """
    # --- Definisi Font ---
    font_title = Font(name="Arial", size=14, bold=True, color="000000")
    font_subtitle = Font(name="Arial", size=11, bold=True, italic=True)
    font_header = Font(name="Arial", size=11, bold=True, color="000000")
    font_body = Font(name="Arial", size=10)
    font_label = Font(name="Arial", size=10, bold=True)
    font_arab = Font(name="Arial", size=12, bold=True)
    font_const = Font(name="Arial", size=13, bold=True, color="FF0000")
    font_footer_label = Font(name="Arial", size=10, bold=True)
    font_footer_body = Font(name="Arial", size=10)
    font_signature = Font(name="Arial", size=9, italic=True, color="666666")
    
    # --- Definisi Alignment ---
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_center_top = Alignment(horizontal="center", vertical="top", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    
    # --- Definisi Border ---
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    
    # --- Definisi Fill ---
    fill_alt_row = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    
    # --- Baris 1: Judul ---
    ws.merge_cells("A1:N1")
    ws["A1"] = "Pola Baptis (Taubatan Nasuha)"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_center
    
    # --- Baris 2: Sub-judul ---
    ws.merge_cells("A2:N2")
    ws["A2"] = "Dosa-dosa dimandikan/dibersihkan dengan air (AIN)"
    ws["A2"].font = font_subtitle
    ws["A2"].alignment = align_center
    
    # --- Baris 3: Info Juz, Total Huruf Ain, Metode ---
    ws["C3"] = f"Juz: {juz_number}"
    ws["C3"].font = font_label
    ws["C3"].alignment = align_center
    
    # ws["F3"] = f"Total Huruf Ain: {total_huruf_ain}"
    ws["F3"].font = font_label
    ws["F3"].alignment = align_center
    
    ws["N3"] = f"Tabel {metode}"
    ws["N3"].font = font_label
    ws["N3"].alignment = align_center
    
    # --- Baris 4: Header Kolom ---
    headers = ["No", "Huruf", "A", "A", "A", "A", "B", "B", "B", "B", "C", "C", "C", "C"]
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_idx, value=header)
        cell.font = font_header
        cell.alignment = align_center
        cell.border = thin_border
    
    # Merge header A (kolom 3-6), B (kolom 7-10), C (kolom 11-14)
    ws.merge_cells("C4:F4")
    ws.merge_cells("G4:J4")
    ws.merge_cells("K4:N4")
    
    # --- Baris 5+: Data (9 huruf × 2 baris = 18 baris) ---
    current_row = 5
    for no_huruf, huruf_arab in enumerate(DAFTAR_HURUF_ARAB, 1):
        # Baris ayat
        ws.cell(row=current_row, column=1, value=no_huruf).font = font_label
        ws.cell(row=current_row, column=1).alignment = align_center
        ws.cell(row=current_row, column=1).border = thin_border
        
        ws.cell(row=current_row, column=2, value=huruf_arab).font = font_arab
        ws.cell(row=current_row, column=2).alignment = align_center
        ws.cell(row=current_row, column=2).border = thin_border
        
        # Baris angka konstanta
        ws.cell(row=current_row + 1, column=1).border = thin_border
        ws.cell(row=current_row + 1, column=2).border = thin_border
        
        # Apply border dan styling untuk kolom ayat (3-14)
        for col in range(3, 15):
            ws.cell(row=current_row, column=col).border = thin_border
            ws.cell(row=current_row, column=col).alignment = align_center_top
            ws.cell(row=current_row, column=col).font = font_body
            
            ws.cell(row=current_row + 1, column=col).border = thin_border
            ws.cell(row=current_row + 1, column=col).alignment = align_center
            ws.cell(row=current_row + 1, column=col).font = font_const
        
        # Warna selang-seling
        if no_huruf % 2 == 0:
            for col in range(1, 15):
                ws.cell(row=current_row, column=col).fill = fill_alt_row
                ws.cell(row=current_row + 1, column=col).fill = fill_alt_row
        
        current_row += 2
    
    # ========================================================
    # FOOTER: Pisahkan merge A-G dan H-N agar tidak konflik
    # ========================================================
    footer_start_row = current_row + 1
    
    # Baris 1 Footer: CATATAN & CARA BACA
    ws.merge_cells(f"A{footer_start_row}:G{footer_start_row}")
    ws[f"A{footer_start_row}"] = "CATATAN:"
    ws[f"A{footer_start_row}"].font = font_footer_label
    ws[f"A{footer_start_row}"].alignment = align_left
    
    ws.merge_cells(f"H{footer_start_row}:N{footer_start_row}")
    ws[f"H{footer_start_row}"] = "CARA BACA:"
    ws[f"H{footer_start_row}"].font = font_footer_label
    ws[f"H{footer_start_row}"].alignment = align_left
    
    # Baris 2 Footer
    ws.merge_cells(f"A{footer_start_row + 1}:G{footer_start_row + 1}")
    ws[f"A{footer_start_row + 1}"] = "KOLOM A : KABAH"
    ws[f"A{footer_start_row + 1}"].font = font_footer_body
    ws[f"A{footer_start_row + 1}"].alignment = align_left
    
    ws.merge_cells(f"H{footer_start_row + 1}:N{footer_start_row + 1}")
    ws[f"H{footer_start_row + 1}"] = "1. A + B : SEKALI DUDUK"
    ws[f"H{footer_start_row + 1}"].font = font_footer_body
    ws[f"H{footer_start_row + 1}"].alignment = align_left
    
    # Baris 3 Footer
    ws.merge_cells(f"A{footer_start_row + 2}:G{footer_start_row + 2}")
    ws[f"A{footer_start_row + 2}"] = "KOLOM B : SUJUD"
    ws[f"A{footer_start_row + 2}"].font = font_footer_body
    ws[f"A{footer_start_row + 2}"].alignment = align_left
    
    ws.merge_cells(f"H{footer_start_row + 2}:N{footer_start_row + 2}")
    ws[f"H{footer_start_row + 2}"] = "2. C : SEKALI DUDUK"
    ws[f"H{footer_start_row + 2}"].font = font_footer_body
    ws[f"H{footer_start_row + 2}"].alignment = align_left
    
    # Baris 4 Footer
    ws.merge_cells(f"A{footer_start_row + 3}:G{footer_start_row + 3}")
    ws[f"A{footer_start_row + 3}"] = "KOLOM C : KONTRUKSI SHALAT"
    ws[f"A{footer_start_row + 3}"].font = font_footer_body
    ws[f"A{footer_start_row + 3}"].alignment = align_left
    
    # Baris 5 Footer
    ws.merge_cells(f"A{footer_start_row + 4}:N{footer_start_row + 4}")
    ws[f"A{footer_start_row + 4}"] = ": HALAMAN JUZ"
    ws[f"A{footer_start_row + 4}"].font = font_footer_body
    ws[f"A{footer_start_row + 4}"].alignment = align_left
    
    # --- Signature ---
    sig_row = footer_start_row + 6
    ws.merge_cells(f"A{sig_row}:N{sig_row}")
    current_month = datetime.now().strftime("%B")
    current_year = datetime.now().strftime("%Y")
    ws[f"A{sig_row}"] = f"dirumuskan oleh: BWMH21 | disusun: di Depok, {current_month} {current_year}"
    ws[f"A{sig_row}"].font = font_signature
    ws[f"A{sig_row}"].alignment = align_center
    
    # --- Set Column Widths ---
    column_widths = {
        "A": 5,   # No
        "B": 8,   # Huruf
        "C": 10,  # A1
        "D": 10,  # A2
        "E": 10,  # A3
        "F": 10,  # A4
        "G": 10,  # B1
        "H": 10,  # B2
        "I": 10,  # B3
        "J": 10,  # B4
        "K": 10,  # C1
        "L": 10,  # C2
        "M": 10,  # C3
        "N": 10,  # C4
    }
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width


# ============================================================
# FUNGSI UTAMA: BUAT SATU SHEET POLA BAPTIS
# ============================================================
def buat_sheet_pola_baptis(wb, sheet_name, juz_number, metode, df_master, total_huruf_ain):
    """
    Buat satu sheet Pola Baptis untuk kombinasi Juz dan Metode tertentu.
    """
    ws = wb.create_sheet(title=sheet_name)
    
    # Apply styling dasar
    apply_pola_baptis_style(ws, juz_number, metode, total_huruf_ain)
    
    # --- Isi data ayat ---
    current_row = 5
    for huruf_arab in DAFTAR_HURUF_ARAB:
        # Ambil 4 ayat dengan susunan: ayat_1, ayat_2, ayat_4, ayat_3
        daftar_kolom = ['ayat_1', 'ayat_2', 'ayat_4', 'ayat_3']
        
        # Ambil ayat untuk kolom A, B, C (semua sama)
        ayat_A = ambil_ayat_dari_ah(juz_number, metode, huruf_arab, daftar_kolom, df_master)
        ayat_B = ambil_ayat_dari_ah(juz_number, metode, huruf_arab, daftar_kolom, df_master)
        ayat_C = ambil_ayat_dari_ah(juz_number, metode, huruf_arab, daftar_kolom, df_master)
        
        # Isi kolom A (kolom 3-6)
        for i, ayat in enumerate(ayat_A):
            ws.cell(row=current_row, column=3 + i, value=ayat)
        
        # Isi kolom B (kolom 7-10)
        for i, ayat in enumerate(ayat_B):
            ws.cell(row=current_row, column=7 + i, value=ayat)
        
        # Isi kolom C (kolom 11-14)
        for i, ayat in enumerate(ayat_C):
            ws.cell(row=current_row, column=11 + i, value=ayat)
        
        # ========================================================
        # ISI ANGKA KONSTANTA (baris current_row + 1)
        # Format baru: (A, B, C2, C3)
        # ========================================================
        const_A, const_B, const_C2, const_C3 = KONSTANTA_ANGKA[huruf_arab]
        
        if const_A is not None:
            # Kolom A: Merge cell kolom 4-5 (D-E), tulis angka di kolom 4
            ws.merge_cells(start_row=current_row + 1, start_column=4, 
                          end_row=current_row + 1, end_column=5)
            ws.cell(row=current_row + 1, column=4, value=const_A)
            
            # Kolom B: Merge cell kolom 8-9 (H-I), tulis angka di kolom 8
            ws.merge_cells(start_row=current_row + 1, start_column=8, 
                          end_row=current_row + 1, end_column=9)
            ws.cell(row=current_row + 1, column=8, value=const_B)
            
            # Kolom C: Tulis 2 angka terpisah di kolom 12 dan 13 (L-M)
            ws.cell(row=current_row + 1, column=12, value=const_C2)
            ws.cell(row=current_row + 1, column=13, value=const_C3)
        
        current_row += 2
    
    return ws


# ============================================================
# FUNGSI UTAMA: GENERATE POLA BAPTIS
# ============================================================
def generate_pola_baptis():
    """
    Fungsi utama untuk menghasilkan file Excel Pola Baptis.
    """
    print("=" * 70)
    print(" GENERATOR POLA BAPTIS (TAUBATAN NASUHA)")
    print("=" * 70)
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # ========================================================
    # TAHAP 1: Pemilihan File Master CSV
    # ========================================================
    print("\n--- PILIH FILE MASTER CSV ---")
    available_masters = list_available_master(script_dir)
    
    if not available_masters:
        print(f"[ERROR] Tidak ada file master CSV (master_tabel_ah_*.csv) di folder: {script_dir}")
        print("        Silakan jalankan 'generate_master_ah.py' terlebih dahulu.")
        sys.exit(1)
    
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
    
    # ========================================================
    # TAHAP 2: Load Master CSV
    # ========================================================
    print(f"\n[INFO] Memuat master CSV: {file_master}")
    df_master = load_master_ah(file_master)
    if df_master is None:
        print("[FATAL] Gagal memuat master CSV. Proses dihentikan.")
        sys.exit(1)
    
    # ========================================================
    # TAHAP 3: Pemilihan Mode Output
    # ========================================================
    print("\n--- PILIH MODE OUTPUT ---")
    print("  a. Per-Juz (30 file: Pola_Baptis_Juz01.xlsx, Pola_Baptis_Juz02.xlsx, ...)")
    print("  b. All-Juz (1 file: Pola_Baptis_AllJuz.xlsx dengan 90 sheet)")
    
    while True:
        mode_choice = input("\nMasukkan pilihan (a/b): ").strip().lower()
        if mode_choice in ['a', 'b']:
            break
        print("[ERROR] Pilihan tidak valid. Masukkan 'a' atau 'b'.")
    
    # ========================================================
    # TAHAP 4: Proses dan Generate Excel
    # ========================================================
    output_dir = os.path.join(script_dir, "dataoutput")
    os.makedirs(output_dir, exist_ok=True)
    
    if mode_choice == 'a':
        # Mode Per-Juz
        print("\n[INFO] Mode: Per-Juz. Memproses 30 Juz × 3 Metode...")
        
        for juz_number in range(1, 31):
            print(f"\n[INFO] === Memproses Juz {juz_number} ===")
            
            from backend1 import PengolahDataJBAH
            
            datasource_dir = os.path.join(script_dir, "datasource")
            source_files = [f for f in os.listdir(datasource_dir) if f.endswith('.xlsx')]
            if not source_files:
                print(f"[ERROR] Tidak ada file sumber di folder datasource.")
                continue
            
            file_sumber = os.path.join(datasource_dir, source_files[0])
            app = PengolahDataJBAH(file_sumber)
            
            if app.filter_data("aq_juz_id", juz_number) is None:
                print(f"[WARN] Tidak ada data untuk Juz {juz_number}. Dilewati.")
                continue
            
            total_huruf_ain = app.total_ayat
            
            wb = Workbook()
            if wb.active is not None:
                wb.remove(wb.active)
            
            for metode in DAFTAR_METODE:
                sheet_name = metode
                print(f"  [INFO] Membuat sheet: {sheet_name}")
                buat_sheet_pola_baptis(wb, sheet_name, juz_number, metode, df_master, total_huruf_ain)
            
            juz_str = f"{juz_number:02d}"
            output_file = os.path.join(output_dir, f"Pola_Baptis_Juz{juz_str}.xlsx")
            
            try:
                wb.save(output_file)
                print(f"  [OK] File disimpan: {output_file}")
            except Exception as e:
                print(f"  [ERROR] Gagal menyimpan file: {e}")
    
    else:
        # Mode All-Juz
        print("\n[INFO] Mode: All-Juz. Memproses 30 Juz × 3 Metode...")
        
        wb = Workbook()
        if wb.active is not None:
            wb.remove(wb.active)
        
        for juz_number in range(1, 31):
            print(f"\n[INFO] === Memproses Juz {juz_number} ===")
            
            from backend1 import PengolahDataJBAH
            datasource_dir = os.path.join(script_dir, "datasource")
            source_files = [f for f in os.listdir(datasource_dir) if f.endswith('.xlsx')]
            if not source_files:
                print(f"[ERROR] Tidak ada file sumber di folder datasource.")
                continue
            
            file_sumber = os.path.join(datasource_dir, source_files[0])
            app = PengolahDataJBAH(file_sumber)
            
            if app.filter_data("aq_juz_id", juz_number) is None:
                print(f"[WARN] Tidak ada data untuk Juz {juz_number}. Dilewati.")
                continue
            
            total_huruf_ain = app.total_ayat
            
            for metode in DAFTAR_METODE:
                juz_str = f"{juz_number:02d}"
                sheet_name = f"Juz{juz_str}_{metode}"
                print(f"  [INFO] Membuat sheet: {sheet_name}")
                buat_sheet_pola_baptis(wb, sheet_name, juz_number, metode, df_master, total_huruf_ain)
        
        output_file = os.path.join(output_dir, "Pola_Baptis_AllJuz.xlsx")
        
        try:
            wb.save(output_file)
            print(f"\n[SUCCESS] File disimpan: {output_file}")
        except Exception as e:
            print(f"\n[ERROR] Gagal menyimpan file: {e}")
    
    print("\n" + "=" * 70)
    print(" PROSES SELESAI")
    print("=" * 70)


# ============================================================
# ENTRY POINT
# ============================================================
if __name__ == "__main__":
    generate_pola_baptis()