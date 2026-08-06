# outputproc.py

import os
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# --- Fungsi untuk ekspor ke HTML ---
def export_and_format_html(processor, output_name, jenis="pengguna", signature_text=None):
    """Ekspor hasil ke file HTML."""
    print(f"[INFO] Mengekspor ke HTML: {output_name} (jenis: {jenis})")

    # Ganti NaN dengan string kosong sebelum konversi ke HTML
    df_init_clean = processor.df_init.fillna("") if hasattr(processor, 'df_init') and processor.df_init is not None else None
    df_p1_clean = processor.df_p1.fillna("") if hasattr(processor, 'df_p1') and processor.df_p1 is not None else None
    df_p2_clean = processor.df_p2.fillna("") if hasattr(processor, 'df_p2') and processor.df_p2 is not None else None
    df_p3_clean = processor.df_p3.fillna("") if hasattr(processor, 'df_p3') and processor.df_p3 is not None else None
    df_final_clean = processor.df_final.fillna("") if hasattr(processor, 'df_final') and processor.df_final is not None else None

    html_parts = []
    # Judul Halaman
    html_parts.append(f"<html><head><title>Output JBAH - {processor.metode_aktif}</title>")
    # CSS Sederhana untuk styling dasar dan Final
    html_parts.append("<style>")
    html_parts.append("body { font-family: Arial, sans-serif; margin: 20px; }")
    html_parts.append("h1, h2 { color: #1F497D; }")
    html_parts.append("table { border-collapse: collapse; width: 100%; margin-bottom: 20px; table-layout: fixed; }") # table-layout: fixed; untuk lebar kolom konsisten
    html_parts.append("th, td { border: 1px solid black; padding: 8px; text-align: center; vertical-align: middle; overflow: hidden; text-overflow: ellipsis; }") # vertical-align: middle; untuk align vertikal
    html_parts.append("th { background-color: #1F497D; color: white; font-weight: bold; }")
    html_parts.append("tr:nth-child(even) { background-color: #f2f2f2; }")
    # Styling khusus untuk sheet Final (mengikuti struktur Final di Excel)
    # Header Konsep
    html_parts.append(".final-concept { font-size: 14pt; font-weight: bold; text-align: center; background-color: #E9F2FF; }")
    # Header Metode/Juz
    html_parts.append(".final-subtitle { font-size: 11pt; font-weight: bold; font-style: italic; text-align: center; background-color: #E9F2FF; }")
    # Sub-header
    html_parts.append(".final-sub-header { font-weight: bold; text-align: center; border: 3px solid black; }")
    # Header kolom utama
    html_parts.append(".final-main-header { background-color: #1F497D !important; }") # Override warna th
    # Kolom Arab
    html_parts.append(".final-arab-col { font-weight: bold; font-size: 12pt; }")
    # Kolom Label (Latin, No AH)
    html_parts.append(".final-label-col { font-weight: bold; }")
    # Footer
    html_parts.append(".signature-title { text-align: center; background-color: #E9F2FF; font-weight: bold; font-size: 12pt; color: #1F497D; border: 1px solid black; }")
    html_parts.append(".signature-stamp, .signature-time { text-align: center; border: 1px solid black; }")
    # Lebar kolom untuk tabel Final
    html_parts.append(".final-table { width: 100%; }")
    html_parts.append(".final-table td:nth-child(1) { width: 5%; }") # Latin
    html_parts.append(".final-table td:nth-child(2) { width: 7%; }") # No AH
    html_parts.append(".final-table td:nth-child(3) { width: 10%; }") # Ayat #3
    html_parts.append(".final-table td:nth-child(4) { width: 4%; }") # No AH
    html_parts.append(".final-table td:nth-child(5) { width: 10%; }") # Ayat #4
    html_parts.append(".final-table td:nth-child(6) { width: 4%; }") # No AH
    html_parts.append(".final-table td:nth-child(7) { width: 10%; }") # Ayat #2
    html_parts.append(".final-table td:nth-child(8) { width: 7%; }") # No AH
    html_parts.append(".final-table td:nth-child(9) { width: 15%; }") # Ayat #1
    html_parts.append(".final-table td:nth-child(10) { width: 5%; }") # Arab
    html_parts.append("</style>")
    html_parts.append("</head><body>")

    # --- Ekspor sheet berdasarkan jenis ---
    if jenis == "pengguna":
        # Hanya Init dan Final
        # --- Sheet Init ---
        if df_init_clean is not None and not df_init_clean.empty:
            html_parts.append(f"<h2>Sheet: Init ({processor.metode_aktif})</h2>")
            html_parts.append(df_init_clean.to_html(index=False, escape=False, classes='standard-table'))

        # --- Sheet Final ---
        if df_final_clean is not None and not df_final_clean.empty:
            html_parts.append(f"<h2>Sheet: Final ({processor.metode_aktif})</h2>")
            # Bangun tabel Final secara manual untuk kontrol lebih baik atas merging cells
            html_parts.append('<table class="final-table">')
            # Baris 1: Konsep Dasar (merge semua kolom)
            html_parts.append(f'<tr class="final-concept-row"><td colspan="10" class="final-concept">{df_final_clean.iloc[0, 0]}</td></tr>')
            # Baris 2: Metode dan Juz (merge kolom 2-10)
            html_parts.append(f'<tr class="final-subtitle-row"><td></td><td colspan="8" class="final-subtitle">{df_final_clean.iloc[1, 1]}</td><td></td></tr>') # Kolom pertama dan terakhir kosong
            # Baris 3: Sub Headers (merge dan styling)
            html_parts.append('<tr class="final-sub-header">')
            html_parts.append(f'<td colspan="3" class="final-sub-header">{df_final_clean.iloc[2, 0]}</td>') # Latin & Hari
            html_parts.append(f'<td colspan="4" class="final-sub-header">{df_final_clean.iloc[2, 3]}</td>') # Ayat Khusus
            html_parts.append(f'<td colspan="3" class="final-sub-header">{df_final_clean.iloc[2, 7]}</td>') # Arab & Dasar
            html_parts.append('</tr>')
            # Baris 4: Header Kolom Utama
            html_parts.append('<tr class="final-main-header">')
            for header in df_final_clean.columns:
                html_parts.append(f'<th class="final-main-header">{header}</th>')
            html_parts.append('</tr>')
            # Baris Data
            for index, row in df_final_clean.iloc[3:].iterrows(): # Lewati baris 0-2 karena sudah dihandle
                html_parts.append('<tr>')
                for i, cell in enumerate(row):
                    cell_class = ""
                    if i == 0: # Kolom Latin
                        cell_class = "final-label-col"
                    elif i in [1, 3, 5, 7]: # Kolom No AH
                        cell_class = "final-label-col"
                    elif i == 9: # Kolom Arab
                        cell_class = "final-arab-col"
                    html_parts.append(f'<td class="{cell_class}">{cell}</td>')
                html_parts.append('</tr>')
            html_parts.append('</table>')

            # Tambahkan footer ke tabel Final
            sig_text = signature_text or processor.signature_text
            html_parts.append(f'<div class="signature-title">__</div>')
            html_parts.append(f'<div class="signature-stamp">: {sig_text}</div>')
            html_parts.append(f'<div class="signature-time">: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</div>')

    elif jenis == "praktisi":
        # Semua sheet
        sheets_data = [
            ("Init", df_init_clean),
            ("Proses 1", df_p1_clean),
            ("Proses 2", df_p2_clean),
            ("Proses 3", df_p3_clean),
            ("Final", df_final_clean),
        ]

        for sheet_name, df in sheets_data:
            if df is not None and not df.empty:
                html_parts.append(f"<h2>Sheet: {sheet_name} ({processor.metode_aktif})</h2>")
                if sheet_name != "Final":
                    html_parts.append(df.to_html(index=False, escape=False, classes='standard-table'))
                else: # Untuk Final, gunakan struktur manual
                    html_parts.append('<table class="final-table">')
                    html_parts.append(f'<tr class="final-concept-row"><td colspan="10" class="final-concept">{df.iloc[0, 0]}</td></tr>')
                    html_parts.append(f'<tr class="final-subtitle-row"><td></td><td colspan="8" class="final-subtitle">{df.iloc[1, 1]}</td><td></td></tr>')
                    html_parts.append('<tr class="final-sub-header">')
                    html_parts.append(f'<td colspan="3" class="final-sub-header">{df.iloc[2, 0]}</td>')
                    html_parts.append(f'<td colspan="4" class="final-sub-header">{df.iloc[2, 3]}</td>')
                    html_parts.append(f'<td colspan="3" class="final-sub-header">{df.iloc[2, 7]}</td>')
                    html_parts.append('</tr>')
                    html_parts.append('<tr class="final-main-header">')
                    for header in df.columns:
                        html_parts.append(f'<th class="final-main-header">{header}</th>')
                    html_parts.append('</tr>')
                    for index, row in df.iloc[3:].iterrows():
                        html_parts.append('<tr>')
                        for i, cell in enumerate(row):
                            cell_class = ""
                            if i == 0:
                                cell_class = "final-label-col"
                            elif i in [1, 3, 5, 7]:
                                cell_class = "final-label-col"
                            elif i == 9:
                                cell_class = "final-arab-col"
                            html_parts.append(f'<td class="{cell_class}">{cell}</td>')
                        html_parts.append('</tr>')
                    html_parts.append('</table>')

                    sig_text = signature_text or processor.signature_text
                    html_parts.append(f'<div class="signature-title">__</div>')
                    html_parts.append(f'<div class="signature-stamp">: {sig_text}</div>')
                    html_parts.append(f'<div class="signature-time">: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</div>')

    else:
        print(f"[WARNING] Jenis output '{jenis}' tidak dikenal untuk HTML. Gunakan 'pengguna' atau 'praktisi'.")
        return # Batalkan ekspor jika jenis tidak valid

    html_parts.append("</body></html>")
    full_html = "".join(html_parts)

    try:
        with open(output_name, 'w', encoding='utf-8') as f:
            f.write(full_html)
        print(f"[SAVED] File HTML {processor.metode_aktif} disimpan sebagai: {output_name}")
    except Exception as e:
        print(f"[ERROR] Gagal menyimpan file HTML '{output_name}': {e}")
        import traceback
        traceback.print_exc()


# --- Modifikasi fungsi write_excel_output ---
def write_excel_output(processor, output_name, jenis="pengguna"):
    """Ekspor DataFrame hasil pemrosesan ke file Excel."""
    with pd.ExcelWriter(output_name, engine="openpyxl") as writer:
        if jenis == "praktisi":
            if hasattr(processor, 'df_init') and processor.df_init is not None:
                processor.df_init.to_excel(writer, sheet_name="Init", index=False)
            if hasattr(processor, 'df_p1') and processor.df_p1 is not None:
                processor.df_p1.to_excel(writer, sheet_name="Proses 1", index=False)
            if hasattr(processor, 'df_p2') and processor.df_p2 is not None:
                processor.df_p2.to_excel(writer, sheet_name="Proses 2", index=False)
            if hasattr(processor, 'df_p3') and processor.df_p3 is not None:
                processor.df_p3.to_excel(writer, sheet_name="Proses 3", index=False)
            if hasattr(processor, 'df_final') and processor.df_final is not None:
                processor.df_final.to_excel(writer, sheet_name="Final", index=False)
        elif jenis == "pengguna":
            if hasattr(processor, 'df_init') and processor.df_init is not None:
                processor.df_init.to_excel(writer, sheet_name="Init", index=False)
            if hasattr(processor, 'df_final') and processor.df_final is not None:
                processor.df_final.to_excel(writer, sheet_name="Final", index=False)
        else:
            print(f"[WARNING] Jenis output '{jenis}' tidak dikenal, menggunakan default 'pengguna'.")
            if hasattr(processor, 'df_init') and processor.df_init is not None:
                processor.df_init.to_excel(writer, sheet_name="Init", index=False)
            if hasattr(processor, 'df_final') and processor.df_final is not None:
                processor.df_final.to_excel(writer, sheet_name="Final", index=False)
    print(f"[SAVED] File Excel {processor.metode_aktif} disimpan sebagai: {output_name}")
    return output_name

# --- Modifikasi fungsi export_and_format_output ---
def export_and_format_output(processor, output_name, jenis="pengguna", format="excel", sheet_name="Final", signature_text=None):
    """Ekspor hasil pemrosesan lalu aplikasikan format Excel atau buat HTML."""
    if format.lower() == "excel":
        write_excel_output(processor, output_name, jenis=jenis)
        formatter = FormatterTabelFinal(output_name, sheet_name=sheet_name, signature_text=signature_text or processor.signature_text)
        formatter.apply_formatting()
    elif format.lower() == "html":
        # Panggil fungsi HTML baru
        export_and_format_html(processor, output_name, jenis=jenis, signature_text=signature_text or processor.signature_text)
        # Note: HTML tidak menerapkan formatting Excel
    else:
        print(f"[ERROR] Format '{format}' tidak didukung. Gunakan 'excel' atau 'html'.")
        return
    return output_name

# --- Kelas FormatterTabelFinal tetap sama (untuk Excel) ---
class FormatterTabelFinal:
    def __init__(self, file_path, sheet_name="Final", signature_text="Digital Signature"):
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.signature_text = signature_text

    def _get_target_sheets(self, wb):
        all_possible_sheets = ["Init", "Proses 1", "Proses 2", "Proses 3", "Final"]
        if self.sheet_name == "Final":
            return [name for name in all_possible_sheets if name in wb.sheetnames]
        elif self.sheet_name == "All":
             return [name for name in wb.sheetnames if name in all_possible_sheets]
        else:
             if self.sheet_name in wb.sheetnames:
                 return [self.sheet_name]
             else:
                 print(f"[WARNING] Sheet '{self.sheet_name}' tidak ditemukan di workbook.")
                 return []

    def _apply_standard_style(self, ws):
        font_header = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        font_body = Font(name="Arial", size=10)
        font_label = Font(name="Arial", size=10, bold=True)
        font_title = Font(name="Arial", size=12, bold=True)
        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
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

    def _apply_final_style(self, ws):
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
        signature_stamp = f": {self.signature_text}"
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

    def apply_formatting(self):
        try:
            wb = openpyxl.load_workbook(self.file_path)
            target_sheets = self._get_target_sheets(wb)

            if not target_sheets:
                print(f"[WARNING] Tidak ada sheet target yang ditemukan untuk formatting: {self.sheet_name}")
                return

            for sheet_name in target_sheets:
                ws = wb[sheet_name]
                if sheet_name == "Final":
                    self._apply_final_style(ws)
                else:
                    self._apply_standard_style(ws)

            wb.save(self.file_path)
            print(f"[OK] Format dan Style berhasil diaplikasikan pada sheet(s): {', '.join(target_sheets)}")
        except Exception as e:
            print(f"[ERROR] Gagal melakukan styling: {e}")
