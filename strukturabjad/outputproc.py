import os
from datetime import datetime

import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


def write_excel_output(processor, output_name, jenis="advance"):
    """Ekspor DataFrame hasil pemrosesan ke file Excel."""
    with pd.ExcelWriter(output_name, engine="openpyxl") as writer:
        if jenis == "advance":
            processor.df_init.to_excel(writer, sheet_name="Init", index=False)
            processor.df_p1.to_excel(writer, sheet_name="Proses 1", index=False)
            processor.df_p2.to_excel(writer, sheet_name="Proses 2", index=False)
            processor.df_p3.to_excel(writer, sheet_name="Proses 3", index=False)
            processor.df_final.to_excel(writer, sheet_name="Final", index=False)
        else:
            processor.df_init.to_excel(writer, sheet_name="Init", index=False)
            processor.df_final.to_excel(writer, sheet_name="Final", index=False)

    print(f"[SAVED] File {processor.metode_aktif} disimpan sebagai: {output_name}")
    return output_name


def export_and_format_output(processor, output_name, jenis="advance", sheet_name="Final", signature_text=None):
    """Ekspor hasil pemrosesan lalu aplikasikan format Excel."""
    write_excel_output(processor, output_name, jenis=jenis)
    formatter = FormatterTabelFinal(output_name, sheet_name=sheet_name, signature_text=signature_text or processor.signature_text)
    formatter.apply_formatting()
    return output_name


class FormatterTabelFinal:
    def __init__(self, file_path, sheet_name="Final", signature_text="Digital Signature"):
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.signature_text = signature_text

    def _get_target_sheets(self, wb):
        if self.sheet_name == "Final":
            return [name for name in ["Init", "Proses 1", "Proses 2", "Proses 3", "Final"] if name in wb.sheetnames]
        if self.sheet_name in wb.sheetnames:
            return [self.sheet_name]
        return []

    def _apply_standard_style(self, ws):
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
                print(f"[WARNING] Sheet '{self.sheet_name}' tidak ditemukan.")
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
