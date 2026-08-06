# pdf_export.py
# Kode ekspor PDF yang sepenuhnya bergantung pada reportlab
# Harus diimpor hanya jika reportlab tersedia

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, PageBreak, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.units import inch # Import inch secara eksplisit
import pandas as pd # Import pandas untuk pd.isna
import io
import os

# --- Canvas kustom untuk penomoran halaman (versi sederhana dan aman) ---
class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        # Tidak perlu inisialisasi variabel khusus untuk penomoran

    def onFirstPage(self, canvas, doc):
        # Fungsi ini dipanggil sekali saat halaman pertama dibuat
        # Biasanya digunakan untuk header halaman pertama
        # Kita abaikan untuk penomoran halaman biasa
        pass

    def onLaterPages(self, canvas, doc):
        # Fungsi ini dipanggil untuk setiap halaman setelah halaman pertama
        # Kita bisa menempatkan nomor halaman di sini jika ingin berbeda untuk hal. 1 dan seterusnya
        # Tapi untuk penomoran standar, kita gunakan showPage
        pass

    def showPage(self):
        # Fungsi ini dipanggil setiap kali halaman selesai dirender
        # Tempat yang ideal untuk menambahkan nomor halaman
        # reportlab secara otomatis menambahkan 1 ke _pageNumber setiap kali showPage dipanggil
        self.draw_page_number()
        canvas.Canvas.showPage(self) # Wajib: panggil fungsi asli showPage

    def draw_page_number(self):
        # Fungsi untuk menggambar nomor halaman
        # Gunakan self._pageNumber yang disediakan oleh reportlab
        # Lokasi nomor halaman
        x_position = 7.5 * inch
        y_position = 0.5 * inch
        # Format nomor halaman (misalnya "Page X of Y")
        # Catatan: self._currentPages disediakan oleh doc, bukan canvas
        # Untuk mendapatkan total halaman, kita perlu trik lain atau menetapkan nilainya nanti
        # Untuk sementara, kita gunakan nomor halaman saat ini saja
        # Kita bisa menetapkan total halaman di canvas saat save
        page_num_text = f"Page {self.getPageNumber()}" # getPageNumber() adalah metode resmi reportlab
        self.setFont("Helvetica", 10)
        self.drawRightString(x_position, y_position, page_num_text)

    def save(self):
        # Fungsi ini dipanggil saat dokumen selesai dibuat
        # Biasanya digunakan untuk membersihkan atau menyimpan state akhir
        # Kita cukup panggil fungsi asli
        canvas.Canvas.save(self)

def clean_for_reportlab(value):
    """Fungsi untuk membersihkan nilai sebelum dimasukkan ke reportlab."""
    if pd.isna(value) or value is None:
        return "" # Ganti NaN atau None dengan string kosong
    # Coba encode/decode untuk memastikan string bisa ditulis ke PDF
    try:
        s_val = str(value)
        cleaned_bytes = s_val.encode('utf-8', errors='replace')
        return cleaned_bytes.decode('utf-8', errors='replace')
    except UnicodeDecodeError:
        print(f"[DEBUG] Karakter aneh terdeteksi: {value}, diganti dengan ''")
        return ""

def df_to_reportlab_table_final(df, processor): # Tambahkan processor sebagai argumen untuk akses metode dan sign
    """Konversi DataFrame Final ke objek Table reportlab dengan styling Final."""
    # Data asli dari df_final
    df_values = df.values.tolist()

    # --- Bangun struktur tabel untuk reportlab ---
    # Header Row (kolom 1-10)
    # Kita gunakan header dari kolom DataFrame, tetapi struktur Final spesifik
    # Kolom 1: Konsep Dasar (span 10 cols)
    # Kolom 2: Metode dan Juz (span 10 cols)
    # Kolom 3: Header sub judul (span 3, 4, 3 cols)
    # Kolom 4: Header kolom utama (10 kolom)
    # Kolom 5 dst: Data (10 kolom)
    # Baris 1: Konsep Dasar
    row_1 = [df.columns[0]] + [""] * 9 # Span ke 10 kolom
    # Baris 2: Metode dan Juz
    row_2 = [processor.metode_aktif] + [""] * 6 + [f'Juz {processor.NILAI_FILTER}'] + [""] * 2 # Sesuaikan span
    # Baris 3: Sub Judul Header
    row_3 = ["Latin & Hari"] + [""] * 2 + ["Ayat Khusus & Titik Kesadaran"] + [""] * 2 + ["Arab & Dasar"] + [""] * 2
    # Baris 4: Header Kolom Utama
    row_4 = df.iloc[0].tolist() # Ambil header dari baris pertama data asli

    # Data Rows
    data_rows = df.iloc[1:].values.tolist() # Ambil baris data mulai dari indeks 1

    # Gabungkan semua baris
    table_data = [row_1, row_2, row_3, row_4] + data_rows

    # Buat objek Table
    t = Table(table_data, colWidths=[5*mm, 7*mm, 10*mm, 4*mm, 10*mm, 4*mm, 10*mm, 7*mm, 15*mm, 5*mm]) # Gunakan mm untuk presisi lebih tinggi

    # --- Definisikan Style ---
    # Warna Header
    header_bg_color = colors.HexColor('#1F497D')
    alt_row_color = colors.lightgrey

    # Style Umum
    common_style = [
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'), # Gunakan TOP untuk seluruh isi
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]

    # Style Header Spesifik (Baris 0-4)
    header_style = [
        ('FONTNAME', (0, 0), (-1, 3), 'Helvetica-Bold'), # Bold untuk header
        ('FONTSIZE', (0, 0), (-1, 0), 14), # Ukuran lebih besar untuk konsep
        ('FONTSIZE', (0, 1), (-1, 1), 11), # Ukuran subtitle
        ('TEXTCOLOR', (0, 0), (-1, 3), colors.whitesmoke), # Warna teks header
        ('BACKGROUND', (0, 0), (-1, 3), header_bg_color), # Warna background header
        # Merge Cells Header
        ('SPAN', (0, 0), (-1, 0)), # Konsep Dasar
        ('SPAN', (1, 1), (-1, 1)), # Metode & Juz - Kolom 2 sd akhir
        ('SPAN', (0, 2), (2, 2)), # Latin & Hari
        ('SPAN', (3, 2), (6, 2)), # Ayat Khusus
        ('SPAN', (7, 2), (-1, 2)), # Arab & Dasar
        # Align teks spesifik header
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'), # Center konsep
        ('ALIGN', (1, 1), (-1, 1), 'CENTER'), # Center metode & juz
        ('ALIGN', (0, 2), (2, 2), 'CENTER'), # Center Latin
        ('ALIGN', (3, 2), (6, 2), 'CENTER'), # Center Ayat Khusus
        ('ALIGN', (7, 2), (-1, 2), 'CENTER'), # Center Arab
    ]

    # Style Baris Data
    data_style = []
    for i in range(4, len(table_data)): # Mulai dari baris data (indeks 4)
        bg_color = alt_row_color if (i - 4) % 2 == 0 else colors.white # Warna selang-seling
        data_style.append(('BACKGROUND', (0, i), (-1, i), bg_color))
        # Style khusus kolom Arab (kolom terakhir, indeks 9)
        data_style.append(('FONTNAME', (9, i), (9, i), 'Helvetica-Bold')) # Bold Arab
        # Style khusus kolom Label (kolom pertama, indeks 0)
        data_style.append(('FONTNAME', (0, i), (0, i), 'Helvetica-Bold')) # Bold Label
        data_style.append(('FONTNAME', (1, i), (8, i), 'Helvetica')) # Regular body
        # Align teks spesifik kolom
        data_style.append(('ALIGN', (0, i), (0, i), 'CENTER')) # Center Label
        data_style.append(('ALIGN', (1, i), (8, i), 'CENTER')) # Center Body
        data_style.append(('ALIGN', (9, i), (9, i), 'CENTER')) # Center Arab

    # Style Footer (Signature)
    # Kita perlu menambahkan baris footer ke table_data dan menambahkan style merge
    # Baris footer akan ditambahkan di akhir fungsi ini
    signature_row_data = ["__", f": {processor.signature_text}", f": {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"]
    # Tambahkan 7 kolom kosong di awal untuk align dengan tabel
    signature_row_full = [""] * 7 + signature_row_data
    table_data.append(signature_row_full)
    footer_style = [
        ('SPAN', (0, -1), (-1, -1)), # Merge seluruh baris signature
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#E9F2FF")), # Warna background signature
        ('GRID', (0, -1), (-1, -1), 1, colors.black), # Border signature
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'), # Bold title
        ('FONTSIZE', (0, -1), (-1, -1), 12), # Ukuran title
        ('TEXTCOLOR', (0, -1), (-1, -1), colors.HexColor('#1F497D')), # Warna title
        ('ALIGN', (0, -1), (-1, -1), 'CENTER'), # Center title
        ('SPAN', (1, -2), (-1, -2)), # Merge stamp row
        ('FONTNAME', (1, -2), (-1, -2), 'Helvetica'), # Regular stamp
        ('FONTSIZE', (1, -2), (-1, -2), 10), # Size stamp
        ('ALIGN', (1, -2), (-1, -2), 'CENTER'), # Center stamp
        ('SPAN', (1, -3), (-1, -3)), # Merge time row
        ('FONTNAME', (1, -3), (-1, -3), 'Helvetica'), # Regular time
        ('FONTSIZE', (1, -3), (-1, -3), 9), # Size time
        ('ITALIC', (1, -3), (-1, -3)), # Italic time
        ('ALIGN', (1, -3), (-1, -3), 'CENTER'), # Center time
    ]

    # Gabungkan semua style
    t.setStyle(TableStyle(common_style + header_style + data_style + footer_style))

    return t

def export_and_format_pdf(processor, output_name, jenis="pengguna", signature_text=None):
    """Ekspor hasil ke PDF dengan styling dan page number."""
    print(f"[INFO] Mengekspor ke PDF: {output_name} (jenis: {jenis})")

    # --- Inisialisasi dokumen di buffer ---
    buffer = io.BytesIO()
    try:
        doc = SimpleDocTemplate(buffer, pagesize=letter)
    except Exception as e:
        print(f"[ERROR] Gagal membuat SimpleDocTemplate: {e}")
        return

    elements = []

    # --- Ekspor sheet berdasarkan jenis ---
    try:
        if jenis == "pengguna":
            # Hanya Init dan Final
            # --- Sheet Init ---
            if hasattr(processor, 'df_init') and processor.df_init is not None and not processor.df_init.empty:
                elements.append(Paragraph(f"Sheet: Init ({processor.metode_aktif})", getSampleStyleSheet()['Heading1']))
                # Bersihkan data df_init sebelum membuat list
                df_init_clean = processor.df_init.fillna("")
                df_init_clean = df_init_clean.apply(lambda col: col.map(clean_for_reportlab, na_action='ignore'))
                df_init_list = [df_init_clean.columns.tolist()] + df_init_clean.values.tolist()
                t_init = Table(df_init_list)
                # Gunakan style umum untuk Init
                table_header_style = [
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                    ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F497D')),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]
                table_body_style = [
                    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 1), (-1, -1), 10),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]
                # Warna Selang-seling Baris
                for i in range(1, len(df_init_list)):
                    bg_color = colors.lightgrey if i % 2 == 0 else colors.white
                    table_body_style.append(('BACKGROUND', (0, i), (-1, i), bg_color))
                t_init.setStyle(TableStyle(table_header_style + table_body_style))
                elements.append(t_init)
                elements.append(PageBreak())

            # --- Sheet Final ---
            if hasattr(processor, 'df_final') and processor.df_final is not None and not processor.df_final.empty:
                elements.append(Paragraph(f"Sheet: Final ({processor.metode_aktif})", getSampleStyleSheet()['Heading1']))
                # Gunakan fungsi khusus untuk membuat tabel Final dengan styling
                t_final = df_to_reportlab_table_final(processor.df_final, processor) # Kirim processor untuk akses sign
                elements.append(t_final)
                # elements.append(PageBreak()) # Hilangkan PageBreak jika tidak ingin dipisah

        elif jenis == "praktisi":
            sheets_data = [
                ("Init", getattr(processor, 'df_init', None)),
                ("Proses 1", getattr(processor, 'df_p1', None)),
                ("Proses 2", getattr(processor, 'df_p2', None)),
                ("Proses 3", getattr(processor, 'df_p3', None)),
                ("Final", getattr(processor, 'df_final', None)),
            ]

            for sheet_name, df in sheets_data:
                if df is not None and not df.empty:
                    elements.append(Paragraph(f"Sheet: {sheet_name} ({processor.metode_aktif})", getSampleStyleSheet()['Heading1']))
                    if sheet_name == "Final":
                         t = df_to_reportlab_table_final(df, processor)
                         elements.append(t)
                    else: # Untuk sheet lain, gunakan pendekatan umum
                        # Bersihkan data sebelum membuat list
                        df_clean = df.fillna("")
                        df_clean = df_clean.apply(lambda col: col.map(clean_for_reportlab, na_action='ignore'))
                        df_list = [df_clean.columns.tolist()] + df_clean.values.tolist()
                        t = Table(df_list)
                        # Gunakan style umum
                        table_header_style = [
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                            ('FONTSIZE', (0, 0), (-1, 0), 10),
                            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                            ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F497D')),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black)
                        ]
                        table_body_style = [
                            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                            ('FONTSIZE', (0, 1), (-1, -1), 10),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black)
                        ]
                        # Warna Selang-seling Baris
                        for i in range(1, len(df_list)):
                            bg_color = colors.lightgrey if i % 2 == 0 else colors.white
                            table_body_style.append(('BACKGROUND', (0, i), (-1, i), bg_color))
                        t.setStyle(TableStyle(table_header_style + table_body_style))
                        elements.append(t)
                    if sheet_name != "Final":
                        elements.append(PageBreak())
        else:
            print(f"[WARNING] Jenis output '{jenis}' tidak dikenal untuk PDF. Gunakan 'pengguna' atau 'praktisi'.")
            return # Batalkan ekspor jika jenis tidak valid
    except Exception as e:
        print(f"[ERROR] Gagal membangun elemen-elemen PDF: {e}")
        import traceback
        traceback.print_exc()
        return

    # --- Build dokumen ke buffer ---
    try:
        doc.build(elements, canvasmaker=NumberedCanvas)
        print(f"[INFO] Build PDF ke buffer selesai.")
    except Exception as e:
        print(f"[ERROR] Gagal membuat PDF di buffer '{output_name}': {e}")
        import traceback
        traceback.print_exc()
        return # Batalkan jika build gagal

    # --- Salin buffer ke file ---
    try:
        buffer.seek(0) # Pastikan posisi awal buffer
        with open(output_name, 'wb') as f:
            f.write(buffer.getvalue())
        print(f"[SAVED] File PDF {processor.metode_aktif} disimpan sebagai: {output_name}")
    except Exception as e:
        print(f"[ERROR] Gagal menulis buffer ke file '{output_name}': {e}")
        import traceback
        traceback.print_exc()
        # Hapus file jika terjadi error saat menulis
        try:
            if os.path.exists(output_name):
                os.remove(output_name)
                print(f"[INFO] File corrupt '{output_name}' telah dihapus.")
        except OSError as oe:
            print(f"[ERROR] Gagal menghapus file corrupt '{output_name}': {oe}")

    finally:
        buffer.close() # Tutup buffer setelah selesai

# Import datetime untuk waktu signature di footer tabel final
from datetime import datetime
# Import mm untuk ukuran kolom
from reportlab.lib.units import mm
