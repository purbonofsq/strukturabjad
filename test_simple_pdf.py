# test_simple_pdf.py
# Script untuk menguji apakah reportlab dapat membuat PDF sederhana tanpa corrupt

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import inch
    import io
    import os
    print("[INFO] Berhasil mengimpor reportlab.")
except ImportError as e:
    print(f"[ERROR] Gagal mengimpor reportlab: {e}")
    exit(1)

# --- Canvas kustom untuk penomoran halaman ---
class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage() # type: ignore # Pylance: _startPage adalah metode resmi canvas.Canvas

    def save(self):
        """add page info to each page (page x of y)"""
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_number(num_pages)
        canvas.Canvas.save(self)

    def draw_page_number(self, page_count):
        self.setFont("Helvetica", 10)
        # Gunakan inch yang telah diimpor
        self.drawRightString(7.5*inch, 0.5*inch, f"Page {self._pageNumber} of {page_count}") # type: ignore # Pylance: _pageNumber adalah atribut resmi canvas.Canvas

def create_simple_pdf(output_name):
    """Buat PDF sederhana hanya dengan teks dan penomoran halaman."""
    print(f"[INFO] Mencoba membuat PDF sederhana: {output_name}")

    # --- Gunakan buffer ---
    buffer = io.BytesIO()
    try:
        doc = SimpleDocTemplate(buffer, pagesize=letter)
    except Exception as e:
        print(f"[ERROR] Gagal membuat SimpleDocTemplate di test: {e}")
        return False

    styles = getSampleStyleSheet()
    elements = []

    # Tambahkan beberapa elemen sederhana
    elements.append(Paragraph("Hello World dari Test PDF!", styles['Title']))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph("Ini adalah file PDF sederhana untuk menguji reportlab.", styles['Normal']))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph("Jika Anda membaca ini, reportlab berfungsi.", styles['Normal']))

    try:
        doc.build(elements, canvasmaker=NumberedCanvas)
        print(f"[INFO] Build PDF sederhana ke buffer selesai.")
    except Exception as e:
        print(f"[ERROR] Gagal membuat PDF sederhana di buffer: {e}")
        import traceback
        traceback.print_exc()
        return False

    # --- Salin buffer ke file ---
    try:
        buffer.seek(0) # Pastikan posisi awal buffer
        with open(output_name, 'wb') as f:
            f.write(buffer.getvalue())
        print(f"[SAVED] File PDF sederhana disimpan sebagai: {output_name}")
        buffer.close()
        return True
    except Exception as e:
        print(f"[ERROR] Gagal menulis buffer ke file '{output_name}': {e}")
        import traceback
        traceback.print_exc()
        buffer.close()
        # Hapus file jika terjadi error saat menulis
        try:
            if os.path.exists(output_name):
                os.remove(output_name)
                print(f"[INFO] File corrupt '{output_name}' telah dihapus.")
        except OSError as oe:
            print(f"[ERROR] Gagal menghapus file corrupt '{output_name}': {oe}")
        return False

if __name__ == "__main__":
    success = create_simple_pdf("test_simple_reportlab.pdf")
    if success:
        print("\n[SUCCESS] File 'test_simple_reportlab.pdf' berhasil dibuat. Silakan coba buka.")
        print("Jika file ini bisa dibuka, maka masalahnya terletak pada data yang diberikan ke reportlab di pdf_export.py.")
    else:
        print("\n[FAILURE] File 'test_simple_reportlab.pdf' gagal dibuat atau korup.")
        print("Ini berarti ada masalah mendasar dengan reportlab atau cara kita menggunakannya.")