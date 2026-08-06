# test_simple_pdf_direct.py
# Script untuk menguji apakah reportlab dapat membuat PDF sederhana TANPA buffer IO dan TANPA canvas kustom

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    # from reportlab.pdfgen import canvas # Jangan impor canvas dulu
    # from reportlab.lib.units import inch # Jangan impor inch dulu
    # import io # Jangan gunakan io.BytesIO
    print("[INFO] Berhasil mengimpor reportlab (kecuali canvas).")
except ImportError as e:
    print(f"[ERROR] Gagal mengimpor reportlab: {e}")
    exit(1)

def create_simple_pdf_direct(output_name):
    """Buat PDF sederhana hanya dengan teks, TANPA penomoran halaman kustom."""
    print(f"[INFO] Mencoba membuat PDF sederhana (direct write, no custom canvas): {output_name}")

    # --- Tulis LANGSUNG ke file ---
    try:
        doc = SimpleDocTemplate(output_name, pagesize=letter)
    except Exception as e:
        print(f"[ERROR] Gagal membuat SimpleDocTemplate di test direct: {e}")
        return False

    styles = getSampleStyleSheet()
    elements = []

    # Tambahkan beberapa elemen sederhana
    elements.append(Paragraph("Hello World dari Test PDF Direct!", styles['Title']))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph("Ini adalah file PDF sederhana untuk menguji reportlab TANPA buffer dan canvas kustom.", styles['Normal']))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph("Jika Anda membaca ini, reportlab bisa menulis ke file secara langsung.", styles['Normal']))

    try:
        doc.build(elements) # <-- JANGAN gunakan canvasmaker
        print(f"[SAVED] File PDF sederhana (direct) disimpan sebagai: {output_name}")
        return True
    except Exception as e:
        print(f"[ERROR] Gagal membuat atau menulis PDF direct '{output_name}': {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = create_simple_pdf_direct("test_simple_reportlab_direct.pdf")
    if success:
        print("\n[SUCCESS] File 'test_simple_reportlab_direct.pdf' berhasil dibuat. Silakan coba buka.")
        print("Jika file ini bisa dibuka, maka masalahnya MUNGKIN besar terkait dengan canvas kustom NumberedCanvas.")
    else:
        print("\n[FAILURE] File 'test_simple_reportlab_direct.pdf' gagal dibuat atau korup.")
        print("Ini berarti ada masalah mendasar yang lebih luas dengan reportlab dan lingkungan Anda.")