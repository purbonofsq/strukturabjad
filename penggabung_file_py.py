import glob

with open("gabungan_kode.txt", "w", encoding="utf-8") as file_keluar:
    # Membaca semua file dengan ekstensi .py
    for nama_file in glob.glob("*.py"):
        file_keluar.write(f"\n\n{'='*40}\n")
        file_keluar.write(f"NAMA FILE: {nama_file}\n")
        file_keluar.write(f"{'='*40}\n\n")
        
        try:
            with open(nama_file, "r", encoding="utf-8") as file_masuk:
                file_keluar.write(file_masuk.read())
        except Exception as e:
            file_keluar.write(f"Gagal membaca file: {e}\n")

print("Selesai! Buka file gabungan_kode.txt")