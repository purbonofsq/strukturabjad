import os
from typing import Optional

import pandas as pd

from dataprep import load_raw_excel, prepare_staging_data
from outputproc import export_and_format_output
from pipeline import (
    build_final_table,
    build_init_table,
    build_process_1_and_2,
    build_process_3,
    build_staging_met,
)

class PengolahDataJBAH:
    def __init__(self, file_path, concept_label="Konsep Dasar (metode jumlah ayat dalam juz)", signature_text="Digital Signature"):
        """Inisialisasi class dan memuat data mentah Excel."""
        self.file_path = file_path
        self.concept_label = concept_label
        self.signature_text = signature_text
        self.raw_data: Optional[pd.DataFrame] = load_raw_excel(self.file_path)
        if self.raw_data is not None:
            print(f"[OK] Data mentah Excel berhasil dimuat dari: {self.file_path}")

    def filter_data(self, filter_col, filter_value):
        """Tahap II.1: Staging Data Utama (Filter dan Inisialisasi Dasar)"""
        if self.raw_data is None:
            print("[ERROR] raw_data is None, cannot filter.")
            return None

        self.NILAI_FILTER = filter_value
        self.staging_data = prepare_staging_data(self.raw_data, filter_col, filter_value)
        if self.staging_data is None:
            return None

        self.total_ayat = len(self.staging_data)
        print(f"[OK] Filter Data Siap: Total {self.total_ayat} ayat.")
        return self.staging_data

    def set_metode(self, metode="T111"):
        """
        Tahap Penentuan Poros (Metode Siklik):
        Menggeser (shifting) array data agar target Blok F jatuh tepat di tengah.
        """
        if self.staging_data is None:
            raise ValueError("Data belum difilter. Jalankan filter_data() terlebih dahulu.")

        self.metode_aktif = metode
        self.staging_data_met, self.ayat_awal, self.ayat_akhir, _ = build_staging_met(self.staging_data, metode)
        print(f"\n[INFO] Set Metode: {metode} | Ayat Awal Baru: {self.ayat_awal} | Ayat Akhir Baru: {self.ayat_akhir}")

    def proses_1_dan_2(self):
        """Tahap II.2.1 & II.2.2: Chunking Baku & Symmetrical Reordering"""
        self.p1_dict, self.df_p1, self.df_p2, self.titik_tengah = build_process_1_and_2(self.staging_data_met, self.metode_aktif)
        print(f"[OK] Proses 1 & Proses 2 ({self.metode_aktif}) Selesai.")

    def proses_3(self):
        """Tahap II.2.3: Flattening, Merger & Indexing Mutlak"""
        self.df_p3, self.at_h, self.at_ah = build_process_3(self.df_p2)

    def generate_tabel_jb_init(self):
        """Tahap II.2.4: Dashboard JB (Init)"""
        self.df_init = build_init_table(
            self,
            self.p1_dict,
            self.titik_tengah,
            self.ayat_awal,
            self.ayat_akhir,
            self.at_h,
            self.at_ah,
        )

    def generate_tabel_ah_final(self):
        """Tahap II.2.5: Tabel Asmaul Husna"""
        self.df_final = build_final_table(self)

    def export_excel(self, output_name, jenis="advance"):
        """Ekspor hasil ke file Excel melalui modul output terpisah."""
        return ExcelExporter.export(self, output_name, jenis=jenis, signature_text=self.signature_text)


class ExcelExporter:
    """Wrapper yang jelas untuk operasi Excel dan styling output."""

    @staticmethod
    def export(processor, output_name, jenis="advance", signature_text=None):
        return export_and_format_output(
            processor,
            output_name,
            jenis=jenis,
            signature_text=signature_text or processor.signature_text,
        )

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
    
    concept_label = input("Masukkan label konsep dasar (default: Konsep Dasar (metode jumlah ayat dalam juz)): ").strip() or "Konsep Dasar (metode jumlah ayat dalam juz)"
    signature_text = input("Masukkan tanda tangan / penandatangan (default: Digital Signature): ").strip() or "Digital Signature"
    app = PengolahDataJBAH(FILE_MENTAH, concept_label=concept_label, signature_text=signature_text)
    
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
            
            app.export_excel(NAMA_FILE_OUTPUT, jenis=JENIS_OUTPUT)
            
    print("\n[OK] SELURUH PROSES SELESAI")