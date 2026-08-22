import os
import sys
import shutil
import subprocess
import time
from pathlib import Path
from datetime import datetime

# =========================================================
# KONFIGURASI - SESUAIKAN DENGAN KOMPUTER ANDA
# =========================================================

# Folder berisi file Excel/Calc
INPUT_FOLDER = Path(r"D:\ExcelFiles")

# Folder tujuan PDF
OUTPUT_FOLDER = Path(r"D:\PDFFiles")

# Folder profile sementara LibreOffice
# Berguna agar proses batch tidak bentrok dengan LibreOffice manual
PROFILE_FOLDER = Path(r"D:\LOProfile")

# Jika kosong, script akan mencari LibreOffice otomatis.
# Jika LibreOffice tidak terdeteksi, isi path lengkap soffice/libreoffice.
#
# Contoh Windows:
# LIBREOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"
#
# Contoh Linux:
# LIBREOFFICE_PATH = "/usr/bin/soffice"
#
# Contoh macOS:
# LIBREOFFICE_PATH = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
LIBREOFFICE_PATH = ""

# Jika True, proses juga masuk ke subfolder
RECURSIVE = False

# Jika True, PDF lama akan ditimpa
# Jika False, PDF yang sudah ada akan dilewati
OVERWRITE_EXISTING_PDF = False

# Batas waktu konversi per file, dalam detik
TIMEOUT_SECONDS = 300

# Jumlah percobaan ulang jika gagal
MAX_RETRY = 2

# Ekstensi file yang akan diproses
EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".ods"}

# =========================================================


def find_libreoffice():
    """
    Cari executable LibreOffice.
    """
    if LIBREOFFICE_PATH:
        p = Path(LIBREOFFICE_PATH).expanduser().resolve()
        if p.exists():
            return p
        raise FileNotFoundError(
            f"LIBREOFFICE_PATH tidak ditemukan: {p}"
        )

    # Coba cari dari PATH
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return Path(found)

    candidates = []

    if os.name == "nt":
        program_files = os.environ.get("PROGRAMFILES", r"C:\Program Files")
        program_files_x86 = os.environ.get(
            "PROGRAMFILES(X86)", r"C:\Program Files (x86)"
        )

        candidates.extend([
            Path(program_files) / "LibreOffice" / "program" / "soffice.exe",
            Path(program_files_x86) / "LibreOffice" / "program" / "soffice.exe",
        ])

    elif sys.platform == "darwin":
        candidates.append(
            Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
        )

    else:
        candidates.extend([
            Path("/usr/bin/soffice"),
            Path("/usr/bin/libreoffice"),
            Path("/usr/local/bin/soffice"),
            Path("/usr/local/bin/libreoffice"),
            Path("/snap/bin/libreoffice"),
        ])

    for candidate in candidates:
        if candidate.exists():
            return candidate

    raise FileNotFoundError(
        "LibreOffice tidak ditemukan. "
        "Silakan isi LIBREOFFICE_PATH di script dengan path lengkap soffice/libreoffice."
    )


def collect_files(input_folder: Path):
    """
    Kumpulkan file Excel/Calc.
    """
    files = []

    patterns = ["*.xlsx", "*.xls", "*.xlsm", "*.ods"]

    for pattern in patterns:
        if RECURSIVE:
            files.extend(input_folder.rglob(pattern))
        else:
            files.extend(input_folder.glob(pattern))

    # Abaikan file temporary Excel
    files = [
        f for f in files
        if f.is_file() and not f.name.startswith("~$")
    ]

    return sorted(files, key=lambda x: str(x).lower())


def get_pdf_path(input_folder: Path, output_folder: Path, excel_file: Path):
    """
    Tentukan path PDF tujuan.
    """
    if RECURSIVE:
        rel = excel_file.relative_to(input_folder)
        pdf_path = output_folder / rel.with_suffix(".pdf")
    else:
        pdf_path = output_folder / (excel_file.stem + ".pdf")

    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    return pdf_path


def run_libreoffice(
    soffice: Path,
    excel_file: Path,
    outdir: Path,
    profile_folder: Path
):
    """
    Jalankan LibreOffice headless untuk convert satu file ke PDF.
    """
    cmd = [
        str(soffice),
        f"-env:UserInstallation={profile_folder.as_uri()}",
        "--headless",
        "--norestore",
        "--nolockcheck",
        "--nologo",
        "--invisible",
        "--convert-to", "pdf",
        "--outdir", str(outdir),
        str(excel_file),
    ]

    kwargs = {
        "capture_output": True,
        "text": True,
        "timeout": TIMEOUT_SECONDS,
    }

    # Agar tidak muncul jendela console tambahan di Windows
    if os.name == "nt" and hasattr(subprocess, "CREATE_NO_WINDOW"):
        kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW

    return subprocess.run(cmd, **kwargs)


def convert_with_retry(
    soffice: Path,
    excel_file: Path,
    pdf_path: Path,
    profile_folder: Path
):
    """
    Convert file dengan retry jika gagal.
    """
    last_error = ""

    for attempt in range(1, MAX_RETRY + 1):
        try:
            result = run_libreoffice(
                soffice=soffice,
                excel_file=excel_file,
                outdir=pdf_path.parent,
                profile_folder=profile_folder
            )

            if result.returncode == 0 and pdf_path.exists():
                return True, ""

            stderr = (result.stderr or "").strip()
            stdout = (result.stdout or "").strip()

            last_error = (
                f"returncode={result.returncode}; "
                f"stderr={stderr[:500]}; "
                f"stdout={stdout[:200]}"
            )

        except subprocess.TimeoutExpired:
            last_error = f"timeout setelah {TIMEOUT_SECONDS} detik"

        except Exception as e:
            last_error = str(e)

        if attempt < MAX_RETRY:
            time.sleep(2)

    return False, last_error


def main():
    input_folder = INPUT_FOLDER.expanduser().resolve()
    output_folder = OUTPUT_FOLDER.expanduser().resolve()
    profile_folder = PROFILE_FOLDER.expanduser().resolve()

    if not input_folder.exists():
        print(f"Folder input tidak ditemukan: {input_folder}")
        return

    output_folder.mkdir(parents=True, exist_ok=True)
    profile_folder.mkdir(parents=True, exist_ok=True)

    try:
        soffice = find_libreoffice()
    except Exception as e:
        print(e)
        return

    files = collect_files(input_folder)
    total = len(files)

    if total == 0:
        print("Tidak ada file Excel/Calc yang ditemukan.")
        return

    log_file = output_folder / f"export_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

    def log(msg):
        print(msg)
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(msg + "\n")

    log("======================================================")
    log("BATCH EXPORT EXCEL/CALC TO PDF - PYTHON + LIBREOFFICE")
    log("======================================================")
    log(f"LibreOffice : {soffice}")
    log(f"Input       : {input_folder}")
    log(f"Output      : {output_folder}")
    log(f"Profile     : {profile_folder}")
    log(f"Total file  : {total}")
    log(f"Recursive   : {RECURSIVE}")
    log(f"Overwrite   : {OVERWRITE_EXISTING_PDF}")
    log("======================================================")

    success = []
    skipped = []
    failed = []

    try:
        for index, excel_file in enumerate(files, start=1):
            pdf_path = get_pdf_path(
                input_folder=input_folder,
                output_folder=output_folder,
                excel_file=excel_file
            )

            try:
                display_name = excel_file.relative_to(input_folder)
            except ValueError:
                display_name = excel_file.name

            # Jika PDF sudah ada dan overwrite dimatikan, lewati
            if pdf_path.exists() and not OVERWRITE_EXISTING_PDF:
                skipped.append(excel_file)
                log(f"[{index}/{total}] SKIP  : {display_name} -> PDF sudah ada")
                continue

            # Jika overwrite dinyalakan, hapus PDF lama agar tidak salah deteksi
            if pdf_path.exists() and OVERWRITE_EXISTING_PDF:
                try:
                    pdf_path.unlink()
                except Exception as e:
                    failed.append((excel_file, f"gagal hapus PDF lama: {e}"))
                    log(f"[{index}/{total}] ERROR : {display_name} -> gagal hapus PDF lama")
                    continue

            ok, error = convert_with_retry(
                soffice=soffice,
                excel_file=excel_file,
                pdf_path=pdf_path,
                profile_folder=profile_folder
            )

            if ok:
                success.append(excel_file)
                log(f"[{index}/{total}] OK    : {display_name} -> {pdf_path.name}")
            else:
                failed.append((excel_file, error))
                log(f"[{index}/{total}] GAGAL : {display_name} -> {error}")

    except KeyboardInterrupt:
        log("\nProses dihentikan manual oleh user.")

    finally:
        log("\n======================================================")
        log("RINGKASAN")
        log("======================================================")
        log(f"Total   : {total}")
        log(f"OK      : {len(success)}")
        log(f"Skip    : {len(skipped)}")
        log(f"Gagal   : {len(failed)}")
        log(f"Log     : {log_file}")

        if failed:
            log("\nDetail file gagal:")
            for file_path, error in failed:
                log(f"{file_path} | {error}")


if __name__ == "__main__":
    main()