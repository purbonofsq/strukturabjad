import os
import sys
import shutil
import subprocess
import time
import tempfile
from pathlib import Path
from datetime import datetime

# =========================================================
# KONFIGURASI PROSES
# =========================================================

# True = proses juga file di dalam subfolder
# False = hanya proses file di folder yang dipilih saja
PROSES_SUBFOLDER = True

# True = PDF lama akan ditimpa
# False = PDF lama dilewati
OVERWRITE_EXISTING_PDF = False

# Batas waktu konversi per file (detik)
TIMEOUT_SECONDS = 300

# Jumlah percobaan ulang jika gagal
MAX_RETRY = 2

# Jika kosong, script mencari LibreOffice otomatis.
# Jika tidak ketemu, isi path lengkap LibreOffice Anda.
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

# Pola file yang diproses
PATTERNS = ["*.xlsx", "*.xls", "*.xlsm", "*.ods"]

# =========================================================


def find_libreoffice():
    """
    Mencari executable LibreOffice.
    """
    if LIBREOFFICE_PATH:
        p = Path(LIBREOFFICE_PATH).expanduser().resolve()
        if p.exists():
            return p
        raise FileNotFoundError(
            f"LIBREOFFICE_PATH tidak ditemukan: {p}"
        )

    # Cari dari PATH
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


def clean_path_string(text: str) -> str:
    """
    Bersihkan input path dari spasi dan tanda kutip.
    """
    return text.strip().strip('"').strip("'")


def list_folders_in_cwd():
    """
    Ambil daftar folder yang ada di folder aktif tempat program berjalan.
    """
    base = Path.cwd().resolve()

    try:
        entries = list(base.iterdir())
    except Exception:
        entries = []

    folders = []

    for entry in entries:
        try:
            # Hanya tampilkan folder
            # Folder tersembunyi (diawali titik) tidak ditampilkan
            if entry.is_dir() and not entry.name.startswith("."):
                folders.append(entry)
        except Exception:
            pass

    folders.sort(key=lambda p: p.name.lower())

    return base, folders


def ask_input_folders():
    """
    Minta user memilih folder dari list,
    atau memasukkan path folder custom.
    """
    base, available_folders = list_folders_in_cwd()

    print("======================================================")
    print("BATCH EXPORT EXCEL/CALC TO PDF - PYTHON + LIBREOFFICE")
    print("======================================================")
    print(f"Daftar folder yang tersedia di:")
    print(f"{base}")
    print("")
    print("[0] . (folder aktif)")

    if available_folders:
        for i, folder in enumerate(available_folders, start=1):
            print(f"[{i}] {folder.name}")
    else:
        print("(tidak ada subfolder lain di folder ini)")

    print("")
    print("Cara memilih:")
    print("- Ketik nomor folder, contoh: 1")
    print("- Ketik beberapa nomor, contoh: 1,3")
    print("- Ketik path custom, contoh: D:\\ExcelFiles")
    print("- Kombinasi nomor dan path, contoh: 1,3,D:\\ExtraFolder")
    print("- Tekan Enter saja untuk memakai folder aktif")
    print("")

    raw = input("Pilih folder yang akan diproses:\n> ").strip()

    # Jika user hanya tekan Enter, gunakan folder aktif
    if not raw:
        return [base]

    # Dukung pemisah koma atau titik koma
    raw = raw.replace(";", ",")

    tokens = [clean_path_string(token) for token in raw.split(",")]

    folders = []
    seen = set()

    for token in tokens:
        if not token:
            continue

        # Jika user memasukkan nomor dari list
        if token.isdigit():
            idx = int(token)

            if idx == 0:
                selected = base
            elif 1 <= idx <= len(available_folders):
                selected = available_folders[idx - 1]
            else:
                print(f"\nERROR: Nomor pilihan tidak valid: {idx}")
                sys.exit(1)

            try:
                selected = selected.resolve()
            except Exception:
                pass

        # Jika user memasukkan path custom
        else:
            selected = Path(token).expanduser()

            try:
                selected = selected.resolve()
            except Exception:
                pass

            if not selected.exists():
                print(f"\nERROR: Folder tidak ditemukan: {selected}")
                sys.exit(1)

            if not selected.is_dir():
                print(f"\nERROR: Bukan folder: {selected}")
                sys.exit(1)

        key = str(selected).lower()

        if key not in seen:
            seen.add(key)
            folders.append(selected)

    if not folders:
        print("\nERROR: Tidak ada folder yang dipilih atau dimasukkan.")
        sys.exit(1)

    return folders


def resolve_output_folder(raw_output: str, input_folders):
    """
    Tentukan folder output.
    Jika user mengosongkan input, buat folder output default.
    """
    raw_output = clean_path_string(raw_output)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Jika kosong, buat default
    if not raw_output:
        if len(input_folders) == 1:
            root = input_folders[0]

            if root.name:
                return root.parent / f"{root.name}_PDF"
            else:
                return Path.cwd() / f"PDF_Output_{timestamp}"
        else:
            return Path.cwd() / f"PDF_Output_{timestamp}"

    p = Path(raw_output).expanduser()

    # Jika user memasukkan path lengkap
    if p.is_absolute():
        return p.resolve()

    # Jika user hanya mengetik nama folder saja
    # dan hanya ada satu folder input,
    # maka folder output dibuat di parent folder input.
    #
    # Contoh:
    # Input folder : D:\ExcelFiles
    # Output diisi : HasilPDF
    # Output menjadi: D:\HasilPDF
    #
    # Jika ada banyak folder input, output dibuat di folder aktif Python.
    if len(input_folders) == 1 and not any(sep in raw_output for sep in (os.sep, "/")):
        return (input_folders[0].parent / p).resolve()

    return (Path.cwd() / p).resolve()


def is_under(path: Path, parent: Path) -> bool:
    """
    Cek apakah path berada di dalam parent.
    """
    try:
        path.relative_to(parent)
        return True
    except ValueError:
        return False


def collect_files(input_folders, recursive: bool, output_folder: Path):
    """
    Kumpulkan file Excel/Calc dari folder input.
    """
    items = []
    seen = set()

    for root in input_folders:
        for pattern in PATTERNS:
            if recursive:
                iterator = root.rglob(pattern)
            else:
                iterator = root.glob(pattern)

            for file_path in iterator:
                if not file_path.is_file():
                    continue

                # Abaikan file temporary
                if file_path.name.startswith("~$"):
                    continue

                # Jangan proses file yang berada di folder output,
                # kecuali folder output memang sama dengan folder input.
                if output_folder != root and is_under(file_path, output_folder):
                    continue

                key = str(file_path).lower()

                if key in seen:
                    continue

                seen.add(key)
                items.append((root, file_path))

    items.sort(key=lambda x: str(x[1]).lower())
    return items


def sanitize_folder_name(name: str) -> str:
    """
    Bersihkan nama folder agar aman dipakai sebagai nama folder output.
    """
    safe = "".join(
        c if c.isalnum() or c in ("-", "_", " ") else "_"
        for c in name
    ).strip()

    return safe or "root"


def get_pdf_path(
    root: Path,
    excel_file: Path,
    output_folder: Path,
    multiple_roots: bool
):
    """
    Tentukan path PDF hasil konversi.
    """
    rel = excel_file.relative_to(root)

    if multiple_roots:
        root_name = root.name

        if not root_name and root.drive:
            root_name = root.drive.strip(":")

        root_name = sanitize_folder_name(root_name)

        pdf_path = output_folder / root_name / rel.with_suffix(".pdf")
    else:
        pdf_path = output_folder / rel.with_suffix(".pdf")

    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    return pdf_path


def run_libreoffice(
    soffice: Path,
    excel_file: Path,
    outdir: Path,
    profile_folder: Path
):
    """
    Menjalankan LibreOffice headless untuk convert satu file ke PDF.
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
    Convert file ke PDF dengan retry jika gagal.
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
    input_folders = ask_input_folders()

    print("\nFolder input yang akan diproses:")
    for folder in input_folders:
        print(f"  - {folder}")

    print("")
    print("Masukkan folder output.")
    print("Jika dikosongkan, akan dibuat folder output default.")
    print("")

    raw_output = input("Folder output [Enter untuk default]:\n> ")

    try:
        output_folder = resolve_output_folder(raw_output, input_folders)
    except Exception as e:
        print(f"\nERROR: Gagal menentukan folder output: {e}")
        return

    try:
        output_folder.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        print(f"\nERROR: Gagal membuat folder output: {output_folder}")
        print(e)
        return

    try:
        soffice = find_libreoffice()
    except Exception as e:
        print(f"\nERROR: {e}")
        return

    profile_folder = Path(tempfile.gettempdir()) / (
        "lo_profile_" + datetime.now().strftime("%Y%m%d_%H%M%S")
    )

    try:
        profile_folder.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        print(f"\nERROR: Gagal membuat folder profile LibreOffice: {e}")
        return

    items = collect_files(
        input_folders=input_folders,
        recursive=PROSES_SUBFOLDER,
        output_folder=output_folder
    )

    total = len(items)

    if total == 0:
        print("\nTidak ada file Excel/Calc yang ditemukan.")
        return

    log_file = output_folder / f"export_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

    def log(msg):
        print(msg)
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(msg + "\n")

    log("======================================================")
    log("BATCH EXPORT EXCEL/CALC TO PDF - PYTHON + LIBREOFFICE")
    log("======================================================")
    log(f"LibreOffice     : {soffice}")
    log(f"Folder input    :")

    for folder in input_folders:
        log(f"  - {folder}")

    log(f"Folder output   : {output_folder}")
    log(f"Profile         : {profile_folder}")
    log(f"Total file      : {total}")
    log(f"Proses subfolder: {PROSES_SUBFOLDER}")
    log(f"Overwrite PDF   : {OVERWRITE_EXISTING_PDF}")
    log("======================================================")

    success = []
    skipped = []
    failed = []

    multiple_roots = len(input_folders) > 1

    try:
        for index, (root, excel_file) in enumerate(items, start=1):
            pdf_path = get_pdf_path(
                root=root,
                excel_file=excel_file,
                output_folder=output_folder,
                multiple_roots=multiple_roots
            )

            try:
                rel_name = excel_file.relative_to(root)
            except ValueError:
                rel_name = excel_file.name

            if multiple_roots:
                root_label = root.name

                if not root_label and root.drive:
                    root_label = root.drive.strip(":")

                root_label = sanitize_folder_name(root_label)
                display_name = f"{root_label}/{rel_name}"
            else:
                display_name = str(rel_name)

            if pdf_path.exists() and not OVERWRITE_EXISTING_PDF:
                skipped.append(excel_file)
                log(f"[{index}/{total}] SKIP  : {display_name} -> PDF sudah ada")
                continue

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