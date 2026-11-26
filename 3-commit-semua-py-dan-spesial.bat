@echo off
title Commit Semua PY + File Spesial
chcp 65001 >nul
echo ==========================================
echo  COMMIT SEMUA .PY + FILE SPESIAL
echo  1 file = 1 commit, biar AI tidak bingung
echo ==========================================
echo.

echo [1/3] Cek LM Studio...
curl -s http://localhost:1234/v1/models | findstr "qwen2.5-coder" >nul
if errorlevel 1 (
  echo.
  echo LM STUDIO BELUM JALAN! Buka LM Studio dulu, klik Start Server.
  echo Lalu klik lagi file .bat ini.
  pause
  exit /b
)
echo LM Studio JALAN. Lanjut...
echo.

echo [2/3] Commit semua file .py satu per satu...
echo Kalau ada pertanyaan dari oco, jawab Yes / No seperti biasa.
echo.

for %%F in (*.py) do (
  git status --porcelain "%%F" | findstr . >nul
  if not errorlevel 1 (
    echo.
    echo ==========================================
    echo  COMMIT: %%F
    echo ==========================================
    git add "%%F"
    oco
  )
)

echo.
echo [3/3] Commit file SPESIAL .xlsx dan .csv...
echo Edit file .bat ini kalau mau tambah / kurangi daftar di bawah.
echo.

REM ============================================================
REM EDIT DI SINI: Tulis nama file spesial yang BOLEH di-commit.
REM Pisahkan dengan spasi. Kalau tidak ada, kosongkan saja.
REM Contoh: set SPESIAL=master_tabel_ah_rawahaq.csv data_penting.xlsx
REM ============================================================
set SPESIAL=master_tabel_ah_rawahaq.csv master_tabel_ah_rawahaq_expanded.csv "master_tabel_ah_Struktur Quran18.csv"

for %%S in (%SPESIAL%) do (
  git status --porcelain "%%~S" | findstr . >nul
  if not errorlevel 1 (
    echo.
    echo ==========================================
    echo  COMMIT SPESIAL: %%~S
    echo ==========================================
    git add "%%~S"
    oco
  ) else (
    echo Skip: %%~S - tidak ada perubahan / tidak ketemu
  )
)

echo.
echo ==========================================
echo  SELESAI SEMUA!
echo ==========================================
git status --short
echo.
pause
