@echo off
title Commit 1 File Pakai AI Lokal
echo ==========================================
echo  COMMIT 1 FILE PAKAI OPENCOMMIT + LM STUDIO
echo ==========================================
echo.
echo File yang belum di-commit:
git status --short
echo.
set /p NAMA_FILE="Ketik nama file yang mau di-commit (contoh: polabaptis.py): "
echo.
echo Menambahkan %NAMA_FILE%...
git add "%NAMA_FILE%"
echo.
echo Meminta AI buatkan pesan commit...
oco
echo.
echo Selesai!
pause
