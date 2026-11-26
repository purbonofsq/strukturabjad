@echo off
title Cek + Nyalakan LM Studio via CLI (lms)
chcp 65001 >nul
echo ==========================================
echo  CEK + NYALAKAN LM STUDIO PAKAI CLI (lms)
echo ==========================================
echo.

REM Cari lms.exe. Kalau belum di-PATH, pakai path lengkapnya.
where lms >nul 2>&1
if errorlevel 1 (
  echo lms belum ada di PATH, pakai path lengkap...
  set LMS=%USERPROFILE%\.lmstudio\bin\lms.exe
) else (
  set LMS=lms
)

echo [1/3] Cek apakah server sudah jalan...
curl -s http://localhost:1234/v1/models | findstr "qwen2.5-coder" >nul
if not errorlevel 1 (
  echo.
  echo SUDAH JALAN. Model qwen2.5-coder-1.5b-instruct kedetek.
  echo Aman lanjut ke oco.
  goto :end
)

echo Server belum jalan. Menyalakan via CLI...
echo.

echo [2/3] lms server start...
"%LMS%" server start

echo.
echo [3/3] Load model qwen2.5-coder-1.5b-instruct...
echo Ini butuh waktu beberapa detik, tunggu ya...
"%LMS%" load qwen2.5-coder-1.5b-instruct --context-length 4096 --yes

echo.
echo Cek lagi...
curl -s http://localhost:1234/v1/models
echo.
echo Kalau di atas muncul qwen2.5-coder-1.5b-instruct, berarti BERHASIL.
echo Sekarang bisa pakai oco.

:end
echo.
pause
