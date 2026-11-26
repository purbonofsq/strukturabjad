@echo off
title Cek LM Studio
echo ==========================================
echo  CEK APAKAH LM STUDIO SUDAH JALAN
echo ==========================================
echo.
curl -s http://localhost:1234/v1/models
echo.
echo.
echo Kalau muncul tulisan qwen2.5-coder-1.5b-instruct di atas,
echo berarti LM Studio SUDAH JALAN. Aman lanjut ke oco.
echo.
echo Kalau muncul error / tidak bisa connect,
echo buka LM Studio dulu, klik Start Server.
echo.
pause
