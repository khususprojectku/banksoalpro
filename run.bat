@echo off
cd /d "%~dp0"
echo ==============================================
echo   BankSoalPro - Sistem Manajemen Bank Soal
echo ==============================================
echo Menjalankan server lokal di http://localhost:8000...
echo Tekan Ctrl+C di jendela ini untuk menghentikan server.
echo.

:: Hentikan proses server lama jika masih berjalan
for /f "tokens=5" %%a in ('netstat -ano ^| findstr ":8000" ^| findstr "LISTENING"') do (
    taskkill /F /PID %%a >nul 2>&1
)

:: Jalankan server Python di background window tersendiri
start "BankSoalPro Server" py server.py

:: Tunggu 2 detik agar server sempat menyala
timeout /t 2 /nobreak >nul

:: Buka browser setelah server siap
start "" "http://localhost:8000"

echo Server berjalan. Jangan tutup jendela "BankSoalPro Server".
echo Tekan tombol apapun untuk keluar dari launcher ini...
pause >nul
