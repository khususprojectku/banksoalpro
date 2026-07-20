@echo off
echo ==============================================
echo   BankSoalPro - Sistem Manajemen Bank Soal
echo ==============================================
echo Menjalankan server lokal di http://localhost:8000...
echo Tekan Ctrl+C di jendela ini untuk menghentikan server.
echo.
start "" "http://localhost:8000"
py -m http.server 8000
