@echo off
title Validasi Export Modul
cd /d "%~dp0"

if exist ".venv\Scripts\activate.bat" goto jalan

echo Pertama kali dijalankan. Membuat environment, mohon tunggu...
python -m venv .venv
if errorlevel 1 goto gagal_python
call ".venv\Scripts\activate.bat"
python -m pip install -r requirements.txt
if errorlevel 1 goto gagal_install
goto jalan

:jalan
call ".venv\Scripts\activate.bat"
echo Aplikasi berjalan di http://localhost:8501
echo Jangan tutup jendela ini selama aplikasi dipakai.
streamlit run app.py
pause
goto :eof

:gagal_python
echo Python tidak ditemukan. Install Python 3.10 atau lebih baru dari python.org.
echo Saat instalasi, centang "Add Python to PATH".
pause
goto :eof

:gagal_install
echo Gagal menginstall library. Periksa koneksi internet lalu coba lagi.
echo Kalau tetap gagal, hapus folder .venv lalu jalankan ulang file ini.
pause
goto :eof
