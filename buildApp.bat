@echo off
title Build Inter-SmartPLS Application
echo =======================================================
echo        MEMULAI PROSES BUILD APLIKASI INTER-SMARTPLS
echo =======================================================
echo.

:: Cek keberadaan virtual environment
if not exist "venv\Scripts\activate.bat" (
    echo [ERROR] Virtual environment venv tidak ditemukan!
    echo Silakan buat venv terlebih dahulu atau pastikan folder venv ada.
    goto end
)

:: Membersihkan folder build dan dist lama
echo Membersihkan folder build dan dist lama jika ada...
if exist build (
    echo Menghapus folder build...
    rmdir /s /q build
)
if exist dist (
    echo Menghapus folder dist...
    rmdir /s /q dist
)
if exist build_old (
    echo Menghapus folder build_old...
    rmdir /s /q build_old
)
if exist dist_old (
    echo Menghapus folder dist_old...
    rmdir /s /q dist_old
)
echo.

:: Mengaktifkan Virtual Environment
echo [1/3] Mengaktifkan virtual environment...
call .\venv\Scripts\activate
if %errorlevel% neq 0 (
    echo [ERROR] Gagal mengaktifkan virtual environment.
    goto end
)

:: Menjalankan PyInstaller
echo [2/3] Memulai kompilasi dengan PyInstaller...
pyinstaller -y --clean Inter-SmartPLS.spec
if %errorlevel% neq 0 (
    echo [ERROR] Proses kompilasi PyInstaller gagal!
    goto deactivate_env
)

:: Sukses
echo.
echo =======================================================
echo [3/3] PROSES BUILD BERHASIL SELESAI!
echo Executable dapat ditemukan di folder: dist\Inter-SmartPLS
echo =======================================================
goto deactivate_env

:deactivate_env
echo Menutup virtual environment...
call deactivate

:end
echo.
echo Tekan tombol apa saja untuk keluar...
pause >nul
