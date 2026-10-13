@echo off
REM ============================================================
REM Build script untuk Pencocokan Transaksi v4.3.2
REM ============================================================

echo.
echo ============================================================
echo   BUILD Pencocokan Transaksi v4.3.2
echo ============================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python tidak ditemukan di PATH.
    echo Install Python 3.9+ dari https://python.org
    pause
    exit /b 1
)

echo [1/5] Upgrade pip...
python -m pip install --upgrade pip

echo.
echo [2/5] Install dependencies...
python -m pip install pyinstaller openpyxl scipy numpy

echo.
echo [3/5] Bersihkan build lama...
if exist build rmdir /S /Q build
if exist dist rmdir /S /Q dist
if exist "PencocokanTransaksi.spec" del /Q "PencocokanTransaksi.spec"

echo.
echo [4/5] Build EXE (sabar, 5-15 menit)...
python -m PyInstaller ^
    --noconfirm ^
    --onefile ^
    --windowed ^
    --name "PencocokanTransaksi" ^
    --hidden-import openpyxl ^
    --hidden-import scipy.optimize ^
    --hidden-import scipy._lib.messagestream ^
    --hidden-import numpy ^
    --collect-all openpyxl ^
    --collect-all scipy ^
    --collect-all numpy ^
    --exclude-module matplotlib ^
    --exclude-module pandas ^
    --exclude-module PIL ^
    --exclude-module IPython ^
    --exclude-module notebook ^
    --exclude-module pytest ^
    pencocokan_transaksi.py

if errorlevel 1 (
    echo.
    echo [ERROR] Build gagal. Periksa pesan di atas.
    pause
    exit /b 1
)

echo.
echo [5/5] Siapkan folder distribusi...
if not exist "dist\PencocokanTransaksi_v4.3.2" mkdir "dist\PencocokanTransaksi_v4.3.2"
copy /Y "dist\PencocokanTransaksi.exe" "dist\PencocokanTransaksi_v4.3.2\PencocokanTransaksi.exe"
copy /Y "README.txt" "dist\PencocokanTransaksi_v4.3.2\README.txt" 2>nul

echo.
echo ============================================================
echo   SELESAI!
echo ============================================================
echo.
echo File EXE ada di:
echo   dist\PencocokanTransaksi.exe
echo.
echo Folder siap kirim ke teman:
echo   dist\PencocokanTransaksi_v4.3.2\
echo.
pause