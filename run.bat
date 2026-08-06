@echo off
chcp 65001 >nul
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
echo ==========================================
echo LUONG 1: TU DONG HOA BAO CAO HAN SU DUNG (CAN DATE)
echo ==========================================
echo.

:: Xac dinh duong dan thu muc venv va requirements.txt ngay trong thu muc nay
set "VENV_DIR=%~dp0venv"
set "REQ_FILE=%~dp0requirements.txt"

:: Kiem tra file requirements.txt ton tai khong
if not exist "%REQ_FILE%" (
    echo LOI: Khong tim thay file "%REQ_FILE%"
    echo Vui long dam bao co file requirements.txt nam cung thu muc voi file nay.
    pause
    exit /b 1
)

:: -----------------------------------------------
:: BUOC 1: Kiem tra Python co trong PATH khong
:: -----------------------------------------------
python --version >nul 2>&1
if errorlevel 1 (
    echo LOI: Khong tim thay Python!
    echo Vui long cai dat Python tai https://python.org
    echo Khi cai dat, nho tick chon "Add Python to PATH".
    pause
    exit /b 1
)
echo [OK] Da tim thay Python.

:: -----------------------------------------------
:: BUOC 2: Tao Virtual Environment moi neu chua co
:: -----------------------------------------------
set "NEED_CREATE_VENV=0"
if not exist "%VENV_DIR%\Scripts\python.exe" set "NEED_CREATE_VENV=1"

if "%NEED_CREATE_VENV%"=="1" goto skip_check

"%VENV_DIR%\Scripts\python.exe" -c "import os, sys; sys.exit(0 if os.path.normpath(sys.prefix).lower() == os.path.normpath(r'%VENV_DIR%').lower() else 1)" >nul 2>&1
if not errorlevel 1 goto skip_check

echo [WARN] Phat hien Virtual Environment cu co duong dan khong hop le (do copy sang may/thu muc khac).
echo Dang xoa venv cu de tu dong tao lai...
rd /s /q "%VENV_DIR%"
set "NEED_CREATE_VENV=1"

:skip_check

if "%NEED_CREATE_VENV%"=="1" (
    echo Dang tao Virtual Environment moi trong venv ...
    python -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo LOI: Khong the tao Virtual Environment.
        pause
        exit /b 1
    )
    echo [OK] Da tao Virtual Environment moi.
) else (
    echo [OK] Virtual Environment da ton tai va hoat dong tot.
)

:: -----------------------------------------------
:: BUOC 3: Kich hoat venv va cai dat thu vien
:: -----------------------------------------------
call "%VENV_DIR%\Scripts\activate.bat"
echo Dang kiem tra va cai dat/cap nhat thu vien Python...
pip install -r "%REQ_FILE%" -q
if errorlevel 1 (
    echo LOI: Khong the cai dat cac thu vien trong requirements.txt.
    pause
    exit /b 1
)
echo [OK] Thu vien da cap nhat.

:: -----------------------------------------------
:: BUOC 4: Chay script Python chinh (lay du lieu ra data.js)
:: -----------------------------------------------
echo.
echo ==========================================
echo [BUOC 5] Dang lay du lieu tu Sabeco Portal va ghi vao file bao cao...
echo ==========================================
echo.
python sabeco_shelflife_automation.py
if errorlevel 1 (
    echo.
    echo ==========================================
    echo LOI: Qua trinh lay du lieu / ghi Excel xay ra loi.
    echo Xem thong bao loi phia tren de biet chi tiet.
    echo ==========================================
    pause
    exit /b 1
)

echo.
echo ==========================================
echo HOAN THANH: Da cap nhat du lieu ra file data.js thanh cong!
echo ==========================================
pause

