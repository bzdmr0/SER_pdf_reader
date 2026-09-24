@echo off
rem Temiz bir sanal ortamda BrutUcretOkuyucu.exe derler. Cikti: dist\BrutUcretOkuyucu.exe
cd /d "%~dp0"
if not exist .venv python -m venv .venv || exit /b 1
.venv\Scripts\python -m pip install -q -r requirements.txt pyinstaller || exit /b 1
.venv\Scripts\python -m PyInstaller --noconfirm --clean --onefile --windowed --name BrutUcretOkuyucu arayuz.py || exit /b 1
echo.
echo Hazir: %~dp0dist\BrutUcretOkuyucu.exe
