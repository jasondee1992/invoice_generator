@echo off
setlocal

set "PROJECT_DIR=%~dp0"
cd /d "%PROJECT_DIR%"

if not exist ".\venv\Scripts\python.exe" (
    echo Virtual environment not found at .\venv\Scripts\python.exe
    exit /b 1
)

echo Installing/updating PyInstaller...
call ".\venv\Scripts\python.exe" -m pip install pyinstaller
if errorlevel 1 exit /b 1

echo Building InvoiceGenerator.exe...
call ".\venv\Scripts\python.exe" -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name InvoiceGenerator ^
  --distpath . ^
  --workpath build ^
  --specpath . ^
  generate_invoice.py
if errorlevel 1 exit /b 1

echo Build complete: "%PROJECT_DIR%InvoiceGenerator.exe"
endlocal
