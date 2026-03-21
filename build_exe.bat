@echo off
setlocal

set "PROJECT_DIR=%~dp0"
cd /d "%PROJECT_DIR%"

if not exist ".\venv\Scripts\python.exe" (
    echo Virtual environment not found at .\venv\Scripts\python.exe
    exit /b 1
)

if not exist ".\icon\icon.png" (
    echo Missing icon file: .\icon\icon.png
    exit /b 1
)

echo Installing/updating PyInstaller...
call ".\venv\Scripts\python.exe" -m pip install pyinstaller
if errorlevel 1 exit /b 1

echo Converting icon\icon.png to icon\icon.ico...
call ".\venv\Scripts\python.exe" -c "from pathlib import Path; from PIL import Image; src=Path(r'.\icon\icon.png'); dst=Path(r'.\icon\icon.ico'); base=Image.open(src).convert('RGBA'); sizes=[(256,256),(128,128),(64,64),(48,48),(32,32),(24,24),(16,16)]; icons=[base.resize(size, Image.Resampling.LANCZOS) for size in sizes]; icons[0].save(dst, format='ICO', append_images=icons[1:])"
if errorlevel 1 exit /b 1

echo Building InvoiceGenerator.exe...
call ".\venv\Scripts\python.exe" -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --icon .\icon\icon.ico ^
  --name InvoiceGenerator ^
  --distpath . ^
  --workpath build ^
  --specpath . ^
  generate_invoice.py
if errorlevel 1 exit /b 1

echo Build complete: "%PROJECT_DIR%InvoiceGenerator.exe"
endlocal
