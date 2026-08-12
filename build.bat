@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ============================================
echo  PassPro2 - clean Windows rebuild
echo ============================================
echo.
echo IMPORTANT:
echo  - Do NOT run anything under the "build" folder.
echo  - The working app is: dist\PassPro2\PassPro2.exe
echo  - Prefer a short local path (e.g. C:\PassPro2), not a deep OneDrive folder.
echo  - Your passwords file assets\data2.xlsx is NEVER uploaded to GitHub.
echo.

where python >nul 2>&1
if errorlevel 1 (
  echo ERROR: Python was not found on PATH.
  echo Install Python 3.11+ from https://www.python.org/downloads/
  echo Make sure "Add python.exe to PATH" is checked.
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo Creating virtual environment...
  python -m venv .venv
  if errorlevel 1 (
    echo ERROR: failed to create .venv
    pause
    exit /b 1
  )
)

call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 (
  echo ERROR: pip install failed
  pause
  exit /b 1
)

echo.
echo Cleaning old build output...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo.
echo Building with PyInstaller...
pyinstaller --noconfirm PassPro2.spec
if errorlevel 1 (
  echo ERROR: PyInstaller build failed
  pause
  exit /b 1
)

REM Keep the user's existing password workbook next to the new EXE (never overwrite).
if not exist "dist\PassPro2\assets" mkdir "dist\PassPro2\assets"
if exist "assets\data2.xlsx" (
  if not exist "dist\PassPro2\assets\data2.xlsx" (
    echo Copying existing assets\data2.xlsx next to the EXE...
    copy /y "assets\data2.xlsx" "dist\PassPro2\assets\data2.xlsx" >nul
  ) else (
    echo dist\PassPro2\assets\data2.xlsx already exists - leaving it untouched.
  )
)

echo.
echo ============================================
echo  BUILD OK
echo ============================================
echo Run this file:
echo   %cd%\dist\PassPro2\PassPro2.exe
echo.
echo If Windows says python311.dll is missing:
echo   1) Run ONLY the EXE above (not build\...)
echo   2) Right-click the PassPro2 folder -^> Always keep on this device (OneDrive)
echo   3) Install VC++ Redistributable x64:
echo      https://aka.ms/vs/17/release/vc_redist.x64.exe
echo   4) Prefer copying dist\PassPro2 to C:\PassPro2 and run from there
echo.

start "" "dist\PassPro2"
pause
