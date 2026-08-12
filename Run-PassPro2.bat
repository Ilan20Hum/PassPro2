@echo off
REM Quick launcher for PassPro2 after you download/unzip the app.
cd /d "%~dp0"
if not exist "PassPro2.exe" (
  echo PassPro2.exe not found in this folder.
  pause
  exit /b 1
)
if not exist "assets\data2.xlsx" (
  if exist "assets\data2.template.xlsx" (
    echo Creating empty assets\data2.xlsx from template...
    if not exist assets mkdir assets
    copy /y "assets\data2.template.xlsx" "assets\data2.xlsx" >nul
  )
)
start "" "PassPro2.exe"
