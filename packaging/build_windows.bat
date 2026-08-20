@echo off
rem ============================================================
rem  XingYiCha (星易查) Windows one-click build script
rem  Run on a Windows machine with Python 3.11 installed.
rem  Produces:
rem    dist\星易查\              portable folder (double-click 星易查.exe)
rem    dist\星易查-便携版.zip    portable zip
rem    dist\星易查-Setup.exe     installer (needs Inno Setup 6, optional)
rem ============================================================
setlocal enabledelayedexpansion
chcp 65001 >nul
cd /d "%~dp0.."

echo [1/6] Creating build venv (.venv-win)...
python -m venv .venv-win || goto :fail
call ".venv-win\Scripts\activate.bat" || goto :fail

echo [2/6] Installing dependencies (this may take a few minutes)...
python -m pip install --upgrade pip || goto :fail
pip install -r packaging\requirements-win.txt || goto :fail

echo [3/6] Generating app icon...
python packaging\make_icon.py || goto :fail

echo [4/6] Downloading Inno Setup Chinese language file (optional)...
powershell -NoProfile -Command "try { Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/jrsoftware/issrc/main/Files/Languages/Unofficial/ChineseSimplified.isl' -OutFile 'packaging\ChineseSimplified.isl' } catch { Write-Host 'download failed - installer UI falls back to English' }"

echo [5/6] Running PyInstaller (this may take 5-10 minutes)...
pyinstaller packaging\star.spec --noconfirm --distpath dist --workpath build\pyinstaller-win || goto :fail

echo [6/6] Packaging portable zip...
if exist "dist\星易查-便携版.zip" del "dist\星易查-便携版.zip"
powershell -NoProfile -Command "Compress-Archive -Path 'dist\星易查' -DestinationPath 'dist\星易查-便携版.zip'" || goto :fail

rem Installer: only if Inno Setup 6 is installed
set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if exist "%ISCC%" (
    echo Building installer with Inno Setup...
    "%ISCC%" "packaging\星易查.iss" || goto :fail
) else (
    echo Inno Setup 6 not found - skipped Setup.exe. Portable zip is ready.
)

echo.
echo ============================================================
echo  BUILD OK
echo    dist\星易查\              portable folder
echo    dist\星易查-便携版.zip    portable zip
if exist "dist\星易查-Setup.exe" echo    dist\星易查-Setup.exe     installer
echo ============================================================
goto :eof

:fail
echo BUILD FAILED - see messages above.
exit /b 1
