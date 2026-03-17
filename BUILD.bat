@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

echo ==========================================
echo Building iPhone_decryptor.exe from main.py
echo ==========================================

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist iPhone_decryptor.spec del /f /q iPhone_decryptor.spec

where py >nul 2>&1
if not errorlevel 1 (
    set "PY_CMD=py -3"
) else (
    where python >nul 2>&1
    if not errorlevel 1 (
        set "PY_CMD=python"
    ) else (
        echo [ERROR] Python was not found in PATH.
        pause
        exit /b 1
    )
)

echo [1/4] Upgrading pip...
%PY_CMD% -m pip install --upgrade pip
if errorlevel 1 goto :fail

echo [2/4] Installing build requirements...
%PY_CMD% -m pip install pyinstaller PySide6
if errorlevel 1 goto :fail

set "ICON_FILE="
if exist "%cd%\phone.ico" set "ICON_FILE=%cd%\phone.ico"
if not defined ICON_FILE if exist "%cd%\icons\phone.ico" set "ICON_FILE=%cd%\icons\phone.ico"
if not defined ICON_FILE if exist "%cd%\wicon.ico" set "ICON_FILE=%cd%\wicon.ico"
if not defined ICON_FILE if exist "%cd%\icons\wicon.ico" set "ICON_FILE=%cd%\icons\wicon.ico"

if defined ICON_FILE (
    echo [INFO] Using icon: "%ICON_FILE%"
    set "ICON_ARG=--icon=%ICON_FILE%"
) else (
    echo [WARNING] No ICO file found. EXE will use default icon.
    set "ICON_ARG="
)

echo [3/4] Running PyInstaller...
%PY_CMD% -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name "iPhone_decryptor" ^
  --add-data "%cd%\icons;icons" ^
  !ICON_ARG! ^
  "%cd%\main.py"
if errorlevel 1 goto :fail

echo [4/4] Build complete.
echo.
if exist "dist\iPhone_decryptor.exe" (
    echo EXE created at:
    echo %cd%\dist\iPhone_decryptor.exe
    echo.
    echo If the icon still looks old in Windows Explorer:
    echo - close File Explorer and reopen it
    echo - rename the EXE once
    echo - or delete build/dist and rebuild again
    start "" "dist"
) else (
    echo [WARNING] Build finished but EXE was not found in dist.
)

echo.
echo Done.
pause
exit /b 0

:fail
echo.
echo [ERROR] Build failed.
pause
exit /b 1
