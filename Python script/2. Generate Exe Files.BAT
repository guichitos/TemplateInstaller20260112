@echo off
setlocal

echo ============================================
echo Cleaning previous build artifacts
echo ============================================

REM --- Remove build folder ---
if exist build (
    echo Deleting build folder...
    rmdir /s /q build
    if exist build (
        echo ERROR: Could not delete build folder
        exit /b 1
    )
)

REM --- Remove dist folder ---
if exist dist (
    echo Deleting dist folder...
    rmdir /s /q dist
    if exist dist (
        echo ERROR: Could not delete dist folder
        exit /b 1
    )
)

echo.
echo ============================================
echo Building installer
echo ============================================

py -m PyInstaller ^
    --onefile ^
    --console ^
    --name "1. Pin templates to Office apps home screen" ^
    installer.py

IF ERRORLEVEL 1 (
    echo ERROR: Failed to build installer
    exit /b 1
)

echo.
echo ============================================
echo Building uninstaller
echo ============================================

py -m PyInstaller ^
    --onefile ^
    --console ^
    --name "2. Unpin templates from Office apps home screen" ^
    uninstaller.py

IF ERRORLEVEL 1 (
    echo ERROR: Failed to build uninstaller
    exit /b 1
)

echo.
echo ============================================
echo Build completed successfully
echo Opening dist folder...
echo ============================================

explorer "%cd%\dist"
