@echo off
title TIFF to DOCX Converter
color 0A

echo ============================================================
echo        TIFF to DOCX Converter - Google Vision OCR
echo ============================================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERREUR] Python n'est pas installe ou pas dans le PATH
    pause
    exit /b 1
)

:: Install dependencies
echo [INFO] Verification des dependances...
pip install -r requirements.txt --quiet

if errorlevel 1 (
    echo [ERREUR] Impossible d'installer les dependances
    pause
    exit /b 1
)

echo.
echo [OK] Dependances installees
echo.
echo ============================================================
echo    Demarrage du serveur...
echo    Ouvrez votre navigateur a l'adresse:
echo.
echo    http://localhost:5000
echo.
echo    Appuyez sur Ctrl+C pour arreter le serveur
echo ============================================================
echo.

:: Launch app
python app.py

pause
