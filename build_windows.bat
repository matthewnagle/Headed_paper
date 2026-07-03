@echo off
REM Build the standalone Windows executable for the Headed Paper PDF Generator.
REM Requires Python 3.10+ installed and on PATH. Run from this folder.

python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller
python -m PyInstaller headed_paper.spec --noconfirm

echo.
echo Done. The standalone app is at dist\HeadedPaperPDF.exe
pause
