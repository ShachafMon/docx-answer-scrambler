@echo off
REM Builds a single-file Windows executable (dist\docx-answer-scrambler.exe).
REM Run this ON WINDOWS, with Python installed. It creates a throwaway venv
REM to build with pyinstaller, then deletes it - the resulting .exe needs no
REM venv or Python installed to run.
cd /d "%~dp0"

python -m venv build_venv
call build_venv\Scripts\activate.bat
pip install --quiet --upgrade pip
pip install --quiet python-docx Pillow pyinstaller

pyinstaller --onefile --name docx-answer-scrambler --clean shuffle_all.py

call build_venv\Scripts\deactivate.bat
rmdir /s /q build_venv
rmdir /s /q build
del docx-answer-scrambler.spec

echo.
echo Done. Executable: dist\docx-answer-scrambler.exe
echo Copy dist\docx-answer-scrambler.exe next to the .docx files and run it there.
echo Requires Inkscape and LibreOffice (soffice) to be installed and on PATH.
pause
