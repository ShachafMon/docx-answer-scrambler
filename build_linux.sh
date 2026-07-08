#!/bin/bash
# Builds a single-file Linux executable (dist/shuffle_exam).
# Run this ON UBUNTU. It creates a throwaway venv to build with pyinstaller,
# then deletes it — the resulting binary needs no venv or Python at all.
set -e
cd "$(dirname "$0")"

python3 -m venv build_venv
source build_venv/bin/activate
pip install --quiet --upgrade pip
pip install --quiet python-docx Pillow pyinstaller

pyinstaller --onefile --name docx-answer-scrambler-linux --clean shuffle_all.py

cp docx-answer-scrambler-linux.desktop dist/
chmod +x dist/docx-answer-scrambler-linux.desktop

deactivate
rm -rf build_venv build docx-answer-scrambler-linux.spec

echo ""
echo "Done. Executable: dist/docx-answer-scrambler-linux"
echo "Copy both dist/docx-answer-scrambler-linux and dist/docx-answer-scrambler-linux.desktop"
echo "next to the .docx files and run it there (double-click the .desktop file to get a"
echo "terminal window with live status)."
echo "Requires 'inkscape' and 'soffice' (LibreOffice) to be installed on the machine that runs it."
