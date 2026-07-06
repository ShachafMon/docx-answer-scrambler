#!/bin/bash
# Builds a single-file macOS executable (dist/shuffle_exam_mac).
# Run this ON A MAC. It creates a throwaway venv to build with pyinstaller,
# then deletes it — the resulting binary needs no venv or Python at all.
# Note: PyInstaller can't cross-compile — a binary built on Apple Silicon
# only runs on Apple Silicon, and a binary built on Intel only runs on Intel.
# Build on whichever Mac architecture your users have (or build on both and
# ship both files).
set -e
cd "$(dirname "$0")"

python3 -m venv build_venv
source build_venv/bin/activate
pip install --quiet --upgrade pip
pip install --quiet python-docx Pillow pyinstaller

pyinstaller --onefile --name shuffle_exam_mac --clean shuffle_final.py

deactivate
rm -rf build_venv build shuffle_exam_mac.spec

echo ""
echo "Done. Executable: dist/shuffle_exam_mac"
echo "Copy dist/shuffle_exam_mac next to the .docx files and run it there."
echo "Requires 'inkscape' and 'soffice' (LibreOffice) to be installed on the machine that runs it"
echo "(brew install --cask inkscape libreoffice)."
echo ""
echo "First run may be blocked by Gatekeeper since the binary isn't notarized/signed."
echo "If macOS refuses to open it: System Settings > Privacy & Security > 'Open Anyway',"
echo "or run: xattr -d com.apple.quarantine dist/shuffle_exam_mac"
