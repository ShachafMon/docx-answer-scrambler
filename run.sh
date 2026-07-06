#!/bin/bash
cd "$(dirname "$0")"

if [ ! -d "venv" ]; then
    echo "יוצר סביבה וירטואלית בפעם הראשונה..."
    python3 -m venv venv
    source venv/bin/activate
    pip install --quiet python-docx Pillow
else
    source venv/bin/activate
fi

echo "מתחיל עיבוד..."
echo ""
python3 shuffle_all.py

echo ""
echo "הסתיים! לחץ Enter לסגירת החלון."
read