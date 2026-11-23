# Smart Vehicle Data Automation

Streamlit app: upload Excel + messages + meter images → OCR → match → verify → updated Excel.

## Features
- Reads vehicle lists from `.xlsx` or `.xls` workbooks.
- Accepts WhatsApp exports or direct text/image uploads.
- Uses Tesseract OCR to extract meter readings and general text from images.
- Writes extracted text to an `Extracted Text` Excel column.
- Embeds the uploaded source image directly into the matching Excel row.
- Keeps an image file path in the `Image` column for reference.
- Supports duplicate skip/update and unmatched-vehicle reporting.

## Deploy (access from anywhere, free)
1. Create a GitHub repo and upload `app.py`, `requirements.txt`, `packages.txt`, `README.md`.
2. Go to https://share.streamlit.io → sign in with GitHub → **New app** → pick the repo, branch `main`, main file `app.py` → **Deploy**.
3. You get a public URL that works on any device.

## Run locally
```
pip install -r requirements.txt
# install Tesseract OCR (Windows: UB-Mannheim installer, Linux: apt install tesseract-ocr)
streamlit run app.py
```

## Notes
- Excel header row = row 1. Columns Date, Shift, Meter Reading, Extracted Text, and Image are added if missing.
- Unique key = vehicle + date + shift; duplicates can be skipped or updated.
- Streamlit Cloud storage is temporary: always download the updated Excel after each session.
