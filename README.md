# 🚖 Holiday Taxis Invoice Extractor

A Streamlit web app that extracts data from Holiday Taxis PDF invoices (in a ZIP) and produces a formatted Excel file.

## Features
- Upload a ZIP of PDF invoices
- Extracts: Tour No, Invoice No, Transfer Date, Guest Name, Agent Ref, Amount
- Handles spaced-out PDF text rendering automatically
- Downloads a formatted, colour-coded Excel file with totals

## Files
```
app.py            ← Main Streamlit application
requirements.txt  ← Python dependencies
```

PDF text extraction uses `pdftotext` (poppler) when that binary is available,
and falls back to the pure-Python `pdfminer.six` otherwise — so no system
packages are needed to deploy.

## Deploy to Streamlit Community Cloud (Free)

1. **Push to GitHub**
   - Create a new GitHub repository (public or private)
   - Upload these 2 files: `app.py`, `requirements.txt`

2. **Deploy on Streamlit Cloud**
   - Go to [share.streamlit.io](https://share.streamlit.io)
   - Sign in with GitHub
   - Click **"New app"**
   - Select your repository and set **Main file path** to `app.py`
   - Click **Deploy** — done! You'll get a public URL like `https://yourapp.streamlit.app`

## Run Locally
```bash
pip install -r requirements.txt

# Optional — sharper extraction if poppler is installed:
# sudo apt install poppler-utils   # Linux / WSL
# brew install poppler             # macOS

streamlit run app.py
```
