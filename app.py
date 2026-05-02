import streamlit as st
import zipfile
import io
import os
import re
import subprocess
import tempfile
import glob
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Holiday Taxis · Invoice Extractor",
    page_icon="🚖",
    layout="centered",
)

# ── Custom CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:wght@300;400;500;600&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

/* Dark navy background */
.stApp {
    background: #0a0f1e;
    color: #e8eaf0;
}

/* Hide default streamlit chrome */
#MainMenu, footer, header { visibility: hidden; }

/* Hero section */
.hero {
    text-align: center;
    padding: 3rem 1rem 2rem;
}
.hero-badge {
    display: inline-block;
    background: rgba(255,180,50,0.12);
    border: 1px solid rgba(255,180,50,0.3);
    color: #ffb432;
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.15em;
    text-transform: uppercase;
    padding: 0.3rem 1rem;
    border-radius: 100px;
    margin-bottom: 1.2rem;
}
.hero h1 {
    font-family: 'DM Serif Display', serif;
    font-size: 2.8rem;
    font-weight: 400;
    color: #ffffff;
    line-height: 1.15;
    margin: 0 0 0.8rem;
}
.hero h1 em {
    font-style: italic;
    color: #ffb432;
}
.hero p {
    color: #8892a4;
    font-size: 1rem;
    font-weight: 300;
    max-width: 480px;
    margin: 0 auto;
    line-height: 1.6;
}

/* Upload card */
.upload-card {
    background: #111827;
    border: 1px solid #1e293b;
    border-radius: 16px;
    padding: 2rem;
    margin: 1.5rem 0;
}

/* Streamlit uploader override */
[data-testid="stFileUploader"] {
    background: #0d1526;
    border: 2px dashed #1e3a5f;
    border-radius: 12px;
    transition: border-color 0.2s;
}
[data-testid="stFileUploader"]:hover {
    border-color: #ffb432;
}
[data-testid="stFileUploader"] label {
    color: #8892a4 !important;
}

/* Stat cards */
.stats-row {
    display: flex;
    gap: 1rem;
    margin: 1.5rem 0;
}
.stat-card {
    flex: 1;
    background: #111827;
    border: 1px solid #1e293b;
    border-radius: 12px;
    padding: 1.2rem 1rem;
    text-align: center;
}
.stat-value {
    font-family: 'DM Serif Display', serif;
    font-size: 2rem;
    color: #ffb432;
    line-height: 1;
    margin-bottom: 0.3rem;
}
.stat-label {
    font-size: 0.78rem;
    color: #8892a4;
    font-weight: 500;
    letter-spacing: 0.05em;
    text-transform: uppercase;
}

/* Progress area */
.log-box {
    background: #060c18;
    border: 1px solid #1e293b;
    border-radius: 10px;
    padding: 1rem 1.2rem;
    font-family: 'DM Sans', monospace;
    font-size: 0.82rem;
    color: #4ade80;
    max-height: 220px;
    overflow-y: auto;
    line-height: 1.7;
}

/* Buttons */
.stButton > button {
    background: linear-gradient(135deg, #ffb432, #ff8c00);
    color: #0a0f1e;
    font-weight: 600;
    font-family: 'DM Sans', sans-serif;
    border: none;
    border-radius: 10px;
    padding: 0.65rem 2rem;
    font-size: 0.95rem;
    width: 100%;
    cursor: pointer;
    transition: opacity 0.2s, transform 0.1s;
}
.stButton > button:hover {
    opacity: 0.9;
    transform: translateY(-1px);
}

/* Download button */
[data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, #22c55e, #16a34a);
    color: white;
    font-weight: 600;
    font-family: 'DM Sans', sans-serif;
    border: none;
    border-radius: 10px;
    padding: 0.65rem 2rem;
    font-size: 0.95rem;
    width: 100%;
}

/* Alert boxes */
.success-box {
    background: rgba(34,197,94,0.08);
    border: 1px solid rgba(34,197,94,0.25);
    border-radius: 10px;
    padding: 1rem 1.2rem;
    color: #4ade80;
    font-size: 0.9rem;
    margin: 1rem 0;
}
.warning-box {
    background: rgba(251,191,36,0.08);
    border: 1px solid rgba(251,191,36,0.25);
    border-radius: 10px;
    padding: 1rem 1.2rem;
    color: #fbbf24;
    font-size: 0.9rem;
    margin: 0.5rem 0;
}
.divider {
    border: none;
    border-top: 1px solid #1e293b;
    margin: 1.5rem 0;
}
.section-title {
    font-size: 0.78rem;
    font-weight: 600;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #8892a4;
    margin-bottom: 0.8rem;
}
</style>
""", unsafe_allow_html=True)

# ── Extraction logic ────────────────────────────────────────────────────────────

def extract_text(pdf_bytes):
    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as f:
        f.write(pdf_bytes)
        tmp = f.name
    result = subprocess.run(['pdftotext', tmp, '-'], capture_output=True, text=True)
    os.unlink(tmp)
    return result.stdout

def is_spaced(s):
    tokens = s.split()
    return len(tokens) >= 3 and sum(1 for t in tokens if len(t) <= 2) / len(tokens) > 0.5

def collapse_spaced(s):
    collapsed = s.replace(' ', '')
    result = re.sub(r'(?<=[a-z])(?=[A-Z])', ' ', collapsed)
    result = re.sub(r'\.\s*', '. ', result)
    return result.strip()

def clean_code(v):
    if re.match(r'[A-Z0-9/\s]+$', v) and is_spaced(v):
        return re.sub(r'\s+', '', v)
    return v

def parse_invoice(text, filename):
    lines = text.split('\n')
    vals = [line.strip()[2:].strip() for line in lines if line.strip().startswith(': ')]

    inv_no = date = tour_no = agent_ref = guest_name = ''
    transfer_on = ''

    for v in vals:
        cv = clean_code(v)
        if not inv_no and re.match(r'I/?[\d/]+', cv.replace(' ', '')):
            inv_no = re.sub(r'\s', '', cv)
        elif not date and re.match(r'\d{2}/\d{2}/\d{4}$', cv):
            date = cv
        elif not tour_no and re.match(r'JWZTP[\d]+', cv.replace(' ', ''), re.I):
            tour_no = re.sub(r'\s', '', cv)
        elif not agent_ref and re.match(r'[A-Z0-9]+-\d+', cv) and 'JWDILH' not in cv:
            agent_ref = cv
        elif not guest_name:
            if re.match(r'(?:Mr|Mrs|Ms|Miss|Herr|Frau|Dhr|M\b|DR|MR|MRS|MS)\b', v, re.I):
                guest_name = collapse_spaced(v) if is_spaced(v) else v
            elif is_spaced(v) and re.match(r'[A-Z][a-z\s\.]+', collapse_spaced(v)):
                candidate = collapse_spaced(v)
                if re.match(r'(?:Herr|Frau|Dhr|Mr|Mrs)', candidate):
                    guest_name = candidate

    if not tour_no:
        tm2 = re.search(r'(?:J\s*W\s*Z\s*T\s*P\s*[\d\s]+)', text)
        if tm2:
            tour_no = re.sub(r'\s', '', tm2.group(0))

    tm = re.search(r'(\d{2}/\d{2}/\d{4})\s+\d{2}/\d{2}/\d{4}', text)
    transfer_on = tm.group(1) if tm else date

    amount = 0.0
    am = re.search(r'(?:Total\s+Amount|Invoice\s*Total)[:\s]*USD\s*([\d,]+\.?\d*)', text)
    if not am: am = re.search(r'\bUSD\b\s*\n\s*([\d,]+\.?\d*)', text)
    if not am: am = re.search(r'USD\s+([\d,]+\.\d{2})\s', text)
    if am:
        try: amount = float(am.group(1).replace(',', ''))
        except: pass

    return {
        'inv_no': inv_no, 'date': date, 'tour_no': tour_no,
        'agent_ref': agent_ref, 'guest_name': guest_name,
        'transfer_on': transfer_on, 'amount': amount,
        'currency': 'USD', 'doc_type': 'TaxInvoice',
        'filename': filename,
    }

def build_excel(results, month_label=""):
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoices"

    hdr_fill = PatternFill("solid", start_color="1F4E79", end_color="1F4E79")
    alt_fill = PatternFill("solid", start_color="EBF3FB", end_color="EBF3FB")
    tot_fill = PatternFill("solid", start_color="BDD7EE", end_color="BDD7EE")
    thin = Side(style='thin', color="9DC3E6")
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)

    title = f"Holiday Taxis – Invoices{' · ' + month_label if month_label else ''}"
    ws.merge_cells('A1:J1')
    c = ws['A1']
    c.value = title
    c.font = Font(bold=True, size=13, color='FFFFFF', name='Calibri')
    c.fill = PatternFill("solid", start_color="0D3B66", end_color="0D3B66")
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    headers = ['#', 'Tour No', 'Invoice No', 'Transfer On', 'Doc. Type',
               'Guest / Group Name', 'Agent Ref', 'Amount', 'Curr.', 'Source File']
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col, value=h)
        cell.font = Font(bold=True, color='FFFFFF', name='Calibri', size=10)
        cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = bdr
    ws.row_dimensions[2].height = 22

    for i, rec in enumerate(results, 1):
        row = i + 2
        vals = [i, rec['tour_no'], rec['inv_no'], rec['transfer_on'],
                rec['doc_type'], rec['guest_name'], rec['agent_ref'],
                rec['amount'], rec['currency'], rec['filename']]
        for col, val in enumerate(vals, 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = Font(name='Calibri', size=10)
            cell.border = bdr
            if i % 2 == 0: cell.fill = alt_fill
            if col == 1: cell.alignment = Alignment(horizontal='center')
            elif col == 8:
                cell.alignment = Alignment(horizontal='right')
                cell.number_format = '#,##0.00'
            else: cell.alignment = Alignment(horizontal='left')

    tr = len(results) + 3
    for col in range(1, 11):
        c2 = ws.cell(row=tr, column=col)
        c2.fill = tot_fill; c2.border = bdr
        c2.font = Font(bold=True, name='Calibri', size=10)
    ws.cell(row=tr, column=7).value = 'TOTAL'
    ws.cell(row=tr, column=7).alignment = Alignment(horizontal='right')
    ws.cell(row=tr, column=8).value = f'=SUM(H3:H{tr-1})'
    ws.cell(row=tr, column=8).number_format = '#,##0.00'
    ws.cell(row=tr, column=8).alignment = Alignment(horizontal='right')
    ws.cell(row=tr, column=9).value = 'USD'

    for i, w in enumerate([5, 15, 20, 14, 12, 32, 20, 12, 8, 30], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A3'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── UI ─────────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="hero">
    <div class="hero-badge">🚖 Jetwing Travels · Automation</div>
    <h1>Invoice <em>Extractor</em></h1>
    <p>Upload a ZIP of Holiday Taxis PDF invoices and get a clean, formatted Excel file in seconds.</p>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="upload-card">', unsafe_allow_html=True)
st.markdown('<div class="section-title">📁 Upload Invoice ZIP</div>', unsafe_allow_html=True)

uploaded = st.file_uploader(
    "Drop your ZIP file here",
    type=["zip"],
    label_visibility="collapsed"
)

month_label = st.text_input(
    "Month label for Excel title (optional)",
    placeholder="e.g. March 2026",
    label_visibility="visible"
)

st.markdown('</div>', unsafe_allow_html=True)

if uploaded:
    st.markdown('<hr class="divider">', unsafe_allow_html=True)

    if st.button("⚡  Extract & Build Excel"):
        results = []
        warnings = []
        log_lines = []

        with zipfile.ZipFile(io.BytesIO(uploaded.read())) as zf:
            pdf_names = sorted([n for n in zf.namelist() if n.lower().endswith('.pdf')])

        if not pdf_names:
            st.error("No PDF files found in the ZIP.")
        else:
            progress = st.progress(0, text="Starting extraction…")
            log_placeholder = st.empty()

            uploaded.seek(0)
            with zipfile.ZipFile(io.BytesIO(uploaded.read())) as zf:
                for idx, name in enumerate(pdf_names):
                    pdf_bytes = zf.read(name)
                    text = extract_text(pdf_bytes)
                    rec = parse_invoice(text, os.path.basename(name))
                    results.append(rec)

                    status = "✓"
                    if not rec['guest_name'] or not rec['inv_no']:
                        status = "⚠"
                        warnings.append(name)

                    log_lines.append(
                        f"{status} {os.path.basename(name):<35} {rec['inv_no']:<22} USD {rec['amount']:>8.2f}"
                    )
                    log_placeholder.markdown(
                        f'<div class="log-box">' + '<br>'.join(log_lines[-12:]) + '</div>',
                        unsafe_allow_html=True
                    )
                    progress.progress((idx + 1) / len(pdf_names),
                                      text=f"Processing {idx+1} of {len(pdf_names)}…")

            progress.progress(1.0, text="Done!")

            # Stats
            total_amt = sum(r['amount'] for r in results)
            missing = len([r for r in results if not r['guest_name'] or not r['inv_no']])

            st.markdown(f"""
            <div class="stats-row">
                <div class="stat-card">
                    <div class="stat-value">{len(results)}</div>
                    <div class="stat-label">Invoices</div>
                </div>
                <div class="stat-card">
                    <div class="stat-value">${total_amt:,.2f}</div>
                    <div class="stat-label">Total USD</div>
                </div>
                <div class="stat-card">
                    <div class="stat-value">{len(results) - missing}</div>
                    <div class="stat-label">Complete</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            if warnings:
                st.markdown(
                    f'<div class="warning-box">⚠️ {len(warnings)} invoice(s) may have incomplete data — '
                    f'please review: {", ".join(os.path.basename(w) for w in warnings)}</div>',
                    unsafe_allow_html=True
                )
            else:
                st.markdown(
                    '<div class="success-box">✅ All invoices extracted successfully — no missing fields.</div>',
                    unsafe_allow_html=True
                )

            # Build & offer Excel download
            excel_buf = build_excel(results, month_label)
            fname = f"Holiday_Taxis_{month_label.replace(' ', '_') if month_label else 'Invoices'}.xlsx"

            st.markdown('<hr class="divider">', unsafe_allow_html=True)
            st.download_button(
                label="⬇️  Download Excel File",
                data=excel_buf,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="text-align:center; padding: 3rem 0 1rem; color: #2d3748; font-size: 0.78rem;">
    Holiday Taxis Invoice Extractor · Powered by Streamlit
</div>
""", unsafe_allow_html=True)
