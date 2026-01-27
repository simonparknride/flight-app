# Flight List Factory - Streamlit app (reverted to v12)
# - ONE-PAGE DOCX (two-column) is produced as a Word .docx file (no DOCX->PDF conversion).
# - Existing two-page DOCX and PDF labels unchanged.
# - Parser, time pickers, year selection, and parser tuning box included.

import streamlit as st
import re
import io
from datetime import datetime, timedelta, time as dtime
from typing import List, Dict, Optional
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

st.set_page_config(page_title="Flight List Factory", layout="centered", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .stApp { background-color: #000000; }
    [data-testid="stSidebar"] { background-color: #111111 !important; }
    .stMarkdown, p, h1, h2, h3, label { color: #ffffff !important; }

    div.stDownloadButton > button {
        background-color: #ffffff !important;
        color: #000000 !important;
        border: 2px solid #ffffff !important;
        border-radius: 8px !important;
        padding: 0.6rem 1.2rem !important;
        font-weight: 800 !important;
        width: 100% !important;
    }
    div.stDownloadButton > button * { color: #000000 !important; }
    div.stDownloadButton > button:hover {
        background-color: #60a5fa !important;
        color: #ffffff !important;
        border: 2px solid #60a5fa !important;
    }
    div.stDownloadButton > button:hover * { color: #ffffff !important; }

    .top-left-container { text-align: left; padding-top: 10px; margin-bottom: 20px; }
    .top-left-container a { font-size: 1.1rem; color: #ffffff !important; text-decoration: underline; display: block; margin-bottom: 5px;}
    .main-title { font-size: 3rem; font-weight: 800; color: #ffffff; line-height: 1.1; margin-bottom: 0.5rem; }
    .sub-title { font-size: 2.5rem; font-weight: 400; color: #60a5fa; }
    </style>
    """, unsafe_allow_html=True)

# --- Parsing patterns ---
TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s?[AP]M)\s+([A-Z0-9]{2,4}\d*[A-Z]?)\s*$", re.IGNORECASE)
DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
IATA_IN_PARENS = re.compile(r"\(([^)]+)\)")
PLANE_TYPES = ['A21N','A20N','A320','32Q','320','73H','737','74Y','77W','B77W','789','B789','359','A359','332','A332','AT76','DH8C','DH3','AT7','388','333','A333','330','76V','77L','B38M','A388','772','B772','32X','77X']
PLANE_TYPE_PATTERN = re.compile(r"\b(" + "|".join(sorted(set(PLANE_TYPES), key=len, reverse=True)) + r")\b", re.IGNORECASE)

NORMALIZE_MAP = {
    '32q': 'A320', '320': 'A320', 'a320': 'A320', '32x': 'A320',
    '789': 'B789', 'b789': 'B789', '772': 'B772', 'b772': 'B772',
    '77w': 'B77W', 'b77w': 'B77W', '332': 'A332', 'a332': 'A332',
    '333': 'A333', 'a333': 'A333', '330': 'A330', 'a330': 'A330',
    '359': 'A359', 'a359': 'A359', '388': 'A388', 'a388': 'A388',
    '737': 'B737', '73h': 'B737', 'at7': 'AT76'
}
ALLOWED_AIRLINES = {"NZ", "QF", "JQ", "CZ", "CA", "SQ", "LA", "IE", "FX"}
NZ_DOMESTIC_IATA = {"AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO","WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"}
REGO_LIKE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9\-–—]*$")

def normalize_type(t: Optional[str]) -> str:
    if not t:
        return ""
    key = t.strip().lower()
    return NORMALIZE_MAP.get(key, t.strip().upper())

def try_parse_date_header(line: str, year: int) -> Optional[datetime.date]:
    candidates = [
        "%A, %b %d %Y",
        "%A, %B %d %Y",
        "%a, %b %d %Y",
        "%a, %B %d %Y",
    ]
    text = line.strip() + f" {year}"
    for fmt in candidates:
        try:
            return datetime.strptime(text, fmt).date()
        except Exception:
            continue
    return None

def parse_raw_lines(lines: List[str], year: int) -> List[Dict]:
    records = []
    current_date = None
    i = 0
    L = len(lines)
    while i < L:
        line = lines[i].strip()
        if DATE_HEADER.match(line):
            parsed = try_parse_date_header(line, year)
            current_date = parsed if parsed else None
            i += 1
            continue

        m = TIME_LINE.match(line)
        if m and current_date is not None:
            time_str_raw, flight_raw = m.groups()
            dest_line = lines[i+1].strip() if i+1 < L else ''
            carrier_line = lines[i+2].rstrip('\n') if i+2 < L else ''
            m2 = IATA_IN_PARENS.search(dest_line)
            dest_iata = (m2.group(1).strip().upper() if m2 else '').upper()
            mtype = PLANE_TYPE_PATTERN.search(carrier_line or '')
            plane_type = normalize_type(mtype.group(1) if mtype else '')
            reg = ''
            parens = IATA_IN_PARENS.findall(carrier_line or '')
            if parens:
                for candidate in reversed(parens):
                    cand = candidate.strip()
                    if REGO_LIKE.match(cand) and ('-' in cand or '–' in cand or '—' in cand):
                        reg = cand
                        break
                if not reg:
                    reg = parens[-1].strip()

            dep_dt = None
            try:
                tnorm = time_str_raw.strip().upper().replace(" ", "")
                if re.match(r"^\d{1,2}:\d{2}[AP]M$", tnorm):
                    dep_dt = datetime.strptime(f"{current_date} {tnorm}", "%Y-%m-%d %I:%M%p")
                else:
                    dep_dt = datetime.strptime(f"{current_date} {time_str_raw.strip()}", "%Y-%m-%d %I:%M %p")
            except Exception:
                dep_dt = None

            records.append({
                'dt': dep_dt,
                'time': time_str_raw.strip(),
                'flight': flight_raw.strip().upper(),
                'dest': dest_iata,
                'type': plane_type,
                'reg': reg
            })
            i += 3
            continue

        i += 1
    return records

def filter_records(records: List[Dict], start_time: dtime, end_time: dtime):
    dates = sorted({r['dt'].date() for r in records if r.get('dt')})
    if not dates:
        return [], None, None
    day1 = dates[0]
    day2 = dates[1] if len(dates) >= 2 else (day1 + timedelta(days=1))
    start_dt = datetime.combine(day1, start_time)
    end_dt = datetime.combine(day2, end_time)
    if end_dt <= start_dt:
        return [], start_dt, end_dt

    def allowed(r):
        if not r.get('dt'):
            return False
        flight_prefix = (r.get('flight') or '')[:2].upper()
        if flight_prefix not in ALLOWED_AIRLINES:
            return False
        dest = (r.get('dest') or '').upper()
        if dest in NZ_DOMESTIC_IATA:
            return False
        return start_dt <= r['dt'] <= end_dt

    out = [r for r in records if allowed(r)]
    out.sort(key=lambda x: x['dt'] or datetime.max)
    return out, start_dt, end_dt

# --- Existing DOCX (two-page / original style) ---
def build_docx_stream(records: List[Dict], start_dt: datetime, end_dt: datetime) -> io.BytesIO:
    doc = Document()
    font_name = 'Air New Zealand Sans'
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Inches(0.3)
    section.left_margin = section.right_margin = Inches(0.5)

    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_f = footer_para.add_run("created by Simon Park'nRide's Flight List Factory 2026")
    run_f.font.name = font_name
    run_f.font.size = Pt(10)
    run_f.font.color.rgb = RGBColor(128, 128, 128)
    rPr_f = run_f._element.get_or_add_rPr()
    rFonts_f = OxmlElement('w:rFonts')
    rFonts_f.set(qn('w:ascii'), font_name); rFonts_f.set(qn('w:hAnsi'), font_name)
    rPr_f.append(rFonts_f)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_head = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}")
    run_head.bold = True
    run_head.font.name = font_name
    run_head.font.size = Pt(16)
    rPr_h = run_head._element.get_or_add_rPr()
    rFonts_h = OxmlElement('w:rFonts')
    rFonts_h.set(qn('w:ascii'), font_name); rFonts_h.set(qn('w:hAnsi'), font_name)
    rPr_h.append(rFonts_h)

    table = doc.add_table(rows=0, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = table._element.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr'); table._element.insert(0, tblPr)
    tblW = OxmlElement('w:tblW'); tblW.set(qn('w:w'), '4000'); tblW.set(qn('w:type'), 'pct'); tblPr.append(tblW)

    for i, r in enumerate(records):
        row = table.add_row()
        try:
            tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        except Exception:
            tdisp = r['time']
        vals = [r['flight'], tdisp, r['dest'], r['type'], r['reg']]
        for j, val in enumerate(vals):
            cell = row.cells[j]
            if i % 2 == 1:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9'); tcPr.append(shd)
            para = cell.paragraphs[0]
            para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            para.paragraph_format.space_before = para.paragraph_format.space_after = Pt(0)
            run = para.add_run(str(val))
            run.font.name = font_name
            run.font.size = Pt(14)
            rPr = run._element.get_or_add_rPr()
            rFonts = OxmlElement('w:rFonts')
            rFonts.set(qn('w:ascii'), font_name); rFonts.set(qn('w:hAnsi'), font_name)
            rPr.append(rFonts)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- ONE-PAGE DOCX (fixed fonts: body 11pt, Reg 9pt) ---
def build_docx_onepage_stream(records: List[Dict], start_dt: datetime, end_dt: datetime) -> io.BytesIO:
    """
    Create a one-page DOCX with two columns.
    Body font is fixed at 11pt; Reg column uses 9pt.
    Other styles (header, striping, fonts) follow the original.
    """
    doc = Document()
    font_name = 'Air New Zealand Sans'
    section = doc.sections[0]

    # tighten margins to help everything fit on a single page
    section.top_margin = section.bottom_margin = Inches(0.2)
    section.left_margin = section.right_margin = Inches(0.4)

    # Footer (same style)
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_f = footer_para.add_run("created by Simon Park'nRide's Flight List Factory 2026")
    run_f.font.name = font_name
    run_f.font.size = Pt(10)
    run_f.font.color.rgb = RGBColor(128, 128, 128)
    rPr_f = run_f._element.get_or_add_rPr()
    rFonts_f = OxmlElement('w:rFonts')
    rFonts_f.set(qn('w:ascii'), font_name); rFonts_f.set(qn('w:hAnsi'), font_name)
    rPr_f.append(rFonts_f)

    # Header centered (same as original)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_head = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}")
    run_head.bold = True
    run_head.font.name = font_name
    run_head.font.size = Pt(16)
    rPr_h = run_head._element.get_or_add_rPr()
    rFonts_h = OxmlElement('w:rFonts')
    rFonts_h.set(qn('w:ascii'), font_name); rFonts_h.set(qn('w:hAnsi'), font_name)
    rPr_h.append(rFonts_h)

    # Split records into two columns
    total = len(records)
    mid = (total + 1) // 2
    left_recs = records[:mid]
    right_recs = records[mid:]

    # Outer 1x2 table to emulate two columns
    outer = doc.add_table(rows=1, cols=2)
    outer.alignment = WD_TABLE_ALIGNMENT.CENTER

    # populate inner tables
    def add_inner_table(cell, recs, start_index=0):
        inner = cell.add_table(rows=1, cols=5)
        hdr_cells = inner.rows[0].cells
        headers = ['Flight', 'Time', 'Dest', 'Type', 'Reg']
        for idx, text in enumerate(headers):
            run = hdr_cells[idx].paragraphs[0].add_run(text)
            run.bold = True
            run.font.size = Pt(11)  # header in inner table
            run.font.name = font_name
            rPr = run._element.get_or_add_rPr()
            rFonts = OxmlElement('w:rFonts')
            rFonts.set(qn('w:ascii'), font_name); rFonts.set(qn('w:hAnsi'), font_name)
            rPr.append(rFonts)
        # add rows
        for i, r in enumerate(recs):
            row = inner.add_row()
            try:
                tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
            except Exception:
                tdisp = r['time']
            vals = [r['flight'], tdisp, r['dest'], r['type'], r['reg']]
            for j, val in enumerate(vals):
                cell_j = row.cells[j]
                if (start_index + i) % 2 == 1:
                    tcPr = cell_j._tc.get_or_add_tcPr()
                    shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9'); tcPr.append(shd)
                para = cell_j.paragraphs[0]
                para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                para.paragraph_format.space_before = para.paragraph_format.space_after = Pt(0)
                run = para.add_run(str(val))
                run.font.name = font_name
                # body font fixed at 11pt; Reg column (index 4) is 9pt
                if j == 4:
                    run.font.size = Pt(9)
                else:
                    run.font.size = Pt(11)
                rPr = run._element.get_or_add_rPr()
                rFonts = OxmlElement('w:rFonts')
                rFonts.set(qn('w:ascii'), font_name); rFonts.set(qn('w:hAnsi'), font_name)
                rPr.append(rFonts)
        # set inner column widths (best-effort)
        col_widths = [Inches(1.3), Inches(0.9), Inches(0.9), Inches(1.0), Inches(1.2)]
        for ci, w in enumerate(col_widths):
            try:
                inner.columns[ci].width = w
            except Exception:
                pass

    left_cell = outer.cell(0, 0)
    right_cell = outer.cell(0, 1)
    add_inner_table(left_cell, left_recs, start_index=0)
    add_inner_table(right_cell, right_recs, start_index=len(left_recs))

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- PDF labels generation (unchanged) ---
def build_labels_stream(records: List[Dict], start_num: int) -> io.BytesIO:
    target = io.BytesIO()
    c = canvas.Canvas(target, pagesize=A4)
    w, h = A4
    margin, gutter = 15*mm, 6*mm
    col_w, row_h = (w - 2*margin - gutter) / 2, (h - 2*margin) / 5
    try:
        start_num = int(start_num)
    except Exception:
        start_num = 1

    for i, r in enumerate(records):
        if i > 0 and i % 10 == 0:
            c.showPage()
        idx = i % 10
        x_left = margin + (idx % 2) * (col_w + gutter)
        y_top = h - margin - (idx // 2) * row_h
        c.setStrokeGray(0.3); c.setLineWidth(0.2); c.rect(x_left, y_top - row_h + 2*mm, col_w, row_h - 4*mm)
        c.setLineWidth(0.5); c.rect(x_left + 3*mm, y_top - 12*mm, 8*mm, 8*mm)
        c.setFont('Helvetica-Bold', 14); c.drawCentredString(x_left + 7*mm, y_top - 9.5*mm, str(start_num + i))
        date_str = r['dt'].strftime('%d %b') if r.get('dt') else ''
        c.setFont('Helvetica-Bold', 18); c.drawRightString(x_left + col_w - 4*mm, y_top - 11*mm, date_str)
        c.setFont('Helvetica-Bold', 38); c.drawString(x_left + 15*mm, y_top - 21*mm, r['flight'])
        c.setFont('Helvetica-Bold', 23); c.drawString(x_left + 15*mm, y_top - 33*mm, r['dest'])
        tdisp = ''
        try:
            tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        except Exception:
            tdisp = r.get('time','')
        c.setFont('Helvetica-Bold', 29); c.drawString(x_left + 15*mm, y_top - 47*mm, tdisp)
        c.setFont('Helvetica', 13); c.drawRightString(x_left + col_w - 6*mm, y_top - row_h + 12*mm, r.get('type',''))
        c.drawRightString(x_left + col_w - 6*mm, y_top - row_h + 7*mm, r.get('reg',''))
    c.save(); target.seek(0)
    return target

# --- Sidebar / inputs (time pickers + year) ---
with st.sidebar:
    st.header("⚙️ Settings")
    year = st.number_input("Year", value=datetime.now().year, min_value=2000, max_value=2100)
    s_time = st.time_input("Start Time", value=dtime(hour=4, minute=55))
    e_time = st.time_input("End Time", value=dtime(hour=5, minute=0))
    label_start = st.number_input("Label Start Number", value=1, min_value=1)

st.markdown('<div class="top-left-container"><a href="https://www.flightradar24.com/data/airports/akl/arrivals" target="_blank">Import Raw Text</a><a href="https://www.flightradar24.com/data/airports/akl/departures" target="_blank">Export Raw Text</a></div>', unsafe_allow_html=True)
st.markdown('<div class="main-title">Simon Park\'nRide\'s<br><span class="sub-title">Flight List Factory</span></div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Raw Text File", type=['txt'])

# Parser tuning area
st.subheader("Parser Tuning")
sample_text = st.text_area("Optional: paste a small sample of your raw text here for parser tuning (include date header + 2–3 flights).", height=160)
if st.button("Run Parser on Sample"):
    if not sample_text.strip():
        st.warning("Please paste some sample lines first.")
    else:
        sample_lines = sample_text.splitlines()
        parsed = parse_raw_lines(sample_lines, year)
        if parsed:
            st.success(f"Parsed {len(parsed)} record(s) from sample")
            st.json(parsed)
        else:
            st.warning("No records parsed from the sample — include the date header and at least one flight block.")

if uploaded_file:
    try:
        content = uploaded_file.read()
        if isinstance(content, bytes):
            lines = content.decode("utf-8", errors="replace").splitlines()
        else:
            lines = str(content).splitlines()
    except Exception as e:
        st.error(f"Failed to read uploaded file: {e}")
        lines = []

    if lines:
        all_recs = parse_raw_lines(lines, year)
        if not all_recs:
            st.warning("No records parsed from the file. Use the Parser Tuning box to paste a sample and iterate.")
        else:
            filtered, s_dt, e_dt = filter_records(all_recs, s_time, e_time)
            if s_dt and e_dt and e_dt <= s_dt:
                st.error("Invalid time window: end time must be later than start time across the two-day window.")
            elif not filtered:
                st.warning("No flights matched the filters (time window, permitted airlines, or destination exclusions).")
            else:
                st.success(f"Processed {len(filtered)} flights (year {year})")
                col1, col_mid, col2 = st.columns([1, 0.9, 1])
                fn = f"List_{s_dt.strftime('%d-%m')}" if s_dt else "List"

                # original (two-page) DOCX
                docx_bytes = build_docx_stream(filtered, s_dt, e_dt).getvalue()
                col1.download_button("📥 Download DOCX List (2 pages)", data=docx_bytes, file_name=f"{fn}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

                # one-page two-column DOCX (v12) - WORD file only (no pdf conversion)
                onepage_bytes = build_docx_onepage_stream(filtered, s_dt, e_dt).getvalue()
                col_mid.download_button("📥 Download DOCX One-Page (2 columns)", data=onepage_bytes, file_name=f"{fn}_onepage.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

                # PDF labels
                pdf_bytes = build_labels_stream(filtered, label_start).getvalue()
                col2.download_button("📥 Download PDF Labels", data=pdf_bytes, file_name=f"Labels_{fn}.pdf", mime="application/pdf")

                # show table preview
                table_rows = []
                for i, r in enumerate(filtered):
                    try:
                        tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
                    except Exception:
                        tdisp = r['time']
                    table_rows.append({'No': label_start + i, 'Flight': r['flight'], 'Time': tdisp, 'Dest': r['dest'], 'Type': r['type'], 'Reg': r['reg']})
                st.table(table_rows)
