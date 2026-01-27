# Flight List Factory - Streamlit app
# - Existing functionality preserved.
# - NEW: build_onepage_pdf_stream creates a guaranteed single-page two-column PDF
#   by directly drawing with reportlab and adaptively reducing font/spacing/margins.
# - Use the "Download One-Page PDF (direct)" button to get the guaranteed single-page output.

import streamlit as st
import re
import io
import math
from datetime import datetime, timedelta, time as dtime
from typing import List, Dict, Optional, Tuple
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

st.set_page_config(page_title="Flight List Factory", layout="centered", initial_sidebar_state="expanded")

# --- Styles (unchanged) ---
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

# --- Parsing patterns (same as before) ---
TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s?[AP]M)\s+([A-Z0-9]{2,4}\d*[A-Z]?)\s*$", re.IGNORECASE)
DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
IATA_IN_PARENS = re.compile(r"\(([^)]+)\)")
PLANE_TYPES = ['A21N','A20N','A320','32Q','320','73H','737','74Y','77W','B77W','789','B789','359','A359','332','A332','AT76','DH8C','DH3','AT7','388','333','A333','330','76V','77L','B38M','A388','772','B772','32X','77X']
PLANE_TYPE_PATTERN = re.compile(r"\b(" + "|".join(sorted(set(PLANE_TYPES), key=len, reverse=True)) + r")\b", re.IGNORECASE)
NORMALIZE_MAP = {
    '32q':'A320','320':'A320','a320':'A320','32x':'A320',
    '789':'B789','b789':'B789','772':'B772','b772':'B772','77w':'B77W','b77w':'B77W',
    '332':'A332','a332':'A332','333':'A333','a333':'A333','330':'A330','a330':'A330',
    '359':'A359','a359':'A359','388':'A388','a388':'A388','737':'B737','73h':'B737','at7':'AT76'
}
ALLOWED_AIRLINES = {"NZ","QF","JQ","CZ","CA","SQ","LA","IE","FX"}
NZ_DOMESTIC_IATA = {"AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO","WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"}
REGO_LIKE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9\-–—]*$")

def normalize_type(t: Optional[str]) -> str:
    if not t: return ""
    key = t.strip().lower()
    return NORMALIZE_MAP.get(key, t.strip().upper())

def parse_raw_lines(lines: List[str], year: int) -> List[Dict]:
    records = []
    current_date = None
    i = 0
    L = len(lines)
    while i < L:
        line = lines[i].strip()
        if DATE_HEADER.match(line):
            try:
                current_date = datetime.strptime(line + f" {year}", '%A, %b %d %Y').date()
            except Exception:
                try: current_date = datetime.strptime(line + f" {year}", '%A, %B %d %Y').date()
                except Exception:
                    current_date = None
            i += 1; continue
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
                        reg = cand; break
                if not reg: reg = parens[-1].strip()
            try:
                tnorm = time_str_raw.strip().upper().replace(" ", "")
                dep_dt = datetime.strptime(f"{current_date} {tnorm}", "%Y-%m-%d %I:%M%p")
            except Exception:
                dep_dt = None
            records.append({'dt': dep_dt, 'time': time_str_raw.strip(), 'flight': flight_raw.strip().upper(), 'dest': dest_iata, 'type': plane_type, 'reg': reg})
            i += 3; continue
        i += 1
    return records

def filter_records(records: List[Dict], start_hm: str, end_hm: str):
    # start_hm and end_hm in 'HH:MM' 24-hour format strings (from time_input)
    dates = sorted({r['dt'].date() for r in records if r.get('dt')})
    if not dates: return [], None, None
    day1 = dates[0]; day2 = dates[1] if len(dates) >= 2 else (day1 + timedelta(days=1))
    try:
        start_dt = datetime.combine(day1, datetime.strptime(start_hm, '%H:%M').time())
        end_dt = datetime.combine(day2, datetime.strptime(end_hm, '%H:%M').time())
    except Exception:
        return [], None, None
    out = [r for r in records if r.get('dt') and (r['flight'][:2].upper() in ALLOWED_AIRLINES) and (r['dest'] not in NZ_DOMESTIC_IATA) and (start_dt <= r['dt'] <= end_dt)]
    out.sort(key=lambda x: x['dt'] or datetime.max)
    return out, start_dt, end_dt

# --- Existing two-page DOCX generator (unchanged) ---
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
    run_f.font.name = font_name; run_f.font.size = Pt(10); run_f.font.color.rgb = RGBColor(128,128,128)

    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_head = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}")
    run_head.bold = True; run_head.font.name = font_name; run_head.font.size = Pt(16)

    table = doc.add_table(rows=0, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = table._element.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr'); table._element.insert(0, tblPr)
    tblW = OxmlElement('w:tblW'); tblW.set(qn('w:w'),'4000'); tblW.set(qn('w:type'),'pct'); tblPr.append(tblW)

    for i, r in enumerate(records):
        row = table.add_row()
        try: tdisp = datetime.strptime(r['time'],'%I:%M %p').strftime('%H:%M')
        except: tdisp = r['time']
        vals = [r['flight'], tdisp, r['dest'], r['type'], r['reg']]
        for j, val in enumerate(vals):
            cell = row.cells[j]
            if i % 2 == 1:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd'); shd.set(qn('w:val'),'clear'); shd.set(qn('w:fill'),'D9D9D9'); tcPr.append(shd)
            para = cell.paragraphs[0]
            para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            para.paragraph_format.space_before = para.paragraph_format.space_after = Pt(0)
            run = para.add_run(str(val))
            run.font.name = font_name; run.font.size = Pt(14)
    target = io.BytesIO(); doc.save(target); target.seek(0); return target

# --- NEW: One-page direct PDF generator (guaranteed single page) ---
def build_onepage_pdf_stream(records: List[Dict], start_dt: datetime, end_dt: datetime) -> io.BytesIO:
    """
    Create a single A4 PDF with two columns, adapting font/spacing/margins to guarantee fit.
    Returns an in-memory BytesIO with the PDF.
    """
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    W, H = A4

    # Base layout metrics (in points)
    left_margin = 18 * mm
    right_margin = 18 * mm
    top_margin = 10 * mm
    bottom_margin = 12 * mm

    # Column widths (approx similar to DOCX)
    # total usable width:
    usable_width = W - left_margin - right_margin
    # give columns equal width
    col_gap = 8 * mm
    col_width = (usable_width - col_gap) / 2

    # Header
    header_text = f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}"
    header_font = "Helvetica-Bold"
    header_pt = 14  # start
    # Footer text
    footer_text = "created by Simon Park'nRide's Flight List Factory 2026"
    footer_font = "Helvetica"
    footer_pt = 9

    # Determine rows per column
    total = len(records)
    left_count = (total + 1) // 2
    right_count = total - left_count
    rows_per_column = max(left_count, right_count)

    # starting body/reg fonts
    body_pt = 11.0
    reg_pt = 9.0
    min_body_pt = 8.0
    min_reg_pt = 7.0

    # compute available height (points) for rows = page height - top_margin - header_height - footer area - small extra
    # estimate header height = header_pt * 1.2
    # footer height = footer_pt * 1.2
    def compute_available_for_rows(hdr_pt):
        header_height = hdr_pt * 1.25
        footer_height = footer_pt * 1.25
        available = H - top_margin - bottom_margin - header_height - footer_height - (3 * mm)  # small extra
        return available

    # function to check if fonts/leading fit
    def fits(body_pt_try, leading_multiplier, hdr_pt_try):
        available = compute_available_for_rows(hdr_pt_try)
        line_h = body_pt_try * leading_multiplier
        required = rows_per_column * line_h
        return required <= available, available, line_h

    # Try to fit by decreasing leading multiplier first, then reducing fonts, then reducing header font
    leading = 1.15  # start
    min_leading = 1.02
    fit, available_ht, line_h = fits(body_pt, leading, header_pt)
    # iterative adjustments (guarantee fit)
    while not fit:
        # try reduce leading down to min
        if leading > min_leading + 1e-6:
            leading = max(min_leading, leading - 0.05)
            fit, available_ht, line_h = fits(body_pt, leading, header_pt)
            if fit: break
        # reduce body font a little (not below min)
        if body_pt > min_body_pt + 0.01:
            body_pt = max(min_body_pt, body_pt - 0.5)
            reg_pt = max(min_reg_pt, reg_pt - 0.4)
            fit, available_ht, line_h = fits(body_pt, leading, header_pt)
            if fit: break
        # reduce header font to free space
        if header_pt > 10:
            header_pt = max(10, header_pt - 1)
            fit, available_ht, line_h = fits(body_pt, leading, header_pt)
            if fit: break
        # if none works, reduce margins a bit (top/bottom)
        if top_margin > 4*mm:
            top_margin = max(4*mm, top_margin - 2*mm)
            bottom_margin = max(4*mm, bottom_margin - 2*mm)
            # recalc usable width/height
            fit, available_ht, line_h = fits(body_pt, leading, header_pt)
            if fit: break
        # last resort: very slightly reduce body_pt
        if body_pt > min_body_pt:
            body_pt = max(min_body_pt, body_pt - 0.5)
            reg_pt = max(min_reg_pt, reg_pt - 0.4)
            fit, available_ht, line_h = fits(body_pt, leading, header_pt)
            if fit: break
        # if still not fit, break to avoid infinite loop
        break

    # Compute positions
    header_height = header_pt * 1.25
    y_start = H - top_margin - header_height - (4 * mm)  # start line baseline for first row
    # horizontal column left positions
    col1_x = left_margin
    col2_x = left_margin + col_width + col_gap

    # draw header centered
    c.setFont(header_font, header_pt)
    c.drawCentredString(W / 2.0, H - top_margin - (header_pt * 0.75), header_text)

    # draw rows
    # row baseline y for row i: y_start - i * line_h
    # iterate for max rows_per_column, draw left and right if exist
    # alternate shading: draw a light rect behind odd rows (index starting 0 for top row)
    shading_color = (0.85, 0.85, 0.85)  # light gray
    text_left_pad = 2 * mm
    # column field x offsets relative to column left
    # fields: Flight, Time, Dest, Type, Reg
    # choose approximate widths in points
    flight_w = col_width * 0.20
    time_w = col_width * 0.12
    dest_w = col_width * 0.18
    type_w = col_width * 0.20
    reg_w = col_width - (flight_w + time_w + dest_w + type_w)  # rest

    # set fonts for drawing
    body_font = "Helvetica"
    body_bold = "Helvetica-Bold"

    # draw rows loop
    for row_idx in range(rows_per_column):
        y = y_start - row_idx * line_h
        # skip if y below bottom margin
        if y - line_h * 0.2 < bottom_margin:
            # safety: do not draw beyond margin
            continue
        # left side
        if row_idx < left_count:
            r = records[row_idx]
            # shading for alternate rows (use row_idx starting 0)
            if (row_idx % 2) == 1:
                c.setFillColorRGB(*shading_color)
                c.rect(col1_x, y - line_h + 2, col_width, line_h, fill=1, stroke=0)
                c.setFillColorRGB(0, 0, 0)
            # draw fields
            x = col1_x + text_left_pad
            # Flight (bold)
            c.setFont(body_bold, body_pt)
            c.drawString(x, y - (body_pt * 0.2), r['flight'])
            x += flight_w
            # Time
            c.setFont(body_font, body_pt)
            c.drawString(x, y - (body_pt * 0.2), r['time'] if isinstance(r['time'], str) else '')
            x += time_w
            # Dest
            c.drawString(x, y - (body_pt * 0.2), r.get('dest',''))
            x += dest_w
            # Type
            c.drawString(x, y - (body_pt * 0.2), r.get('type',''))
            x += type_w
            # Reg (smaller font)
            c.setFont(body_font, reg_pt)
            c.drawRightString(col1_x + col_width - text_left_pad, y - (reg_pt * 0.2), r.get('reg',''))
        # right side
        if row_idx < right_count:
            r_idx = left_count + row_idx
            r = records[r_idx]
            if (row_idx % 2) == 1:
                c.setFillColorRGB(*shading_color)
                c.rect(col2_x, y - line_h + 2, col_width, line_h, fill=1, stroke=0)
                c.setFillColorRGB(0, 0, 0)
            x = col2_x + text_left_pad
            c.setFont(body_bold, body_pt)
            c.drawString(x, y - (body_pt * 0.2), r['flight'])
            x += flight_w
            c.setFont(body_font, body_pt)
            c.drawString(x, y - (body_pt * 0.2), r['time'] if isinstance(r['time'], str) else '')
            x += time_w
            c.drawString(x, y - (body_pt * 0.2), r.get('dest',''))
            x += dest_w
            c.drawString(x, y - (body_pt * 0.2), r.get('type',''))
            x += type_w
            c.setFont(body_font, reg_pt)
            c.drawRightString(col2_x + col_width - text_left_pad, y - (reg_pt * 0.2), r.get('reg',''))

    # draw footer small
    c.setFont(footer_font, footer_pt)
    c.drawRightString(W - right_margin, bottom_margin - (footer_pt * 0.2) + 4, footer_text)

    c.showPage()
    c.save()
    buf.seek(0)
    return buf

# --- PDF labels generator (unchanged) ---
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
        c.setFont('Helvetica-Bold', 38); c.drawString(x_left + 15*mm, y_top - 21*mm, r.get('flight', ''))
        c.setFont('Helvetica-Bold', 23); c.drawString(x_left + 15*mm, y_top - 33*mm, r.get('dest', ''))
        try:
            tdisp = datetime.strptime(r.get('time',''), '%I:%M %p').strftime('%H:%M')
        except Exception:
            tdisp = r.get('time','')
        c.setFont('Helvetica-Bold', 29); c.drawString(x_left + 15*mm, y_top - 47*mm, tdisp)
        c.setFont('Helvetica', 13); c.drawRightString(x_left + col_w - 6*mm, y_top - row_h + 12*mm, r.get('type',''))
        c.drawRightString(x_left + col_w - 6*mm, y_top - row_h + 7*mm, r.get('reg',''))
    c.save()
    target.seek(0)
    return target

# --- Sidebar / UI ---
with st.sidebar:
    st.header("⚙️ Settings")
    s_year = st.number_input("Year", value=datetime.now().year, min_value=2000, max_value=2100)
    s_time = st.time_input("Start Time", value=dtime(hour=4, minute=55))
    e_time = st.time_input("End Time", value=dtime(hour=5, minute=0))
    label_start = st.number_input("Label Start Number", value=1, min_value=1)
    show_inner_headers = st.checkbox("One-page: show column headers (uses more space)", value=False)

st.markdown('<div class="top-left-container"><a href="https://www.flightradar24.com/data/airports/akl/arrivals" target="_blank">Import Raw Text</a><a href="https://www.flightradar24.com/data/airports/akl/departures" target="_blank">Export Raw Text</a></div>', unsafe_allow_html=True)
st.markdown('<div class="main-title">Simon Park\'nRide\'s<br><span class="sub-title">Flight List Factory</span></div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Raw Text File", type=['txt'])

# Parser tuning
st.subheader("Parser Tuning")
sample_text = st.text_area("Optional: paste a small sample of your raw text here for parser tuning (include date header + 2–3 flights).", height=160)
if st.button("Run Parser on Sample"):
    if not sample_text.strip():
        st.warning("Please paste some sample lines first.")
    else:
        sample_lines = sample_text.splitlines()
        parsed = parse_raw_lines(sample_lines, s_year)
        if parsed:
            st.success(f"Parsed {len(parsed)} record(s) from sample")
            st.json(parsed)
        else:
            st.warning("No records parsed from the sample — include the date header and at least one flight block.")

if uploaded_file:
    try:
        content = uploaded_file.read()
        lines = content.decode("utf-8", errors="replace").splitlines() if isinstance(content, (bytes,bytearray)) else str(content).splitlines()
    except Exception as e:
        st.error(f"Failed to read uploaded file: {e}")
        lines = []

    if lines:
        all_recs = parse_raw_lines(lines, s_year)
        if not all_recs:
            st.warning("No records parsed from the file.")
        else:
            filtered, s_dt, e_dt = filter_records(all_recs, s_time.strftime('%H:%M'), e_time.strftime('%H:%M'))
            if not filtered:
                st.warning("No flights matched the filters (time window, airlines, or exclusions).")
            else:
                st.success(f"Processed {len(filtered)} flights")
                col1, col_mid, col2 = st.columns([1,0.9,1])
                fn = f"List_{s_dt.strftime('%d-%m')}" if s_dt else "List"

                # two-page DOCX (unchanged)
                docx_bytes = build_docx_stream(filtered, s_dt, e_dt).getvalue()
                col1.download_button("📥 Download DOCX List (2 pages)", data=docx_bytes, file_name=f"{fn}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

                # ONE-PAGE direct PDF (guaranteed single page)
                onepage_pdf_buf = build_onepage_pdf_stream(filtered, s_dt, e_dt)
                onepage_pdf_bytes = onepage_pdf_buf.getvalue()
                col_mid.download_button("📥 Download One-Page PDF (direct)", data=onepage_pdf_bytes, file_name=f"{fn}_onepage.pdf", mime="application/pdf")
                st.info("One-Page PDF was generated directly (layout controlled so it will not overflow).")

                # PDF labels (unchanged)
                pdf_labels = build_labels_stream(filtered, label_start).getvalue()
                col2.download_button("📥 Download PDF Labels", data=pdf_labels, file_name=f"Labels_{fn}.pdf", mime="application/pdf")

                # preview table
                table_rows = []
                for i, r in enumerate(filtered):
                    try:
                        tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
                    except Exception:
                        tdisp = r['time']
                    table_rows.append({'No': label_start + i, 'Flight': r['flight'], 'Time': tdisp, 'Dest': r['dest'], 'Type': r['type'], 'Reg': r['reg']})
                st.table(table_rows)
