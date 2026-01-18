import streamlit as st
import re
import io
from datetime import datetime, timedelta
from typing import List, Dict

# --- Core Logic & Patterns ---
TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s[AP]M)\t([A-Z]{2}\d+[A-Z]?)\s*$")
DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
IATA_IN_PAREns = re.compile(r"\(([^)]+)\)")
PLANE_TYPES = ['A21N','A20N','A320','32Q','320','73H','737','74Y','77W','B77W','789','B789','359','A359','332','A332','AT76','DH8C','DH3','AT7','388','333','A333','330','76V','77L','B38M','A388','772','B772','32X','77X']
PLANE_TYPE_PATTERN = re.compile(r"\b(" + "|".join(sorted(set(PLANE_TYPES), key=len, reverse=True)) + r")\b", re.IGNORECASE)
NORMALIZE_MAP = {'32q': 'A320', '320': 'A320', 'a320': 'A320', '32x': 'A320', '789': 'B789', 'b789': 'B789', '772': 'B772', 'b772': 'B772', '77w': 'B77W', 'b77w': 'B77W', '332': 'A332', 'a332': 'A332', '333': 'A333', 'a333': 'A333', '330': 'A330', 'a330': 'A330', '359': 'A359', 'a359': 'A359', '388': 'A388', 'a388': 'A388', '737': 'B737', '73h': 'B737', 'at7': 'AT76'}
ALLOWED_AIRLINES = {"NZ","QF","JQ","CZ","CA","SQ","LA","IE"}
NZ_DOMESTIC_IATA = {"AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO","WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"}
REGO_LIKE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9\-–—]*$")

def normalize_type(t: str) -> str:
    key = (t or '').strip().lower()
    return NORMALIZE_MAP.get(key, t)

def parse_raw_lines(lines: List[str]) -> List[Dict]:
    records = []
    current_date = None
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if DATE_HEADER.match(line):
            try: current_date = datetime.strptime(line + ' 2026', '%A, %b %d %Y').date()
            except: current_date = None
            i += 1; continue
        m = TIME_LINE.match(line)
        if m and current_date is not None:
            time_str, flight = m.groups()
            dest_line = lines[i+1].strip() if i+1 < len(lines) else ''
            m2 = IATA_IN_PAREns.search(dest_line)
            dest_iata = (m2.group(1).strip() if m2 else '').upper()
            carrier_line = lines[i+2].rstrip('\n') if i+2 < len(lines) else ''
            mtype = PLANE_TYPE_PATTERN.search(carrier_line)
            plane_type = normalize_type(mtype.group(1).upper() if mtype else '')
            reg = ''
            parens = IATA_IN_PAREns.findall(carrier_line)
            if parens:
                for candidate in reversed(parens):
                    cand = candidate.strip()
                    if REGO_LIKE.match(cand) and '-' in cand:
                        reg = cand; break
                if not reg: reg = parens[-1].strip()
            try: dep_dt = datetime.strptime(f"{current_date} {time_str}", '%Y-%m-%d %I:%M %p')
            except: dep_dt = None
            records.append({'dt': dep_dt, 'time': time_str, 'flight': flight, 'dest': dest_iata, 'type': plane_type, 'reg': reg})
            i += 4; continue
        i += 1
    return records

def filter_records(records: List[Dict], start_hm: str, end_hm: str):
    dates = sorted({r['dt'].date() for r in records if r.get('dt')})
    if not dates: return [], None, None
    day1 = dates[0]
    day2 = dates[1] if len(dates) >= 2 else (day1 + timedelta(days=1))
    start_dt = datetime.combine(day1, datetime.strptime(start_hm, '%H:%M').time())
    end_dt = datetime.combine(day2, datetime.strptime(end_hm, '%H:%M').time())
    out = [r for r in records if r.get('dt') and r['flight'][:2] in ALLOWED_AIRLINES and r['dest'] not in NZ_DOMESTIC_IATA and (start_dt <= r['dt'] <= end_dt)]
    out.sort(key=lambda x: x['dt'])
    return out, start_dt, end_dt

# --- DOCX Generation (Footer 추가됨) ---
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn

def build_docx_stream(records, start_dt, end_dt, reg_placeholder):
    doc = Document()
    section = doc.sections[0]
    section.top_margin, section.bottom_margin = Inches(0.3), Inches(0.5) # 푸터 공간 확보
    section.left_margin, section.right_margin = Inches(0.5), Inches(0.5)
    
    # 푸터 추가
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_f = footer_para.add_run("created by Simon Park'nRide")
    run_f.font.size = Pt(10)
    run_f.font.color.rgb = RGBColor(0x80, 0x80, 0x80) # 회색

    heading_text = f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}"
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(heading_text); run.bold = True; run.font.size = Pt(16)
    
    row_count = len(records)
    fs, ls = (12.0, 0.65) if row_count > 60 else (13.5, 0.7) if row_count > 45 else (15.0, 0.8)
    
    table = doc.add_table(rows=0, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = table._element.find(qn('w:tblPr'))
    tblW = OxmlElement('w:tblW'); tblW.set(qn('w:w'), '5000'); tblW.set(qn('w:type'), 'pct'); tblPr.append(tblW)

    for i, r in enumerate(records):
        row = table.add_row()
        tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        vals = [r['flight'], tdisp, r['dest'], r['type'], r['reg'] or reg_placeholder]
        for j, val in enumerate(vals):
            cell = row.cells[j]
            if i % 2 == 1:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9'); tcPr.append(shd)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT if j < 4 else WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.line_spacing = ls
            run = para.add_run(val); run.font.size = Pt(fs)
    target = io.BytesIO()
    doc.save(target); target.seek(0)
    return target

# --- PDF Label Generation (중앙 정렬 유지) ---
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

def build_labels_stream(records, start_dt, end_dt, start_num, reg_placeholder):
    target = io.BytesIO()
    c = canvas.Canvas(target, pagesize=A4)
    w, h = A4
    margin, gutter = 15*mm, 6*mm
    col_w = (w - 2*margin - gutter) / 2
    row_h = (h - 2*margin) / 5
    for i, r in enumerate(records):
        if i > 0 and i % 10 == 0: c.showPage()
        idx = i % 10
        row_idx, col_idx = idx // 2, idx % 2
        x_left = margin + col_idx * (col_w + gutter)
        y_top = h - margin - row_idx * row_h
        c.setStrokeGray(0.3); c.setLineWidth(0.2)
        c.rect(x_left, y_top - row_h + 2*mm, col_w, row_h - 4*mm)
        
        # 1. 왼쪽 상단: 정사각형 번호 박스
        c.setLineWidth(0.5)
        c.rect(x_left + 3*mm, y_top - 12*mm, 8*mm, 8*mm)
        c.setFont('Helvetica-Bold', 14)
        c.drawCentredString(x_left + 7*mm, y_top - 9.5*mm, str(start_num + i))
        
        # 2. 오른쪽 상단: 날짜 (13mm 하향 유지)
        c.setFont('Helvetica-Bold', 18)
        c.drawRightString(x_left + col_w - 4*mm, y_top - 13*mm, r['dt'].strftime('%d %b'))
        
        # 3. 중앙 정보 (정중앙 정렬 유지)
        content_x = x_left + 15*mm
        c.setFont('Helvetica-Bold', 29)
        c.drawString(content_x, y_top - 21*mm, r['flight'])
        c.setFont('Helvetica-Bold', 23)
        c.drawString(content_x, y_top - 33*mm, r['dest'])
        tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        c.setFont('Helvetica-Bold', 29)
        c.drawString(content_x, y_top - 47*mm, tdisp)
        
        # 4. 오른쪽 하단: Plane Type & Reg
        c.setFont('Helvetica', 13)
        c.drawRightString(x_left + col_w - 6*mm, y_top - row_h + 12*mm, r['type'])
        c.drawRightString(x_left + col_w - 6*mm, y_top - row_h + 7*mm, r['reg'] or reg_placeholder)
        
        # 5. 라벨 최하단 문구 (필요시 라벨에도 추가 가능)
        # c.setFont('Helvetica', 7); c.setFillGray(0.5)
        # c.drawRightString(x_left + col_w - 2*mm, y_top - row_h + 3*mm, "created by Simon Park'nRide")
        # c.setFillGray(0)
        
    c.save(); target.seek(0)
    return target

# --- Streamlit UI ---
st.set_page_config(page_title="Easy Flight List", layout="centered")
st.markdown("""
    <style>
    .stApp { background-color: #000000; }
    .header-link { font-size: 1.75rem; color: #ffffff !important; text-decoration: underline; font-weight: 300; display: block; margin-bottom: 0px; }
    .header-link:hover { color: #60a5fa !important; }
    .main-title { font-size: 3rem; font-weight: 800; color: #ffffff; line-height: 1.1; margin-top: 5px; margin-bottom: 0.5rem; }
    .sub-title { font-size: 2.5rem; font-weight: 400; color: #60a5fa; }
    [data-testid="stSidebar"] { background-color: #111111; }
    .stMarkdown, p, h1, h2, h3, label { color: #ffffff !important; }
    </style>
    """, unsafe_allow_html=True)
st.markdown('<a href="https://www.flightradar24.com/data/airports/akl/arrivals" target="_blank" class="header-link">get a Raw Text File here</a>', unsafe_allow_html=True)
st.markdown('<div class="main-title">Simon Park\'nRide\'s<br><span class="sub-title">Easy Flight List</span></div>', unsafe_allow_html=True)
uploaded_file = st.file_uploader("Upload Raw Text File", type=['txt'])
with st.sidebar:
    st.header("Settings")
    s_time = st.text_input("Start Time (Day 1)", value="05:00")
    e_time = st.text_input("End Time (Day 2)", value="04:55")
    reg_p = ""
    label_start = st.number_input("Label Start Number", value=1)
if uploaded_file:
    content = uploaded_file.read().decode("utf-8").splitlines()
    records_all = parse_raw_lines(content)
    if records_all:
        filtered, s_dt, e_dt = filter_records(records_all, s_time, e_time)
        if filtered:
            st.success(f"Successfully processed {len(filtered)} flights.")
            col1, col2 = st.columns(2)
            fn_date = f"{s_dt.strftime('%d')}-{e_dt.strftime('%d')}_{s_dt.strftime('%b')}"
            docx_data = build_docx_stream(filtered, s_dt, e_dt, reg_p)
            col1.download_button("📥 Download DOCX List", docx_data, f"Flight_List_{fn_date}.docx")
            pdf_data = build_labels_stream(filtered, s_dt, e_dt, label_start, reg_p)
            col2.download_button("📥 Download PDF Labels", pdf_data, f"Labels_{fn_date}.pdf")
            st.write("### Preview")
            st.table([{'No': label_start+i, 'Flight': r['flight'], 'Time': r['time'], 'Dest': r['dest'], 'Reg': r['reg']} for i, r in enumerate(filtered)])
