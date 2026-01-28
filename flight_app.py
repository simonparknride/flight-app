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

# --- 1. 페이지 설정 및 디자인 ---
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
    div.stDownloadButton > button:hover { background-color: #60a5fa !important; color: #ffffff !important; }
    .top-left-container { text-align: left; padding-top: 10px; margin-bottom: 20px; }
    .top-left-container a { font-size: 1.1rem; color: #ffffff !important; text-decoration: underline; display: block; margin-bottom: 5px;}
    .main-title { font-size: 3rem; font-weight: 800; color: #ffffff; line-height: 1.1; margin-bottom: 0.5rem; }
    .sub-title { font-size: 2.5rem; font-weight: 400; color: #60a5fa; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 데이터 파싱 로직 ---
TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s*[AP]M)\s+([A-Z0-9]{2,4}\d*[A-Z]?)\s*$", re.IGNORECASE)
DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
IATA_IN_PARENS = re.compile(r"\(([^)]+)\)")
PLANE_TYPES = ['A21N','A20N','A320','32Q','320','73H','737','74Y','77W','B77W','789','B789','359','A359','332','A332','AT76','DH8C','DH3','AT7','388','333','A333','330','76V','77L','B38M','A388','772','B772','32X','77X']
PLANE_TYPE_PATTERN = re.compile(r"\b(" + "|".join(sorted(set(PLANE_TYPES), key=len, reverse=True)) + r")\b", re.IGNORECASE)

NORMALIZE_MAP = {'32q': 'A320', '320': 'A320', 'a320': 'A320', '32x': 'A320', '789': 'B789', 'b789': 'B789', '772': 'B772', 'b772': 'B772', '77w': 'B77W', 'b77w': 'B77W', '332': 'A332', 'a332': 'A332', '333': 'A333', 'a333': 'A333', '330': 'A330', 'a330': 'A330', '359': 'A359', 'a359': 'A359', '388': 'A388', 'a388': 'A388', '737': 'B737', '73h': 'B737', 'at7': 'AT76'}
ALLOWED_AIRLINES = {"NZ", "QF", "JQ", "CZ", "CA", "SQ", "LA", "IE", "FX"}
NZ_DOMESTIC_IATA = {"AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO","WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"}
REGO_LIKE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9\-–—]*$")

def normalize_type(t: Optional[str]) -> str:
    if not t: return ""
    return NORMALIZE_MAP.get(t.strip().lower(), t.strip().upper())

def try_parse_date_header(line: str, year: int) -> Optional[datetime.date]:
    for fmt in ["%A, %b %d %Y", "%A, %B %d %Y", "%a, %b %d %Y", "%a, %B %d %Y"]:
        try: return datetime.strptime(line.strip() + f" {year}", fmt).date()
        except: continue
    return None

def parse_raw_lines(lines: List[str], year: int) -> List[Dict]:
    records = []
    current_date = None
    i, L = 0, len(lines)
    while i < L:
        line = lines[i].strip()
        if DATE_HEADER.match(line):
            current_date = try_parse_date_header(line, year)
            i += 1
            continue
        m = TIME_LINE.match(line)
        if m and current_date:
            time_raw, flight_raw = m.groups()
            dest_line = lines[i+1].strip() if i+1 < L else ''
            carrier_line = lines[i+2].rstrip('\n') if i+2 < L else ''
            m2 = IATA_IN_PARENS.search(dest_line)
            dest_iata = (m2.group(1).strip().upper() if m2 else '').upper()
            mtype = PLANE_TYPE_PATTERN.search(carrier_line)
            plane_type = normalize_type(mtype.group(1) if mtype else '')
            reg = ''
            parens = IATA_IN_PARENS.findall(carrier_line)
            if parens:
                for cand in reversed(parens):
                    cand = cand.strip()
                    if REGO_LIKE.match(cand) and any(x in cand for x in ['-', '–', '—']):
                        reg = cand; break
                if not reg: reg = parens[-1].strip()
            dep_dt = None
            try:
                tnorm = time_raw.strip().upper().replace(" ", "")
                fmt = "%Y-%m-%d %I:%M%p" if len(tnorm) <= 7 else "%Y-%m-%d %I:%M %p"
                dep_dt = datetime.strptime(f"{current_date} {time_raw.strip()}", fmt)
            except: pass
            records.append({'dt': dep_dt, 'time': time_raw.strip(), 'flight': flight_raw.strip().upper(), 'dest': dest_iata, 'type': plane_type, 'reg': reg})
            i += 3
            continue
        i += 1
    return records

def filter_records(records: List[Dict], start_time: dtime, end_time: dtime):
    valid = [r for r in records if r['dt']]
    if not valid: return [], None, None
    day1 = valid[0]['dt'].date()
    day2 = (day1 + timedelta(days=1))
    start_dt, end_dt = datetime.combine(day1, start_time), datetime.combine(day2, end_time)
    def is_ok(r):
        return (r['flight'][:2] in ALLOWED_AIRLINES) and (r['dest'] not in NZ_DOMESTIC_IATA) and (start_dt <= r['dt'] <= end_dt)
    res = sorted([r for r in valid if is_ok(r)], key=lambda x: x['dt'])
    return res, start_dt, end_dt

# --- 3. 문서 생성 로직 ---

def build_docx_stream(records: List[Dict], start_dt: datetime, end_dt: datetime) -> io.BytesIO:
    """TWO-PAGE DOCX: 지시대로 절대 변경 금지 (Spacing 0pt 유지)"""
    doc = Document()
    font_name = 'Air New Zealand Sans'
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Inches(0.3)
    section.left_margin = section.right_margin = Inches(0.5)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_head = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}")
    run_head.bold = True
    run_head.font.size = Pt(16)
    
    table = doc.add_table(rows=0, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, r in enumerate(records):
        row = table.add_row()
        vals = [r['flight'], (r['dt'].strftime('%H:%M') if r['dt'] else r['time']), r['dest'], r['type'], r['reg']]
        for j, v in enumerate(vals):
            cell = row.cells[j]
            if i % 2 == 1:
                shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9'); cell._tc.get_or_add_tcPr().append(shd)
            para = cell.paragraphs[0]
            para.paragraph_format.space_before = para.paragraph_format.space_after = Pt(0)
            run = para.add_run(str(v)); run.font.name = font_name; run.font.size = Pt(14)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0); return buf

def build_docx_onepage_stream(records: List[Dict], start_dt: datetime, end_dt: datetime) -> io.BytesIO:
    """ONE-PAGE DOCX: 테이블 선택 후 Layout-Spacing에서 Before/After 2.2 설정 반영"""
    doc = Document()
    font_name = 'Air New Zealand Sans'
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Inches(0.2)
    section.left_margin = section.right_margin = Inches(0.4)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_head = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}")
    run_head.bold = True
    run_head.font.size = Pt(16)
    
    mid = (len(records) + 1) // 2
    outer = doc.add_table(rows=1, cols=2)
    outer.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    def fill_column(cell, recs, s_idx):
        inner = cell.add_table(rows=1, cols=5)
        for idx, txt in enumerate(['Flight', 'Time', 'Dest', 'Type', 'Reg']):
            h_run = inner.rows[0].cells[idx].paragraphs[0].add_run(txt)
            h_run.bold = True; h_run.font.size = Pt(11)
            
        for i, r in enumerate(recs):
            row = inner.add_row()
            vals = [r['flight'], (r['dt'].strftime('%H:%M') if r['dt'] else r['time']), r['dest'], r['type'], r['reg']]
            for j, v in enumerate(vals):
                c = row.cells[j]
                if (s_idx + i) % 2 == 1:
                    shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9'); c._tc.get_or_add_tcPr().append(shd)
                
                para = c.paragraphs[0]
                # --- 지시하신 핵심 코드: Before/After 2.2pt ---
                para.paragraph_format.space_before = Pt(2.2)
                para.paragraph_format.space_after = Pt(2.2)
                
                r_run = para.add_run(str(v))
                r_run.font.name = font_name
                r_run.font.size = Pt(9) if j == 4 else Pt(11)

    fill_column(outer.cell(0, 0), records[:mid], 0)
    fill_column(outer.cell(0, 1), records[mid:], mid)
    
    buf = io.BytesIO(); doc.save(buf); buf.seek(0); return buf

def build_labels_stream(records: List[Dict], start_num: int) -> io.BytesIO:
    buf = io.BytesIO(); c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4; margin, gut = 15*mm, 6*mm; cw, rh = (w - 2*margin - gut)/2, (h - 2*margin)/5
    for i, r in enumerate(records):
        if i > 0 and i % 10 == 0: c.showPage()
        idx = i % 10; x, y = margin + (idx % 2)*(cw + gut), h - margin - (idx // 2)*rh
        c.setStrokeGray(0.3); c.rect(x, y - rh + 2*mm, cw, rh - 4*mm)
        c.setFont('Helvetica-Bold', 14); c.drawCentredString(x + 7*mm, y - 9.5*mm, str(start_num + i))
        c.setFont('Helvetica-Bold', 38); c.drawString(x + 15*mm, y - 21*mm, r['flight'])
        c.setFont('Helvetica-Bold', 29); c.drawString(x + 15*mm, y - 47*mm, r['dt'].strftime('%H:%M') if r['dt'] else r['time'])
    c.save(); buf.seek(0); return buf

# --- 4. 메인 인터페이스 ---
with st.sidebar:
    st.header("⚙️ Settings")
    year = st.number_input("Year", value=datetime.now().year)
    s_time = st.time_input("Start Time", value=dtime(5, 0))
    e_time = st.time_input("End Time", value=dtime(4, 55))
    label_start = st.number_input("Label Start Number", value=1)

st.markdown('<div class="top-left-container"><a href="https://www.flightradar24.com/data/airports/akl/arrivals" target="_blank">Import Raw Text</a><a href="https://www.flightradar24.com/data/airports/akl/departures" target="_blank">Export Raw Text</a></div>', unsafe_allow_html=True)
st.markdown('<div class="main-title">Air New Zealand Cargo<br><span class="sub-title">Flight List</span></div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Raw Text File", type=['txt'])

if uploaded_file is not None:
    try:
        content = uploaded_file.read().decode("utf-8", errors="replace")
        all_recs = parse_raw_lines(content.splitlines(), year)
        
        if all_recs:
            filtered, s_dt, e_dt = filter_records(all_recs, s_time, e_time)
            if filtered:
                st.success(f"{len(filtered)} Flights Found")
                c1, cm, c2 = st.columns(3)
                fn = f"List_{s_dt.strftime('%d-%m')}"
                
                c1.download_button("📥 2-Page DOCX", build_docx_stream(filtered, s_dt, e_dt).getvalue(), f"{fn}.docx")
                cm.download_button("📥 1-Page DOCX", build_docx_onepage_stream(filtered, s_dt, e_dt).getvalue(), f"{fn}_onepage.docx")
                c2.download_button("📥 PDF Labels", build_labels_stream(filtered, label_start).getvalue(), f"Labels_{fn}.pdf")
                
                preview = []
                for i, r in enumerate(filtered):
                    preview.append({'No': label_start+i, 'Flight': r['flight'], 'Time': (r['dt'].strftime('%H:%M') if r['dt'] else r['time']), 'Dest': r['dest'], 'Reg': r['reg']})
                st.table(preview)
            else:
                st.warning("데이터는 있지만 필터(시간)에 걸리는 정보가 없습니다. 사이드바의 시간을 조절해 보세요.")
        else:
            st.error("파일에서 비행 정보를 파싱하지 못했습니다. 형식을 확인해 주세요.")
    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")
