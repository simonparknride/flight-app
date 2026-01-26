import streamlit as st
import re
import io
from datetime import datetime, timedelta
from typing import List, Dict
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# --- 1. UI 설정 (버튼 스타일 및 웹 폰트 크기 복구) ---
st.set_page_config(page_title="Flight List Factory", layout="centered", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .stApp { background-color: #000000; }
    [data-testid="stSidebar"] { background-color: #111111 !important; }
    .stMarkdown, p, h1, h2, h3, label { color: #ffffff !important; }
    
    /* 상단 링크 및 타이틀 크기 복구 */
    .top-left-container { text-align: left; padding-top: 10px; margin-bottom: 20px; }
    .top-left-container a { font-size: 1.1rem !important; color: #ffffff !important; text-decoration: underline; display: block; margin-bottom: 5px; }
    .main-title { font-size: 3.2rem !important; font-weight: 800; color: #ffffff; line-height: 1.1; margin-bottom: 0.5rem; }
    .sub-title { font-size: 2.6rem !important; font-weight: 400; color: #60a5fa; }

    /* [복구] 다운로드 버튼 스타일: 검은 배경에 흰색 글자 */
    div.stDownloadButton > button {
        background-color: #000000 !important; 
        color: #ffffff !important;           
        border: 1px solid #ffffff !important;
        border-radius: 4px !important;
        padding: 0.5rem 1rem !important;
        width: 100% !important;
    }
    div.stDownloadButton > button:hover {
        background-color: #ffffff !important; 
        color: #000000 !important;           
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 파싱 로직 ---
TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s[AP]M)\t([A-Z]{2}\d+[A-Z]?)\s*$")
DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
IATA_IN_PAREns = re.compile(r"\(([^)]+)\)")
PLANE_TYPES = ['A21N','A20N','A320','32Q','320','73H','737','74Y','77W','B77W','789','B789','359','A359','332','A332','AT76','DH8C','DH3','AT7','388','333','A333','330','76V','77L','B38M','A388','772','B772','32X','77X']
PLANE_TYPE_PATTERN = re.compile(r"\b(" + "|".join(sorted(set(PLANE_TYPES), key=len, reverse=True)) + r")\b", re.IGNORECASE)
NORMALIZE_MAP = {'32q': 'A320', '320': 'A320', 'a320': 'A320', '32x': 'A320', '789': 'B789', 'b789': 'B789', '772': 'B772', 'b772': 'B772', '77w': 'B77W', 'b77w': 'B77W', '332': 'A332', 'a332': 'A332', '333': 'A333', 'a333': 'A333', '330': 'A330', 'a330': 'A330', '359': 'A359', 'a359': 'A359', '388': 'A388', 'a388': 'A388', '737': 'B737', '73h': 'B737', 'at7': 'AT76'}
ALLOWED_AIRLINES = {"NZ","QF","JQ","CZ","CA","SQ","LA","IE"}
NZ_DOMESTIC_IATA = {"AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO","WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"}

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
            plane_type = NORMALIZE_MAP.get(mtype.group(1).lower() if mtype else '', mtype.group(1).upper() if mtype else '')
            reg = ''
            parens = IATA_IN_PAREns.findall(carrier_line)
            if parens: reg = parens[-1].strip()
            try: dep_dt = datetime.strptime(f"{current_date} {time_str}", '%Y-%m-%d %I:%M %p')
            except: dep_dt = None
            records.append({'dt': dep_dt, 'time': time_str, 'flight': flight, 'dest': dest_iata, 'type': plane_type, 'reg': reg})
            i += 4; continue
        i += 1
    return records

def filter_records(records, start_hm, end_hm):
    dates = sorted({r['dt'].date() for r in records if r.get('dt')})
    if not dates: return [], None, None
    day1, day2 = dates[0], dates[1] if len(dates) >= 2 else (dates[0] + timedelta(days=1))
    start_dt = datetime.combine(day1, datetime.strptime(start_hm, '%H:%M').time())
    end_dt = datetime.combine(day2, datetime.strptime(end_hm, '%H:%M').time())
    out = [r for r in records if r.get('dt') and r['flight'][:2] in ALLOWED_AIRLINES and r['dest'] not in NZ_DOMESTIC_IATA and (start_dt <= r['dt'] <= end_dt)]
    out.sort(key=lambda x: x['dt'])
    return out, start_dt, end_dt

# --- 3. DOCX 생성 (One Page 최적화) ---
def build_docx_stream(records, start_dt, end_dt, mode='Two Pages'):
    doc = Document()
    font_name = 'Air New Zealand Sans'
    section = doc.sections[0]
    section.left_margin = section.right_margin = Inches(0.5)

    if mode == 'One Page':
        section.top_margin = section.bottom_margin = Inches(0.1)
        font_size = Pt(7.5)      # [수정] 7.5pt 적용
        table_width = '4200'
        header_size = Pt(11)
        header_align = WD_ALIGN_PARAGRAPH.LEFT # [수정] 날짜 왼쪽 정렬
    else:
        section.top_margin = section.bottom_margin = Inches(0.3)
        font_size = Pt(14)       # [보존] Two Pages 기존 설정 유지
        table_width = '4000'
        header_size = Pt(16)
        header_align = WD_ALIGN_PARAGRAPH.CENTER

    # 헤더
    p = doc.add_paragraph()
    p.alignment = header_align
    run_head = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}")
    run_head.bold = True
    run_head.font.name = font_name
    run_head.font.size = header_size

    # 테이블
    table = doc.add_table(rows=0, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER # [수정] 표는 중앙 정렬
    tblPr = table._element.find(qn('w:tblPr'))
    tblW = OxmlElement('w:tblW'); tblW.set(qn('w:w'), table_width); tblW.set(qn('w:type'), 'pct'); tblPr.append(tblW)

    for i, r in enumerate(records):
        row = table.add_row()
        if mode == 'One Page': row.height = Inches(0.1)
        tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        vals = [r['flight'], tdisp, r['dest'], r['type'], r['reg']]
        for j, val in enumerate(vals):
            cell = row.cells[j]
            if i % 2 == 1:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9'); tcPr.append(shd)
            para = cell.paragraphs[0]
            para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            run = para.add_run(str(val))
            run.font.name = font_name
            run.font.size = font_size
    
    target = io.BytesIO()
    doc.save(target); target.seek(0)
    return target

# --- 4. 앱 실행 ---
with st.sidebar:
    st.header("⚙️ Settings")
    s_time = st.text_input("Start Time", value="05:00")
    e_time = st.text_input("End Time", value="04:55")
    label_start = st.number_input("Label Start Number", value=1, min_value=1)

st.markdown('<div class="top-left-container"><a href="...">Import Raw Text</a><a href="...">Export Raw Text</a></div>', unsafe_allow_html=True)
st.markdown('<div class="main-title">Simon Park\'nRide\'s<br><span class="sub-title">Flight List Factory</span></div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Raw Text File", type=['txt'])

if uploaded_file:
    lines = uploaded_file.read().decode("utf-8").splitlines()
    all_recs = parse_raw_lines(lines)
    if all_recs:
        filtered, s_dt, e_dt = filter_records(all_recs, s_time, e_time)
        if filtered:
            st.success(f"Processed {len(filtered)} flights (2026 Updated)")
            col1, col2 = st.columns(2)
            fn = f"List_{s_dt.strftime('%d-%m')}"
            col1.download_button("📥 One Page DOCX", build_docx_stream(filtered, s_dt, e_dt, mode='One Page'), f"{fn}_1P.docx")
            col2.download_button("📥 Two Pages DOCX", build_docx_stream(filtered, s_dt, e_dt, mode='Two Pages'), f"{fn}_2P.docx")
