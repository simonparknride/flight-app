import streamlit as st
import re
import io
from datetime import datetime, timedelta
from typing import List, Dict

# --- UI 및 디자인 설정 ---
st.set_page_config(page_title="Flight List Factory", layout="centered")

st.markdown("""
    <style>
    /* 전체 배경색 검정 */
    .stApp {
        background-color: #000000;
    }

    /* 왼쪽 상단 링크 컨테이너 */
    .top-left-container { 
        text-align: left; 
        padding-top: 10px;
        margin-bottom: 20px;
    }
    .top-left-container a { 
        font-size: 1.1rem; 
        color: #ffffff !important; 
        text-decoration: underline; 
        font-weight: 300;
        display: block;
        margin-bottom: 5px;
    }
    .top-left-container a:hover { color: #60a5fa !important; }
    
    /* 메인 타이틀 스타일 */
    .main-title { font-size: 3rem; font-weight: 800; color: #ffffff; line-height: 1.1; margin-top: 5px; margin-bottom: 0.5rem; }
    .sub-title { font-size: 2.5rem; font-weight: 400; color: #60a5fa; }
    
    /* 사이드바 및 텍스트 색상 설정 */
    [data-testid="stSidebar"] { background-color: #111111; }
    .stMarkdown, p, h1, h2, h3, label { color: #ffffff !important; }
    
    /* 가독성을 위한 표 배경 처리 */
    .stTable {
        background-color: rgba(255, 255, 255, 0.05);
        border-radius: 10px;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 상단 링크 영역 ---
st.markdown("""
    <div class="top-left-container">
        <a href="https://www.flightradar24.com/data/airports/akl/arrivals" target="_blank">Import Raw Text File</a>
        <a href="https://www.flightradar24.com/data/airports/akl/departures" target="_blank">Export Raw TExt File</a>
    </div>
    """, unsafe_allow_html=True)

# --- 메인 타이틀 (홍보 문구 삭제됨) ---
st.markdown('<div class="main-title">Simon Park\'nRide\'s<br><span class="sub-title">Flight List Factory</span></div>', unsafe_allow_html=True)

# --- 핵심 로직 및 패턴 설정 ---
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

# --- DOCX 생성 (Width 80%) ---
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn

def build_docx_stream(records, start_dt, end_dt, reg_placeholder):
    doc = Document()
    section = doc.sections[0]
    section.top_margin, section.bottom_margin = Inches(0.3), Inches(0.5)
    section.left_margin, section.right_margin = Inches(0.5), Inches(0.5)
    
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_f = footer_para.add_run("created by Simon Park'nRide")
    run_f.font.size = Pt(10)
    run_f.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    heading_text = f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}"
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(heading_text); run.bold = True; run.font.size = Pt(16)
    
    table = doc.add_table(rows=0, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = table._element.find(qn('w:tblPr'))
    tblW = OxmlElement('w:tblW'); tblW.set(qn('w:w'), '4000'); tblW.set(qn('w:type'), 'pct'); tblPr.append(tblW)

    for i, r in enumerate(records):
        row = table
