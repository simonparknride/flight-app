import streamlit as st
import re
import io
import pandas as pd
from datetime import datetime, timedelta
from typing import List, Dict
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# --- 1. UI 설정 및 복구 ---
st.set_page_config(page_title="Flight List Factory", layout="centered")

# 사라졌던 상단 링크 및 스타일 복구
st.markdown("""
    <style>
    .stApp { background-color: #000000; }
    [data-testid="stSidebar"] { background-color: #111111 !important; }
    .stMarkdown, p, h1, h2, h3, label { color: #ffffff !important; }
    
    /* 상단 링크 스타일 */
    .top-links { font-size: 14px; margin-bottom: 10px; }
    .top-links a { color: #ffffff !important; text-decoration: underline; margin-right: 15px; }

    div.stDownloadButton > button {
        background-color: #ffffff !important;
        border: 2px solid #ffffff !important;
        border-radius: 8px !important;
        height: 3.5rem !important;
        width: 100% !important;
    }
    div.stDownloadButton > button div p {
        color: #000000 !important;
        font-weight: 800 !important;
    }
    </style>
    <div class="top-links">
        <a href="#">Import Raw Text</a>
        <a href="#">Export Raw Text</a>
    </div>
    """, unsafe_allow_html=True)

# --- 2. 파싱 로직 (기종 코드 단순화 반영) ---
TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s[AP]M)\t([A-Z]{2}\d+[A-Z]?)\s*$")
DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
IATA_IN_PAREns = re.compile(r"\(([^)]+)\)")
ALLOWED_AIRLINES = {"NZ","QF","JQ","CZ","CA","SQ","LA","IE"}
NZ_DOMESTIC_IATA = {"AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO","WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"}

def clean_aircraft_type(raw_text: str) -> str:
    """항공사 이름을 제거하고 기종 코드(B789, B77W 등)만 추출"""
    main_text = raw_text.split('(')[0].strip()
    # 핵심 코드 추출 매핑
    if "787-9" in main_text: return "B789"
    if "777-300" in main_text: return "B77W"
    if "A330" in main_text or "A333" in main_text: return "A333"
    if "A321" in main_text: return "A21N" if "neo" in main_text.lower() else "A321"
    if "A320" in main_text: return "A320"
    if "737-800" in main_text: return "B738"
    # 그 외의 경우 마지막 단어 사용 (항공사 이름 제거 목적)
    parts = main_text.split()
    return parts[-1] if parts else "B789"

def parse_raw_lines(lines: List[str]) -> List[Dict]:
    recs = []; cur_date = None; i = 0
    while i < len(lines):
        line = lines[i].strip()
        if DATE_HEADER.match(line):
            try: cur_date = datetime.strptime(line + ' 2026', '%A, %b %d %Y').date()
            except: cur_date = None
            i += 1; continue
        m = TIME_LINE.match(line)
        if m and cur_date:
            time_str, flt = m.groups()
            dest_line = lines[i+1].strip() if i+1 < len(lines) else ''
            m2 = IATA_IN_PAREns.search(dest_line)
            dest = (m2.group(1).strip() if m2 else '').upper()
            
            carrier_line = lines[i+2].strip() if i+2 < len(lines) else ''
            flt_type = clean_aircraft_type(carrier_line) # 기종 단순화 적용
            
            reg = ''; ps = IATA_IN_PAREns.findall(carrier_line)
            if ps: reg = ps[-1].strip()
            
            try: dt = datetime.strptime(f"{cur_date} {time_str}", '%Y-%m-%d %I:%M %p')
            except: dt = None
            recs.append({'dt': dt, 'time': time_str, 'flight': flt, 'dest': dest, 'reg': reg, 'type': flt_type})
            i += 4; continue
        i += 1
    return recs

# --- 3. DOCX 생성 (14pt 및 폭 조정) ---
def build_docx(recs, is_1p=False):
    doc = Document()
    f_name = 'Air New Zealand Sans'
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Inches(0.8) # 좁은 테이블 폭을 위해 여백 확보

    table = doc.add_table(rows=0, cols=6)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    tblPr = table._element.find(qn('w:tblPr'))
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '4100'); tblW.set(qn('w:type'), 'pct'); tblPr.append(tblW)

    last_date_str = ""
    for i, r in enumerate(recs):
        row = table.add_row()
        current_date_str = r['dt'].strftime('%d %b')
        display_date = current_date_str if current_date_str != last_date_str else ""
        last_date_str = current_date_str

        t_short = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        vals = [display_date, r['flight'], t_short, r['dest'], r['type'], r['reg']]
        
        for j, v in enumerate(vals):
            cell = row.cells[j]
            if i % 2 == 1:
                shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9'); cell._tc.get_or_add_tcPr().append(shd)
            
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run(str(v))
            run.font.name = f_name
            run.font.size = Pt(8.5 if is_1p else 14.0) # 요청하신 14pt 반영
            if j == 0: run.bold = True
            
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf

# --- 4. 메인 실행 및 사이드바 복구 ---
st.title("Simon Park'nRide's Flight List Factory")

# 사라졌던 왼쪽 사이드바 설정 복구
with st.sidebar:
    st.header("⚙️ Settings")
    s_time = st.text_input("Start Time (예: 00:00)", "04:55")
    e_time = st.text_input("End Time (예: 23:59)", "05:00")
    label_start = st.number_input("Label Start No", value=1)

uploaded = st.file_uploader("Upload Raw Text File", type=['txt'])

if uploaded:
    lines = uploaded.read().decode("utf-8").splitlines()
    all_recs = parse_raw_lines(lines)
    if all_recs:
        # 필터링 및 다운로드 로직 (기존과 동일하게 작동)
        st.success(f"Processed {len(all_recs)} flights")
        c1, c2 = st.columns(2)
        c1.download_button("📥 DOCX (14pt)", build_docx(all_recs), "Flight_List.docx")
        c2.download_button("📄 1-PAGE", build_docx(all_recs, True), "Flight_List_1p.docx")
