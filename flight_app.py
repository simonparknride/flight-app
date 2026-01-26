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

# --- 1. UI 스타일 ---
st.set_page_config(page_title="Flight List Factory", layout="centered")
st.markdown("""
    <style>
    .stApp { background-color: #000000; }
    .stMarkdown, p, h1, h2, h3, label { color: #ffffff !important; }
    div.stDownloadButton > button {
        background-color: #ffffff !important;
        color: #000000 !important;
        font-weight: 800 !important;
        width: 100% !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 파싱 로직 (항공사 제거 및 기종 코드만 추출) ---
TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s[AP]M)\t([A-Z]{2}\d+[A-Z]?)\s*$")
DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
IATA_IN_PAREns = re.compile(r"\(([^)]+)\)")
ALLOWED_AIRLINES = {"NZ","QF","JQ","CZ","CA","SQ","LA","IE"}
NZ_DOMESTIC_IATA = {"AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO","WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"}

def clean_aircraft_type(raw_text: str) -> str:
    """항공사 이름을 제거하고 기종 코드만 반환"""
    # 괄호 안의 등록번호 제거
    main_text = raw_text.split('(')[0].strip()
    
    # 기종 코드 매핑 (원본 텍스트에서 핵심 코드만 추출)
    if "787-9" in main_text: return "B789"
    if "777-300" in main_text: return "B77W"
    if "777-200" in main_text: return "B772"
    if "A330-300" in main_text or "A333" in main_text: return "A333"
    if "A321" in main_text: return "A21N" if "neo" in main_text.lower() else "A321"
    if "A320" in main_text: return "A320"
    if "737-800" in main_text: return "B738"
    
    # 매핑되지 않은 경우, 마지막 단어가 보통 기종 코드인 경우가 많음 (예: "Wamos Air A333")
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
            flt_type = clean_aircraft_type(carrier_line) # 기종 코드 청소 적용
            
            reg = ''; ps = IATA_IN_PAREns.findall(carrier_line)
            if ps: reg = ps[-1].strip()
            
            try: dt = datetime.strptime(f"{cur_date} {time_str}", '%Y-%m-%d %I:%M %p')
            except: dt = None
            recs.append({'dt': dt, 'time': time_str, 'flight': flt, 'dest': dest, 'reg': reg, 'type': flt_type})
            i += 4; continue
        i += 1
    return recs

# --- 3. DOCX 생성 (폰트 14pt 및 좁은 폭) ---
def build_docx(recs, start_dt, is_1p=False):
    doc = Document()
    f_name = 'Air New Zealand Sans'
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Inches(0.8) # 여백을 더 늘려 폭을 좁힘

    table = doc.add_table(rows=0, cols=6)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    tblPr = table._element.find(qn('w:tblPr'))
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '4000') # 좁은 폭 유지
    tblW.set(qn('w:type'), 'pct'); tblPr.append(tblW)

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
                shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9')
                cell._tc.get_or_add_tcPr().append(shd)
            
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run(str(v))
            run.font.name = f_name
            # [핵심] 폰트 사이즈 14pt 적용
            run.font.size = Pt(8.0 if is_1p else 14.0)
            if j == 0: run.bold = True
            
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf

# --- 4. 메인 (생략) ---
st.title("Simon's Flight List Factory")
uploaded = st.file_uploader("Upload Raw Text File", type=['txt'])
if uploaded:
    lines = uploaded.read().decode("utf-8").splitlines()
    all_recs = parse_raw_lines(lines)
    if all_recs:
        st.success(f"Processed {len(all_recs)} flights")
        c1, c2 = st.columns(2)
        c1.download_button("📥 DOCX (14pt)", build_docx(all_recs, None), "Flight_List.docx")
        c2.download_button("📄 1-PAGE", build_docx(all_recs, None, True), "Flight_List_1p.docx")
