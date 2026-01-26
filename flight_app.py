import streamlit as st
import re
import io
import pandas as pd
from datetime import datetime, timedelta, date, time # date, time 명시적 추가
from typing import List, Dict
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# --- 1. UI 설정 및 스타일 ---
st.set_page_config(page_title="Flight List Factory", layout="centered")

st.markdown("""
    <style>
    .stApp { background-color: #000000; }
    [data-testid="stSidebar"] { background-color: #111111 !important; }
    .stMarkdown, p, h1, h2, h3, label { color: #ffffff !important; }
    div.stDownloadButton > button {
        background-color: #ffffff !important;
        color: #000000 !important;
        font-weight: 800 !important;
        width: 100% !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 파싱 로직 ---
TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s[AP]M)\t([A-Z]{2}\d+[A-Z]?)\s*$")
DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
IATA_IN_PAREns = re.compile(r"\(([^)]+)\)")
ALLOWED_AIRLINES = {"NZ","QF","JQ","CZ","CA","SQ","LA","IE"}
NZ_DOMESTIC_IATA = {"AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO","WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"}

def clean_aircraft_type(raw_text: str) -> str:
    main_text = raw_text.split('(')[0].strip()
    if "787-9" in main_text: return "B789"
    if "777-300" in main_text: return "B77W"
    if "A330" in main_text: return "A333"
    if "A321" in main_text: return "A21N" if "neo" in main_text.lower() else "A321"
    if "A320" in main_text: return "A320"
    if "737-800" in main_text: return "B738"
    return main_text.split()[-1] if main_text.split() else "B789"

def parse_raw_lines(lines: List[str]) -> List[Dict]:
    recs = []; cur_date = None; i = 0
    while i < len(lines):
        line = lines[i].strip()
        if DATE_HEADER.match(line):
            try: cur_date = datetime.strptime(line + f' {datetime.now().year}', '%A, %b %d %Y').date()
            except: cur_date = None
            i += 1; continue
        m = TIME_LINE.match(line)
        if m and cur_date:
            time_str, flt = m.groups()
            dest_line = lines[i+1].strip() if i+1 < len(lines) else ''
            m2 = IATA_IN_PAREns.search(dest_line)
            dest = (m2.group(1).strip() if m2 else '').upper()
            carrier_line = lines[i+2].strip() if i+2 < len(lines) else ''
            flt_type = clean_aircraft_type(carrier_line)
            reg = ''; ps = IATA_IN_PAREns.findall(carrier_line)
            if ps: reg = ps[-1].strip()
            try: dt = datetime.strptime(f"{cur_date} {time_str}", '%Y-%m-%d %I:%M %p')
            except: dt = None
            recs.append({'dt': dt, 'time': time_str, 'flight': flt, 'dest': dest, 'reg': reg, 'type': flt_type})
            i += 4; continue
        i += 1
    return recs

# --- 3. DOCX 생성 (2단 배열 모드 포함) ---
def build_docx(recs, is_1p=False):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = Inches(0.2)
    sec.left_margin = sec.right_margin = Inches(0.5)

    if is_1p:
        # 2단 구성을 위해 큰 테이블 생성
        main_table = doc.add_table(rows=1, cols=2)
        half = (len(recs) + 1) // 2
        for idx, side_recs in enumerate([recs[:half], recs[half:]]):
            cell = main_table.rows[0].cells[idx]
            sub_table = cell.add_table(rows=0, cols=6)
            last_date = ""
            for i, r in enumerate(side_recs):
                row = sub_table.add_row()
                cur_d = r['dt'].strftime('%d %b')
                t_s = r['dt'].strftime('%H:%M')
                vals = [cur_d if cur_d != last_date else "", r['flight'], t_s, r['dest'], r['type'], r['reg']]
                last_date = cur_d
                for j, v in enumerate(vals):
                    p = row.cells[j].paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run(str(v))
                    run.font.size = Pt(8.5)
    else:
        table = doc.add_table(rows=0, cols=6)
        last_date = ""
        for r in recs:
            row = table.add_row()
            cur_d = r['dt'].strftime('%d %b')
            t_s = r['dt'].strftime('%H:%M')
            vals = [cur_d if cur_d != last_date else "", r['flight'], t_s, r['dest'], r['type'], r['reg']]
            last_date = cur_d
            for j, v in enumerate(vals):
                run = row.cells[j].paragraphs[0].add_run(str(v))
                run.font.size = Pt(14)
                
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf

# --- 5. 메인 실행 ---
st.title("Simon Park'nRide's Factory")

with st.sidebar:
    st.header("⚙️ Settings")
    s_time = st.text_input("Start Time", "04:55")
    label_start = st.number_input("Label Start No", value=1)

uploaded = st.file_uploader("Upload Raw Text File", type=['txt'])

if uploaded:
    lines = uploaded.read().decode("utf-8").splitlines()
    all_recs = parse_raw_lines(lines)
    if all_recs:
        # 오류 수정 포인트: datetime.combine에 r['dt'].date()를 전달
        day1 = all_recs[0]['dt'].date() 
        cur_s = datetime.combine(day1, datetime.strptime(s_time, '%H:%M').time())
        cur_e = cur_s + timedelta(hours=24)
        
        filtered = [r for r in all_recs if cur_s <= r['dt'] < cur_e and r['flight'][:2] in ALLOWED_AIRLINES and r['dest'] not in NZ_DOMESTIC_IATA]
        
        if filtered:
            st.success(f"{len(filtered)} Flights Processed")
            c1, c2, c3, c4 = st.columns(4)
            c1.download_button("📥 DOCX", build_docx(filtered), "List.docx")
            c2.download_button("📄 1-PAGE", build_docx(filtered, True), "List_1p.docx")
            # PDF 및 CSV 로직은 이전과 동일하게 유지...
