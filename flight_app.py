import streamlit as st
import re
import io
import pandas as pd
from datetime import datetime, timedelta, date
from typing import List, Dict
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn

# --- 1. UI 설정 및 상단 스타일 ---
st.set_page_config(page_title="Flight List Factory", layout="centered")
st.markdown("""
    <style>
    .stApp { background-color: #000000; }
    [data-testid="stSidebar"] { background-color: #111111 !important; }
    .stMarkdown, p, h1, h2, h3, label { color: #ffffff !important; }
    .top-links { font-size: 14px; margin-bottom: 10px; }
    .top-links a { color: #ffffff !important; text-decoration: underline; margin-right: 15px; }
    div.stDownloadButton > button {
        background-color: #ffffff !important;
        color: #000000 !important;
        font-weight: 800 !important;
        width: 100% !important;
    }
    </style>
    <div class="top-links">
        <a href="#">Import Raw Text</a>
        <a href="#">Export Raw Text</a>
    </div>
    """, unsafe_allow_html=True)

# --- 2. 파싱 로직 (기종 단순화) ---
def clean_aircraft_type(raw_text: str) -> str:
    main_text = raw_text.split('(')[0].strip()
    mapping = {"787-9": "B789", "777-300": "B77W", "A330": "A333", "737-800": "B738"}
    for k, v in mapping.items():
        if k in main_text: return v
    if "A321" in main_text: return "A21N" if "neo" in main_text.lower() else "A321"
    if "A320" in main_text: return "A320"
    return main_text.split()[-1] if main_text.split() else "B789"

def parse_raw_lines(lines: List[str]) -> List[Dict]:
    recs = []; cur_date = None; i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line: i += 1; continue
        if re.match(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$", line):
            try: cur_date = datetime.strptime(line + f' {datetime.now().year}', '%A, %b %d %Y').date()
            except: cur_date = None
            i += 1; continue
        m = re.match(r"^(\d{1,2}:\d{2}\s[AP]M)\t([A-Z]{2}\d+[A-Z]?)\s*$", line)
        if m and cur_date:
            try:
                dest_line = lines[i+1].strip()
                dest = (re.search(r"\(([^)]+)\)", dest_line).group(1)).upper()
                carrier_line = lines[i+2].strip()
                flt_type = clean_aircraft_type(carrier_line)
                reg = re.findall(r"\(([^)]+)\)", carrier_line)[-1]
                dt = datetime.strptime(f"{cur_date} {m.group(1)}", '%Y-%m-%d %I:%M %p')
                recs.append({'dt': dt, 'flight': m.group(2), 'dest': dest, 'reg': reg, 'type': flt_type})
                i += 4; continue
            except: pass
        i += 1
    return recs

# --- 3. DOCX 생성 (제브라 무늬 및 2단 배열) ---
def build_docx(recs, is_1p=False):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = Inches(0.25)
    sec.left_margin = sec.right_margin = Inches(0.4)

    def add_zebra(cell):
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:fill'), 'D9D9D9') # 회색 배경
        cell._tc.get_or_add_tcPr().append(shd)

    if is_1p: # 1-PAGE 모드 (9pt, 2단 배열)
        main_table = doc.add_table(rows=1, cols=2)
        half = (len(recs) + 1) // 2
        for idx, side_data in enumerate([recs[:half], recs[half:]]):
            sub_table = main_table.rows[0].cells[idx].add_table(rows=0, cols=6)
            last_d = ""
            for i, r in enumerate(side_data):
                row = sub_table.add_row()
                d_str = r['dt'].strftime('%d %b')
                vals = [d_str if d_str != last_d else "", r['flight'], r['dt'].strftime('%H:%M'), r['dest'], r['type'], r['reg']]
                last_d = d_str
                for j, v in enumerate(vals):
                    cell = row.cells[j]
                    if i % 2 == 1: add_zebra(cell) # 제브라 무늬 적용
                    p = cell.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(0)
                    run = p.add_run(str(v))
                    run.font.size = Pt(8.5)
    else: # 일반 DOCX 모드 (14pt, 단일 배열)
        table = doc.add_table(rows=0, cols=6)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        last_d = ""
        for i, r in enumerate(recs):
            row = table.add_row()
            d_str = r['dt'].strftime('%d %b')
            vals = [d_str if d_str != last_d else "", r['flight'], r['dt'].strftime('%H:%M'), r['dest'], r['type'], r['reg']]
            last_d = d_str
            for j, v in enumerate(vals):
                cell = row.cells[j]
                if i % 2 == 1: add_zebra(cell) # 제브라 무늬 적용
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(2)
                run = p.add_run(str(v))
                run.font.size = Pt(14)
                if j == 0: run.bold = True

    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf

# --- 4. 메인 실행부 ---
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
        day1 = all_recs[0]['dt'].date()
        cur_s = datetime.combine(day1, datetime.strptime(s_time, '%H:%M').time())
        cur_e = cur_s + timedelta(hours=24)
        
        ALLOWED = {"NZ","QF","JQ","CZ","CA","SQ","LA","IE"}
        DOMESTIC = {"AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO","WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"}
        
        filtered = [r for r in all_recs if cur_s <= r['dt'] < cur_e and r['flight'][:2] in ALLOWED and r['dest'] not in DOMESTIC]
        
        if filtered:
            st.success(f"Processed {len(filtered)} flights")
            c1, c2, c3, c4 = st.columns(4)
            c1.download_button("📥 DOCX", build_docx(filtered), "List.docx")
            c2.download_button("📄 1-PAGE", build_docx(filtered, True), "List_1p.docx")
            # Labels 및 CSV 생략
