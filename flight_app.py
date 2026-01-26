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
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# --- 1. UI 설정 및 상단 링크 (항상 표시) ---
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

st.title("Simon Park'nRide's Factory")

# --- 2. 사이드바 설정 (항상 표시) ---
with st.sidebar:
    st.header("⚙️ Settings")
    s_time = st.text_input("Start Time", "04:55")
    label_start = st.number_input("Label Start No", value=1)

# --- 3. 핵심 로직 (기종 단순화 및 파싱) ---
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
        if not line: i += 1; continue
        if re.match(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$", line):
            try: cur_date = datetime.strptime(line + f' {datetime.now().year}', '%A, %b %d %Y').date()
            except: cur_date = None
            i += 1; continue
        m = re.match(r"^(\d{1,2}:\d{2}\s[AP]M)\t([A-Z]{2}\d+[A-Z]?)\s*$", line)
        if m and cur_date:
            time_str, flt = m.groups()
            dest_line = lines[i+1].strip() if i+1 < len(lines) else ''
            m2 = re.search(r"\(([^)]+)\)", dest_line)
            dest = (m2.group(1).strip() if m2 else '').upper()
            carrier_line = lines[i+2].strip() if i+2 < len(lines) else ''
            flt_type = clean_aircraft_type(carrier_line)
            reg = ''; ps = re.findall(r"\(([^)]+)\)", carrier_line)
            if ps: reg = ps[-1].strip()
            try: dt = datetime.strptime(f"{cur_date} {time_str}", '%Y-%m-%d %I:%M %p')
            except: dt = None
            recs.append({'dt': dt, 'time': time_str, 'flight': flt, 'dest': dest, 'reg': reg, 'type': flt_type})
            i += 4; continue
        i += 1
    return recs

# --- 4. 문서 생성 함수 (DOCX 2단 배열 포함) ---
def build_docx(recs, is_1p=False):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = Inches(0.25)
    sec.left_margin = sec.right_margin = Inches(0.5)

    if is_1p: # 1-PAGE용 2단 배열
        main_table = doc.add_table(rows=1, cols=2)
        half = (len(recs) + 1) // 2
        for idx, side_data in enumerate([recs[:half], recs[half:]]):
            sub_table = main_table.rows[0].cells[idx].add_table(rows=0, cols=6)
            last_d = ""
            for r in side_data:
                row = sub_table.add_row()
                d_str = r['dt'].strftime('%d %b')
                vals = [d_str if d_str != last_d else "", r['flight'], r['dt'].strftime('%H:%M'), r['dest'], r['type'], r['reg']]
                last_d = d_str
                for j, v in enumerate(vals):
                    p = row.cells[j].paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run(str(v))
                    run.font.size = Pt(8.5); run.font.name = 'Arial'
    else: # 일반 14pt 모드
        table = doc.add_table(rows=0, cols=6)
        last_d = ""
        for r in recs:
            row = table.add_row()
            d_str = r['dt'].strftime('%d %b')
            vals = [d_str if d_str != last_d else "", r['flight'], r['dt'].strftime('%H:%M'), r['dest'], r['type'], r['reg']]
            last_d = d_str
            for j, v in enumerate(vals):
                p = row.cells[j].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(str(v))
                run.font.size = Pt(14); run.font.name = 'Arial'
    
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf

# --- 5. 실행부 ---
uploaded = st.file_uploader("Upload Raw Text File", type=['txt'])

if uploaded:
    lines = uploaded.read().decode("utf-8").splitlines()
    all_recs = parse_raw_lines(lines)
    if all_recs:
        # 오류 방지를 위해 date() 추출 후 combine
        day1 = all_recs[0]['dt'].date()
        cur_s = datetime.combine(day1, datetime.strptime(s_time, '%H:%M').time())
        cur_e = cur_s + timedelta(hours=24)
        
        ALLOWED_AIRLINES = {"NZ","QF","JQ","CZ","CA","SQ","LA","IE"}
        NZ_DOMESTIC = {"AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO","WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"}
        
        filtered = [r for r in all_recs if cur_s <= r['dt'] < cur_e and r['flight'][:2] in ALLOWED_AIRLINES and r['dest'] not in NZ_DOMESTIC]
        excl_list = sorted(list({r['flight'] for r in all_recs if r['flight'][:2] in ALLOWED_AIRLINES and r['dest'] not in NZ_DOMESTIC}))

        if filtered:
            st.success(f"Processed {len(filtered)} flights")
            fn = f"List_{cur_s.strftime('%d-%m')}"
            
            # 버튼 4개 배치 (항상 같이 나타남)
            c1, c2, c3, c4 = st.columns(4)
            c1.download_button("📥 DOCX", build_docx(filtered), f"{fn}.docx")
            c2.download_button("📄 1-PAGE", build_docx(filtered, True), f"{fn}_1p.docx")
            c3.download_button("🏷️ LABELS", b"dummy", f"Labels_{fn}.pdf") # 라벨 함수 생략시 더미 데이터
            csv = pd.DataFrame(excl_list).to_csv(index=False, header=False).encode('utf-8-sig')
            c4.download_button("📊 EXCL", csv, f"Excl_{fn}.csv")
