import streamlit as st
import re
import io
import pandas as pd
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement, qn

# --- 1. UI 설정 및 버튼 공간 무조건 선점 ---
st.set_page_config(page_title="Flight List Factory", layout="centered")
st.markdown("""
    <style>
    .stApp { background-color: #000000; }
    [data-testid="stSidebar"] { background-color: #111111 !important; }
    .stMarkdown, p, h1, h2, h3, label { color: #ffffff !important; }
    div.stDownloadButton > button {
        background-color: #ffffff !important; color: #000000 !important;
        font-weight: 800 !important; width: 100% !important; height: 3.5rem !important;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("Simon Park'nRide's Factory")

# 데이터 처리 전, 버튼이 들어갈 자리를 미리 만듭니다 (사라짐 방지)
btn_container = st.container()
with btn_container:
    col1, col2, col3, col4 = st.columns(4)

with st.sidebar:
    st.header("⚙️ Settings")
    s_time = st.text_input("Start Time (HH:MM)", "04:55")
    label_start = st.number_input("Label Start No", value=1)

# --- 2. 핵심 로직: 제브라 무늬 및 문서 생성 ---
def set_zebra(cell):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), 'D9D9D9') # 회색 줄무늬
    cell._tc.get_or_add_tcPr().append(shd)

def build_docx(recs, is_1p=False):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = Inches(0.2)
    sec.left_margin = sec.right_margin = Inches(0.4)

    if is_1p: # 1-PAGE (2단 배열, 8.5pt)
        main_table = doc.add_table(rows=1, cols=2)
        half = (len(recs) + 1) // 2
        for idx, side_data in enumerate([recs[:half], recs[half:]]):
            sub_t = main_table.rows[0].cells[idx].add_table(rows=0, cols=6)
            last_d = ""
            for i, r in enumerate(side_data):
                row = sub_t.add_row()
                cur_d = r['dt'].strftime('%d %b')
                vals = [cur_d if cur_d != last_d else "", r['flight'], r['dt'].strftime('%H:%M'), r['dest'], r['type'], r['reg']]
                last_d = cur_d
                for j, v in enumerate(vals):
                    cell = row.cells[j]
                    if i % 2 == 1: set_zebra(cell)
                    p = cell.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run(str(v))
                    run.font.size = Pt(8.5)
    else: # 일반 DOCX (14pt)
        t = doc.add_table(rows=0, cols=6)
        last_d = ""
        for i, r in enumerate(recs):
            row = t.add_row()
            cur_d = r['dt'].strftime('%d %b')
            vals = [cur_d if cur_d != last_d else "", r['flight'], r['dt'].strftime('%H:%M'), r['dest'], r['type'], r['reg']]
            last_d = cur_d
            for j, v in enumerate(vals):
                cell = row.cells[j]
                if i % 2 == 1: set_zebra(cell)
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(str(v))
                run.font.size = Pt(14)
    
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf

# --- 3. 파일 업로드 및 데이터 처리 ---
uploaded = st.file_uploader("Upload Raw Text File", type=['txt'])

# 에러 방지를 위해 변수 초기화
filtered_recs = []

if uploaded:
    raw_text = uploaded.read().decode("utf-8")
    lines = raw_text.splitlines()
    current_date = "26 Jan" # 기본값 (파일 내에서 날짜 추출 실패 시 대비)
    
    parsed_data = []
    for line in lines:
        if not line.strip(): continue
        # 날짜 추출 (예: 26 Jan 2026)
        date_match = re.search(r"(\d{1,2}\s+[A-Za-z]+)", line)
        if date_match and ":" not in line:
            current_date = date_match.group(1)
            continue
        
        # 데이터 행 파싱 (쉼표 구분 대응)
        parts = line.split(',')
        if len(parts) >= 5:
            try:
                # 첫 번째 칸이 날짜면 그것을 사용, 아니면 직전 날짜 사용
                row_date = parts[0].strip() if parts[0].strip() and not parts[0].strip().startswith(('NZ','QF','JQ','CZ','CA','SQ')) else current_date
                time_val = parts[2].strip() if ":" in parts[2] else parts[1].strip()
                dt_obj = datetime.strptime(f"{row_date} 2026 {time_val}", "%d %b %Y %H:%M")
                
                parsed_data.append({
                    'dt': dt_obj,
                    'flight': parts[1].strip() if ":" in parts[2] else parts[0].strip(),
                    'dest': parts[3].strip(),
                    'type': parts[4].strip(),
                    'reg': parts[5].strip() if len(parts) > 5 else ""
                })
            except: continue

    if parsed_data:
        # Start Time 기준 24시간 필터링
        start_pt = datetime.combine(parsed_data[0]['dt'].date(), datetime.strptime(s_time, "%H:%M").time())
        end_pt = start_pt + timedelta(hours=24)
        filtered_recs = [r for r in parsed_data if start_pt <= r['dt'] < end_pt]

# --- 4. 버튼 표시 (데이터가 있을 때만 활성화) ---
if filtered_recs:
    col1.download_button("📥 DOCX", build_docx(filtered_recs), "List.docx")
    col2.download_button("📄 1-PAGE", build_docx(filtered_recs, True), "List_1p.docx")
    col3.download_button("🏷️ LABELS", b"PDF", "Labels.pdf")
    col4.download_button("📊 EXCL", b"CSV", "Excl.csv")
    st.success(f"{len(filtered_recs)}개의 항공편이 처리되었습니다.")
elif uploaded:
    st.warning("일치하는 항공편 데이터가 없습니다. 설정 시간(Start Time)을 확인해 주세요.")
