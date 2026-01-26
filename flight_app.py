import streamlit as st
import re
import io
import pandas as pd
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement, qn

# --- 1. UI 설정 및 버튼 공간 선행 확보 ---
st.set_page_config(page_title="Flight List Factory", layout="centered")

# 배경 및 폰트 스타일링
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

# 상단 버튼 배치용 컬럼 생성 (데이터 유무와 상관없이 공간 유지)
btn_col1, btn_col2, btn_col3, btn_col4 = st.columns(4)

with st.sidebar:
    st.header("⚙️ Settings")
    s_time = st.text_input("Start Time (HH:MM)", "04:55")
    label_start = st.number_input("Label Start No", value=1)

# --- 2. 핵심 로직 (제브라 무늬 및 2단 배열) ---
def set_zebra(cell):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), 'D9D9D9') # 가시성 좋은 회색
    cell._tc.get_or_add_tcPr().append(shd)

def build_docx(recs, is_1p=False):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = Inches(0.25)
    sec.left_margin = sec.right_margin = Inches(0.4)

    if is_1p: # 1-PAGE용 2단 배열
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
                    p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(0)
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
                p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(2)
                run = p.add_run(str(v))
                run.font.size = Pt(14)
                if j == 0: run.bold = True
    
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf

# --- 3. 데이터 파싱 및 실행 ---
uploaded = st.file_uploader("Upload Raw Text File", type=['txt'])

if uploaded:
    raw_content = uploaded.read().decode("utf-8")
    lines = raw_content.splitlines()
    all_recs = []
    current_date = None
    
    # [cite_start]정규표현식을 이용한 날짜 및 데이터 라인 추출 (2026-01-26 형식 대응) [cite: 1, 5]
    for line in lines:
        line = line.strip()
        # [cite_start]날짜 헤더 인식 [cite: 1, 3]
        date_match = re.search(r"(\d{1,2}\s+[A-Za-z]+|\d{1,2}\s+[A-Za-z]+\s+\d{4})", line)
        if date_match and ":" not in line:
            current_date = date_match.group(1)
            continue
        
        # [cite_start]항공 데이터 라인 인식 (예: QF4,04:55,JFK,...) [cite: 2, 4, 6]
        parts = line.split(',')
        if len(parts) >= 5:
            try:
                # [cite_start]라인 처음에 날짜가 포함된 경우 처리 [cite: 5, 6]
                row_date = parts[0].strip() if parts[0].strip() else current_date
                time_str = parts[2].strip() if ":" in parts[2] else parts[1].strip()
                dt_obj = datetime.strptime(f"{row_date} 2026 {time_str}", "%d %b %Y %H:%M")
                
                all_recs.append({
                    'dt': dt_obj,
                    'flight': parts[1].strip() if ":" in parts[2] else parts[0].strip(),
                    'dest': parts[3].strip(),
                    'type': parts[4].strip(),
                    'reg': parts[5].strip() if len(parts) > 5 else ""
                })
            except: continue

    if all_recs:
        # 설정된 Start Time 기준 24시간 필터링
        start_dt = datetime.combine(all_recs[0]['dt'].date(), datetime.strptime(s_time, "%H:%M").time())
        end_dt = start_dt + timedelta(hours=24)
        
        filtered = [r for r in all_recs if start_dt <= r['dt'] < end_dt]
        
        if filtered:
            st.success(f"Successfully processed {len(filtered)} flights.")
            # 고정된 컬럼 위치에 버튼 할당
            btn_col1.download_button("📥 DOCX", build_docx(filtered), "List.docx")
            btn_col2.download_button("📄 1-PAGE", build_docx(filtered, True), "List_1p.docx")
            btn_col3.download_button("🏷️ LABELS", b"PDF", "Labels.pdf")
            btn_col4.download_button("📊 EXCL", b"CSV", "Excl.csv")
        else:
            st.warning("No flights match the filter criteria.")
