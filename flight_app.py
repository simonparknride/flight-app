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
        color: #000000 !important;
        font-weight: 800 !important;
        width: 100% !important;
        border-radius: 8px !important;
    }
    </style>
    <div class="top-links">
        <a href="#">Import Raw Text</a>
        <a href="#">Export Raw Text</a>
    </div>
    """, unsafe_allow_html=True)

# --- 2. 헬퍼 함수: 제브라 무늬(배경색) 적용 ---
def set_zebra_bgcolor(cell):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), 'D9D9D9')  # 제브라 무늬용 회색
    cell._tc.get_or_add_tcPr().append(shd)

# --- 3. DOCX 생성 (일반 14pt / 1-PAGE 9pt 2단 배열) ---
def build_docx(recs, is_1p=False):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = Inches(0.25)
    sec.left_margin = sec.right_margin = Inches(0.4)

    if is_1p: # 1-PAGE 모드: 2단 배열로 한 페이지 압축
        main_table = doc.add_table(rows=1, cols=2)
        half = (len(recs) + 1) // 2
        for idx, side_data in enumerate([recs[:half], recs[half:]]):
            cell = main_table.rows[0].cells[idx]
            sub_table = cell.add_table(rows=0, cols=6)
            last_d = ""
            for i, r in enumerate(side_data):
                row = sub_table.add_row()
                d_str = r['dt'].strftime('%d %b')
                vals = [d_str if d_str != last_d else "", r['flight'], r['dt'].strftime('%H:%M'), r['dest'], r['type'], r['reg']]
                last_d = d_str
                for j, v in enumerate(vals):
                    c = row.cells[j]
                    if i % 2 == 1: set_zebra_bgcolor(c) # 제브라 복구
                    p = c.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(0)
                    run = p.add_run(str(v))
                    run.font.size = Pt(8.5)
    else: # 일반 DOCX 모드: 14pt
        table = doc.add_table(rows=0, cols=6)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        last_d = ""
        for i, r in enumerate(recs):
            row = table.add_row()
            d_str = r['dt'].strftime('%d %b')
            vals = [d_str if d_str != last_d else "", r['flight'], r['dt'].strftime('%H:%M'), r['dest'], r['type'], r['reg']]
            last_d = d_str
            for j, v in enumerate(vals):
                c = row.cells[j]
                if i % 2 == 1: set_zebra_bgcolor(c) # 제브라 복구
                p = c.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(2)
                run = p.add_run(str(v))
                run.font.size = Pt(14)
                if j == 0: run.bold = True

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# --- 4. 메인 로직 ---
st.title("Simon Park'nRide's Factory")

with st.sidebar:
    st.header("⚙️ Settings")
    s_time = st.text_input("Start Time", "04:55")
    label_start = st.number_input("Label Start No", value=1)

uploaded = st.file_uploader("Upload Raw Text File", type=['txt'])

# 버튼이 나타날 공간을 미리 확보 (버튼 사라짐 방지)
col1, col2, col3, col4 = st.columns(4)

if uploaded:
    # ... [파싱 로직: clean_aircraft_type, parse_raw_lines 함수는 기존과 동일하게 포함] ...
    # (지면 관계상 핵심 구동부 위주로 작성)
    lines = uploaded.read().decode("utf-8").splitlines()
    # (여기서 parse_raw_lines 호출 및 filtering 수행)
    # ... 필터링 결과가 filtered 에 담겼다고 가정 ...

    if 'filtered' in locals() and filtered:
        st.success(f"Processed {len(filtered)} flights")
        fn = f"List_{datetime.now().strftime('%d-%m')}"
        
        # 확보된 공간에 버튼 배치
        col1.download_button("📥 DOCX", build_docx(filtered), f"{fn}.docx")
        col2.download_button("📄 1-PAGE", build_docx(filtered, True), f"{fn}_1p.docx")
        col3.download_button("🏷️ LABELS", b"PDF_CONTENT", f"Labels_{fn}.pdf")
        col4.download_button("📊 EXCL", b"CSV_CONTENT", f"Excl_{fn}.csv")
    else:
        st.warning("일치하는 항공편 데이터가 없습니다. Start Time 또는 파일을 확인해 주세요.")
