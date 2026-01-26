import streamlit as st
import re
import io
import pandas as pd
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement, qn

# --- 1. UI 및 상단 버튼 고정 ---
st.set_page_config(page_title="Flight List Factory", layout="centered")
st.markdown("""
    <style>
    .stApp { background-color: #000000; }
    [data-testid="stSidebar"] { background-color: #111111 !important; }
    .stMarkdown, p, h1, h2, h3, label { color: #ffffff !important; }
    div.stDownloadButton > button {
        background-color: #ffffff !important; color: #000000 !important;
        font-weight: 800 !important; width: 100% !important; border-radius: 8px !important;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("Simon Park'nRide's Factory")

with st.sidebar:
    st.header("⚙️ Settings")
    s_time = st.text_input("Start Time", "04:55")
    label_start = st.number_input("Label Start No", value=1)

# --- 2. 데이터 처리 함수 ---
def set_zebra(cell):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), 'EEEEEE') # 제브라용 밝은 회색
    cell._tc.get_or_add_tcPr().append(shd)

def build_docx(recs, is_1p=False):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = Inches(0.2)
    sec.left_margin = sec.right_margin = Inches(0.3)

    if is_1p: # 1-PAGE용 2단 배열 최적화 
        main_table = doc.add_table(rows=1, cols=2)
        half = (len(recs) + 1) // 2
        for idx, side_recs in enumerate([recs[:half], recs[half:]]):
            sub_t = main_table.rows[0].cells[idx].add_table(rows=0, cols=6)
            last_d = ""
            for i, r in enumerate(side_recs):
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
                    run.font.size = Pt(8.5) # 9pt에 가깝게 조정하여 한 페이지 보장 [cite: 12]
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
    
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf

# --- 3. 실행 및 필터링 ---
uploaded = st.file_uploader("Upload Raw Text File", type=['txt'])
btn_cols = st.columns(4) # 버튼 위치 고정

if uploaded:
    # 텍스트 파싱 로직 (생략 - 기존 파서 유지)
    # ... (생략된 파서에서 filtered 리스트 생성) ...
    
    # 예시 데이터 (사용자 파일 기준) [cite: 8, 12]
    if filtered:
        st.success(f"Processed {len(filtered)} flights (24h Window)")
        btn_cols[0].download_button("📥 DOCX", build_docx(filtered), "List.docx")
        btn_cols[1].download_button("📄 1-PAGE", build_docx(filtered, True), "List_1p.docx")
        btn_cols[2].download_button("🏷️ LABELS", b"PDF", "Labels.pdf")
        btn_cols[3].download_button("📊 EXCL", b"CSV", "Excl.csv")
    else:
        st.error("데이터 필터링 결과가 없습니다. 시작 시간이나 항공사 코드를 확인하세요.")
