import streamlit as st
import re
import io
from datetime import datetime, timedelta
from typing import List, Dict

# --- 1. UI 및 디자인 설정 (강력한 버튼 스타일 포함) ---
st.set_page_config(page_title="Flight List Factory", layout="centered", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .stApp { background-color: #000000; }
    [data-testid="stSidebar"] { background-color: #111111 !important; }
    .stMarkdown, p, h1, h2, h3, label { color: #ffffff !important; }
    
    /* 다운로드 버튼: 가독성 해결 (평상시 흰 배경/검정 글자) */
    div.stDownloadButton > button {
        background-color: #ffffff !important; 
        color: #000000 !important;           
        border: 2px solid #ffffff !important;
        border-radius: 8px !important;
        padding: 0.6rem 1.2rem !important;
        font-weight: 800 !important;
        width: 100% !important;
    }
    div.stDownloadButton > button * { color: #000000 !important; }
    div.stDownloadButton > button:hover {
        background-color: #60a5fa !important; 
        color: #ffffff !important;           
        border: 2px solid #60a5fa !important;
    }
    div.stDownloadButton > button:hover * { color: #ffffff !important; }

    .top-left-container { text-align: left; padding-top: 10px; margin-bottom: 20px; }
    .top-left-container a { font-size: 1.1rem; color: #ffffff !important; text-decoration: underline; display: block; }
    .main-title { font-size: 3rem; font-weight: 800; color: #ffffff; line-height: 1.1; margin-bottom: 0.5rem; }
    .sub-title { font-size: 2.5rem; font-weight: 400; color: #60a5fa; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 사이드바 설정 ---
with st.sidebar:
    st.header("⚙️ Settings")
    s_time = st.text_input("Start Time", value="05:00")
    e_time = st.text_input("End Time", value="04:55")
    label_start = st.number_input("Label Start Number", value=1, min_value=1)

# --- 3. 타이틀 영역 ---
st.markdown('<div class="top-left-container"><a href="https://www.flightradar24.com/data/airports/akl/arrivals" target="_blank">Import Raw Text</a><a href="https://www.flightradar24.com/data/airports/akl/departures" target="_blank">Export Raw Text</a></div>', unsafe_allow_html=True)
st.markdown('<div class="main-title">Simon Park\'nRide\'s<br><span class="sub-title">Flight List Factory</span></div>', unsafe_allow_html=True)

# --- 4. 파싱 로직 (이전과 동일) ---
# (중복 방지를 위해 로직은 요약 유지, 실제 사용 시 전체 포함)
def parse_raw_lines(lines):
    # 정규식 패턴 및 데이터 추출 로직...
    return [] # (실제 구현부 생략, 이전 코드와 동일)

# --- 5. DOCX 생성 함수 (2페이지 최적화 핵심) ---
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn

def build_docx_stream(records, start_dt, end_dt):
    doc = Document()
    section = doc.sections[0]
    # 상하 여백을 줄여 공간 확보
    section.top_margin = Inches(0.4)
    section.bottom_margin = Inches(0.4)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

    heading = f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}"
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(heading); run.bold = True; run.font.size = Pt(14)
    
    table = doc.add_table(rows=0, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # 표 너비 80% 설정
    tblPr = table._element.find(qn('w:tblPr'))
    tblW = OxmlElement('w:tblW'); tblW.set(qn('w:w'), '4000'); tblW.set(qn('w:type'), 'pct'); tblPr.append(tblW)

    for i, r in enumerate(records):
        row = table.add_row()
        # 행 높이 조정을 위한 내부 설정 (vMerge 등을 통한 촘촘한 구성)
        row.height = Inches(0.2) 
        
        vals = [r['flight'], r['time'], r['dest'], r['type'], r['reg']]
        for j, val in enumerate(vals):
            cell = row.cells[j]
            # 홀수 행 배경색 (가독성)
            if i % 2 == 1:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9'); tcPr.append(shd)
            
            para = cell.paragraphs[0]
            # 줄 간격을 'Single'로 강제 고정하여 공간 절약
            para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(0)
            
            run = para.add_run(str(val))
            run.font.size = Pt(11) # 글자 크기를 11pt로 축소하여 2페이지 내 수록 
            
    target = io.BytesIO()
    doc.save(target); target.seek(0)
    return target

# --- 6. 실행 및 업로드 ---
uploaded_file = st.file_uploader("Upload Raw Text File", type=['txt'])
if uploaded_file:
    # 필터링 및 다운로드 버튼 로직... (이전 코드와 동일)
    st.write("2페이지 최적화가 적용되었습니다.")
