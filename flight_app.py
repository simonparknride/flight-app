import streamlit as st
import re
import io
from datetime import datetime, timedelta
from typing import List, Dict
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# --- 1. UI 설정 및 버튼 스타일 ---
st.set_page_config(page_title="Flight List Factory", layout="centered", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .stApp { background-color: #000000; }
    [data-testid="stSidebar"] { background-color: #111111 !important; }
    .stMarkdown, p, h1, h2, h3, label { color: #ffffff !important; }
    
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
    .top-left-container a { font-size: 1.1rem; color: #ffffff !important; text-decoration: underline; display: block; margin-bottom: 5px;}
    .main-title { font-size: 3rem; font-weight: 800; color: #ffffff; line-height: 1.1; margin-bottom: 0.5rem; }
    .sub-title { font-size: 2.5rem; font-weight: 400; color: #60a5fa; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 파싱 및 로직 (기존 유지) ---
# [생략된 파싱 코드는 이전과 동일하게 적용됩니다]
# (parse_raw_lines, filter_records 함수 포함)
def parse_raw_lines(lines):
    records = []
    # ... (기존 정규식 로직 동일하게 적용)
    return records # 실제 구현 시 이전 코드의 내용을 그대로 넣으시면 됩니다.

def filter_records(records, s, e):
    # ... (기존 필터링 로직 동일)
    return records, None, None

# --- 3. DOCX 생성 함수 (One Page vs Two Pages 선택 가능) ---
def build_docx_stream(records, start_dt, end_dt, mode='Two Pages'):
    doc = Document()
    font_name = 'Air New Zealand Sans'
    section = doc.sections[0]
    
    # 공통 여백 최소화
    section.left_margin = section.right_margin = Inches(0.4)
    
    # 모드별 설정
    if mode == 'One Page':
        # 70행을 한 페이지에 넣기 위한 극단적 설정
        section.top_margin = section.bottom_margin = Inches(0.2)
        font_size = Pt(8.5)   # 글자 크기 대폭 축소
        table_width = '3000'  # 표 너비 축소 (pct 단위)
        header_size = Pt(11)
    else:
        # 기존 Two Pages 설정
        section.top_margin = section.bottom_margin = Inches(0.3)
        font_size = Pt(14)
        table_width = '4000'
        header_size = Pt(16)

    # Footer
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_f = footer_para.add_run("created by Simon Park'nRide's Flight List Factory 2026")
    run_f.font.size = Pt(8 if mode == 'One Page' else 10)
    run_f.font.color.rgb = RGBColor(128, 128, 128)

    # Title
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_head = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}")
    run_head.bold = True
    run_head.font.size = header_size

    # Table
    table = doc.add_table(rows=0, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = table._element.find(qn('w:tblPr'))
    tblW = OxmlElement('w:tblW'); tblW.set(qn('w:w'), table_width); tblW.set(qn('w:type'), 'pct'); tblPr.append(tblW)

    for i, r in enumerate(records):
        row = table.add_row()
        # 1페이지 모드일 때 행 높이 고정하여 압축
        if mode == 'One Page':
            row.height = Inches(0.12)
            
        vals = [r['flight'], r['time'], r['dest'], r['type'], r['reg']]
        for j, val in enumerate(vals):
            cell = row.cells[j]
            if i % 2 == 1:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9'); tcPr.append(shd)
            
            para = cell.paragraphs[0]
            para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            para.paragraph_format.space_before = para.paragraph_format.space_after = Pt(0)
            
            run = para.add_run(str(val))
            run.font.name = font_name
            run.font.size = font_size
            
    target = io.BytesIO()
    doc.save(target); target.seek(0)
    return target

# --- 4. 메인 실행부 ---
with st.sidebar:
    st.header("⚙️ Settings")
    s_time = st.text_input("Start Time", value="05:00")
    e_time = st.text_input("End Time", value="04:55")
    label_start = st.number_input("Label Start Number", value=1, min_value=1)

st.markdown('<div class="top-left-container"><a href="...">Import Raw Text</a><a href="...">Export Raw Text</a></div>', unsafe_allow_html=True)
st.markdown('<div class="main-title">Simon Park\'nRide\'s<br><span class="sub-title">Flight List Factory</span></div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Raw Text File", type=['txt'])

if uploaded_file:
    # [파싱 로직 수행 부분]
    # ...
    # if filtered:
    st.success(f"Processed flights (2026 Updated)")
    
    # 버튼 레이아웃 3분할 (One Page / Two Pages / PDF)
    col1, col2, col3 = st.columns(3)
    fn = "Flight_List"
    
    # 1. One Page DOCX 버튼
    col1.download_button(
        "📥 One Page DOCX", 
        build_docx_stream(filtered, s_dt, e_dt, mode='One Page'), 
        f"{fn}_OnePage.docx"
    )
    
    # 2. Two Pages DOCX 버튼 (기존)
    col2.download_button(
        "📥 Two Pages DOCX", 
        build_docx_stream(filtered, s_dt, e_dt, mode='Two Pages'), 
        f"{fn}_TwoPages.docx"
    )
    
    # 3. PDF Labels 버튼
    col3.download_button(
        "📥 PDF Labels", 
        build_labels_stream(filtered, label_start), 
        f"Labels_{fn}.pdf"
    )
