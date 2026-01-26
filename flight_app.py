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

# --- 1. UI 설정 ---
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
        padding: 0.6rem 0.8rem !important;
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
    </style>
    """, unsafe_allow_html=True)

# --- 2. 파싱 및 로직 (기존 유지) ---
# [파싱 함수들은 이전과 동일하게 유지됩니다]
def parse_raw_lines(lines: List[str]) -> List[Dict]:
    # ... (생략된 기존 파싱 로직)
    return records

def filter_records(records, start_hm, end_hm):
    # ... (생략된 기존 필터링 로직)
    return out, start_dt, end_dt

# --- 3. DOCX 생성 (One Page 극한 압축 모드) ---
def build_docx_stream(records, start_dt, end_dt, mode='Two Pages'):
    doc = Document()
    font_name = 'Air New Zealand Sans'
    section = doc.sections[0]
    
    # 공통 가로 여백
    section.left_margin = section.right_margin = Inches(0.5)

    if mode == 'One Page':
        # [최적화 1] 상단 여백을 거의 0에 가깝게 (0.05인치)
        section.top_margin = Inches(0.05)
        section.bottom_margin = Inches(0.1)
        font_size = Pt(8.2)   # 폰트 소폭 추가 축소
        table_width = '3300'  # 표 너비 미세 조정
        header_size = Pt(10)
        # [최적화 2] 단락 여백 완전 제거
        para_space_before = Pt(0)
        para_space_after = Pt(0)
    else:
        section.top_margin = section.bottom_margin = Inches(0.3)
        font_size = Pt(14)
        table_width = '4000'
        header_size = Pt(16)
        para_space_before = Pt(0)
        para_space_after = Pt(12)

    # Footer
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_f = footer_para.add_run("created by Simon Park'nRide's Flight List Factory 2026")
    run_f.font.size = Pt(7 if mode == 'One Page' else 10)
    run_f.font.color.rgb = RGBColor(128, 128, 128)

    # 헤더 (날짜)
    p = doc.add_paragraph()
    # [최적화 3] One Page일 때 왼쪽 정렬 및 여백 제거
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT if mode == 'One Page' else WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = para_space_before
    p.paragraph_format.space_after = para_space_after
    p.paragraph_format.left_indent = Inches(-0.02) # 표의 테두리와 시각적으로 맞추기 위해 미세 조정
    
    run_head = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}")
    run_head.bold = True
    run_head.font.name = font_name
    run_head.font.size = header_size

    # 테이블
    table = doc.add_table(rows=0, cols=5)
    # [최적화 4] 표를 왼쪽으로 정렬하여 헤더와 맞춤
    table.alignment = WD_TABLE_ALIGNMENT.LEFT if mode == 'One Page' else WD_TABLE_ALIGNMENT.CENTER
    
    tblPr = table._element.find(qn('w:tblPr'))
    tblW = OxmlElement('w:tblW'); tblW.set(qn('w:w'), table_width); tblW.set(qn('w:type'), 'pct'); tblPr.append(tblW)

    for i, r in enumerate(records):
        row = table.add_row()
        if mode == 'One Page':
            # [최적화 5] 행 높이 절대적 압축
            row.height = Inches(0.08)
            
        tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        vals = [r['flight'], tdisp, r['dest'], r['type'], r['reg']]
        for j, val in enumerate(vals):
            cell = row.cells[j]
            if i % 2 == 1:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9'); tcPr.append(shd)
            
            para = cell.paragraphs[0]
            para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(0)
            
            run = para.add_run(str(val))
            run.font.name = font_name
            run.font.size = font_size
            
            # 폰트 강제 적용
            rPr = run._element.get_or_add_rPr()
            rFonts = OxmlElement('w:rFonts')
            rFonts.set(qn('w:ascii'), font_name); rFonts.set(qn('w:hAnsi'), font_name); rPr.append(rFonts)

    target = io.BytesIO()
    doc.save(target); target.seek(0)
    return target

# --- 4. 메인 앱 실행부 ---
# (생략된 기존 앱 UI 및 필터 호출 로직)
# if filtered:
#     st.download_button("📥 One Page DOCX", build_docx_stream(filtered, s_dt, e_dt, mode='One Page'), ...)
