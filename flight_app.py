import streamlit as st
import re
import io
import pandas as pd
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

# --- 1. UI 및 스타일 (기존 유지) ---
st.set_page_config(page_title="Flight List Factory", layout="centered")
st.markdown("""
    <style>
    .stApp { background-color: #000000; }
    [data-testid="stSidebar"] { background-color: #111111 !important; }
    .stMarkdown, p, h1, h2, h3, label { color: #ffffff !important; }
    div.stDownloadButton > button {
        background-color: #ffffff !important; color: #000000 !important;
        border: 2px solid #ffffff !important; border-radius: 8px !important;
        font-weight: 800 !important; width: 100% !important;
    }
    div.stDownloadButton > button:hover {
        background-color: #60a5fa !important; color: #ffffff !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 파싱 로직 (기존 유지) ---
TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s[AP]M)\t([A-Z]{2}\d+[A-Z]?)\s*$")
DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
IATA_IN_PAREns = re.compile(r"\(([^)]+)\)")
PLANE_TYPES = ['A21N','A20N','A320','32Q','320','73H','737','74Y','77W','B77W','789','B789','359','A359','332','A332','AT76','388','333','A333','330']
PLANE_TYPE_PATTERN = re.compile(r"\b(" + "|".join(PLANE_TYPES) + r")\b", re.IGNORECASE)

def parse_raw_lines(lines: List[str]) -> List[Dict]:
    records = []
    current_date = None
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if DATE_HEADER.match(line):
            try: current_date = datetime.strptime(line + ' 2026', '%A, %b %d %Y').date()
            except: current_date = None
            i += 1; continue
        m = TIME_LINE.match(line)
        if m and current_date:
            time_str, flight = m.groups()
            dest_line = lines[i+1].strip() if i+1 < len(lines) else ''
            dest_iata = (IATA_IN_PAREns.search(dest_line).group(1) if IATA_IN_PAREns.search(dest_line) else '').upper()
            carrier_line = lines[i+2].strip() if i+2 < len(lines) else ''
            ptype = PLANE_TYPE_PATTERN.search(carrier_line)
            plane_type = ptype.group(1).upper() if ptype else ''
            reg = ''
            parens = IATA_IN_PAREns.findall(carrier_line)
            if parens: reg = parens[-1].strip()
            try: dt = datetime.strptime(f"{current_date} {time_str}", '%Y-%m-%d %I:%M %p')
            except: dt = None
            records.append({'dt': dt, 'date_label': current_date.strftime('%d %b'), 'time': time_str, 'flight': flight, 'dest': dest_iata, 'type': plane_type, 'reg': reg})
            i += 4; continue
        i += 1
    return records

# --- 3. PDF Labels (가이드 PDF 레이아웃 정밀 재현) ---
def build_labels_stream(records, start_num):
    target = io.BytesIO()
    c = canvas.Canvas(target, pagesize=A4)
    w, h = A4
    margin, gutter = 15*mm, 6*mm
    col_w, row_h = (w - 2*margin - gutter) / 2, (h - 2*margin) / 5
    
    for i, r in enumerate(records):
        if i > 0 and i % 10 == 0: c.showPage()
        idx = i % 10
        x = margin + (idx % 2) * (col_w + gutter)
        y = h - margin - (idx // 2 + 1) * row_h
        
        # 1. 칸 테두리
        c.setStrokeGray(0.8); c.setLineWidth(0.1)
        c.rect(x, y + 2*mm, col_w, row_h - 4*mm)
        
        # 2. 순번 (좌상단) & 날짜 (우상단)
        c.setFont('Helvetica-Bold', 12)
        c.drawString(x + 5*mm, y + row_h - 11*mm, str(start_num + i))
        c.setFont('Helvetica', 10)
        c.drawRightString(x + col_w - 5*mm, y + row_h - 11*mm, r['date_label'])
        
        # 3. 비행 편명 (중앙 상단)
        c.setFont('Helvetica-Bold', 32)
        c.drawCentredString(x + col_w/2, y + row_h - 26*mm, r['flight'])
        
        # 4. 목적지 (중앙)
        c.setFont('Helvetica-Bold', 22)
        c.drawCentredString(x + col_w/2, y + row_h - 38*mm, r['dest'])
        
        # 5. 시간 (중앙 하단)
        tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        c.setFont('Helvetica-Bold', 28)
        c.drawCentredString(x + col_w/2, y + 16*mm, tdisp)
        
        # 6. 기종 & 등록번호 (맨 아래)
        c.setFont('Helvetica', 9)
        info_txt = f"{r['type']}   {r['reg']}".strip()
        c.drawCentredString(x + col_w/2, y + 8*mm, info_txt)
        
    c.save(); target.seek(0); return target

# --- 4. DOCX (기존 설정 유지하되 1-Page 최적화 포함) ---
def build_docx_stream(records, start_dt, end_dt):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Inches(0.4)
    section.left_margin = section.right_margin = Inches(1.2)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_head = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}")
    run_head.bold = True
    run_head.font.size = Pt(7.5) # 요청하셨던 7.5pt
    
    table = doc.add_table(rows=0, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, r in enumerate(records):
        row = table.add_row()
        tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        vals = [r['flight'], tdisp, r['dest'], r['type'], r['reg']]
        for j, val in enumerate(vals):
            cell = row.cells[j]
            if i % 2 == 1:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9'); tcPr.append(shd)
            para = cell.paragraphs[0]
            run = para.add_run(str(val))
            run.font.size = Pt(7.5)
    target = io.BytesIO(); doc.save(target); target.seek(0); return target

# --- 5. 사이드바 및 실행 ---
with st.sidebar:
    st.header("⚙️ Settings")
    s_val = st.text_input("Start Time", "04:55")
    e_val = st.text_input("End Time", "05:00")
    label_start = st.number_input("Label Start No", value=1)

st.title("Simon Park'nRide's Factory")
uploaded_file = st.file_uploader("Upload Raw Text", type=['txt'])

if uploaded_file:
    lines = uploaded_file.read().decode("utf-8").splitlines()
    all_recs = parse_raw_lines(lines)
    if all_recs:
        # 필터링 로직
        dates = sorted({r['dt'].date() for r in all_recs if r.get('dt')})
        if dates:
            day1 = dates[0]
            day2 = dates[1] if len(dates) >= 2 else (day1 + timedelta(days=1))
            start_dt = datetime.combine(day1, datetime.strptime(s_val, '%H:%M').time())
            end_dt = datetime.combine(day2, datetime.strptime(e_val, '%H:%M').time())
            filtered = [r for r in all_recs if r.get('dt') and (start_dt <= r['dt'] <= end_dt)]
            
            if filtered:
                col1, col2 = st.columns(2)
                fn = f"List_{start_dt.strftime('%d-%m')}"
                col1.download_button("📥 DOCX (1-Page)", build_docx_stream(filtered, start_dt, end_dt), f"{fn}.docx")
                col2.download_button("📥 PDF Labels", build_labels_stream(filtered, label_start), f"Labels_{fn}.pdf")
                st.table([{'No': label_start+i, 'Flight': r['flight'], 'Time': r['time'], 'Dest': r['dest']} for i, r in enumerate(filtered)])
