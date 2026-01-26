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

# --- 1. UI 설정 (사용자 원본 스타일 유지) ---
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
    div.stDownloadButton > button:hover { background-color: #60a5fa !important; color: #ffffff !important; }
    .top-left-container a { font-size: 1.1rem; color: #ffffff !important; text-decoration: underline; display: block; margin-bottom: 5px;}
    .main-title { font-size: 3rem; font-weight: 800; color: #ffffff; line-height: 1.1; }
    .sub-title { font-size: 2.5rem; font-weight: 400; color: #60a5fa; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 파싱 로직 (사용자 원본 유지) ---
TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s[AP]M)\t([A-Z]{2}\d+[A-Z]?)\s*$")
DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
IATA_IN_PAREns = re.compile(r"\(([^)]+)\)")
ALLOWED_AIRLINES = {"NZ","QF","JQ","CZ","CA","SQ","LA","IE"}
NZ_DOMESTIC_IATA = {"AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO","WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"}

def parse_raw_lines(lines: List[str]) -> List[Dict]:
    records = []; current_date = None; i = 0
    while i < len(lines):
        line = lines[i].strip()
        if DATE_HEADER.match(line):
            try: current_date = datetime.strptime(line + ' 2026', '%A, %b %d %Y').date()
            except: current_date = None
            i += 1; continue
        m = TIME_LINE.match(line)
        if m and current_date is not None:
            time_str, flight = m.groups()
            dest_line = lines[i+1].strip() if i+1 < len(lines) else ''
            m2 = IATA_IN_PAREns.search(dest_line)
            dest_iata = (m2.group(1).strip() if m2 else '').upper()
            carrier_line = lines[i+2].rstrip('\n') if i+2 < len(lines) else ''
            reg = ''; ps = IATA_IN_PAREns.findall(carrier_line)
            if ps: reg = ps[-1].strip()
            try: dep_dt = datetime.strptime(f"{current_date} {time_str}", '%Y-%m-%d %I:%M %p')
            except: dep_dt = None
            records.append({'dt': dep_dt, 'time': time_str, 'flight': flight, 'dest': dest_iata, 'reg': reg, 'type': 'B789'})
            i += 4; continue
        i += 1
    return records

# --- 3. DOCX 생성 (1-Page 기능 통합) ---
def build_docx_stream(records, start_dt, end_dt, is_1page=False):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Inches(0.3)
    f_size = Pt(7.5) if is_1page else Pt(14)
    
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_h = p.add_run(f"{start_dt.strftime('%d %b')} - {end_dt.strftime('%d %b')}")
    run_h.bold = True
    run_h.font.size = Pt(7.5) if is_1page else Pt(16)
    
    table = doc.add_table(rows=0, cols=5); table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = table._element.find(qn('w:tblPr'))
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '3500' if is_1page else '4000') # 70% vs 80%
    tblW.set(qn('w:type'), 'pct'); tblPr.append(tblW)
    
    for i, r in enumerate(records):
        row = table.add_row()
        t_disp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        vals = [r['flight'], t_disp, r['dest'], r['type'], r['reg']]
        for cell, val in zip(row.cells, vals):
            if i % 2 == 1:
                shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9'); cell._tc.get_or_add_tcPr().append(shd)
            para = cell.paragraphs[0]; para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run(str(val)); run.font.size = f_size
            
    target = io.BytesIO(); doc.save(target); target.seek(0)
    return target

# --- 4. PDF 레이블 (에러 수정 및 겹침 방지) ---
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
        y = h - margin - (idx // 2) * row_h
        
        c.setStrokeGray(0.3) # 에러 해결 핵심 함수
        c.setLineWidth(0.2)
        c.rect(x, y - row_h + 2*mm, col_w, row_h - 4*mm)
        
        c.setFillColorRGB(0, 0, 0)
        c.setFont('Helvetica-Bold', 12); c.drawString(x + 5*mm, y - 10*mm, str(start_num + i))
        c.setFont('Helvetica', 10); c.drawRightString(x + col_w - 5*mm, y - 10*mm, r['dt'].strftime('%d %b'))
        
        c.setFont('Helvetica-Bold', 34); c.drawCentredString(x + col_w/2, y - 28*mm, r['flight'])
        c.setFont('Helvetica-Bold', 20); c.drawCentredString(x + col_w/2, y - 42*mm, r['dest'])
        t_disp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        c.setFont('Helvetica-Bold', 28); c.drawCentredString(x + col_w/2, y - 56*mm, t_disp)
        
    c.save(); target.seek(0)
    return target

# --- 5. 앱 구동부 ---
st.markdown('<div class="top-left-container"><a href="https://www.flightradar24.com/" target="_blank">Import Raw Text</a><a href="https://www.google.com/" target="_blank">Export Raw Text</a></div>', unsafe_allow_html=True)
st.markdown('<div class="main-title">Simon Park\'nRide\'s<br><span class="sub-title">Flight List Factory</span></div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("⚙️ Settings")
    s_time = st.text_input("Start Time", "05:00")
    e_time = st.text_input("End Time", "04:55")
    label_start = st.number_input("Label Start No", value=1)

uploaded_file = st.file_uploader("Upload Raw Text File", type=['txt'])

if uploaded_file:
    lines = uploaded_file.read().decode("utf-8").splitlines()
    all_recs = parse_raw_lines(lines)
    if all_recs:
        day1 = all_recs[0]['dt'].date()
        start_dt = datetime.combine(day1, datetime.strptime(s_time, '%H:%M').time())
        end_dt = start_dt + timedelta(hours=24)
        filtered = [r for r in all_recs if start_dt <= r['dt'] <= end_dt and r['flight'][:2] in ALLOWED_AIRLINES and r['dest'] not in NZ_DOMESTIC_IATA]
        
        if filtered:
            st.success(f"Processed {len(filtered)} flights")
            c1, c2, c3, c4 = st.columns(4)
            fn = f"List_{start_dt.strftime('%d-%m')}"
            
            c1.download_button("📥 DOCX", build_docx_stream(filtered, start_dt, end_dt), f"{fn}.docx")
            c2.download_button("📄 1-PAGE", build_docx_stream(filtered, start_dt, end_dt, True), f"{fn}_1p.docx")
            c3.download_button("🏷️ PDF", build_labels_stream(filtered, label_start), f"Labels_{fn}.pdf")
            
            excl_list = sorted(list({r['flight'] for r in all_recs if r['flight'][:2] in ALLOWED_AIRLINES and r['dest'] not in NZ_DOMESTIC_IATA}))
            csv = pd.DataFrame(excl_list).to_csv(index=False, header=False).encode('utf-8-sig')
            c4.download_button("📊 EXCL", csv, f"Excl_{fn}.csv", "text/csv")
