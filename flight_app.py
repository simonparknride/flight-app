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
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# --- 1. UI 및 버튼 스타일 (완벽 고정) ---
st.set_page_config(page_title="Flight List Factory", layout="centered")
st.markdown("""
    <style>
    .stApp { background-color: #000000; }
    [data-testid="stSidebar"] { background-color: #111111 !important; }
    .stMarkdown, p, h1, h2, h3, label { color: #ffffff !important; }
    
    /* 버튼 스타일 */
    div.stDownloadButton > button {
        background-color: #ffffff !important; color: #000000 !important;
        border: 2px solid #ffffff !important; border-radius: 8px !important;
        padding: 0.6rem 1rem !important; font-weight: 800 !important;
        width: 100% !important; transition: all 0.3s ease;
    }
    div.stDownloadButton > button p { color: #000000 !important; }
    div.stDownloadButton > button:hover {
        background-color: #60a5fa !important; color: #ffffff !important;
        border: 2px solid #60a5fa !important;
    }
    div.stDownloadButton > button:hover p { color: #ffffff !important; }

    .top-left-container { text-align: left; padding-top: 10px; margin-bottom: 20px; }
    .top-left-container a { font-size: 1.1rem; color: #ffffff !important; text-decoration: underline; display: block; margin-bottom: 5px;}
    .main-title { font-size: 3rem; font-weight: 800; color: #ffffff; line-height: 1.1; margin-bottom: 0.5rem; }
    .sub-title { font-size: 2.5rem; font-weight: 400; color: #60a5fa; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 파싱 로직 ---
TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s[AP]M)\t([A-Z]{2}\d+[A-Z]?)\s*$")
DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
IATA_IN_PAREns = re.compile(r"\(([^)]+)\)")
PLANE_TYPES = ['A21N','A20N','A320','32Q','320','73H','737','74Y','77W','B77W','789','B789','359','A359','332','A332','AT76','388','333','A333','330']
PLANE_TYPE_PATTERN = re.compile(r"\b(" + "|".join(PLANE_TYPES) + r")\b", re.IGNORECASE)

def parse_raw_lines(lines):
    records = []; current_date = None; i = 0
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
            reg = ''; parens = IATA_IN_PAREns.findall(carrier_line)
            if parens: reg = parens[-1].strip()
            try: dt = datetime.strptime(f"{current_date} {time_str}", '%Y-%m-%d %I:%M %p')
            except: dt = None
            records.append({'dt': dt, 'date_label': current_date.strftime('%d %b'), 'time': time_str, 'flight': flight, 'dest': dest_iata, 'type': (ptype.group(1).upper() if ptype else ''), 'reg': reg})
            i += 4; continue
        i += 1
    return records

# --- 3. PDF Labels (가이드 PDF와 1:1 매칭 좌표) ---
def build_labels_stream(records, start_num):
    target = io.BytesIO()
    c = canvas.Canvas(target, pagesize=A4)
    w, h = A4
    margin_x, margin_y = 10*mm, 15*mm
    gutter = 5*mm
    col_w = (w - 2*margin_x - gutter) / 2
    row_h = (h - 2*margin_y) / 5
    
    for i, r in enumerate(records):
        if i > 0 and i % 10 == 0: c.showPage()
        idx = i % 10
        col, row = idx % 2, idx // 2
        x = margin_x + col * (col_w + gutter)
        y_top = h - margin_y - row * row_h
        
        # 1. 테두리
        c.setStrokeColorGray(0.8); c.setLineWidth(0.1)
        c.rect(x, y_top - row_h + 2*mm, col_w, row_h - 4*mm)
        
        # 2. 텍스트 배치 (가이드와 동일한 세로 위치 강제 지정)
        c.setFillColorGray(0)
        
        # 순번 & 날짜 (Top)
        c.setFont('Helvetica-Bold', 12); c.drawString(x + 4*mm, y_top - 10*mm, str(start_num + i))
        c.setFont('Helvetica', 10); c.drawRightString(x + col_w - 4*mm, y_top - 10*mm, r['date_label'])
        
        # 편명 (Center Upper)
        c.setFont('Helvetica-Bold', 34)
        c.drawCentredString(x + col_w/2, y_top - 25*mm, r['flight'])
        
        # 목적지 (Center Middle)
        c.setFont('Helvetica-Bold', 24)
        c.drawCentredString(x + col_w/2, y_top - 36*mm, r['dest'])
        
        # 시간 (Center Lower)
        tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        c.setFont('Helvetica-Bold', 30)
        c.drawCentredString(x + col_w/2, y_top - 48*mm, tdisp)
        
        # 기종 & 등록번호 (Bottom)
        c.setFont('Helvetica', 10)
        info = f"{r['type']}  {r['reg']}".strip()
        c.drawCentredString(x + col_w/2, y_top - row_h + 8*mm, info)
        
    c.save(); target.seek(0); return target

# --- 4. DOCX 생성 ---
def build_docx_stream(records, start_dt, end_dt, is_1p=False):
    doc = Document()
    sec = doc.sections[0]
    if is_1p: 
        sec.top_margin = sec.bottom_margin = Inches(0.4)
        sec.left_margin = sec.right_margin = Inches(0.5)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}")
    run.bold = True; run.font.size = Pt(7.5 if is_1p else 16)
    
    table = doc.add_table(rows=0, cols=5); table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, r in enumerate(records):
        row = table.add_row()
        tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        for j, v in enumerate([r['flight'], tdisp, r['dest'], r['type'], r['reg']]):
            cell = row.cells[j]
            if i % 2 == 1:
                shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9'); cell._tc.get_or_add_tcPr().append(shd)
            run_c = cell.paragraphs[0].add_run(str(v)); run_c.font.size = Pt(7.5 if is_1p else 14)
    target = io.BytesIO(); doc.save(target); target.seek(0); return target

# --- 5. 화면 출력 ---
st.markdown('<div class="top-left-container"><a href="https://www.flightradar24.com/data/airports/akl/arrivals" target="_blank">Import Raw Text</a><a href="https://www.flightradar24.com/data/airports/akl/departures" target="_blank">Export Raw Text</a></div>', unsafe_allow_html=True)
st.markdown('<div class="main-title">Simon Park\'nRide\'s<br><span class="sub-title">Flight List Factory</span></div>', unsafe_allow_html=True)

with st.sidebar:
    s_val = st.text_input("Start Time", "04:55")
    e_val = st.text_input("End Time", "05:00")
    label_start = st.number_input("Label Start No", value=1)

uploaded_file = st.file_uploader("Upload Raw Text", type=['txt'])

if uploaded_file:
    lines = uploaded_file.read().decode("utf-8").splitlines()
    recs = parse_raw_lines(lines)
    if recs:
        dates = sorted({r['dt'].date() for r in recs if r.get('dt')})
        if dates:
            day1 = dates[0]
            day2 = dates[1] if len(dates) >= 2 else (day1 + timedelta(days=1))
            start_dt = datetime.combine(day1, datetime.strptime(s_val, '%H:%M').time())
            end_dt = datetime.combine(day2, datetime.strptime(e_val, '%H:%M').time())
            filtered = [r for r in recs if r.get('dt') and (start_dt <= r['dt'] <= end_dt)]
            
            if filtered:
                st.success(f"Processed {len(filtered)} flights")
                c1, c2, c3, c4 = st.columns(4)
                fn = f"List_{start_dt.strftime('%d-%m')}"
                c1.download_button("📥 DOCX (Orig)", build_docx_stream(filtered, start_dt, end_dt), f"{fn}_orig.docx")
                c2.download_button("📄 DOCX (1-Page)", build_docx_stream(filtered, start_dt, end_dt, True), f"{fn}_1p.docx")
                c3.download_button("📥 PDF Labels", build_labels_stream(filtered, label_start), f"Labels_{fn}.pdf")
                c4.download_button("📊 Only Flights", pd.DataFrame([r['flight'] for r in filtered]).to_csv(index=False, header=False).encode('utf-8-sig'), f"{fn}.csv")
                
                st.table([{'No': label_start+i, 'Flight': r['flight'], 'Time': r['time'], 'Dest': r['dest']} for i, r in enumerate(filtered)])
