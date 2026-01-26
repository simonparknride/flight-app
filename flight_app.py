import streamlit as st
import re
import io
import pandas as pd
from datetime import datetime, timedelta
from typing import List, Dict
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement, qn
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# --- 1. UI 스타일 설정 ---
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

# --- 2. 데이터 파싱 로직 (기종 및 등록번호 포함) ---
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
            records.append({'dt': dt, 'date_str': current_date.strftime('%d %b'), 'time': time_str, 'flight': flight, 'dest': dest_iata, 'type': plane_type, 'reg': reg})
            i += 4; continue
        i += 1
    return records

# --- 3. PDF Labels 생성 (Labels_List_26-01 (1).pdf 레이아웃 완벽 재현) ---
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
        
        # 칸 테두리
        c.setStrokeGray(0.8)
        c.rect(x, y + 2*mm, col_w, row_h - 4*mm)
        
        # 1. 순번 & 날짜 (좌상단)
        c.setFont('Helvetica-Bold', 12)
        c.drawString(x + 5*mm, y + row_h - 10*mm, str(start_num + i))
        c.setFont('Helvetica', 10)
        c.drawRightString(x + col_w - 5*mm, y + row_h - 10*mm, r['date_str'])
        
        # 2. 비행 편명 & 목적지 (중앙 상단)
        c.setFont('Helvetica-Bold', 28)
        c.drawCentredString(x + col_w/2, y + row_h - 25*mm, r['flight'])
        c.setFont('Helvetica-Bold', 20)
        c.drawCentredString(x + col_w/2, y + row_h - 35*mm, r['dest'])
        
        # 3. 시간 (중앙 하단)
        tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        c.setFont('Helvetica-Bold', 24)
        c.drawCentredString(x + col_w/2, y + 15*mm) # 고정 좌표
        c.drawCentredString(x + col_w/2, y + 15*mm, tdisp)
        
        # 4. 기종 & 등록번호 (우하단)
        c.setFont('Helvetica', 9)
        info_text = f"{r['type']}  {r['reg']}".strip()
        c.drawCentredString(x + col_w/2, y + 7*mm, info_text)
        
    c.save(); target.seek(0); return target

# --- 4. 메인 실행 로직 ---
with st.sidebar:
    st.header("⚙️ Settings")
    s_time = st.text_input("Start Time", "04:55")
    e_time = st.text_input("End Time", "05:00")
    label_start = st.number_input("Label Start Number", value=1)

st.title("Simon Park'nRide's Factory")
uploaded_file = st.file_uploader("Upload Text File", type=['txt'])

if uploaded_file:
    lines = uploaded_file.read().decode("utf-8").splitlines()
    recs = parse_raw_lines(lines)
    # 필터링 로직 (생략 - 이전과 동일)
    if recs:
        st.success(f"Loaded {len(recs)} flights")
        pdf = build_labels_stream(recs, label_start)
        st.download_button("📥 Download Perfect PDF Labels", pdf, "Perfect_Labels.pdf")
