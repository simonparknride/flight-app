import streamlit as st
import re
import io
import pandas as pd
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# --- 1. UI 디자인 (사진속 깨짐 현상 강제 수정) ---
st.set_page_config(page_title="Flight List Factory", layout="wide")

st.markdown("""
    <style>
    /* 전체 배경 검정 및 기본 폰트 설정 */
    .stApp { background-color: #000000; color: #ffffff; }
    
    /* 상단 영역: 사진처럼 구석에 박히지 않게 중앙 정렬 및 여백 확보 */
    .header-box {
        padding: 30px;
        text-align: center;
        border-bottom: 2px solid #333;
        margin-bottom: 40px;
    }
    .header-box a {
        color: #60a5fa !important;
        text-decoration: underline;
        font-size: 1.2rem;
        margin: 0 15px;
        font-weight: bold;
    }
    .main-title { font-size: 3.5rem; font-weight: 900; color: #ffffff; margin-top: 15px; }
    .sub-title { font-size: 2.8rem; font-weight: 300; color: #60a5fa; }

    /* 버튼 스타일 강제 주입: 흰색 배경, 검정 글자 */
    div.stDownloadButton > button {
        background-color: #ffffff !important;
        color: #000000 !important;
        border: 2px solid #ffffff !important;
        border-radius: 10px !important;
        height: 55px !important;
        width: 100% !important;
        font-weight: 900 !important;
        font-size: 1.1rem !important;
        box-shadow: 0 4px 6px rgba(0,0,0,0.3);
    }
    
    /* 마우스 호버 시: 파란색 배경 */
    div.stDownloadButton > button:hover {
        background-color: #60a5fa !important;
        border-color: #60a5fa !important;
        color: #ffffff !important;
    }

    /* 사이드바 및 입력창 색상 고정 */
    [data-testid="stSidebar"] { background-color: #111111 !important; }
    label { color: #ffffff !important; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 데이터 파싱 ---
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

# --- 3. PDF Labels (가이드와 100% 동일한 좌표) ---
def build_labels_stream(records, start_num):
    target = io.BytesIO()
    c = canvas.Canvas(target, pagesize=A4)
    w, h = A4
    margin_x, margin_y = 10*mm, 15*mm
    gutter = 6*mm
    col_w = (w - 2*margin_x - gutter) / 2
    row_h = (h - 2*margin_y) / 5
    
    for i, r in enumerate(records):
        if i > 0 and i % 10 == 0: c.showPage()
        idx = i % 10
        col, row = idx % 2, idx // 2
        x = margin_x + col * (col_w + gutter)
        y_top = h - margin_y - row * row_h
        
        c.setStrokeColorGray(0.8); c.setLineWidth(0.1)
        c.rect(x, y_top - row_h + 2*mm, col_w, row_h - 4*mm)
        
        c.setFillColorGray(0)
        # 1. 순번 (Top Left) / 날짜 (Top Right)
        c.setFont('Helvetica-Bold', 12); c.drawString(x + 5*mm, y_top - 12*mm, str(start_num + i))
        c.setFont('Helvetica', 11); c.drawRightString(x + col_w - 5*mm, y_top - 12*mm, r['date_label'])
        
        # 2. 편명 (가장 크게)
        c.setFont('Helvetica-Bold', 38); c.drawCentredString(x + col_w/2, y_top - 28*mm, r['flight'])
        
        # 3. 목적지
        c.setFont('Helvetica-Bold', 24); c.drawCentredString(x + col_w/2, y_top - 42*mm, r['dest'])
        
        # 4. 시간 (강조)
        tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        c.setFont('Helvetica-Bold', 32); c.drawCentredString(x + col_w/2, y_top - 56*mm, tdisp)
        
        # 5. 기종 & 등록번호 (맨 아래)
        c.setFont('Helvetica', 10); c.drawCentredString(x + col_w/2, y_top - row_h + 8*mm, f"{r['type']}  {r['reg']}")
        
    c.save(); target.seek(0); return target

# --- 4. DOCX 생성 (7.5pt 1-Page 포함) ---
def build_docx_stream(records, start_dt, end_dt, is_1p=False):
    doc = Document()
    sec = doc.sections[0]
    if is_1p: sec.top_margin = sec.bottom_margin = Inches(0.4); sec.left_margin = sec.right_margin = Inches(0.5)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}")
    run.bold = True; run.font.size = Pt(7.5 if is_1p else 16)
    table = doc.add_table(rows=0, cols=5); table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, r in enumerate(records):
        row = table.add_row()
        vals = [r['flight'], datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M'), r['dest'], r['type'], r['reg']]
        for j, v in enumerate(vals):
            cell = row.cells[j]
            if i % 2 == 1:
                shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9'); cell._tc.get_or_add_tcPr().append(shd)
            run_c = cell.paragraphs[0].add_run(str(v)); run_c.font.size = Pt(7.5 if is_1p else 14)
    target = io.BytesIO(); doc.save(target); target.seek(0); return target

# --- 5. 화면 출력 (중앙 정렬 레이아웃) ---
st.markdown(f"""
    <div class="header-box">
        <a href="https://www.flightradar24.com/data/airports/akl/arrivals" target="_blank">✈️ Import Raw Text</a>
        <a href="https://www.flightradar24.com/data/airports/akl/departures" target="_blank">🛫 Export Raw Text</a>
        <div class="main-title">Simon Park'nRide's</div>
        <div class="sub-title">Flight List Factory</div>
    </div>
    """, unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### ⚙️ SETTINGS")
    s_val = st.text_input("Start Time", "04:55")
    e_val = st.text_input("End Time", "05:00")
    label_start = st.number_input("Label No Start", value=1)

uploaded_file = st.file_uploader("Upload Raw Text File", type=['txt'])

if uploaded_file:
    lines = uploaded_file.read().decode("utf-8").splitlines()
    recs = parse_raw_lines(lines)
    if recs:
        dates = sorted({r['dt'].date() for r in recs if r.get('dt')})
        if dates:
            day1 = dates[0]; day2 = dates[1] if len(dates) >= 2 else (day1 + timedelta(days=1))
            start_dt = datetime.combine(day1, datetime.strptime(s_val, '%H:%M').time())
            end_dt = datetime.combine(day2, datetime.strptime(e_val, '%H:%M').time())
            filtered = [r for r in recs if r.get('dt') and (start_dt <= r['dt'] <= end_dt)]
            
            if filtered:
                st.success(f"Successfully processed {len(filtered)} flights!")
                col_btn1, col_btn2, col_btn3, col_btn4 = st.columns(4)
                fn = f"List_{start_dt.strftime('%d-%m')}"
                with col_btn1: st.download_button("📥 DOCX (Orig)", build_docx_stream(filtered, start_dt, end_dt), f"{fn}_orig.docx")
                with col_btn2: st.download_button("📄 DOCX (1-Page)", build_docx_stream(filtered, start_dt, end_dt, True), f"{fn}_1p.docx")
                with col_btn3: st.download_button("🏷️ PDF LABELS", build_labels_stream(filtered, label_start), f"Labels_{fn}.pdf")
                with col_btn4: st.download_button("📊 CSV FLIGHTS", pd.DataFrame([r['flight'] for r in filtered]).to_csv(index=False, header=False).encode('utf-8-sig'), f"{fn}.csv")
                
                st.write("---")
                st.table([{'No': label_start+i, 'Flight': r['flight'], 'Time': r['time'], 'Dest': r['dest'], 'Reg': r['reg']} for i, r in enumerate(filtered)])
