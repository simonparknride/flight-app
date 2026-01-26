import streamlit as st
import re
import io
import pandas as pd
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# --- 1. UI 설정 및 버튼 스타일 복구 ---
st.set_page_config(page_title="Flight List Factory", layout="centered")

st.markdown("""
    <style>
    .stApp { background-color: #000000; color: #ffffff; }
    [data-testid="stSidebar"] { background-color: #111111 !important; }
    div.stDownloadButton > button {
        background-color: #ffffff !important; color: #000000 !important;
        border: 2px solid #ffffff !important; border-radius: 8px !important;
        height: 48px !important; font-weight: 800 !important; width: 100% !important;
        transition: 0.3s;
    }
    div.stDownloadButton > button:hover {
        background-color: #60a5fa !important; color: #ffffff !important;
        border: 2px solid #60a5fa !important;
    }
    .link-box { margin-bottom: 20px; }
    .link-box a { color: #60a5fa !important; margin-right: 20px; text-decoration: underline; font-size: 1rem; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 데이터 파싱 로직 ---
TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s[AP]M)\t([A-Z]{2}\d+[A-Z]?)\s*$")
DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
IATA_IN_PAREns = re.compile(r"\(([^)]+)\)")
ALLOWED_AIRLINES = {"NZ","QF","JQ","CZ","CA","SQ","LA","IE"}
NZ_DOMESTIC_IATA = {"AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO","WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"}

def parse_raw_text(lines):
    records = []; current_date = None; i = 0
    while i < len(lines):
        line = lines[i].strip()
        if DATE_HEADER.match(line):
            try: current_date = datetime.strptime(line + ' 2026', '%A, %b %d %Y').date()
            except: pass
            i += 1; continue
        m = TIME_LINE.match(line)
        if m and current_date:
            time_str, flight = m.groups()
            dest_line = lines[i+1].strip() if i+1 < len(lines) else ''
            dest_iata = (IATA_IN_PAREns.search(dest_line).group(1) if IATA_IN_PAREns.search(dest_line) else '').upper()
            carrier_line = lines[i+2].strip() if i+2 < len(lines) else ''
            reg = ''; ps = IATA_IN_PAREns.findall(carrier_line)
            if ps: reg = ps[-1].strip()
            try: dt = datetime.strptime(f"{current_date} {time_str}", '%Y-%m-%d %I:%M %p')
            except: dt = None
            records.append({'dt': dt, 'date': current_date.strftime('%d %b'), 'time': time_str, 'flight': flight, 'dest': dest_iata, 'reg': reg, 'type': 'B789'})
            i += 4; continue
        i += 1
    return records

# --- 3. DOCX 생성 (1-Page 70% 폭 + 7.5pt) ---
def build_docx(records, start_dt, end_dt, is_1p=False):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Inches(0.3)
    
    f_size = Pt(7.5) if is_1p else Pt(14)
    
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_h = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d %b')}")
    run_h.bold = True; run_h.font.size = Pt(7.5) if is_1p else Pt(16)
    
    table = doc.add_table(rows=0, cols=5); table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # 테이블 폭 설정 (1-Page는 70% 폭)
    tblPr = table._element.find(qn('w:tblPr'))
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '3500' if is_1p else '5000') # 3500이 약 70%
    tblW.set(qn('w:type'), 'pct')
    tblPr.append(tblW)
    
    for i, r in enumerate(records):
        row = table.add_row()
        t_disp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        vals = [r['flight'], t_disp, r['dest'], r['type'], r['reg']]
        for j, v in enumerate(vals):
            cell = row.cells[j]
            if i % 2 == 1:
                shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9'); cell._tc.get_or_add_tcPr().append(shd)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run(str(v)); run.font.size = f_size
            
    target = io.BytesIO(); doc.save(target); target.seek(0)
    return target

# --- 4. PDF Labels (이미지 속 겹침 완벽 해결) ---
def build_labels(records, start_num):
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
        
        c.setStrokeColorGray(0.8); c.rect(x, y - row_h + 2*mm, col_w, row_h - 4*mm)
        c.setFillColorGray(0)
        c.setFont('Helvetica-Bold', 12); c.drawString(x + 5*mm, y - 10*mm, str(start_num + i))
        c.setFont('Helvetica', 10); c.drawRightString(x + col_w - 5*mm, y - 10*mm, r['date'])
        
        # 텍스트 간격 재조정 (겹침 방지)
        c.setFont('Helvetica-Bold', 36); c.drawCentredString(x + col_w/2, y - 28*mm, r['flight'])
        c.setFont('Helvetica-Bold', 22); c.drawCentredString(x + col_w/2, y - 42*mm, r['dest'])
        t_disp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        c.setFont('Helvetica-Bold', 30); c.drawCentredString(x + col_w/2, y - 57*mm, t_disp)
    c.save(); target.seek(0)
    return target

# --- 5. 앱 실행 및 버튼 배치 ---
st.markdown('<div class="link-box"><a href="https://www.flightradar24.com/data/airports/akl/arrivals" target="_blank">✈️ Import Raw Text</a><a href="https://www.flightradar24.com/data/airports/akl/departures" target="_blank">🛫 Export Raw Text</a></div>', unsafe_allow_html=True)
st.title("Simon Park'nRide's Flight List Factory")

with st.sidebar:
    st.header("⚙️ SETTINGS")
    s_time = st.text_input("Start Time", "05:00")
    e_time = st.text_input("End Time", "04:55")
    label_start = st.number_input("Label Start No", value=1)

uploaded_file = st.file_uploader("Upload Raw Text File", type=['txt'])

if uploaded_file:
    lines = uploaded_file.read().decode("utf-8").splitlines()
    all_recs = parse_raw_text(lines)
    if all_recs:
        day1 = all_recs[0]['dt'].date()
        s_dt = datetime.combine(day1, datetime.strptime(s_time, '%H:%M').time())
        e_dt = s_dt + timedelta(hours=24)
        filtered = [r for r in all_recs if s_dt <= r['dt'] <= e_dt and r['flight'][:2] in ALLOWED_AIRLINES and r['dest'] not in NZ_DOMESTIC_IATA]
        
        st.success(f"Processed {len(filtered)} Flights")
        
        c1, c2, c3, c4 = st.columns(4)
        fn = f"List_{s_dt.strftime('%d-%m')}"
        
        c1.download_button("📥 DOCX", build_docx(filtered, s_dt, e_dt), f"{fn}.docx")
        c2.download_button("📄 1-PAGE", build_docx(filtered, s_dt, e_dt, True), f"{fn}_1p.docx")
        c3.download_button("🏷️ LABELS", build_labels(filtered, label_start), f"Labels_{fn}.pdf")
        
        # EXCL Excel (에러 방지용 CSV 방식 우회 또는 표준 엑셀)
        excl_list = sorted(list({r['flight'] for r in all_recs if r['flight'][:2] in ALLOWED_AIRLINES and r['dest'] not in NZ_DOMESTIC_IATA}))
        df_excl = pd.DataFrame(excl_list)
        excel_io = io.BytesIO()
        df_excl.to_csv(excel_io, index=False, header=False, encoding='utf-8-sig')
        c4.download_button("📊 EXCL CSV", excel_io.getvalue(), f"Excl_{fn}.csv")
