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

# --- 1. UI 설정 및 버튼 호버 효과 ---
st.set_page_config(page_title="Flight List Factory", layout="centered")
st.markdown("""
    <style>
    .stApp { background-color: #000000; color: #ffffff; }
    [data-testid="stSidebar"] { background-color: #111111 !important; }
    
    /* 버튼 호버 및 디자인 복구 */
    div.stDownloadButton > button {
        background-color: #ffffff !important; color: #000000 !important;
        border: 2px solid #ffffff !important; border-radius: 8px !important;
        height: 45px !important; font-weight: 800 !important; width: 100% !important;
    }
    div.stDownloadButton > button:hover {
        background-color: #60a5fa !important; color: #ffffff !important;
        border: 2px solid #60a5fa !important;
    }
    
    /* 상단 링크 복구 */
    .link-box { margin-bottom: 25px; }
    .link-box a { color: #60a5fa !important; margin-right: 20px; text-decoration: underline; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 파싱 로직 ---
TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s[AP]M)\t([A-Z]{2}\d+[A-Z]?)\s*$")
DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
IATA_IN_PAREns = re.compile(r"\(([^)]+)\)")
ALLOWED_AIRLINES = {"NZ","QF","JQ","CZ","CA","SQ","LA","IE"}
NZ_DOMESTIC_IATA = {"AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO","WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"}

def parse_data(lines):
    recs = []; cur_date = None; i = 0
    while i < len(lines):
        l = lines[i].strip()
        if DATE_HEADER.match(l):
            try: cur_date = datetime.strptime(l + ' 2026', '%A, %b %d %Y').date()
            except: pass
            i += 1; continue
        m = TIME_LINE.match(l)
        if m and cur_date:
            time_str, flight = m.groups()
            dest_l = lines[i+1].strip() if i+1 < len(lines) else ''
            dest = (IATA_IN_PAREns.search(dest_l).group(1) if IATA_IN_PAREns.search(dest_l) else '').upper()
            c_line = lines[i+2].strip() if i+2 < len(lines) else ''
            reg = ''; ps = IATA_IN_PAREns.findall(c_line)
            if ps: reg = ps[-1].strip()
            try: dt = datetime.strptime(f"{cur_date} {time_str}", '%Y-%m-%d %I:%M %p')
            except: dt = None
            recs.append({'dt': dt, 'date': cur_date.strftime('%d %b'), 'time': time_str, 'flight': flight, 'dest': dest, 'reg': reg})
            i += 4; continue
        i += 1
    return recs

# --- 3. DOCX (1-Page 70% 폭 + 7.5pt 적용) ---
def make_docx(recs, start_dt, end_dt, is_1p=False):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = Inches(0.3)
    
    # 1-Page일 경우 7.5pt, 아닐 경우 14pt
    f_size = Pt(7.5) if is_1p else Pt(14)
    
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    header_text = f"{start_dt.strftime('%d')}-{end_dt.strftime('%d %b')}"
    run_h = p.add_run(header_text)
    run_h.bold = True
    run_h.font.size = Pt(7.5) if is_1p else Pt(16)
    
    table = doc.add_table(rows=0, cols=5); table.alignment = WD_TABLE_ALIGNMENT.CENTER
    # 테이블 폭 조정 (1-Page는 70% 폭)
    tblPr = table._element.find(qn('w:tblPr'))
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '3500' if is_1p else '5000') # 3500/5000 = 70%
    tblW.set(qn('w:type'), 'pct')
    tblPr.append(tblW)
    
    for i, r in enumerate(recs):
        row = table.add_row()
        t_short = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        vals = [r['flight'], t_short, r['dest'], "B789", r['reg']] # 기종은 파싱값 대체 가능
        for cell, v in zip(row.cells, vals):
            if i % 2 == 1:
                shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9'); cell._tc.get_or_add_tcPr().append(shd)
            run = cell.paragraphs[0].add_run(str(v))
            run.font.size = f_size
    
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf

# --- 4. PDF Labels (이미지 속 겹침 현상 해결 좌표) ---
def make_labels(recs, start_num):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    m, g = 12*mm, 6*mm
    cw, rh = (w - 2*m - g)/2, (h - 2*m)/5
    for i, r in enumerate(recs):
        if i > 0 and i % 10 == 0: c.showPage()
        idx = i % 10
        x = m + (idx % 2) * (cw + g)
        y = h - m - (idx // 2) * rh
        c.setStrokeColorGray(0.8); c.rect(x, y - rh + 2*mm, cw, rh - 4*mm, stroke=1)
        c.setFillColorGray(0)
        c.setFont("Helvetica-Bold", 12); c.drawString(x + 5*mm, y - 10*mm, str(start_num + i))
        c.setFont("Helvetica", 10); c.drawRightString(x + cw - 5*mm, y - 10*mm, r['date'])
        # 겹치지 않게 간격 조정
        c.setFont("Helvetica-Bold", 36); c.drawCentredString(x + cw/2, y - 28*mm, r['flight'])
        c.setFont("Helvetica-Bold", 20); c.drawCentredString(x + cw/2, y - 42*mm, r['dest'])
        t_short = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        c.setFont("Helvetica-Bold", 30); c.drawCentredString(x + cw/2, y - 56*mm, t_short)
    c.save(); buf.seek(0)
    return buf

# --- 5. 메인 레이아웃 ---
st.markdown('<div class="link-box"><a href="https://www.flightradar24.com/data/airports/akl/arrivals" target="_blank">✈️ Import Raw Text</a><a href="https://www.flightradar24.com/data/airports/akl/departures" target="_blank">🛫 Export Raw Text</a></div>', unsafe_allow_html=True)
st.title("Simon Park'nRide's Flight List Factory")

with st.sidebar:
    st.header("⚙️ SETTINGS")
    s_time = st.text_input("Start Time", "05:00")
    e_time = st.text_input("End Time", "04:55")
    label_no = st.number_input("Label Start No", value=1)

uploaded = st.file_uploader("Upload Text File", type=['txt'])

if uploaded:
    lines = uploaded.read().decode("utf-8").splitlines()
    all_recs = parse_data(lines)
    if all_recs:
        day1 = all_recs[0]['dt'].date()
        start_bound = datetime.combine(day1, datetime.strptime(s_time, '%H:%M').time())
        end_bound = start_bound + timedelta(hours=24)
        filtered = [r for r in all_recs if start_bound <= r['dt'] <= end_bound and r['flight'][:2] in ALLOWED_AIRLINES and r['dest'] not in NZ_DOMESTIC_IATA]
        
        st.success(f"Processed {len(filtered)} flights")
        
        # 버튼 4개 배치
        cols = st.columns(4)
        fn = f"List_{start_bound.strftime('%d-%m')}"
        
        cols[0].download_button("📥 DOCX", make_docx(filtered, start_bound, end_bound), f"{fn}.docx")
        cols[1].download_button("📄 1-PAGE", make_docx(filtered, start_bound, end_bound, True), f"{fn}_1p.docx")
        cols[2].download_button("🏷️ LABELS", make_labels(filtered, label_no), f"Labels_{fn}.pdf")
        
        # EXCL EXCEL (24시간 모든 국제선 편명)
        excl_data = sorted(list({r['flight'] for r in all_recs if r['flight'][:2] in ALLOWED_AIRLINES and r['dest'] not in NZ_DOMESTIC_IATA}))
        df_excl = pd.DataFrame(excl_data)
        csv = df_excl.to_csv(index=False, header=False).encode('utf-8-sig')
        cols[3].download_button("📊 EXCL (CSV)", csv, f"Excl_{fn}.csv", "text/csv")

        st.divider()
        st.table(pd.DataFrame(filtered)[['flight', 'time', 'dest', 'reg']])
