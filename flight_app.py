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

# --- UI 설정 ---
st.set_page_config(page_title="Simon Flight Factory", layout="centered")
st.markdown("""
    <style>
    .stApp { background-color: #000000; color: #ffffff; }
    div.stDownloadButton > button {
        background-color: #ffffff !important; color: #000000 !important;
        border-radius: 8px !important; font-weight: 800 !important; width: 100% !important;
    }
    div.stDownloadButton > button:hover { background-color: #60a5fa !important; color: #ffffff !important; }
    </style>
    """, unsafe_allow_html=True)

# --- 파싱 로직 ---
TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s[AP]M)\t([A-Z]{2}\d+[A-Z]?)\s*$")
DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
IATA_IN_PAREns = re.compile(r"\(([^)]+)\)")
ALLOWED_AIRLINES = {"NZ","QF","JQ","CZ","CA","SQ","LA","IE"}
NZ_DOMESTIC = {"AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO","WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"}

def parse_txt(lines):
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

# --- DOCX (1-Page 70% 폭 / 7.5pt 적용) ---
def make_docx(recs, start_dt, end_dt, is_1p=False):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = Inches(0.3)
    f_size = Pt(7.5) if is_1p else Pt(13)
    
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_h = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d %b')}")
    run_h.bold = True; run_h.font.size = Pt(7.5 if is_1p else 15)
    
    table = doc.add_table(rows=0, cols=5); table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblPr = table._element.find(qn('w:tblPr'))
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '3500' if is_1p else '5000') # 3500=70% 폭
    tblW.set(qn('w:type'), 'pct'); tblPr.append(tblW)
    
    for i, r in enumerate(recs):
        row = table.add_row()
        t_short = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        vals = [r['flight'], t_short, r['dest'], "B789", r['reg']]
        for cell, v in zip(row.cells, vals):
            if i % 2 == 1:
                shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9'); cell._tc.get_or_add_tcPr().append(shd)
            p_cell = cell.paragraphs[0]; p_cell.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p_cell.add_run(str(v)); run.font.size = f_size
    
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf

# --- PDF Labels (좌표 전면 재수정 - 겹침 방지) ---
def make_labels(recs, start_num):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    m, g = 15*mm, 6*mm
    cw, rh = (w - 2*m - g)/2, (h - 2*m)/5
    for i, r in enumerate(recs):
        if i > 0 and i % 10 == 0: c.showPage()
        idx = i % 10
        x = m + (idx % 2) * (cw + g)
        y = h - m - (idx // 2) * rh
        
        # 테두리 그리기 (에러 방지를 위해 간단한 방식으로 변경)
        c.setStrokeColorRGB(0.7, 0.7, 0.7)
        c.rect(x, y - rh + 3*mm, cw, rh - 6*mm)
        
        # 텍스트 배치 (y좌표를 각각 다르게 주어 겹침 방지)
        c.setFillColorRGB(0,0,0)
        # 1. 상단: 순번(좌) 및 날짜(우)
        c.setFont("Helvetica-Bold", 12); c.drawString(x + 5*mm, y - 10*mm, str(start_num + i))
        c.setFont("Helvetica", 10); c.drawRightString(x + cw - 5*mm, y - 10*mm, r['date'])
        
        # 2. 중앙 상단: 편명 (가장 크게)
        c.setFont("Helvetica-Bold", 34); c.drawCentredString(x + cw/2, y - 25*mm, r['flight'])
        
        # 3. 중앙 하단: 목적지
        c.setFont("Helvetica-Bold", 20); c.drawCentredString(x + cw/2, y - 40*mm, r['dest'])
        
        # 4. 하단: 시간 (크게)
        t_short = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        c.setFont("Helvetica-Bold", 28); c.drawCentredString(x + cw/2, y - 55*mm, t_short)
        
    c.save(); buf.seek(0)
    return buf

# --- 메인 실행부 ---
st.title("Simon Park'nRide's Flight List Factory")
st.markdown('[✈️ Import Raw Text](https://www.flightradar24.com/) | [🛫 Export Raw Text](https://www.google.com/)')

with st.sidebar:
    s_time = st.text_input("Start Time", "05:00")
    e_time = st.text_input("End Time", "04:55")
    label_start = st.number_input("Label Start No", value=1)

uploaded = st.file_uploader("Upload Raw Text", type=['txt'])

if uploaded:
    lines = uploaded.read().decode("utf-8").splitlines()
    all_recs = parse_txt(lines)
    if all_recs:
        day1 = all_recs[0]['dt'].date()
        start_bound = datetime.combine(day1, datetime.strptime(s_time, '%H:%M').time())
        end_bound = start_bound + timedelta(hours=24)
        filtered = [r for r in all_recs if start_bound <= r['dt'] <= end_bound and r['flight'][:2] in ALLOWED_AIRLINES and r['dest'] not in NZ_DOMESTIC]
        
        st.write(f"✅ Processed {len(filtered)} flights")
        c1, c2, c3, c4 = st.columns(4)
        fn = f"List_{start_bound.strftime('%d-%m')}"
        
        c1.download_button("📥 DOCX", make_docx(filtered, start_bound, end_bound), f"{fn}.docx")
        c2.download_button("📄 1-PAGE", make_docx(filtered, start_bound, end_bound, True), f"{fn}_1p.docx")
        c3.download_button("🏷️ LABELS", make_labels(filtered, label_start), f"Labels_{fn}.pdf")
        
        # 엑셀 에러 해결: 엔진을 명시하지 않고 가장 안전한 CSV/Excel 방식으로 처리
        excl_list = sorted(list({r['flight'] for r in all_recs if r['flight'][:2] in ALLOWED_AIRLINES and r['dest'] not in NZ_DOMESTIC}))
        df_excl = pd.DataFrame(excl_list)
        csv_data = df_excl.to_csv(index=False, header=False).encode('utf-8-sig')
        c4.download_button("📊 EXCL CSV", csv_data, f"Excl_{fn}.csv", "text/csv")
