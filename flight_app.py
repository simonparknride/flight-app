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

# --- 1. UI 및 버튼 스타일 ---
st.set_page_config(page_title="Flight List Factory", layout="centered")

st.markdown("""
    <style>
    .stApp { background-color: #000000; }
    [data-testid="stSidebar"] { background-color: #111111 !important; }
    .stMarkdown, p, h1, h2, h3, label { color: #ffffff !important; }
    
    div.stDownloadButton > button {
        background-color: #ffffff !important;
        border: 2px solid #ffffff !important;
        border-radius: 8px !important;
        height: 3.5rem !important;
        width: 100% !important;
    }
    div.stDownloadButton > button div p {
        color: #000000 !important;
        font-weight: 800 !important;
        font-size: 1rem !important;
    }
    div.stDownloadButton > button:hover {
        background-color: #60a5fa !important;
        border: 2px solid #60a5fa !important;
    }
    div.stDownloadButton > button:hover div p {
        color: #ffffff !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 파싱 로직 (기종 데이터 추출 방식 전면 개편) ---
TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s[AP]M)\t([A-Z]{2}\d+[A-Z]?)\s*$")
DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
IATA_IN_PAREns = re.compile(r"\(([^)]+)\)")
ALLOWED_AIRLINES = {"NZ","QF","JQ","CZ","CA","SQ","LA","IE"}
NZ_DOMESTIC_IATA = {"AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO","WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"}

def parse_raw_lines(lines: List[str]) -> List[Dict]:
    recs = []; cur_date = None; i = 0
    while i < len(lines):
        line = lines[i].strip()
        if DATE_HEADER.match(line):
            try: cur_date = datetime.strptime(line + ' 2026', '%A, %b %d %Y').date()
            except: cur_date = None
            i += 1; continue
        m = TIME_LINE.match(line)
        if m and cur_date:
            time_str, flt = m.groups()
            dest_line = lines[i+1].strip() if i+1 < len(lines) else ''
            m2 = IATA_IN_PAREns.search(dest_line)
            dest = (m2.group(1).strip() if m2 else '').upper()
            
            # 기종(Type) 파싱: 세 번째 줄에서 괄호 앞부분의 텍스트를 기종으로 인식
            carrier_line = lines[i+2].strip() if i+2 < len(lines) else ''
            flt_type = "B789" # 기본값
            if "(" in carrier_line:
                # 괄호 전까지의 문자열을 가져옴 (예: Boeing 787-9 -> B789, Airbus A320 -> A320)
                raw_type = carrier_line.split("(")[0].strip()
                if "787-9" in raw_type: flt_type = "B789"
                elif "777-300" in raw_type: flt_type = "B77W"
                elif "737-800" in raw_type: flt_type = "B738"
                elif "A321" in raw_type: flt_type = "A321"
                elif "A320" in raw_type: flt_type = "A320"
                else: flt_type = raw_type # 매핑되지 않은 경우 원본 텍스트 사용
            
            # 등록번호(Reg) 파싱
            reg = ''; ps = IATA_IN_PAREns.findall(carrier_line)
            if ps: reg = ps[-1].strip()
            
            try: dt = datetime.strptime(f"{cur_date} {time_str}", '%Y-%m-%d %I:%M %p')
            except: dt = None
            recs.append({'dt': dt, 'time': time_str, 'flight': flt, 'dest': dest, 'reg': reg, 'type': flt_type})
            i += 4; continue
        i += 1
    return recs

# --- 3. DOCX 생성 (폰트 14pt 및 폭 조정) ---
def build_docx(recs, start_dt, is_1p=False):
    doc = Document()
    f_name = 'Air New Zealand Sans'
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = Inches(0.2)
    sec.left_margin = sec.right_margin = Inches(0.5)

    table = doc.add_table(rows=0, cols=6)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # 테이블 폭을 보내주신 (4).docx 파일 스타일로 좁게 설정
    tblPr = table._element.find(qn('w:tblPr'))
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '4200') 
    tblW.set(qn('w:type'), 'pct'); tblPr.append(tblW)

    last_date_str = ""
    for i, r in enumerate(recs):
        row = table.add_row()
        tr = row._tr
        trPr = tr.get_or_add_trPr()
        trHeight = OxmlElement('w:trHeight')
        # 폰트가 커짐에 따라 행 높이도 조정
        trHeight.set(qn('w:val'), '240' if is_1p else '360')
        trHeight.set(qn('w:hRule'), 'atLeast')
        trPr.append(trHeight)

        current_date_str = r['dt'].strftime('%d %b')
        display_date = current_date_str if current_date_str != last_date_str else ""
        last_date_str = current_date_str

        t_short = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        # 실제 파싱된 r['type'] 반영
        vals = [display_date, r['flight'], t_short, r['dest'], r['type'], r['reg']]
        
        for j, v in enumerate(vals):
            cell = row.cells[j]
            if i % 2 == 1:
                shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9'); cell._tc.get_or_add_tcPr().append(shd)
            
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.space_before = para.paragraph_format.space_after = Pt(0)
            para.paragraph_format.line_spacing = 1.0
            
            run = para.add_run(str(v))
            run.font.name = f_name
            if j == 0: run.bold = True
            # [수정] 폰트 사이즈를 14pt로 변경 (1-PAGE는 가독성을 위해 9pt로 설정)
            run.font.size = Pt(9.0 if is_1p else 14.0)
            
            rPr = run._element.get_or_add_rPr()
            rFonts = OxmlElement('w:rFonts'); rFonts.set(qn('w:ascii'), f_name); rFonts.set(qn('w:hAnsi'), f_name)
            rPr.append(rFonts)
            
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf

# --- 4. PDF Labels (기존 로직 유지) ---
def build_labels(recs, start_num):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    margin, gutter = 15*mm, 6*mm
    col_w, row_h = (w - 2*margin - gutter) / 2, (h - 2*margin) / 5
    for i, r in enumerate(recs):
        if i > 0 and i % 10 == 0: c.showPage()
        idx = i % 10
        x = margin + (idx % 2) * (col_w + gutter)
        y = h - margin - (idx // 2) * row_h
        c.setStrokeGray(0.3); c.setLineWidth(0.2); c.rect(x, y - row_h + 2*mm, col_w, row_h - 4*mm)
        c.setLineWidth(0.5); c.rect(x + 3*mm, y - 11*mm, 8*mm, 8*mm)
        c.setFillColorRGB(0,0,0)
        c.setFont('Helvetica-Bold', 14); c.drawCentredString(x + 7*mm, y - 9.5*mm, str(start_num + i))
        c.setFont('Helvetica-Bold', 18); c.drawRightString(x + col_w - 4*mm, y - 10*mm, r['dt'].strftime('%d %b'))
        c.setFont('Helvetica-Bold', 38); c.drawString(x + 15*mm, y - 22*mm, r['flight'])
        c.setFont('Helvetica-Bold', 23); c.drawString(x + 15*mm, y - 34*mm, r['dest'])
        t_disp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        c.setFont('Helvetica-Bold', 29); c.drawString(x + 15*mm, y - 48*mm, t_disp)
        c.setFont('Helvetica', 13); c.drawRightString(x + col_w - 6*mm, y - row_h + 10*mm, r['reg'])
    c.save(); buf.seek(0)
    return buf

# --- 5. 실행부 ---
st.markdown('<div class="main-title">Simon Park\'nRide\'s Flight List Factory</div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("⚙️ Settings")
    s_time = st.text_input("Start Time", "04:55")
    e_time = st.text_input("End Time", "05:00")
    label_start = st.number_input("Label Start No", value=1)

uploaded = st.file_uploader("Upload Raw Text File", type=['txt'])

if uploaded:
    lines = uploaded.read().decode("utf-8").splitlines()
    all_recs = parse_raw_lines(lines)
    if all_recs:
        day1 = all_recs[0]['dt'].date()
        cur_s = datetime.combine(day1, datetime.strptime(s_time, '%H:%M').time())
        cur_e = cur_s + timedelta(hours=24)
        
        filtered = [r for r in all_recs if cur_s <= r['dt'] < cur_e and r['flight'][:2] in ALLOWED_AIRLINES and r['dest'] not in NZ_DOMESTIC_IATA]
        
        excl_s = datetime.combine(day1, datetime.min.time())
        excl_e = excl_s + timedelta(hours=24)
        excl_data = sorted(list({r['flight'] for r in all_recs if excl_s <= r['dt'] < excl_e and r['flight'][:2] in ALLOWED_AIRLINES and r['dest'] not in NZ_DOMESTIC_IATA}))

        if filtered:
            st.success(f"Processed {len(filtered)} flights")
            c1, c2, c3, c4 = st.columns(4)
            fn = f"List_{cur_s.strftime('%d-%m')}"
            
            c1.download_button("📥 DOCX", build_docx(filtered, cur_s), f"{fn}.docx")
            c2.download_button("📄 1-PAGE", build_docx(filtered, cur_s, True), f"{fn}_1p.docx")
            c3.download_button("🏷️ LABELS", build_labels(filtered, label_start), f"Labels_{fn}.pdf")
            
            csv = pd.DataFrame(excl_data).to_csv(index=False, header=False).encode('utf-8-sig')
            c4.download_button("📊 EXCL", csv, f"Excl_{fn}.csv", "text/csv")
