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

# --- 1. UI 및 스타일 설정 ---
st.set_page_config(page_title="Flight List Factory", layout="centered")
st.markdown("""
    <style>
    .stApp { background-color: #000000; }
    [data-testid="stSidebar"] { background-color: #111111 !important; }
    .stMarkdown, p, h1, h2, h3, label { color: #ffffff !important; }
    div.stDownloadButton > button {
        background-color: #ffffff !important; color: #000000 !important;
        border: 2px solid #ffffff !important; border-radius: 8px !important;
        padding: 0.6rem 1.2rem !important; font-weight: 800 !important; width: 100% !important;
    }
    div.stDownloadButton > button:hover { background-color: #60a5fa !important; color: #ffffff !important; }
    .main-title { font-size: 3rem; font-weight: 800; color: #ffffff; line-height: 1.1; margin-bottom: 0.5rem; }
    .sub-title { font-size: 2.5rem; font-weight: 400; color: #60a5fa; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 데이터 파싱 로직 ---
TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s[AP]M)\t([A-Z]{2}\d+[A-Z]?)\s*$")
DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
IATA_IN_PAREns = re.compile(r"\(([^)]+)\)")
ALLOWED_AIRLINES = {"NZ","QF","JQ","CZ","CA","SQ","LA","IE"}
NZ_DOMESTIC_IATA = {"AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO","WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"}

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
            reg = ''; ptype = 'B789' # 기본값
            parens = IATA_IN_PAREns.findall(carrier_line)
            if parens: reg = parens[-1].strip()
            # 기종 추출 (간단 버전)
            for pt in ['A21N','A320','B789','B77W','A359','A333']:
                if pt in carrier_line.upper(): ptype = pt; break
            try: dt = datetime.strptime(f"{current_date} {time_str}", '%Y-%m-%d %I:%M %p')
            except: dt = None
            records.append({'dt': dt, 'date_label': current_date.strftime('%d %b'), 'time': time_str, 'flight': flight, 'dest': dest_iata, 'type': ptype, 'reg': reg})
            i += 4; continue
        i += 1
    return records

# --- 3. DOCX 생성 (1-Page 특화 설정 적용) ---
def build_docx_stream(records, start_dt, end_dt, is_1page=False):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Inches(0.3)
    section.left_margin = section.right_margin = Inches(0.5)

    # Footer
    footer = section.footer
    run_f = footer.paragraphs[0].add_run("created by Simon Park'nRide's Flight List Factory 2026")
    run_f.font.size = Pt(7.5 if is_1page else 10)
    run_f.font.color.rgb = RGBColor(128, 128, 128)

    # Header
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_h = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d %b')}")
    run_h.bold = True
    run_h.font.size = Pt(7.5 if is_1page else 16)

    # Table
    table = doc.add_table(rows=0, cols=5); table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # [핵심] 1-Page일 때 테이블 폭을 70%로 제한
    tblPr = table._element.xpath('w:tblPr')[0]
    tblW = OxmlElement('w:tblW')
    if is_1page:
        tblW.set(qn('w:w'), '3500') # 5000이 100%이므로 3500은 70%
    else:
        tblW.set(qn('w:w'), '5000')
    tblW.set(qn('w:type'), 'pct')
    tblPr.append(tblW)

    for i, r in enumerate(records):
        row = table.add_row()
        tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        vals = [r['flight'], tdisp, r['dest'], r['type'], r['reg']]
        for j, v in enumerate(vals):
            cell = row.cells[j]
            if i % 2 == 1: # 줄무늬 배경
                shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9'); cell._tc.get_or_add_tcPr().append(shd)
            para = cell.paragraphs[0]
            para.paragraph_format.space_before = para.paragraph_format.space_after = Pt(0)
            run = para.add_run(str(v))
            run.font.size = Pt(7.5 if is_1page else 14)
    
    target = io.BytesIO(); doc.save(target); target.seek(0)
    return target

# --- 4. PDF Labels (가이드 좌표 고정) ---
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
        c.setStrokeGray(0.3); c.rect(x, y - row_h + 2*mm, col_w, row_h - 4*mm)
        # 순번 & 날짜
        c.setFont('Helvetica-Bold', 12); c.drawString(x + 5*mm, y - 10*mm, str(start_num + i))
        c.setFont('Helvetica', 10); c.drawRightString(x + col_w - 5*mm, y - 10*mm, r['date_label'])
        # 정보들
        c.setFont('Helvetica-Bold', 36); c.drawCentredString(x + col_w/2, y - 25*mm, r['flight'])
        c.setFont('Helvetica-Bold', 22); c.drawCentredString(x + col_w/2, y - 38*mm, r['dest'])
        tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        c.setFont('Helvetica-Bold', 30); c.drawCentredString(x + col_w/2, y - 52*mm, tdisp)
        c.setFont('Helvetica', 10); c.drawCentredString(x + col_w/2, y - row_h + 8*mm, f"{r['type']} {r['reg']}")
    c.save(); target.seek(0); return target

# --- 5. 화면 구성 및 실행 ---
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
        # 시간 필터링 (업로드된 데이터의 첫날 기준)
        day1 = all_recs[0]['dt'].date()
        start_bound = datetime.combine(day1, datetime.strptime(s_time, '%H:%M').time())
        end_bound = start_bound + timedelta(hours=23, minutes=55)
        
        filtered = [r for r in all_recs if start_bound <= r['dt'] <= end_bound and r['flight'][:2] in ALLOWED_AIRLINES and r['dest'] not in NZ_DOMESTIC_IATA]
        
        # Flights Excl용 (모든 국제선 편명)
        excl_list = sorted(list({r['flight'] for r in all_recs if r['flight'][:2] in ALLOWED_AIRLINES and r['dest'] not in NZ_DOMESTIC_IATA}))

        st.success(f"Successfully processed {len(filtered)} flights")
        
        c1, c2, c3, c4 = st.columns(4)
        fn = f"List_{start_bound.strftime('%d-%m')}"
        
        c1.download_button("📥 DOCX", build_docx_stream(filtered, start_bound, end_bound), f"{fn}.docx")
        c2.download_button("📄 1-PAGE", build_docx_stream(filtered, start_bound, end_bound, True), f"{fn}_1p.docx")
        c3.download_button("🏷️ LABELS", build_labels_stream(filtered, label_start), f"Labels_{fn}.pdf")
        
        # Excel 생성 (Flights Excl)
        df_excl = pd.DataFrame(excl_list)
        excel_io = io.BytesIO()
        with pd.ExcelWriter(excel_io, engine='xlsxwriter') as writer:
            df_excl.to_excel(writer, index=False, header=False)
        c4.download_button("📊 EXCL EXCEL", excel_io.getvalue(), f"Excl_{fn}.xlsx")

        st.divider()
        st.table(pd.DataFrame(filtered)[['flight', 'time', 'dest', 'reg']])
