import streamlit as st
import re
import io
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

# --- 1. UI 설정 (버튼 복구 및 Hover 시인성) ---
st.set_page_config(page_title="Flight List Factory", layout="centered")

st.markdown("""
    <style>
    .stApp { background-color: #000000; }
    [data-testid="stSidebar"] { background-color: #111111 !important; }
    .stMarkdown, p, h1, h2, h3, label { color: #ffffff !important; }
    
    .top-left-container { text-align: left; padding-top: 10px; margin-bottom: 20px; }
    .top-left-container a { font-size: 1.1rem !important; color: #ffffff !important; text-decoration: underline; display: block; margin-bottom: 5px; }
    .main-title { font-size: 3.2rem !important; font-weight: 800; color: #ffffff; line-height: 1.1; margin-bottom: 0.5rem; }
    .sub-title { font-size: 2.6rem !important; font-weight: 400; color: #60a5fa; }

    /* 버튼 스타일 및 마우스 오버 시 반전 효과 */
    div.stDownloadButton > button {
        background-color: #000000 !important; color: #ffffff !important;           
        border: 1px solid #ffffff !important; border-radius: 4px !important;
        padding: 0.5rem 1rem !important; width: 100% !important; transition: all 0.2s ease;
    }
    div.stDownloadButton > button:hover {
        background-color: #ffffff !important; color: #000000 !important;
    }
    div.stDownloadButton > button:hover p { color: #000000 !important; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 파싱 로직 ---
TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s[AP]M)\t([A-Z]{2}\d+[A-Z]?)\s*$")
DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
IATA_IN_PAREns = re.compile(r"\(([^)]+)\)")
PLANE_TYPES = ['A21N','A20N','A320','32Q','320','73H','737','74Y','77W','B77W','789','B789','359','A359','332','A332','AT76','388','333','A333','330','772','B772']
PLANE_TYPE_PATTERN = re.compile(r"\b(" + "|".join(sorted(set(PLANE_TYPES), key=len, reverse=True)) + r")\b", re.IGNORECASE)
NORMALIZE_MAP = {'32q': 'A320', '320': 'A320', '789': 'B789', '77w': 'B77W', '359': 'A359', '388': 'A388'}
ALLOWED_AIRLINES = {"NZ","QF","JQ","CZ","CA","SQ","LA","IE"}
NZ_DOMESTIC_IATA = {"AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO"}

def parse_raw_lines(lines):
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
        if m and current_date is not None:
            time_str, flight = m.groups()
            dest_line = lines[i+1].strip() if i+1 < len(lines) else ''
            m2 = IATA_IN_PAREns.search(dest_line)
            dest_iata = (m2.group(1).strip() if m2 else '').upper()
            carrier_line = lines[i+2].rstrip('\n') if i+2 < len(lines) else ''
            mtype = PLANE_TYPE_PATTERN.search(carrier_line)
            plane_type = NORMALIZE_MAP.get(mtype.group(1).lower() if mtype else '', mtype.group(1).upper() if mtype else '')
            reg = ''
            parens = IATA_IN_PAREns.findall(carrier_line)
            if parens: reg = parens[-1].strip()
            try: dep_dt = datetime.strptime(f"{current_date} {time_str}", '%Y-%m-%d %I:%M %p')
            except: dep_dt = None
            records.append({'dt': dep_dt, 'time': time_str, 'flight': flight, 'dest': dest_iata, 'type': plane_type, 'reg': reg})
            i += 4; continue
        i += 1
    return records

def filter_records(records, start_hm, end_hm):
    dates = sorted({r['dt'].date() for r in records if r.get('dt')})
    if not dates: return [], None, None
    day1, day2 = dates[0], dates[1] if len(dates) >= 2 else (dates[0] + timedelta(days=1))
    start_dt = datetime.combine(day1, datetime.strptime(start_hm, '%H:%M').time())
    end_dt = datetime.combine(day2, datetime.strptime(end_hm, '%H:%M').time())
    out = [r for r in records if r.get('dt') and r['flight'][:2] in ALLOWED_AIRLINES and r['dest'] not in NZ_DOMESTIC_IATA and (start_dt <= r['dt'] <= end_dt)]
    out.sort(key=lambda x: x['dt'])
    return out, start_dt, end_dt

# --- 3. DOCX 생성 (강력한 레이아웃 고정) ---
def build_docx_stream(records, start_dt, end_dt, mode='Two Pages'):
    doc = Document()
    section = doc.sections[0]
    section.left_margin = section.right_margin = Inches(0.5)

    if mode == 'One Page':
        section.top_margin = section.bottom_margin = Inches(0.15)
        f_size, h_size = Pt(7.5), Pt(11)
        # 날짜 헤더 왼쪽 끝으로 밀기 (음수 인덴트)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.left_indent = Inches(-0.05)
    else:
        section.top_margin = section.bottom_margin = Inches(0.5)
        f_size, h_size = Pt(14), Pt(16)
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}")
    run.bold = True; run.font.size = h_size

    # 표 중앙 정렬 및 행 높이 강제 제한
    table = doc.add_table(rows=0, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # XML로 표 너비 강제 고정 (5400 dxa = 약 7.5인치)
    tblPr = table._tbl.xpath('w:tblPr')[0]
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '5400'); tblW.set(qn('w:type'), 'dxa')
    tblPr.append(tblW)

    for i, r in enumerate(records):
        row = table.add_row()
        if mode == 'One Page':
            # 행 높이를 7.5pt에 맞춰 아주 작게 고정
            tr = row._tr
            trPr = tr.get_or_add_trPr()
            trHeight = OxmlElement('w:trHeight')
            trHeight.set(qn('w:val'), '180') # 1/1440 inch 단위
            trHeight.set(qn('w:hRule'), 'atLeast')
            trPr.append(trHeight)

        tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        for j, val in enumerate([r['flight'], tdisp, r['dest'], r['type'], r['reg']]):
            cell = row.cells[j]
            if i % 2 == 1:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9'); tcPr.append(shd)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.space_before = para.paragraph_format.space_after = Pt(0)
            para.paragraph_format.line_spacing = 1.0
            run_c = para.add_run(str(val))
            run_c.font.size = f_size
    
    target = io.BytesIO()
    doc.save(target); target.seek(0)
    return target

# --- 4. PDF LABEL (복구 완료) ---
def build_labels_stream(records, start_num):
    target = io.BytesIO()
    c = canvas.Canvas(target, pagesize=A4)
    w, h = A4
    margin, gutter = 15*mm, 6*mm
    col_w, row_h = (w - 2*margin - gutter) / 2, (h - 2*margin) / 5
    for i, r in enumerate(records):
        if i > 0 and i % 10 == 0: c.showPage()
        idx = i % 10
        x_left = margin + (idx % 2) * (col_w + gutter)
        y_top = h - margin - (idx // 2) * row_h
        c.setStrokeGray(0.3); c.rect(x_left, y_top - row_h + 2*mm, col_w, row_h - 4*mm)
        c.setFont('Helvetica-Bold', 14); c.drawString(x_left + 4*mm, y_top - 10*mm, str(start_num + i))
        c.setFont('Helvetica-Bold', 35); c.drawString(x_left + 15*mm, y_top - 22*mm, r['flight'])
        tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        c.setFont('Helvetica-Bold', 25); c.drawString(x_left + 15*mm, y_top - 45*mm, f"{tdisp}  {r['dest']}")
    c.save(); target.seek(0)
    return target

# --- 5. 앱 실행 ---
with st.sidebar:
    st.header("⚙️ Settings")
    s_time = st.text_input("Start Time", value="05:00")
    e_time = st.text_input("End Time", value="04:55")
    label_start = st.number_input("Label Start Number", value=1, min_value=1)

st.markdown('<div class="top-left-container"><a href="...">Import Raw Text</a><a href="...">Export Raw Text</a></div>', unsafe_allow_html=True)
st.markdown('<div class="main-title">Simon Park\'nRide\'s<br><span class="sub-title">Flight List Factory</span></div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Raw Text File", type=['txt'])

if uploaded_file:
    lines = uploaded_file.read().decode("utf-8").splitlines()
    all_recs = parse_raw_lines(lines)
    if all_recs:
        filtered, s_dt, e_dt = filter_records(all_recs, s_time, e_time)
        if filtered:
            st.success(f"Processed {len(filtered)} flights")
            # [복구] PDF Labels를 포함한 3개 버튼 배치
            col1, col2, col3 = st.columns(3)
            fn = f"List_{s_dt.strftime('%d-%m')}"
            col1.download_button("📥 One Page", build_docx_stream(filtered, s_dt, e_dt, mode='One Page'), f"{fn}_1P.docx")
            col2.download_button("📥 Two Pages", build_docx_stream(filtered, s_dt, e_dt, mode='Two Pages'), f"{fn}_2P.docx")
            col3.download_button("📥 PDF Labels", build_labels_stream(filtered, label_start), f"Labels_{fn}.pdf")
