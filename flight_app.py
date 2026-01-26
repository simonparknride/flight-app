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

# --- 1. UI 설정 및 스타일 (기존 유지) ---
st.set_page_config(page_title="Flight List Factory", layout="centered", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .stApp { background-color: #000000; }
    [data-testid="stSidebar"] { background-color: #111111 !important; }
    .stMarkdown, p, h1, h2, h3, label { color: #ffffff !important; }
    
    div.stDownloadButton > button {
        background-color: #ffffff !important; 
        color: #000000 !important;           
        border: 2px solid #ffffff !important;
        border-radius: 8px !important;
        padding: 0.5rem 1rem !important;
        font-weight: 800 !important;
        width: 100% !important;
        font-size: 0.85rem !important;
    }
    div.stDownloadButton > button * { color: #000000 !important; }
    div.stDownloadButton > button:hover {
        background-color: #60a5fa !important; 
        color: #ffffff !important;           
        border: 2px solid #60a5fa !important;
    }
    div.stDownloadButton > button:hover * { color: #ffffff !important; }

    .top-left-container { text-align: left; padding-top: 10px; margin-bottom: 20px; }
    .top-left-container a { font-size: 1.1rem; color: #ffffff !important; text-decoration: underline; display: block; margin-bottom: 5px;}
    .main-title { font-size: 3rem; font-weight: 800; color: #ffffff; line-height: 1.1; margin-bottom: 0.5rem; }
    .sub-title { font-size: 2.5rem; font-weight: 400; color: #60a5fa; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 파싱 및 필터링 로직 (절대 수정 금지 원칙 준수) ---
TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s[AP]M)\t([A-Z]{2}\d+[A-Z]?)\s*$")
DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
IATA_IN_PAREns = re.compile(r"\(([^)]+)\)")
PLANE_TYPES = ['A21N','A20N','A320','32Q','320','73H','737','74Y','77W','B77W','789','B789','359','A359','332','A332','AT76','DH8C','DH3','AT7','388','333','A333','330','76V','77L','B38M','A388','772','B772','32X','77X']
PLANE_TYPE_PATTERN = re.compile(r"\b(" + "|".join(sorted(set(PLANE_TYPES), key=len, reverse=True)) + r")\b", re.IGNORECASE)
NORMALIZE_MAP = {'32q': 'A320', '320': 'A320', 'a320': 'A320', '32x': 'A320', '789': 'B789', 'b789': 'B789', '772': 'B772', 'b772': 'B772', '77w': 'B77W', 'b77w': 'B77W', '332': 'A332', 'a332': 'A332', '333': 'A333', 'a333': 'A333', '330': 'A330', 'a330': 'A330', '359': 'A359', 'a359': 'A359', '388': 'A388', 'a388': 'A388', '737': 'B737', '73h': 'B737', 'at7': 'AT76'}
ALLOWED_AIRLINES = {"NZ","QF","JQ","CZ","CA","SQ","LA","IE","FX"}
NZ_DOMESTIC_IATA = {"AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO","WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"}
REGO_LIKE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9\-–—]*$")

def normalize_type(t: str) -> str:
    key = (t or '').strip().lower()
    return NORMALIZE_MAP.get(key, t)

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
        if m and current_date is not None:
            time_str, flight = m.groups()
            dest_line = lines[i+1].strip() if i+1 < len(lines) else ''
            m2 = IATA_IN_PAREns.search(dest_line)
            dest_iata = (m2.group(1).strip() if m2 else '').upper()
            carrier_line = lines[i+2].rstrip('\n') if i+2 < len(lines) else ''
            mtype = PLANE_TYPE_PATTERN.search(carrier_line)
            plane_type = normalize_type(mtype.group(1).upper() if mtype else '')
            reg = ''
            parens = IATA_IN_PAREns.findall(carrier_line)
            if parens:
                for candidate in reversed(parens):
                    cand = candidate.strip()
                    if REGO_LIKE.match(cand) and '-' in cand:
                        reg = cand; break
                if not reg: reg = parens[-1].strip()
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
    if start_hm == "00:00" and end_hm == "00:00":
        start_dt = datetime.combine(day1, datetime.min.time())
        end_dt = datetime.combine(day1, datetime.max.time())
    else:
        start_dt = datetime.combine(day1, datetime.strptime(start_hm, '%H:%M').time())
        end_dt = datetime.combine(day2, datetime.strptime(end_hm, '%H:%M').time())
    out = [r for r in records if r.get('dt') and r['flight'][:2] in ALLOWED_AIRLINES and r['dest'] not in NZ_DOMESTIC_IATA and (start_dt <= r['dt'] <= end_dt)]
    out.sort(key=lambda x: x['dt'])
    return out, start_dt, end_dt

# --- 3. DOCX 생성 함수 (기존 2페이지용 & 신규 1페이지용) ---
def build_docx_stream(records, start_dt, end_dt, compact=False):
    doc = Document()
    font_name = 'Air New Zealand Sans'
    section = doc.sections[0]
    
    # [수정] compact 모드일 때 여백을 대폭 줄임
    margin = 0.3 if compact else 0.5
    section.top_margin = section.bottom_margin = Inches(margin)
    section.left_margin = section.right_margin = Inches(margin)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_head = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}")
    run_head.bold = True
    run_head.font.name = font_name
    run_head.font.size = Pt(14 if compact else 16)
    
    table = doc.add_table(rows=0, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    for i, r in enumerate(records):
        row = table.add_row()
        tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        vals = [r['flight'], tdisp, r['dest'], r['type'], r['reg']]
        for j, val in enumerate(vals):
            cell = row.cells[j]
            para = cell.paragraphs[0]
            # [수정] 줄 간격 좁게 설정
            para.paragraph_format.line_spacing = 1.0
            para.paragraph_format.space_after = Pt(0)
            run = para.add_run(str(val))
            run.font.name = font_name
            # [수정] 글자 크기 조정 (컴팩트 모드 11.5pt)
            run.font.size = Pt(11.5 if compact else 14)
            
    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- 4. PDF 및 CSV (기존 유지) ---
def build_labels_stream(records, start_num):
    target = io.BytesIO(); c = canvas.Canvas(target, pagesize=A4); w, h = A4
    margin, gutter = 15*mm, 6*mm; col_w, row_h = (w - 2*margin - gutter) / 2, (h - 2*margin) / 5
    for i, r in enumerate(records):
        if i > 0 and i % 10 == 0: c.showPage()
        idx = i % 10; x_left = margin + (idx % 2) * (col_w + gutter); y_top = h - margin - (idx // 2) * row_h
        c.setFont('Helvetica-Bold', 38); c.drawString(x_left + 15*mm, y_top - 21*mm, r['flight'])
    c.save(); target.seek(0); return target

def build_excel_stream_24h(all_records):
    filtered_24h, _, _ = filter_records(all_records, "00:00", "00:00")
    flights = [r['flight'] for r in filtered_24h]
    df = pd.DataFrame(flights, columns=["Flight"])
    csv_data = df.to_csv(index=False, header=False).encode('utf-8-sig')
    return io.BytesIO(csv_data)

# --- 5. 사이드바 (기존 유지) ---
with st.sidebar:
    st.header("⚙️ Settings")
    s_val = st.text_input("Start Time", value="04:55")
    e_val = st.text_input("End Time", value="05:00")
    label_start = st.number_input("Label Start Number", value=1, min_value=1)

st.markdown('<div class="top-left-container"><a href="https://www.flightradar24.com/data/airports/akl/arrivals" target="_blank">Import Raw Text</a><a href="https://www.flightradar24.com/data/airports/akl/departures" target="_blank">Export Raw Text</a></div>', unsafe_allow_html=True)
st.markdown('<div class="main-title">Simon Park\'nRide\'s<br><span class="sub-title">Flight List Factory</span></div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Raw Text File", type=['txt'])

if uploaded_file:
    lines = uploaded_file.read().decode("utf-8").splitlines()
    all_recs = parse_raw_lines(lines)
    if all_recs:
        filtered, s_dt, e_dt = filter_records(all_recs, s_val, e_val)
        if filtered:
            st.success(f"Processed {len(filtered)} flights (2026 Updated)")
            
            # [수정] 버튼 4개 배치를 위해 columns(4)로 확장
            col1, col2, col3, col4 = st.columns(4)
            fn = f"List_{s_dt.strftime('%d-%m')}"
            
            with col1:
                st.download_button("📥 DOCX List (Orig)", build_docx_stream(filtered, s_dt, e_dt), f"{fn}_original.docx")
            with col2:
                # [신규 추가] 1페이지 최적화 버전 버튼
                st.download_button("📄 DOCX (1-Page)", build_docx_stream(filtered, s_dt, e_dt, compact=True), f"{fn}_1page.docx")
            with col3:
                st.download_button("📥 PDF Labels", build_labels_stream(filtered, label_start), f"Labels_{fn}.pdf")
            with col4:
                st.download_button("📊 Only Flights", build_excel_stream_24h(all_recs), f"Flights_24H_{fn}.csv", "text/csv")
            
            st.table([{'No': label_start+i, 'Date': r['dt'].strftime('%d %b'), 'Flight': r['flight'], 'Time': r['time'], 'Dest': r['dest']} for i, r in enumerate(filtered)])
