import streamlit as st
import re
import io
from datetime import datetime, timedelta, time as dtime
from typing import List, Dict, Optional
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# --- 페이지 설정 ---
st.set_page_config(page_title="Flight List Factory", layout="centered", initial_sidebar_state="expanded")

# --- 스타일 설정 ---
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
        padding: 0.6rem 1.2rem !important;
        font-weight: 800 !important;
        width: 100% !important;
    }
    div.stDownloadButton > button * { color: #000000 !important; }
    div.stDownloadButton > button:hover {
        background-color: #60a5fa !important;
        color: #ffffff !important;
        border: 2px solid #60a5fa !important;
    }

    .top-left-container { text-align: left; padding-top: 10px; margin-bottom: 20px; }
    .top-left-container a { font-size: 1.1rem; color: #ffffff !important; text-decoration: underline; display: block; margin-bottom: 5px;}
    .main-title { font-size: 3rem; font-weight: 800; color: #ffffff; line-height: 1.1; margin-bottom: 0.5rem; }
    .sub-title { font-size: 2.5rem; font-weight: 400; color: #60a5fa; }
    </style>
    """, unsafe_allow_html=True)

# --- 파싱 로직 및 상수 ---
TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s?[AP]M)\s+([A-Z0-9]{2,4}\d*[A-Z]?)\s*$", re.IGNORECASE)
DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
IATA_IN_PARENS = re.compile(r"\(([^)]+)\)")
PLANE_TYPES = ['A21N','A20N','A320','32Q','320','73H','737','74Y','77W','B77W','789','B789','359','A359','332','A332','AT76','DH8C','DH3','AT7','388','333','A333','330','76V','77L','B38M','A388','772','B772','32X','77X']
PLANE_TYPE_PATTERN = re.compile(r"\b(" + "|".join(sorted(set(PLANE_TYPES), key=len, reverse=True)) + r")\b", re.IGNORECASE)

NORMALIZE_MAP = {
    '32q': 'A320', '320': 'A320', 'a320': 'A320', '32x': 'A320',
    '789': 'B789', 'b789': 'B789', '772': 'B772', 'b772': 'B772',
    '77w': 'B77W', 'b77w': 'B77W', '332': 'A332', 'a332': 'A332',
    '333': 'A333', 'a333': 'A333', '330': 'A330', 'a330': 'A330',
    '359': 'A359', 'a359': 'A359', '388': 'A388', 'a388': 'A388',
    '737': 'B737', '73h': 'B737', 'at7': 'AT76'
}
ALLOWED_AIRLINES = {"NZ", "QF", "JQ", "CZ", "CA", "SQ", "LA", "IE", "FX"}
NZ_DOMESTIC_IATA = {"AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO","WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"}
REGO_LIKE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9\-–—]*$")

def normalize_type(t: Optional[str]) -> str:
    if not t: return ""
    key = t.strip().lower()
    return NORMALIZE_MAP.get(key, t.strip().upper())

def try_parse_date_header(line: str, year: int) -> Optional[datetime.date]:
    candidates = ["%A, %b %d %Y", "%A, %B %d %Y", "%a, %b %d %Y", "%a, %B %d %Y"]
    text = line.strip() + f" {year}"
    for fmt in candidates:
        try: return datetime.strptime(text, fmt).date()
        except: continue
    return None

def parse_raw_lines(lines: List[str], year: int) -> List[Dict]:
    records = []
    current_date = None
    i, L = 0, len(lines)
    while i < L:
        line = lines[i].strip()
        if DATE_HEADER.match(line):
            parsed = try_parse_date_header(line, year)
            current_date = parsed if parsed else None
            i += 1
            continue
        m = TIME_LINE.match(line)
        if m and current_date is not None:
            time_str_raw, flight_raw = m.groups()
            dest_line = lines[i+1].strip() if i+1 < L else ''
            carrier_line = lines[i+2].rstrip('\n') if i+2 < L else ''
            m2 = IATA_IN_PARENS.search(dest_line)
            dest_iata = (m2.group(1).strip().upper() if m2 else '').upper()
            mtype = PLANE_TYPE_PATTERN.search(carrier_line or '')
            plane_type = normalize_type(mtype.group(1) if mtype else '')
            reg = ''
            parens = IATA_IN_PARENS.findall(carrier_line or '')
            if parens:
                for cand in reversed(parens):
                    cand = cand.strip()
                    if REGO_LIKE.match(cand) and any(x in cand for x in ['-', '–', '—']):
                        reg = cand
                        break
                if not reg: reg = parens[-1].strip()
            dep_dt = None
            try:
                tnorm = time_str_raw.strip().upper().replace(" ", "")
                fmt = "%Y-%m-%d %I:%M%p" if re.match(r"^\d{1,2}:\d{2}[AP]M$", tnorm) else "%Y-%m-%d %I:%M %p"
                dep_dt = datetime.strptime(f"{current_date} {time_str_raw.strip()}", fmt)
            except: dep_dt = None
            records.append({'dt': dep_dt, 'time': time_str_raw.strip(), 'flight': flight_raw.strip().upper(), 'dest': dest_iata, 'type': plane_type, 'reg': reg})
            i += 3
            continue
        i += 1
    return records

def filter_records(records: List[Dict], start_time: dtime, end_time: dtime):
    dates = sorted({r['dt'].date() for r in records if r.get('dt')})
    if not dates: return [], None, None
    day1 = dates[0]
    day2 = dates[1] if len(dates) >= 2 else (day1 + timedelta(days=1))
    start_dt = datetime.combine(day1, start_time)
    end_dt = datetime.combine(day2, end_time)
    if end_dt <= start_dt: return [], start_dt, end_dt
    def allowed(r):
        if not r.get('dt'): return False
        if (r.get('flight') or '')[:2].upper() not in ALLOWED_AIRLINES: return False
        if (r.get('dest') or '').upper() in NZ_DOMESTIC_IATA: return False
        return start_dt <= r['dt'] <= end_dt
    out = sorted([r for r in records if allowed(r)], key=lambda x: x['dt'] or datetime.max)
    return out, start_dt, end_dt

# --- 워드 생성 함수 ---

def build_docx_stream(records: List[Dict], start_dt: datetime, end_dt: datetime) -> io.BytesIO:
    """TWO-PAGE DOCX: 기존 설정을 절대 유지 (0pt Spacing)"""
    doc = Document()
    font_name = 'Air New Zealand Sans'
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Inches(0.3)
    section.left_margin = section.right_margin = Inches(0.5)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_head = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}")
    run_head.bold = True
    run_head.font.name = font_name
    run_head.font.size = Pt(16)
    
    table = doc.add_table(rows=0, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    for i, r in enumerate(records):
        row = table.add_row()
        try:
            tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        except:
            tdisp = r['time']
            
        vals = [r['flight'], tdisp, r['dest'], r['type'], r['reg']]
        for j, val in enumerate(vals):
            cell = row.cells[j]
            if i % 2 == 1:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd')
                shd.set(qn('w:val'), 'clear')
                shd.set(qn('w:fill'), 'D9D9D9')
                tcPr.append(shd)
            
            para = cell.paragraphs[0]
            para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            # 2페이지 리스트는 기존 설정(0pt) 유지
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(0)
            
            run = para.add_run(str(val))
            run.font.name = font_name
            run.font.size = Pt(14)
            
    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

def build_docx_onepage_stream(records: List[Dict], start_dt: datetime, end_dt: datetime) -> io.BytesIO:
    """ONE-PAGE DOCX (Two Columns): 지시대로 행 간격 2.2pt 적용"""
    doc = Document()
    font_name = 'Air New Zealand Sans'
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Inches(0.2)
    section.left_margin = section.right_margin = Inches(0.4)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_head = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}")
    run_head.bold = True
    run_head.font.name = font_name
    run_head.font.size = Pt(16)
    
    total = len(records)
    mid = (total + 1) // 2
    left_recs = records[:mid]
    right_recs = records[mid:]
    
    outer = doc.add_table(rows=1, cols=2)
    outer.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    def add_inner_table(cell, recs, start_index):
        inner = cell.add_table(rows=1, cols=5)
        headers = ['Flight', 'Time', 'Dest', 'Type', 'Reg']
        for idx, text in enumerate(headers):
            h_para = inner.rows[0].cells[idx].paragraphs[0]
            h_run = h_para.add_run(text)
            h_run.bold = True
            h_run.font.size = Pt(11)
            h_run.font.name = font_name
            
        for i, r in enumerate(recs):
            row = inner.add_row()
            try:
                tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
            except:
                tdisp = r['time']
            
            vals = [r['flight'], tdisp, r['dest'], r['type'], r['reg']]
            for j, val in enumerate(vals):
                cell_j = row.cells[j]
                if (start_index + i) % 2 == 1:
                    tcPr = cell_j._tc.get_or_add_tcPr()
                    shd = OxmlElement('w:shd')
                    shd.set(qn('w:val'), 'clear')
                    shd.set(qn('w:fill'), 'D9D9D9')
                    tcPr.append(shd)
                
                para = cell_j.paragraphs[0]
                para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                
                # 지시사항 반영: One-page는 2.2pt로 표를 키움
                para.paragraph_format.space_before = Pt(2.2)
                para.paragraph_format.space_after = Pt(2.2)
                
                run = para.add_run(str(val))
                run.font.name = font_name
                run.font.size = Pt(9) if j == 4 else Pt(11)
                
    add_inner_table(outer.cell(0, 0), left_recs, 0)
    add_inner_table(outer.cell(0, 1), right_recs, mid)
    
    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- PDF 라벨 생성 ---
def build_labels_stream(records: List[Dict], start_num: int) -> io.BytesIO:
    target = io.BytesIO()
    c = canvas.Canvas(target, pagesize=A4)
    w, h = A4
    margin, gutter = 15*mm, 6*mm
    col_w = (w - 2*margin - gutter) / 2
    row_h = (h - 2*margin) / 5
    
    for i, r in enumerate(records):
        if i > 0 and i % 10 == 0:
            c.showPage()
        idx = i % 10
        x_left = margin + (idx % 2) * (col_w + gutter)
        y_top = h - margin - (idx // 2) * row_h
        
        c.setStrokeGray(0.3)
        c.rect(x_left, y_top - row_h + 2*mm, col_w, row_h - 4*mm)
        c.setFont('Helvetica-Bold', 14)
        c.drawCentredString(x_left + 7*mm, y_top - 9.5*mm, str(start_num + i))
        
        c.setFont('Helvetica-Bold', 38)
        c.drawString(x_left + 15*mm, y_top - 21*mm, r['flight'])
        
        try:
            tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        except:
            tdisp = r['time']
            
        c.setFont('Helvetica-Bold', 29)
        c.drawString(x_left + 15*mm, y_top - 47*mm, tdisp)
        
    c.save()
    target.seek(0)
    return target

# --- 메인 애플리케이션 ---

with st.sidebar:
    st.header("⚙️ 설정")
    year = st.number_input("Year", value=datetime.now().year)
    s_time = st.time_input("Start Time", value=dtime(5, 0))
    e_time = st.time_input("End Time", value=dtime(4, 55))
    label_start = st.number_input("라벨 시작 번호", value=1)

st.markdown('<div class="top-left-container"><a href="https://www.flightradar24.com/data/airports/akl/arrivals" target="_blank">Import Raw Text</a><a href="https://www.flightradar24.com/data/airports/akl/departures" target="_blank">Export Raw Text</a></div>', unsafe_allow_html=True)
st.markdown('<div class="main-title">Air New Zealand Cargo<br><span class="sub-title">Flight List</span></div>', unsafe_allow_html=True)

# 파일 업로더
uploaded_file = st.file_uploader("Upload Raw Text File", type=['txt'])

# 파일이 업로드되었을 때만 실행
if uploaded_file is not None:
    try:
        # 파일 내용 읽기
        content = uploaded_file.read().decode("utf-8", errors="replace")
        lines = content.splitlines()
        
        # 파싱 실행
        all_recs = parse_raw_lines(lines, year)
        
        if not all_recs:
            st.warning("데이터를 파싱할 수 없습니다. 파일 형식을 확인해 주세요.")
        else:
            # 필터링 실행
            filtered, s_dt, e_dt = filter_records(all_recs, s_time, e_time)
            
            if not filtered:
                st.warning("필터 설정에 맞는 비행 정보가 없습니다.")
            else:
                st.success(f"{len(filtered)}개의 비행 정보를 찾았습니다.")
                
                # 다운로드 버튼 영역
                c1, cm, c2 = st.columns([1, 1, 1])
                fn = f"List_{s_dt.strftime('%d-%m')}"
                
                c1.download_button("📥 2-Page DOCX", build_docx_stream(filtered, s_dt, e_dt).getvalue(), f"{fn}.docx")
                cm.download_button("📥 1-Page DOCX (Two Columns)", build_docx_onepage_stream(filtered, s_dt, e_dt).getvalue(), f"{fn}_onepage.docx")
                c2.download_button("📥 PDF Labels", build_labels_stream(filtered, label_start).getvalue(), f"Labels_{fn}.pdf")
                
                # 미리보기 테이블
                table_display = []
                for i, r in enumerate(filtered):
                    try:
                        t = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
                    except:
                        t = r['time']
                    table_display.append({
                        'No': label_start + i,
                        'Flight': r['flight'],
                        'Time': t,
                        'Dest': r['dest'],
                        'Type': r['type'],
                        'Reg': r['reg']
                    })
                st.table(table_display)
    except Exception as e:
        st.error(f"파일 처리 중 오류가 발생했습니다: {e}")
