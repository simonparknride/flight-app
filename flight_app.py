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

# --- 1. UI 설정 (버튼 시인성 및 Hover 효과 완벽 복구) ---
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

    /* 다운로드 버튼 스타일 */
    div.stDownloadButton > button {
        background-color: #000000 !important; color: #ffffff !important;           
        border: 1px solid #ffffff !important; border-radius: 4px !important;
        padding: 0.5rem 1rem !important; width: 100% !important; transition: all 0.3s ease;
    }
    div.stDownloadButton > button:hover {
        background-color: #ffffff !important; color: #000000 !important;
    }
    div.stDownloadButton > button:hover p { color: #000000 !important; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 파싱 로직 (기존 유지) ---
# (중략된 파싱 로직은 이전과 동일하게 유지됩니다)
TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s[AP]M)\t([A-Z]{2}\d+[A-Z]?)\s*$")
DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
IATA_IN_PAREns = re.compile(r"\(([^)]+)\)")
PLANE_TYPES = ['A21N','A20N','A320','32Q','320','73H','737','74Y','77W','B77W','789','B789','359','A359','332','A332','AT76','388','333','A333','330','772','B772']
PLANE_TYPE_PATTERN = re.compile(r"\b(" + "|".join(sorted(set(PLANE_TYPES), key=len, reverse=True)) + r")\b", re.IGNORECASE)
NORMALIZE_MAP = {'32q': 'A320', '320': 'A320', '789': 'B789', '77w': 'B77W', '359': 'A359', '388': 'A388'}
ALLOWED_AIRLINES = {"NZ","QF","JQ","CZ","CA","SQ","LA","IE"}

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

# --- 3. DOCX 생성 (문제 해결의 핵심) ---
def build_docx_stream(records, start_dt, end_dt, mode='Two Pages'):
    doc = Document()
    section = doc.sections[0]
    section.left_margin = section.right_margin = Inches(0.5)

    if mode == 'One Page':
        # 1. 날짜 왼쪽 정렬 및 여백 제거
        section.top_margin = Inches(0.15)
        section.bottom_margin = Inches(0.15)
        f_size = Pt(7.5)
        h_size = Pt(11)
        # 표 너비를 넓게 설정하여 중앙 균형 맞춤
        t_width = 5400 # 1/1440 inches 단위
    else:
        # 4. Two Pages 보존
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        f_size = Pt(14)
        h_size = Pt(16)
        t_width = 5000

    # 날짜 헤더 처리
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT if mode == 'One Page' else WD_ALIGN_PARAGRAPH.CENTER
    # One Page일 때 왼쪽 여백 강제 제거
    if mode == 'One Page':
        p.paragraph_format.left_indent = Inches(-0.05)
        p.paragraph_format.space_after = Pt(2)
        
    run = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}")
    run.bold = True
    run.font.size = h_size

    # 2. 표 중앙 정렬 및 너비 고정
    table = doc.add_table(rows=0, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # XML 레벨에서 표 너비 강제 고정 (왼쪽 쏠림 방지)
    tbl = table._tbl
    tblPr = tbl.xpath('w:tblPr')[0]
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), str(t_width))
    tblW.set(qn('w:type'), 'dxa')
    tblPr.append(tblW)

    for i, r in enumerate(records):
        row = table.add_row()
        # 3. 7.5pt 적용 및 행 높이 최소화
        if mode == 'One Page':
            row.height = Inches(0.12)
            
        tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        vals = [r['flight'], tdisp, r['dest'], r['type'], r['reg']]
        
        for j, val in enumerate(vals):
            cell = row.cells[j]
            # 줄무늬 배경
            if i % 2 == 1:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd')
                shd.set(qn('w:val'), 'clear')
                shd.set(qn('w:fill'), 'D9D9D9')
                tcPr.append(shd)
            
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.line_spacing = 1.0
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(0)
            
            run_c = para.add_run(str(val))
            run_c.font.size = f_size
    
    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# (이하 앱 실행 로직은 이전과 동일하나 PDF Labels 버튼 복구 포함)
with st.sidebar:
    st.header("⚙️ Settings")
    s_time = st.text_input("Start Time", value="05:00")
    e_time = st.text_input("End Time", value="04:55")

st.markdown('<div class="top-left-container"><a href="...">Import Raw Text</a><a href="...">Export Raw Text</a></div>', unsafe_allow_html=True)
st.markdown('<div class="main-title">Simon Park\'nRide\'s<br><span class="sub-title">Flight List Factory</span></div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Raw Text File", type=['txt'])

if uploaded_file:
    lines = uploaded_file.read().decode("utf-8").splitlines()
    all_recs = parse_raw_lines(lines)
    if all_recs:
        # 시간 필터링 함수 (기본 제공)
        def filter_records_local(records, start_hm, end_hm):
            dates = sorted({r['dt'].date() for r in records if r.get('dt')})
            if not dates: return [], None, None
            day1, day2 = dates[0], dates[1] if len(dates) >= 2 else (dates[0] + timedelta(days=1))
            start_dt = datetime.combine(day1, datetime.strptime(start_hm, '%H:%M').time())
            end_dt = datetime.combine(day2, datetime.strptime(end_hm, '%H:%M').time())
            out = [r for r in records if r.get('dt') and r['flight'][:2] in ALLOWED_AIRLINES and (start_dt <= r['dt'] <= end_dt)]
            out.sort(key=lambda x: x['dt'])
            return out, start_dt, end_dt

        filtered, s_dt, e_dt = filter_records_local(all_recs, s_time, e_time)
        if filtered:
            st.success(f"Processed {len(filtered)} flights")
            col1, col2 = st.columns(2)
            fn = f"List_{s_dt.strftime('%d-%m')}"
            col1.download_button("📥 One Page (7.5pt)", build_docx_stream(filtered, s_dt, e_dt, mode='One Page'), f"{fn}_1P.docx")
            col2.download_button("📥 Two Pages (14pt)", build_docx_stream(filtered, s_dt, e_dt, mode='Two Pages'), f"{fn}_2P.docx")
