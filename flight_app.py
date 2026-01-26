import streamlit as st
import re
import io
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# --- 1. UI 설정 (디테일 완벽 고정) ---
st.set_page_config(page_title="Flight List Factory", layout="centered")

st.markdown("""
    <style>
    /* 1. 배경 및 사이드바 색상 고정 */
    .stApp { background-color: #000000; }
    [data-testid="stSidebar"] {
        background-color: #111111 !important;
        min-width: 300px !important;
    }
    
    /* 2. 사이드바 글자 크기 확대 (사라지지 않게 강화) */
    [data-testid="stSidebar"] label p, [data-testid="stSidebar"] span {
        font-size: 1.25rem !important; /* 더 크게 조정 */
        color: #ffffff !important;
        font-weight: 700 !important;
    }
    [data-testid="stSidebar"] input {
        font-size: 1.15rem !important;
        background-color: #333333 !important;
        color: #ffffff !important;
    }
    
    /* 3. 다운로드 버튼 스타일 및 애니메이션 */
    div.stDownloadButton > button {
        background-color: #000000 !important; color: #ffffff !important;           
        border: 1px solid #ffffff !important;
        width: 100% !important; height: 3.8rem !important;
        font-size: 1.1rem !important; font-weight: bold;
        transition: all 0.2s ease;
    }
    div.stDownloadButton > button:hover {
        background-color: #ffffff !important; color: #000000 !important;
    }
    div.stDownloadButton > button:hover p { color: #000000 !important; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 파싱 로직 (기존 정규식 유지) ---
TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s[AP]M)\t([A-Z]{2}\d+[A-Z]?)\s*$")
DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
IATA_IN_PAREns = re.compile(r"\(([^)]+)\)")
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
            # 기종 판별
            plane_type = "B789" if "789" in carrier_line else ("A320" if any(x in carrier_line for x in ["320","32Q","A20N","A21N"]) else "")
            reg = ""
            parens = IATA_IN_PAREns.findall(carrier_line)
            if parens: reg = parens[-1].strip()
            try: dep_dt = datetime.strptime(f"{current_date} {time_str}", '%Y-%m-%d %I:%M %p')
            except: dep_dt = None
            records.append({'dt': dep_dt, 'time': time_str, 'flight': flight, 'dest': dest_iata, 'type': plane_type, 'reg': reg})
            i += 4; continue
        i += 1
    return records

# --- 3. DOCX 생성 (One Page 최적화 및 Two Pages 복구 완벽 구분) ---
def build_docx_stream(records, start_dt, end_dt, mode='Two Pages'):
    doc = Document()
    section = doc.sections[0]
    
    if mode == 'One Page':
        # [요청사항] 7.5pt, 날짜 왼쪽, 표 중앙, 좁은 여백
        section.left_margin = section.right_margin = Inches(0.5)
        section.top_margin = section.bottom_margin = Inches(0.15)
        f_size, h_size = Pt(7.5), Pt(11)
        align_h = WD_ALIGN_PARAGRAPH.LEFT
        align_t = WD_TABLE_ALIGNMENT.CENTER
    else:
        # [복구] List_22-01.docx 스타일: 14pt, 왼쪽 정렬, 기본 여백
        section.left_margin = section.right_margin = Inches(1.0)
        section.top_margin = section.bottom_margin = Inches(1.0)
        f_size, h_size = Pt(14), Pt(16)
        align_h = WD_ALIGN_PARAGRAPH.LEFT
        align_t = WD_TABLE_ALIGNMENT.LEFT

    p = doc.add_paragraph()
    p.alignment = align_h
    run = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}")
    run.bold = True; run.font.size = h_size

    table = doc.add_table(rows=0, cols=5)
    table.alignment = align_t
    
    for i, r in enumerate(records):
        row = table.add_row()
        if mode == 'One Page':
            # 행 높이 강제 압축 (7.5pt용)
            tr = row._tr; trPr = tr.get_or_add_trPr()
            trH = OxmlElement('w:trHeight'); trH.set(qn('w:val'), '180'); trH.set(qn('w:hRule'), 'atLeast'); trPr.append(trH)
        
        tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        for j, val in enumerate([r['flight'], tdisp, r['dest'], r['type'], r['reg']]):
            cell = row.cells[j]
            if i % 2 == 1:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9'); tcPr.append(shd)
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_c = para.add_run(str(val))
            run_c.font.size = f_size
    
    target = io.BytesIO(); doc.save(target); target.seek(0)
    return target

# --- 4. PDF LABEL & 메인 로직 ---
def build_labels_stream(records, start_num):
    target = io.BytesIO(); c = canvas.Canvas(target, pagesize=A4)
    w, h = A4; margin, gutter = 15*mm, 6*mm
    col_w, row_h = (w - 2*margin - gutter) / 2, (h - 2*margin) / 5
    for i, r in enumerate(records):
        if i > 0 and i % 10 == 0: c.showPage()
        idx = i % 10; x_left = margin + (idx % 2) * (col_w + gutter); y_top = h - margin - (idx // 2) * row_h
        c.setStrokeGray(0.3); c.rect(x_left, y_top - row_h + 2*mm, col_w, row_h - 4*mm)
        c.setFont('Helvetica-Bold', 14); c.drawString(x_left + 4*mm, y_top - 10*mm, str(start_num + i))
        c.setFont('Helvetica-Bold', 35); c.drawString(x_left + 15*mm, y_top - 22*mm, r['flight'])
        tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        c.setFont('Helvetica-Bold', 25); c.drawString(x_left + 15*mm, y_top - 45*mm, f"{tdisp}  {r['dest']}")
    c.save(); target.seek(0); return target

# --- 사이드바 고정 시간 설정 ---
with st.sidebar:
    st.header("⚙️ Settings")
    # [고정] 사용자님이 지정하신 기본값 유지
    s_time = st.text_input("Start Time", value="04:55")
    e_time = st.text_input("End Time", value="05:00")
    label_start = st.number_input("Label Start Number", value=1, min_value=1)

st.markdown('<h1 style="color:white; font-size: 2.8rem; font-weight: bold;">Flight List Factory</h1>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Raw Text File", type=['txt'])
if uploaded_file:
    lines = uploaded_file.read().decode("utf-8").splitlines()
    all_recs = parse_raw_lines(lines)
    if all_recs:
        dates = sorted({r['dt'].date() for r in all_recs if r.get('dt')})
        day1, day2 = dates[0], dates[1] if len(dates) >= 2 else (dates[0] + timedelta(days=1))
        
        try:
            s_dt = datetime.combine(day1, datetime.strptime(s_time, '%H:%M').time())
            e_dt = datetime.combine(day2, datetime.strptime(e_time, '%H:%M').time())
        except:
            st.error("Check Time Format (HH:MM)")
            st.stop()
            
        filtered = [r for r in all_recs if r.get('dt') and (s_dt <= r['dt'] <= e_dt)]
        filtered.sort(key=lambda x: x['dt'])
        
        if filtered:
            st.success(f"Processed {len(filtered)} flights")
            col1, col2, col3 = st.columns(3)
            fn = f"List_{s_dt.strftime('%d-%m')}"
            col1.download_button("📥 One Page", build_docx_stream(filtered, s_dt, e_dt, mode='One Page'), f"{fn}_1P.docx")
            col2.download_button("📥 Two Pages", build_docx_stream(filtered, s_dt, e_dt, mode='Two Pages'), f"{fn}_2P.docx")
            col3.download_button("📥 PDF Labels", build_labels_stream(filtered, label_start), f"Labels_{fn}.pdf")
