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

# --- 1. UI 설정 (링크 복구 및 입력창 활성화) ---
st.set_page_config(page_title="Flight List Factory", layout="centered")

st.markdown("""
    <style>
    /* 배경색 고정 */
    .stApp { background-color: #000000; }
    
    /* 사이드바 스타일 및 입력 활성화 */
    [data-testid="stSidebar"] {
        background-color: #111111 !important;
        min-width: 320px !important;
    }
    
    /* 사라졌던 상단 링크 스타일 */
    .top-links {
        padding: 10px 0;
        margin-bottom: 20px;
        border-bottom: 1px solid #333;
    }
    .top-links a {
        color: #60a5fa !important;
        font-size: 1.1rem;
        text-decoration: underline;
        margin-right: 15px;
        display: block;
        margin-bottom: 8px;
    }

    /* 사이드바 라벨 크기 및 색상 */
    [data-testid="stSidebar"] label p {
        font-size: 1.3rem !important; /* 글자 크기 더 확대 */
        color: #ffffff !important;
        font-weight: bold !important;
    }
    
    /* 입력창 내부: 클릭 가능하도록 레이어 조정 및 가독성 확보 */
    [data-testid="stSidebar"] input {
        font-size: 1.2rem !important;
        color: #000000 !important;
        background-color: #ffffff !important;
        height: 45px !important;
        pointer-events: auto !important; /* 클릭 방해 요소 제거 */
    }

    /* 다운로드 버튼 */
    div.stDownloadButton > button {
        background-color: #000000 !important; color: #ffffff !important;           
        border: 1px solid #ffffff !important;
        width: 100% !important; height: 3.5rem !important;
        font-weight: bold;
    }
    div.stDownloadButton > button:hover {
        background-color: #ffffff !important; color: #000000 !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 상단 링크 및 타이틀 ---
# 사라졌던 링크를 메인 화면 최상단에 다시 배치했습니다.
st.markdown("""
    <div class="top-links">
        <a href="#">🔗 Import Raw Text</a>
        <a href="#">🔗 Export Raw Text</a>
    </div>
    """, unsafe_allow_html=True)

st.markdown('<h1 style="color:white; font-size: 2.8rem;">Flight List Factory</h1>', unsafe_allow_html=True)

# --- 3. 사이드바 설정 (값 고정 및 수정 가능) ---
with st.sidebar:
    st.header("⚙️ Settings")
    # 시간 값을 04:55 / 05:00으로 다시 고정했습니다.
    s_time = st.text_input("Start Time", value="04:55", key="st_time")
    e_time = st.text_input("End Time", value="05:00", key="et_time")
    label_start = st.number_input("Label Start Number", value=1, min_value=1, key="l_start")

# --- 4. 파싱 및 문서 생성 로직 (기능 유지) ---
def parse_raw_lines(lines):
    TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s[AP]M)\t([A-Z]{2}\d+[A-Z]?)\s*$")
    DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
    IATA_IN_PAREns = re.compile(r"\(([^)]+)\)")
    records = []
    current_date = None
    for i, line in enumerate(lines):
        line = line.strip()
        if DATE_HEADER.match(line):
            try: current_date = datetime.strptime(line + ' 2026', '%A, %b %d %Y').date()
            except: current_date = None
            continue
        m = TIME_LINE.match(line)
        if m and current_date is not None:
            time_str, flight = m.groups()
            dest_line = lines[i+1].strip() if i+1 < len(lines) else ''
            m2 = IATA_IN_PAREns.search(dest_line)
            dest_iata = (m2.group(1).strip() if m2 else '').upper()
            carrier_line = lines[i+2].rstrip('\n') if i+2 < len(lines) else ''
            plane_type = "B789" if "789" in carrier_line else ("A320" if any(x in carrier_line for x in ["320","32Q"]) else "")
            reg = ""
            parens = IATA_IN_PAREns.findall(carrier_line)
            if parens: reg = parens[-1].strip()
            try: dep_dt = datetime.strptime(f"{current_date} {time_str}", '%Y-%m-%d %I:%M %p')
            except: dep_dt = None
            records.append({'dt': dep_dt, 'time': time_str, 'flight': flight, 'dest': dest_iata, 'type': plane_type, 'reg': reg})
    return records

def build_docx_stream(records, start_dt, end_dt, mode='Two Pages'):
    doc = Document()
    section = doc.sections[0]
    if mode == 'One Page':
        section.left_margin = section.right_margin = Inches(0.5); section.top_margin = section.bottom_margin = Inches(0.15)
        f_size, h_size = Pt(7.5), Pt(11); align_h, align_t = WD_ALIGN_PARAGRAPH.LEFT, WD_TABLE_ALIGNMENT.CENTER
    else:
        section.left_margin = section.right_margin = Inches(1.0); section.top_margin = section.bottom_margin = Inches(1.0)
        f_size, h_size = Pt(14), Pt(16); align_h, align_t = WD_ALIGN_PARAGRAPH.LEFT, WD_TABLE_ALIGNMENT.LEFT
    p = doc.add_paragraph(); p.alignment = align_h
    run = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}"); run.bold = True; run.font.size = h_size
    table = doc.add_table(rows=0, cols=5); table.alignment = align_t
    for i, r in enumerate(records):
        row = table.add_row()
        if mode == 'One Page':
            tr = row._tr; trPr = tr.get_or_add_trPr(); trH = OxmlElement('w:trHeight'); trH.set(qn('w:val'), '180'); trH.set(qn('w:hRule'), 'atLeast'); trPr.append(trH)
        tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        for j, val in enumerate([r['flight'], tdisp, r['dest'], r['type'], r['reg']]):
            cell = row.cells[j]
            if i % 2 == 1:
                tcPr = cell._tc.get_or_add_tcPr(); shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9'); tcPr.append(shd)
            para = cell.paragraphs[0]; para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_c = para.add_run(str(val)); run_c.font.size = f_size
    target = io.BytesIO(); doc.save(target); target.seek(0); return target

def build_labels_stream(records, start_num):
    target = io.BytesIO(); c = canvas.Canvas(target, pagesize=A4)
    w, h = A4; margin, gutter = 15*mm, 6*mm; col_w, row_h = (w - 2*margin - gutter) / 2, (h - 2*margin) / 5
    for i, r in enumerate(records):
        if i > 0 and i % 10 == 0: c.showPage()
        idx = i % 10; x_left = margin + (idx % 2) * (col_w + gutter); y_top = h - margin - (idx // 2) * row_h
        c.setStrokeGray(0.3); c.rect(x_left, y_top - row_h + 2*mm, col_w, row_h - 4*mm)
        c.setFont('Helvetica-Bold', 14); c.drawString(x_left + 4*mm, y_top - 10*mm, str(start_num + i))
        c.setFont('Helvetica-Bold', 35); c.drawString(x_left + 15*mm, y_top - 22*mm, r['flight'])
        tdisp = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
        c.setFont('Helvetica-Bold', 25); c.drawString(x_left + 15*mm, y_top - 45*mm, f"{tdisp}  {r['dest']}")
    c.save(); target.seek(0); return target

# --- 5. 파일 업로드 및 결과 출력 ---
uploaded_file = st.file_uploader("Upload Raw Text File", type=['txt'])
if uploaded_file:
    lines = uploaded_file.read().decode("utf-8").splitlines()
    all_recs = parse_raw_lines(lines)
    if all_recs:
        dates = sorted({r['dt'].date() for r in all_recs if r.get('dt')})
        day1 = dates[0]; day2 = dates[1] if len(dates) >= 2 else (dates[0] + timedelta(days=1))
        try:
            s_dt = datetime.combine(day1, datetime.strptime(s_time, '%H:%M').time())
            e_dt = datetime.combine(day2, datetime.strptime(e_time, '%H:%M').time())
            filtered = [r for r in all_recs if r.get('dt') and (s_dt <= r['dt'] <= e_dt)]
            filtered.sort(key=lambda x: x['dt'])
            if filtered:
                st.success(f"Processed {len(filtered)} flights")
                col1, col2, col3 = st.columns(3)
                fn = f"List_{s_dt.strftime('%d-%m')}"
                col1.download_button("📥 One Page", build_docx_stream(filtered, s_dt, e_dt, mode='One Page'), f"{fn}_1P.docx")
                col2.download_button("📥 Two Pages", build_docx_stream(filtered, s_dt, e_dt, mode='Two Pages'), f"{fn}_2P.docx")
                col3.download_button("📥 PDF Labels", build_labels_stream(filtered, label_start), f"Labels_{fn}.pdf")
        except ValueError:
            st.warning("Please enter time in HH:MM format.")
