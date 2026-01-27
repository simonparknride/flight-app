import streamlit as st
import re
import io
from math import ceil
from datetime import datetime, timedelta
from typing import List, Dict
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# ---------------------------
# 0. 설정 / 상수
# ---------------------------
st.set_page_config(page_title="Flight List Factory", layout="centered", initial_sidebar_state="expanded")

# 사용자께서 올려주신 정규식/맵 블록을 포함하고, 파서/필터에서 사용하도록 통합했습니다.
TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s[AP]M)\t([A-Z]{1,3}\d+[A-Z]?)\s*$", re.I)
DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
IATA_IN_PARENS = re.compile(r"\(([^)]+)\)")
PLANE_TYPES = ['A21N','A20N','A320','32Q','320','73H','737','74Y','77W','B77W','789','B789','359','A359','332','A332','333','A333','330','76V','77L','B38M','A388','772','B772','32X','77X','AT76','DH8C','SF3','32N']
PLANE_TYPE_PATTERN = re.compile(r"\b(" + "|".join(sorted(set(PLANE_TYPES), key=len, reverse=True)) + r")\b", re.IGNORECASE)
NORMALIZE_MAP = {
    '32q': 'A320', '320': 'A320', 'a320': 'A320', '32x': 'A320',
    '789': 'B789', 'b789': 'B789', '772': 'B772', 'b772': 'B772',
    '77w': 'B77W', 'b77w': 'B77W', '332': 'A332', 'a332': 'A332',
    '333': 'A333', 'a333': 'A333', '330': 'A330', 'a330': 'A330',
    '359': 'A359', 'a359': 'A359', '388': 'A388', 'a388': 'A388',
    '737': 'B737', '73h': 'B737', 'at7': 'AT76'
}
ALLOWED_AIRLINES = {"NZ","QF","JQ","CZ","CA","SQ","LA","IE","FX"}
NZ_DOMESTIC_IATA = {"AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO","WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"}
REGO_LIKE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9\-–—]*$")

# 스타일 (간단히)
st.markdown("""
    <style>
    .stApp { background-color: #000000; }
    [data-testid="stSidebar"] { background-color: #111111 !important; color: #fff; }
    .stMarkdown, p, h1, h2, h3, label { color: #ffffff !important; }
    div.stDownloadButton > button { background-color: #ffffff !important; color: #000000 !important; border-radius: 8px !important; padding: 0.6rem 1.2rem !important; font-weight: 800 !important; width: 100% !important; }
    div.stDownloadButton > button:hover { background-color: #60a5fa !important; color: #ffffff !important; border: 2px solid #60a5fa !important; }
    </style>
    """, unsafe_allow_html=True)

# ---------------------------
# 1. 파서: flightradar 스타일 텍스트
#    - TIME_LINE 정규식 적용
#    - 항공기 타입은 PLANE_TYPE_PATTERN으로 찾고 NORMALIZE_MAP으로 정규화
# ---------------------------
def parse_fdr_lines(lines: List[str]) -> List[Dict]:
    records = []
    current_date = None
    i = 0
    clean_lines = [l.rstrip() for l in lines]

    while i < len(clean_lines):
        line = clean_lines[i].strip()

        # 날짜 헤더 인식 (예: "Wednesday, Jan 28")
        if DATE_HEADER.match(line):
            m = re.search(r"([A-Za-z]{3})\s+(\d{1,2})", line)
            if m:
                # convert "Jan 28" -> "28 Jan"
                try:
                    day = int(m.group(2))
                    month = m.group(1)
                    current_date = f"{day} {month}"
                except:
                    current_date = None
            i += 1
            continue

        # TIME_LINE 사용하여 시간 및 편명 인식
        m = TIME_LINE.match(line)
        if m:
            time_str = m.group(1).strip()
            flight_code = m.group(2).strip().upper()
            dest = ""
            plane_type = ""
            reg = ""

            # 목적지: 다음 줄에서 괄호 안 IATA 우선 추출
            if i + 1 < len(clean_lines):
                dest_line = clean_lines[i+1].strip()
                ip = IATA_IN_PARENS.search(dest_line)
                if ip:
                    dest = ip.group(1).strip().upper()
                else:
                    # 괄호가 없으면 전체 목적지 텍스트에서 마지막 토큰 괄호식 IATA가 없을 때는 전체를 사용
                    dest = dest_line

            # 항공기 타입/레지스트리: 그다음 줄에서 추출
            if i + 2 < len(clean_lines):
                at_line = clean_lines[i+2].strip()
                # 타입 패턴 찾기
                tmatch = PLANE_TYPE_PATTERN.search(at_line)
                if tmatch:
                    raw_type = tmatch.group(1)
                    plane_type = NORMALIZE_MAP.get(raw_type.lower(), raw_type.upper())
                # 레지 찾기: 괄호 안 우선, 없으면 패턴으로 추출 시도
                rp = IATA_IN_PARENS.search(at_line)
                if rp:
                    candidate = rp.group(1).strip()
                    if REGO_LIKE.match(candidate):
                        reg = candidate
                else:
                    # fallback: 마지막 토큰에 하이픈 포함 등 레지형태 가능성 검사
                    tokens = re.split(r"[ \t]+", at_line)
                    if tokens:
                        last = tokens[-1]
                        if REGO_LIKE.match(last):
                            reg = last

            # datetime 생성: current_date가 있으면 사용, 없으면 오늘(임시)
            if current_date:
                try:
                    dt_obj = datetime.strptime(f"{current_date} 2026 {time_str}", "%d %b %Y %I:%M %p")
                except Exception:
                    # fallback: 시간만 파싱해서 임의 날짜 부여
                    try:
                        dt_temp = datetime.strptime(time_str, "%I:%M %p")
                        dt_obj = dt_temp.replace(year=2026, month=1, day=1)
                    except:
                        dt_obj = datetime.now()
            else:
                try:
                    dt_temp = datetime.strptime(time_str, "%I:%M %p")
                    dt_obj = dt_temp.replace(year=2026, month=1, day=1)
                except:
                    dt_obj = datetime.now()

            records.append({
                'dt': dt_obj,
                'time': dt_obj.strftime("%H:%M"),
                'flight': flight_code,
                'dest': dest,
                'type': plane_type,
                'reg': reg
            })

            # advance: 보통 3줄(시간, 목적지, 항공사+기종), 추가로 상태 줄이 있으면 1줄 더 건너뜀
            i += 3
            if i < len(clean_lines):
                status_line = clean_lines[i].strip()
                if status_line and re.search(r"(Scheduled|Estimated|Delayed|Canceled|Cancelled|\d{1,2}:\d{2})", status_line, re.I):
                    i += 1
            continue

        i += 1

    return records

# 콤마 형식 파서(보존, fallback)
def parse_lines(lines: List[str]) -> List[Dict]:
    records = []
    current_date = "26 Jan"
    for line in lines:
        line = line.strip()
        if not line:
            continue
        date_match = re.search(r"(\d{1,2}\s+[A-Za-z]{3})", line)
        if date_match and ":" not in line:
            current_date = date_match.group(1)
            continue
        parts = line.split(',')
        if len(parts) >= 5:
            try:
                row_date = parts[0].strip() if parts[0].strip() and parts[0].strip()[0].isdigit() else current_date
                time_val = parts[2].strip() if ":" in parts[2] else parts[1].strip()
                dt_obj = datetime.strptime(f"{row_date} 2026 {time_val}", "%d %b %Y %H:%M")
                records.append({
                    'dt': dt_obj,
                    'time': time_val,
                    'flight': parts[1].strip() if ":" in parts[2] else parts[0].strip(),
                    'dest': parts[3].strip(),
                    'type': parts[4].strip(),
                    'reg': parts[5].strip() if len(parts) > 5 else ""
                })
            except:
                continue
    return records

# ---------------------------
# 2. 필터링 (시간 + 항공사 + 국내선 옵션)
# ---------------------------
def filter_records(records: List[Dict], start_hm: str, end_hm: str, exclude_nz_domestic: bool = True):
    if not records:
        return [], None, None

    records.sort(key=lambda x: x['dt'])
    day1 = records[0]['dt'].date()
    start_dt = datetime.combine(day1, datetime.strptime(start_hm, '%H:%M').time())
    end_time_obj = datetime.strptime(end_hm, '%H:%M').time()
    end_dt = datetime.combine(day1, end_time_obj)
    if end_dt <= start_dt:
        end_dt += timedelta(days=1)

    def allowed(r):
        # 항공사 코드로 필터 (앞 두 글자)
        airline_prefix = re.match(r"^([A-Z]{2,3})", r['flight'])
        if not airline_prefix:
            return False
        prefix = airline_prefix.group(1).upper()
        if prefix not in ALLOWED_AIRLINES:
            return False
        # 시간 범위
        if not (start_dt <= r['dt'] < end_dt):
            return False
        # 국내선 제외 옵션
        if exclude_nz_domestic:
            # dest가 IATA 코드이면 uppercase 비교
            d = (r['dest'] or "").upper()
            # 일부 파서에서 목적지를 "Wellington (WLG)"처럼 전체 텍스트로 남겼을 수 있으므로,
            # 괄호를 파싱해 두었지만 만약 괄호 없으면 마지막 토큰이 IATA일 가능성 검사
            if len(d) == 3 and d in NZ_DOMESTIC_IATA:
                return False
            # 만약 "Wellington"처럼 풀네임이 온 경우엔 체크 불가 -> 허용(혹은 개선 필요)
        return True

    filtered = [r for r in records if allowed(r)]
    filtered.sort(key=lambda x: x['dt'])
    return filtered, start_dt, end_dt

# ---------------------------
# 3. DOCX / PDF 생성기 (기존과 동일한 레이아웃)
# ---------------------------
def build_docx_two_pages_stream(records, start_dt, end_dt):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Inches(0.3)
    section.left_margin = section.right_margin = Inches(0.5)

    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_f = footer_para.add_run("created by Simon Park'nRide's Flight List Factory 2026")
    run_f.font.size = Pt(10)
    run_f.font.color.rgb = RGBColor(128, 128, 128)

    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title_para.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}")
    run_title.bold = True
    run_title.font.size = Pt(16)

    n = len(records)
    if n == 0:
        target = io.BytesIO()
        doc.save(target)
        target.seek(0)
        return target

    half = ceil(n / 2)
    first = records[:half]
    second = records[half:]

    def add_table_for_block(block):
        table = doc.add_table(rows=0, cols=5)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, r in enumerate(block):
            row = table.add_row()
            vals = [r['flight'], r['time'], r['dest'], r['type'], r['reg']]
            for j, val in enumerate(vals):
                cell = row.cells[j]
                if i % 2 == 1:
                    tcPr = cell._tc.get_or_add_tcPr()
                    shd = OxmlElement('w:shd')
                    shd.set(qn('w:val'), 'clear')
                    shd.set(qn('w:fill'), 'D9D9D9')
                    tcPr.append(shd)
                para = cell.paragraphs[0]
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = para.add_run(str(val))
                run.font.size = Pt(14)

    add_table_for_block(first)
    doc.add_page_break()
    title2 = doc.add_paragraph()
    title2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title2 = title2.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}")
    run_title2.bold = True
    run_title2.font.size = Pt(16)
    add_table_for_block(second)

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

def build_docx_onepage_stream(records, start_dt, end_dt):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Inches(0.3)
    section.left_margin = section.right_margin = Inches(0.4)

    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_f = footer_para.add_run("created by Simon Park'nRide's Flight List Factory 2026")
    run_f.font.size = Pt(9)
    run_f.font.color.rgb = RGBColor(128, 128, 128)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_head = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}")
    run_head.bold = True
    run_head.font.size = Pt(16)

    n = len(records)
    if n == 0:
        target = io.BytesIO()
        doc.save(target)
        target.seek(0)
        return target

    half = ceil(n / 2)
    left = records[:half]
    right = records[half:]
    rows = max(len(left), len(right))

    table = doc.add_table(rows=rows, cols=10)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for row_idx in range(rows):
        # Left block
        if row_idx < len(left):
            rec = left[row_idx]
            vals = [rec['flight'], rec['time'], rec['dest'], rec['type'], rec['reg']]
            for col_offset, val in enumerate(vals):
                cell = table.rows[row_idx].cells[col_offset]
                if row_idx % 2 == 1:
                    tcPr = cell._tc.get_or_add_tcPr()
                    shd = OxmlElement('w:shd')
                    shd.set(qn('w:val'), 'clear')
                    shd.set(qn('w:fill'), 'D9D9D9')
                    tcPr.append(shd)
                para = cell.paragraphs[0]
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = para.add_run(str(val))
                run.font.size = Pt(11)
        else:
            for col_offset in range(5):
                cell = table.rows[row_idx].cells[col_offset]
                cell.text = ""

        # Right block
        if row_idx < len(right):
            rec = right[row_idx]
            vals = [rec['flight'], rec['time'], rec['dest'], rec['type'], rec['reg']]
            for col_offset, val in enumerate(vals):
                cell = table.rows[row_idx].cells[5 + col_offset]
                if (half + row_idx) % 2 == 1:
                    tcPr = cell._tc.get_or_add_tcPr()
                    shd = OxmlElement('w:shd')
                    shd.set(qn('w:val'), 'clear')
                    shd.set(qn('w:fill'), 'D9D9D9')
                    tcPr.append(shd)
                para = cell.paragraphs[0]
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = para.add_run(str(val))
                run.font.size = Pt(11)
        else:
            for col_offset in range(5):
                cell = table.rows[row_idx].cells[5 + col_offset]
                cell.text = ""

    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

def build_labels_stream(records, start_num):
    target = io.BytesIO()
    c = canvas.Canvas(target, pagesize=A4)
    w, h = A4
    margin, gutter = 15*mm, 6*mm
    col_w, row_h = (w - 2*margin - gutter) / 2, (h - 2*margin) / 5

    for i, r in enumerate(records):
        if i > 0 and i % 10 == 0:
            c.showPage()
        idx = i % 10
        x_left = margin + (idx % 2) * (col_w + gutter)
        y_top = h - margin - (idx // 2) * row_h

        c.setStrokeGray(0.3); c.setLineWidth(0.2); c.rect(x_left, y_top - row_h + 2*mm, col_w, row_h - 4*mm)
        c.setLineWidth(0.5); c.rect(x_left + 3*mm, y_top - 12*mm, 8*mm, 8*mm)
        c.setFont('Helvetica-Bold', 14); c.drawCentredString(x_left + 7*mm, y_top - 9.5*mm, str(start_num + i))
        c.setFont('Helvetica-Bold', 18); c.drawRightString(x_left + col_w - 4*mm, y_top - 11*mm, r['dt'].strftime('%d %b'))
        c.setFont('Helvetica-Bold', 38); c.drawString(x_left + 15*mm, y_top - 21*mm, r['flight'])
        c.setFont('Helvetica-Bold', 23); c.drawString(x_left + 15*mm, y_top - 33*mm, r['dest'])
        c.setFont('Helvetica-Bold', 29); c.drawString(x_left + 15*mm, y_top - 47*mm, r['time'])
        c.setFont('Helvetica', 13); c.drawRightString(x_left + col_w - 6*mm, y_top - row_h + 12*mm, r['type'])
        c.drawRightString(x_left + col_w - 6*mm, y_top - row_h + 7*mm, r['reg'])

    c.save()
    target.seek(0)
    return target

# ---------------------------
# 4. UI: 사이드바, 업로드 처리, 버튼: 기존 Word(two pages) / NEW One-page / PDF
# ---------------------------
with st.sidebar:
    st.header("⚙️ Settings")
    s_time = st.text_input("Start Time", value="04:55")
    e_time = st.text_input("End Time", value="04:50")
    label_start = st.number_input("Label Start Number", value=1, min_value=1)
    exclude_nz_domestic = st.checkbox("Exclude NZ domestic flights (국내선 제외)", value=True)

st.markdown('<div class="top-left-container"><a href="#" target="_blank">Import Raw Text</a><a href="#" target="_blank">Export Raw Text</a></div>', unsafe_allow_html=True)
st.markdown('<div class="main-title">Simon Park\'nRide\'s<br><span class="sub-title">Flight List Factory</span></div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Raw Text File", type=['txt', 'docx'])

if uploaded_file:
    uploaded_file.seek(0)
    content_text = ""
    if uploaded_file.name.lower().endswith('.docx'):
        try:
            raw_bytes = uploaded_file.read()
            docx_stream = io.BytesIO(raw_bytes)
            docx_obj = Document(docx_stream)
            paragraphs = [p.text for p in docx_obj.paragraphs]
            content_text = "\n".join(paragraphs)
        except Exception:
            # fallback to text decode
            uploaded_file.seek(0)
            raw_bytes = uploaded_file.read()
            try:
                content_text = raw_bytes.decode("utf-8")
            except:
                content_text = raw_bytes.decode("latin-1", errors="ignore")
    else:
        try:
            content_text = uploaded_file.read().decode("utf-8")
        except Exception:
            uploaded_file.seek(0)
            content_text = uploaded_file.read().decode("latin-1", errors="ignore")

    lines = content_text.splitlines()
    all_recs = parse_fdr_lines(lines)
    if not all_recs:
        all_recs = parse_lines(lines)

    if all_recs:
        filtered, s_dt, e_dt = filter_records(all_recs, s_time, e_time, exclude_nz_domestic=exclude_nz_domestic)
        if filtered:
            st.success(f"Processed {len(filtered)} flights. (exclude_nz_domestic={exclude_nz_domestic})")
            col1, col2, col3 = st.columns([1,1,1])
            fn = f"List_{s_dt.strftime('%d-%m')}"
            # Left: 기존 Word (Two Pages)
            col1.download_button("📥 Download DOCX (Two Pages)", build_docx_two_pages_stream(filtered, s_dt, e_dt), f"{fn}_twopages.docx")
            # Middle: One-page two-column (새 버튼)
            col2.download_button("📄 Download DOCX (One Page, 2 Columns)", build_docx_onepage_stream(filtered, s_dt, e_dt), f"{fn}_onepage.docx")
            # Right: PDF labels
            col3.download_button("🏷️ Download PDF Labels", build_labels_stream(filtered, label_start), f"Labels_{fn}.pdf")

            st.table([{'No': label_start+i, 'Flight': r['flight'], 'Time': r['time'], 'Dest': r['dest'], 'Reg': r['reg'], 'Type': r['type']} for i, r in enumerate(filtered)])
        else:
            st.warning("No flights match the filter criteria. Please check Start/End Time or uncheck the domestic-exclude option.")
    else:
        st.error("Could not parse data. Please check the file format.")
