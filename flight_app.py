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

# --- UI 설정 ---
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
    div.stDownloadButton > button:hover * { color: #ffffff !important; }

    .top-left-container { text-align: left; padding-top: 10px; margin-bottom: 20px; }
    .top-left-container a { font-size: 1.1rem; color: #ffffff !important; text-decoration: underline; display: block; margin-bottom: 5px;}
    .main-title { font-size: 3rem; font-weight: 800; color: #ffffff; line-height: 1.1; margin-bottom: 0.5rem; }
    .sub-title { font-size: 2.5rem; font-weight: 400; color: #60a5fa; }
    </style>
    """, unsafe_allow_html=True)

ALLOWED_AIRLINES = {"NZ","QF","JQ","CZ","CA","SQ","LA","IE","3C","DL","MU","QR","AA","UA","KE","CX","EK","MH","SB","FJ","HU","LA"}

# --- flightradar24 스타일 텍스트 파서 (제공 샘플 포맷) ---
def parse_fdr_lines(lines: List[str]) -> List[Dict]:
    records = []
    current_date = None  # e.g., "28 Jan"
    i = 0
    clean_lines = [l.rstrip() for l in lines]
    while i < len(clean_lines):
        line = clean_lines[i].strip()
        # 날짜 헤더 예: "Wednesday, Jan 28" 또는 "Thursday, Jan 29"
        dm = re.search(r"([A-Za-z]{3,9}),?\s*([A-Za-z]{3})\s+(\d{1,2})", line)
        if dm:
            try:
                current_date = f"{int(dm.group(3))} {dm.group(2)}"
            except:
                current_date = None
            i += 1
            continue

        tf_match = re.match(r"^(\d{1,2}:\d{2}\s*(?:AM|PM))\s*[ \t]+([A-Z0-9]+)", line, re.I)
        if tf_match:
            time_str = tf_match.group(1).strip()
            flight_code = tf_match.group(2).strip()
            dest = ""
            typ = ""
            reg = ""
            if i + 1 < len(clean_lines):
                dest_line = clean_lines[i + 1].strip()
                m = re.search(r"\((\w{3})\)", dest_line)
                if m:
                    dest = m.group(1)
                else:
                    dest = dest_line
            if i + 2 < len(clean_lines):
                at_line = clean_lines[i + 2].strip()
                parts = re.split(r"\t+", at_line)
                if len(parts) >= 2:
                    typ = parts[1].split('(')[0].strip()
                    rm = re.search(r"\(([^)]+)\)", at_line)
                    if rm:
                        reg = rm.group(1)
                else:
                    rm = re.search(r"\(([^)]+)\)", at_line)
                    if rm:
                        reg = rm.group(1)
                    tm = re.search(r"([A-Za-z0-9]{2,5}\d{0,2})", at_line)
                    if tm:
                        typ = tm.group(0)
            if current_date:
                try:
                    dt_obj = datetime.strptime(f"{current_date} 2026 {time_str}", "%d %b %Y %I:%M %p")
                except Exception:
                    try:
                        dt_obj = datetime.strptime(time_str, "%I:%M %p")
                        dt_obj = dt_obj.replace(year=2026, month=1, day=1)
                    except:
                        dt_obj = datetime.now()
            else:
                try:
                    dt_obj = datetime.strptime(time_str, "%I:%M %p")
                    dt_obj = dt_obj.replace(year=2026, month=1, day=1)
                except:
                    dt_obj = datetime.now()

            records.append({
                'dt': dt_obj,
                'time': dt_obj.strftime("%H:%M"),
                'flight': flight_code,
                'dest': dest,
                'type': typ,
                'reg': reg
            })
            # Advance index: 3 lines typical, and optionally a status line
            i += 3
            if i < len(clean_lines):
                status_line = clean_lines[i].strip()
                if status_line and re.search(r"(Scheduled|Estimated|Delayed|Canceled|Cancelled|\d{1,2}:\d{2})", status_line, re.I):
                    i += 1
            continue

        i += 1

    return records

# --- comma-separated parser (fallback) ---
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

# --- 필터링 ---
def filter_records(records, start_hm, end_hm):
    if not records:
        return [], None, None
    records.sort(key=lambda x: x['dt'])
    day1 = records[0]['dt'].date()
    start_dt = datetime.combine(day1, datetime.strptime(start_hm, '%H:%M').time())
    end_time_obj = datetime.strptime(end_hm, '%H:%M').time()
    end_dt = datetime.combine(day1, end_time_obj)
    if end_dt <= start_dt:
        end_dt += timedelta(days=1)
    filtered = [r for r in records if (r['flight'][:2] in ALLOWED_AIRLINES) and (start_dt <= r['dt'] < end_dt)]
    filtered.sort(key=lambda x: x['dt'])
    return filtered, start_dt, end_dt

# --- DOCX: Two Pages ---
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

# --- DOCX: One-Page Two-Column ---
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
        # Left block (cols 0-4)
        if row_idx < len(left):
            rec = left[row_idx]
            orig_index = row_idx
            vals = [rec['flight'], rec['time'], rec['dest'], rec['type'], rec['reg']]
            for col_offset, val in enumerate(vals):
                cell = table.rows[row_idx].cells[col_offset]
                if orig_index % 2 == 1:
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

        # Right block (cols 5-9)
        if row_idx < len(right):
            rec = right[row_idx]
            orig_index = half + row_idx
            vals = [rec['flight'], rec['time'], rec['dest'], rec['type'], rec['reg']]
            for col_offset, val in enumerate(vals):
                cell = table.rows[row_idx].cells[5 + col_offset]
                if orig_index % 2 == 1:
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

# --- PDF 레이블 생성 ---
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

# --- 사이드바 및 업로드 처리 ---
with st.sidebar:
    st.header("⚙️ Settings")
    s_time = st.text_input("Start Time", value="04:55")
    e_time = st.text_input("End Time", value="04:50")
    label_start = st.number_input("Label Start Number", value=1, min_value=1)

st.markdown('<div class="top-left-container"><a href="#" target="_blank">Import Raw Text</a><a href="#" target="_blank">Export Raw Text</a></div>', unsafe_allow_html=True)
st.markdown('<div class="main-title">Simon Park\'nRide\'s<br><span class="sub-title">Flight List Factory</span></div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Raw Text File", type=['txt', 'docx'])

if uploaded_file:
    content_text = ""
    # Make sure we read bytes once and use a BytesIO for python-docx when needed
    try:
        uploaded_file.seek(0)
    except Exception:
        pass

    if uploaded_file.name.lower().endswith('.docx'):
        # Use BytesIO to feed python-docx
        try:
            uploaded_file.seek(0)
            raw_bytes = uploaded_file.read()
            docx_stream = io.BytesIO(raw_bytes)
            docx_obj = Document(docx_stream)
            paragraphs = [p.text for p in docx_obj.paragraphs]
            content_text = "\n".join(paragraphs)
        except Exception as e:
            # fallback: try decode as text
            try:
                uploaded_file.seek(0)
                content_text = raw_bytes.decode("utf-8")
            except Exception:
                content_text = raw_bytes.decode("latin-1", errors="ignore")
    else:
        try:
            uploaded_file.seek(0)
            content_text = uploaded_file.read().decode("utf-8")
        except Exception:
            uploaded_file.seek(0)
            content_text = uploaded_file.read().decode("latin-1", errors="ignore")

    lines = content_text.splitlines()
    all_recs = parse_fdr_lines(lines)
    if not all_recs:
        all_recs = parse_lines(lines)

    if all_recs:
        filtered, s_dt, e_dt = filter_records(all_recs, s_time, e_time)
        if filtered:
            st.success(f"Processed {len(filtered)} flights.")
            col1, col2, col3 = st.columns([1,1,1])
            fn = f"List_{s_dt.strftime('%d-%m')}"
            col1.download_button("📥 Download DOCX (Two Pages)", build_docx_two_pages_stream(filtered, s_dt, e_dt), f"{fn}_twopages.docx")
            col2.download_button("📄 Download DOCX (One Page, 2 Columns)", build_docx_onepage_stream(filtered, s_dt, e_dt), f"{fn}_onepage.docx")
            col3.download_button("🏷️ Download PDF Labels", build_labels_stream(filtered, label_start), f"Labels_{fn}.pdf")
            st.table([{'No': label_start+i, 'Flight': r['flight'], 'Time': r['time'], 'Dest': r['dest'], 'Reg': r['reg']} for i, r in enumerate(filtered)])
        else:
            st.warning("No flights match the filter criteria. Please check Start/End Time.")
    else:
        st.error("Could not parse data. Please check the file format.")
