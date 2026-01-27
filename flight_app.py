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

# --- 1. UI 설정 및 버튼 가독성 강화 ---
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

# --- 2. 파싱 및 필터링 로직 (쉼표 구분 데이터 완벽 지원) ---
ALLOWED_AIRLINES = {"NZ","QF","JQ","CZ","CA","SQ","LA","IE"}
NZ_DOMESTIC_IATA = {"AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO","WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"}

def parse_lines(lines: List[str]) -> List[Dict]:
    records = []
    current_date = "26 Jan" # 기본값
    
    for line in lines:
        line = line.strip()
        if not line: continue
        
        # 날짜 헤더 인식 (예: 26 Jan 2026)
        date_match = re.search(r"(\d{1,2}\s+[A-Za-z]{3})", line)
        if date_match and ":" not in line:
            current_date = date_match.group(1)
            continue
            
        # 쉼표 구분 데이터 파싱 (List_26-01_1p 형식 지원)
        parts = line.split(',')
        if len(parts) >= 5:
            try:
                # 첫 칸이 날짜면 업데이트, 아니면 유지
                row_date = parts[0].strip() if parts[0].strip() and parts[0].strip()[0].isdigit() else current_date
                # 시간 위치 확인 (04:55 형태)
                time_val = parts[2].strip() if ":" in parts[2] else parts[1].strip()
                
                # datetime 객체 생성
                dt_obj = datetime.strptime(f"{row_date} 2026 {time_val}", "%d %b %Y %H:%M")
                
                records.append({
                    'dt': dt_obj,
                    'time': time_val,
                    'flight': parts[1].strip() if ":" in parts[2] else parts[0].strip(),
                    'dest': parts[3].strip(),
                    'type': parts[4].strip(),
                    'reg': parts[5].strip() if len(parts) > 5 else ""
                })
            except: continue
    return records

def filter_records(records, start_hm, end_hm):
    if not records: return [], None, None
    
    # 설정 시간으로 필터링
    day1 = records[0]['dt'].date()
    start_dt = datetime.combine(day1, datetime.strptime(start_hm, '%H:%M').time())
    
    end_time_obj = datetime.strptime(end_hm, '%H:%M').time()
    end_dt = datetime.combine(day1, end_time_obj)
    if end_dt <= start_dt:
        end_dt += timedelta(days=1)
        
    filtered = [r for r in records if r['flight'][:2] in ALLOWED_AIRLINES and (start_dt <= r['dt'] < end_dt)]
    filtered.sort(key=lambda x: x['dt'])
    return filtered, start_dt, end_dt

# --- 3. DOCX 생성 (Footer 및 제브라 무늬) ---

def build_single_page_docx_stream(records, start_dt, end_dt):
    doc = Document()
    font_name = 'Arial' # 범용 폰트로 설정 (시스템에 따라 조정 가능)
    
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Inches(0.3)
    section.bottom_margin = Inches(0.3)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

    # 2단 설정 (Column setting)
    sectPr = section._sectPr
    cols = OxmlElement('w:cols')
    cols.set(qn('w:num'), '2')
    cols.set(qn('w:space'), '360') # 0.25 inch space
    sectPr.append(cols)

    # Footer 설정: Simon Park'nRide's Factory 2026
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_f = footer_para.add_run("created by Simon Park'nRide's Flight List Factory 2026")
    run_f.font.size = Pt(8) # 폰트 크기 축소
    run_f.font.color.rgb = RGBColor(128, 128, 128)

    # 타이틀
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_head = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')} (Single Page)")
    run_head.bold = True
    run_head.font.size = Pt(14) # 폰트 크기 축소

    # 표 생성
    # 5열 대신 4열로 변경 (Reg. 제거)
    table = doc.add_table(rows=0, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # 표 스타일 조정 (셀 간격 줄이기)
    table.style = 'Table Grid'
    
    for i, r in enumerate(records):
        row = table.add_row()
        # Reg. (r['reg']) 제거
        vals = [r['flight'], r['time'], r['dest'], r['type']]
        for j, val in enumerate(vals):
            cell = row.cells[j]
            # 제브라 무늬 (홀수 행 배경색)
            if i % 2 == 1:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd')
                shd.set(qn('w:val'), 'clear')
                shd.set(qn('w:fill'), 'EFEFEF') # 더 밝은 회색으로 변경
                tcPr.append(shd)
            
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.space_before = Pt(0)
            para.space_after = Pt(0)
            para.line_spacing_rule = WD_LINE_SPACING.SINGLE
            
            run = para.add_run(str(val))
            run.font.size = Pt(10) # 폰트 크기 축소
            
    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target


def build_docx_stream(records, start_dt, end_dt):
    doc = Document()
    font_name = 'Arial' # 범용 폰트로 설정 (시스템에 따라 조정 가능)
    
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Inches(0.3)
    section.left_margin = section.right_margin = Inches(0.5)

    # Footer 설정: Simon Park'nRide's Factory 2026
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_f = footer_para.add_run("created by Simon Park'nRide's Flight List Factory 2026")
    run_f.font.size = Pt(10)
    run_f.font.color.rgb = RGBColor(128, 128, 128)

    # 타이틀
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_head = p.add_run(f"{start_dt.strftime('%d')}-{end_dt.strftime('%d')} {start_dt.strftime('%b')}")
    run_head.bold = True
    run_head.font.size = Pt(16)

    # 표 생성
    table = doc.add_table(rows=0, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    for i, r in enumerate(records):
        row = table.add_row()
        vals = [r['flight'], r['time'], r['dest'], r['type'], r['reg']]
        for j, val in enumerate(vals):
            cell = row.cells[j]
            # 제브라 무늬 (홀수 행 배경색)
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
            
    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- 4. PDF 레이블 생성 ---
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

# --- 5. 실행 및 사이드바 ---
with st.sidebar:
    st.header("⚙️ Settings")
    s_time = st.text_input("Start Time", value="04:55")
    e_time = st.text_input("End Time", value="04:50")
    label_start = st.number_input("Label Start Number", value=1, min_value=1)

st.markdown('<div class="top-left-container"><a href="https://www.flightradar24.com/data/airports/akl/arrivals" target="_blank">Import Raw Text</a><a href="https://www.flightradar24.com/data/airports/akl/departures" target="_blank">Export Raw Text</a></div>', unsafe_allow_html=True)
st.markdown('<div class="main-title">Simon Park\'nRide\'s<br><span class="sub-title">Flight List Factory</span></div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Raw Text File", type=['txt', 'docx'])

if uploaded_file:
    # 텍스트 데이터 추출
    raw_content = uploaded_file.read().decode("utf-8")
    all_recs = parse_lines(raw_content.splitlines())
    
    if all_recs:
        filtered, s_dt, e_dt = filter_records(all_recs, s_time, e_time)
        if filtered:
            st.success(f"Processed {len(filtered)} flights.")
            col1, col2, col3 = st.columns(3)
            fn = f"List_{s_dt.strftime('%d-%m')}"
            
            col1.download_button("📥 Download DOCX List", build_docx_stream(filtered, s_dt, e_dt), f"{fn}.docx")
            col2.download_button("📄 Download 1-Page DOCX", build_single_page_docx_stream(filtered, s_dt, e_dt), f"1Page_{fn}.docx")
            col3.download_button("🏷️ Download PDF Labels", build_labels_stream(filtered, label_start), f"Labels_{fn}.pdf")
            
            st.table([{'No': label_start+i, 'Flight': r['flight'], 'Time': r['time'], 'Dest': r['dest'], 'Reg': r['reg']} for i, r in enumerate(filtered)])
        else:
            st.warning("No flights match the filter criteria. Please check Start/End Time.")
    else:
        st.error("Could not parse data. Please check the file format.")
