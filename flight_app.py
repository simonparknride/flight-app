import streamlit as st
import re
import io
from datetime import datetime, timedelta, time as dtime
from typing import List, Dict
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# --- 1. UI 디자인 및 테마 설정 ---
st.set_page_config(page_title="Flight List Factory", layout="centered", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .stApp { background-color: #000000; }
    [data-testid="stSidebar"] { background-color: #111111 !important; }
    .stMarkdown, p, h1, h2, h3, label { color: #ffffff !important; }
    
    /* 다운로드 버튼 스타일 */
    div.stDownloadButton > button {
        background-color: #ffffff !important;
        color: #000000 !important;
        border: 2px solid #ffffff !important;
        border-radius: 8px !important;
        padding: 0.6rem 1.2rem !important;
        font-weight: 800 !important;
        width: 100% !important;
    }
    div.stDownloadButton > button:hover {
        background-color: #60a5fa !important;
        color: #ffffff !important;
        border: 2px solid #60a5fa !important;
    }

    .main-title { font-size: 3rem; font-weight: 800; color: #ffffff; line-height: 1.1; margin-bottom: 0.5rem; }
    .sub-title { font-size: 2.5rem; font-weight: 400; color: #60a5fa; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 데이터 파싱 및 필터링 (쉼표 데이터 지원) ---
ALLOWED_AIRLINES = {"NZ", "QF", "JQ", "CZ", "CA", "SQ", "LA", "IE", "FX"}
NZ_DOMESTIC_IATA = {"AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO","WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"}

def parse_raw_lines(lines: List[str], year: int) -> List[Dict]:
    """쉼표로 구분된 데이터 행을 읽어 리스트로 변환합니다."""
    records = []
    current_date_str = ""
    
    for line in lines:
        line = line.strip()
        if not line: continue
        
        parts = line.split(',')
        
        # 1. 날짜 헤더 인식 (예: 26 Jan)
        if len(parts) > 0 and parts[0] and not any(c.isdigit() for c in parts[0].split()[-1]):
            current_date_str = parts[0].strip()
            continue
            
        # 2. 항공편 데이터 파싱 (쉼표 5개 이상일 때)
        if len(parts) >= 5:
            try:
                # 첫 칸에 날짜가 있으면 사용, 없으면 직전 날짜 사용
                row_date = parts[0].strip() if parts[0].strip() and parts[0].strip()[0].isdigit() else current_date_str
                flight = parts[1].strip()
                time_val = parts[2].strip()
                dest = parts[3].strip()
                p_type = parts[4].strip()
                reg = parts[5].strip() if len(parts) > 5 else ""
                
                # 날짜+시간 객체 생성
                dt_obj = datetime.strptime(f"{row_date} {year} {time_val}", "%d %b %Y %H:%M")
                
                records.append({
                    'dt': dt_obj, 
                    'time': time_val, 
                    'flight': flight.upper(),
                    'dest': dest.upper(), 
                    'type': p_type, 
                    'reg': reg
                })
            except: continue
    return records

def filter_records(records, s_time, e_time):
    """설정된 시간 범위와 필터 조건에 맞춰 데이터를 거릅니다."""
    if not records: return [], None, None
    
    day1 = records[0]['dt'].date()
    start_dt = datetime.combine(day1, s_time)
    end_dt = datetime.combine(day1 + timedelta(days=1), e_time)
    
    out = [r for r in records if r['flight'][:2] in ALLOWED_AIRLINES and 
           r['dest'] not in NZ_DOMESTIC_IATA and (start_dt <= r['dt'] < end_dt)]
    out.sort(key=lambda x: x['dt'])
    return out, start_dt, end_dt

# --- 3. DOCX 리스트 생성 (제브라 무늬 및 Footer) ---
def build_docx_stream(records, start_dt, end_dt):
    doc = Document()
    font_name = 'Arial'
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = Inches(0.3)
    section.left_margin = section.right_margin = Inches(0.5)

    # Footer 설정
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_f = footer_para.add_run("created by Simon Park'nRide's Flight List Factory 2026")
    run_f.font.size = Pt(10)
    run_f.font.color.rgb = RGBColor(128, 128, 128)

    # 타이틀
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_head = p.add_run(f"{start_dt.strftime('%d %b')} - {end_dt.strftime('%d %b')} FLIGHT LIST")
    run_head.bold = True
    run_head.font.size = Pt(16)

    # 테이블 생성
    table = doc.add_table(rows=0, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    for i, r in enumerate(records):
        row = table.add_row()
        vals = [r['flight'], r['time'], r['dest'], r['type'], r['reg']]
        for j, val in enumerate(vals):
            cell = row.cells[j]
            # 제브라 무늬 적용 (홀수 행 배경색)
            if i % 2 == 1:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd')
                shd.set(qn('w:val'), 'clear')
                shd.set(qn('w:fill'), 'D9D9D9')
                tcPr.append(shd)
            
            para = cell.paragraphs[0]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = para.add_run(str(val))
            run.font.size = Pt(13)
            
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

# --- 5. 사이드바 및 메인 화면 ---
with st.sidebar:
    st.header("⚙️ Settings")
    year = st.number_input("Year", value=2026) # NameError 해결
    s_time = st.time_input("Start Time (Day 1)", value=dtime(5, 0)) # UI 복구
    e_time = st.time_input("End Time (Day 2)", value=dtime(4, 55))
    label_start = st.number_input("Label Start Number", value=1, min_value=1)

st.markdown('<div class="main-title">Simon Park\'nRide\'s<br><span class="sub-title">Flight List Factory</span></div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Raw Text File (.txt)", type=['txt'])

if uploaded_file:
    # 파일 읽기 및 처리
    try:
        content = uploaded_file.read().decode("utf-8", errors="replace").splitlines()
        all_recs = parse_raw_lines(content, year)
        
        if all_recs:
            filtered, s_dt, e_dt = filter_records(all_recs, s_time, e_time)
            if filtered:
                st.success(f"✅ {len(filtered)}개의 항공편을 처리했습니다.")
                
                col1, col2 = st.columns(2)
                fn = f"List_{s_dt.strftime('%d-%m')}"
                
                col1.download_button("📥 Download DOCX List", build_docx_stream(filtered, s_dt, e_dt), f"{fn}.docx")
                col2.download_button("🏷️ Download PDF Labels", build_labels_stream(filtered, label_start), f"Labels_{fn}.pdf")
                
                # 결과 테이블 미리보기
                preview = []
                for i, r in enumerate(filtered):
                    preview.append({'No': label_start + i, 'Flight': r['flight'], 'Time': r['time'], 'Dest': r['dest'], 'Reg': r['reg']})
                st.table(preview)
            else:
                st.warning("⚠️ 필터 조건(시간, 항공사 등)에 맞는 데이터가 없습니다.")
        else:
            st.error("❌ 데이터를 읽지 못했습니다. 파일 내용이 쉼표(,) 구분 형식인지 확인해주세요.")
    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")
