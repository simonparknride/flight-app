import streamlit as st
import re
import io
import pandas as pd
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement, qn

# --- 1. UI 설정 및 스타일 ---
st.set_page_config(page_title="Flight List Factory", layout="centered")
st.markdown("""
    <style>
    .stApp { background-color: #000000; }
    [data-testid="stSidebar"] { background-color: #111111 !important; }
    .stMarkdown, p, h1, h2, h3, label { color: #ffffff !important; }
    div.stDownloadButton > button {
        background-color: #ffffff !important; color: #000000 !important;
        font-weight: 800 !important; width: 100% !important; height: 3.5rem !important;
        border-radius: 8px !important;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("Simon Park'nRide's Factory")

# --- 2. 사이드바 (End Time 복구) ---
with st.sidebar:
    st.header("⚙️ Settings")
    s_time_input = st.text_input("Start Time (HH:MM)", "04:55")
    e_time_input = st.text_input("End Time (HH:MM)", "04:50") # 사라졌던 End Time 복구
    label_start = st.number_input("Label Start No", value=1)

# --- 3. 제브라 무늬 및 문서 생성 ---
def set_zebra(cell):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), 'D9D9D9') # 가독성 좋은 회색
    cell._tc.get_or_add_tcPr().append(shd)

def build_docx(recs, is_1p=False):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = Inches(0.2)
    sec.left_margin = sec.right_margin = Inches(0.35)

    if is_1p: # 1-PAGE (2단 배열, 8.5pt로 한 장 고정)
        main_table = doc.add_table(rows=1, cols=2)
        half = (len(recs) + 1) // 2
        for idx, side_data in enumerate([recs[:half], recs[half:]]):
            cell = main_table.rows[0].cells[idx]
            sub_t = cell.add_table(rows=0, cols=6)
            last_d = ""
            for i, r in enumerate(side_data):
                row = sub_t.add_row()
                d_str = r['dt'].strftime('%d %b')
                vals = [d_str if d_str != last_d else "", r['flight'], r['dt'].strftime('%H:%M'), r['dest'], r['type'], r['reg']]
                last_d = d_str
                for j, v in enumerate(vals):
                    c = row.cells[j]
                    if i % 2 == 1: set_zebra(c)
                    p = c.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run(str(v))
                    run.font.size = Pt(8.5)
    else: # 일반 DOCX (14pt, 2페이지 이내)
        t = doc.add_table(rows=0, cols=6)
        last_d = ""
        for i, r in enumerate(recs):
            row = t.add_row()
            d_str = r['dt'].strftime('%d %b')
            vals = [d_str if d_str != last_d else "", r['flight'], r['dt'].strftime('%H:%M'), r['dest'], r['type'], r['reg']]
            last_d = d_str
            for j, v in enumerate(vals):
                c = row.cells[j]
                if i % 2 == 1: set_zebra(c)
                p = c.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(str(v))
                run.font.size = Pt(14)
                if j == 0: run.bold = True
    
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf

# --- 4. 메인 로직 ---
uploaded = st.file_uploader("Upload Raw Text File", type=['txt'])
btn_cols = st.columns(4)

if uploaded:
    raw_text = uploaded.read().decode("utf-8")
    lines = raw_text.splitlines()
    all_recs = []
    current_date = "26 Jan" # 기본값
    
    for line in lines:
        if not line.strip(): continue
        # 날짜 헤더 감지 (예: 26 Jan 2026)
        date_m = re.search(r"(\d{1,2}\s+[A-Za-z]{3})", line)
        if date_m and ":" not in line:
            current_date = date_m.group(1)
            continue
        
        # 쉼표 구분 데이터 파싱
        parts = line.split(',')
        if len(parts) >= 5:
            try:
                # 첫 칸이 날짜가 아니면 저장된 current_date 사용
                row_date = parts[0].strip() if parts[0].strip() and parts[0].strip()[0].isdigit() else current_date
                time_str = parts[2].strip() if ":" in parts[2] else parts[1].strip()
                dt_obj = datetime.strptime(f"{row_date} 2026 {time_str}", "%d %b %Y %H:%M")
                
                all_recs.append({
                    'dt': dt_obj,
                    'flight': parts[1].strip() if ":" in parts[2] else parts[0].strip(),
                    'dest': parts[3].strip(),
                    'type': parts[4].strip(),
                    'reg': parts[5].strip() if len(parts) > 5 else ""
                })
            except: continue

    if all_recs:
        try:
            # 사용자 입력 시간으로 필터링 범위 설정
            s_dt = datetime.combine(all_recs[0]['dt'].date(), datetime.strptime(s_time_input, "%H:%M").time())
            # End Time이 Start Time보다 빠르면 다음 날로 처리
            e_time = datetime.strptime(e_time_input, "%H:%M").time()
            e_dt = datetime.combine(all_recs[0]['dt'].date(), e_time)
            if e_dt <= s_dt: e_dt += timedelta(days=1)
            
            filtered = [r for r in all_recs if s_dt <= r['dt'] < e_dt]
            
            if filtered:
                st.success(f"{len(filtered)} Flights Processed")
                btn_cols[0].download_button("📥 DOCX", build_docx(filtered), "List.docx")
                btn_cols[1].download_button("📄 1-PAGE", build_docx(filtered, True), "List_1p.docx")
                btn_cols[2].download_button("🏷️ LABELS", b"PDF", "Labels.pdf")
                btn_cols[3].download_button("📊 EXCL", b"CSV", "Excl.csv")
            else:
                st.warning("일치하는 항공편이 없습니다. 시간을 확인하세요.")
        except Exception as e:
            st.error(f"시간 설정 오류: {e}")
