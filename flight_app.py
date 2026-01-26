import streamlit as st
import re
import io
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement, qn

# 1. UI 및 스타일 설정
st.set_page_config(page_title="Flight List Factory", layout="centered")
st.markdown("""
    <style>
    .stApp { background-color: #000000; }
    [data-testid="stSidebar"] { background-color: #111111 !important; }
    .stMarkdown, p, h1, h2, h3, label { color: #ffffff !important; }
    div.stDownloadButton > button {
        background-color: #ffffff !important; color: #000000 !important;
        font-weight: 800 !important; width: 100% !important; height: 3.5rem !important;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("Simon Park'nRide's Factory")

# 2. 버튼 공간 미리 확보 (사라짐 방지)
btn_cols = st.columns(4)

# 3. 사이드바 (End Time 필드 강제 고정)
with st.sidebar:
    st.header("⚙️ Settings")
    s_time = st.text_input("Start Time (HH:MM)", "04:55")
    e_time = st.text_input("End Time (HH:MM)", "04:50") 
    label_start = st.number_input("Label Start No", value=1)

# 4. 데이터 파싱 및 필터링 로직
uploaded = st.file_uploader("Upload Raw Text File", type=['txt'])
filtered_recs = [] # NameError 방지를 위한 초기화

if uploaded:
    content = uploaded.read().decode("utf-8")
    lines = content.splitlines()
    current_date = "26 Jan" 
    parsed = []
    
    for line in lines:
        if not line.strip(): continue
        # 날짜 헤더 인식
        dt_match = re.search(r"(\d{1,2}\s+[A-Za-z]{3})", line)
        if dt_match and ":" not in line:
            current_date = dt_match.group(1)
            continue
        
        # 쉼표 구분 데이터 처리 [cite: 2, 4, 6]
        parts = line.split(',')
        if len(parts) >= 5:
            try:
                row_date = parts[0].strip() if parts[0].strip() and parts[0].strip()[0].isdigit() else current_date
                time_str = parts[2].strip() if ":" in parts[2] else parts[1].strip()
                dt_obj = datetime.strptime(f"{row_date} 2026 {time_str}", "%d %b %Y %H:%M")
                
                parsed.append({
                    'dt': dt_obj,
                    'flight': parts[1].strip() if ":" in parts[2] else parts[0].strip(),
                    'dest': parts[3].strip(),
                    'type': parts[4].strip(),
                    'reg': parts[5].strip() if len(parts) > 5 else ""
                })
            except: continue

    if parsed:
        try:
            start_dt = datetime.combine(parsed[0]['dt'].date(), datetime.strptime(s_time, "%H:%M").time())
            end_t = datetime.strptime(e_time, "%H:%M").time()
            end_dt = datetime.combine(parsed[0]['dt'].date(), end_t)
            if end_dt <= start_dt: end_dt += timedelta(days=1)
            
            filtered_recs = [r for r in parsed if start_dt <= r['dt'] < end_dt]
        except: st.error("시간 형식을 확인하세요 (HH:MM)")

# 5. 결과 버튼 활성화
if filtered_recs:
    st.success(f"준비 완료: {len(filtered_recs)}건")
    # 여기에 build_docx 함수를 연결한 버튼 생성 (생략)
    btn_cols[0].download_button("📥 DOCX", b"file", "List.docx")
    btn_cols[1].download_button("📄 1-PAGE", b"file", "List_1p.docx")
    btn_cols[2].download_button("🏷️ LABELS", b"PDF", "Labels.pdf")
    btn_cols[3].download_button("📊 EXCL", b"CSV", "Excl.csv")
elif uploaded:
    st.warning("데이터가 없거나 필터링에 실패했습니다.")
