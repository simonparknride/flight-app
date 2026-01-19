import streamlit as st
import re
import io
from datetime import datetime, timedelta
from typing import List, Dict

# --- 1. UI 및 사이드바 설정 ---
st.set_page_config(
    page_title="Flight List Factory", 
    layout="centered",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    /* 전체 배경 및 기본 텍스트 */
    .stApp { background-color: #000000; }
    [data-testid="stSidebar"] { background-color: #111111 !important; }
    .stMarkdown, p, h1, h2, h3, label { color: #ffffff !important; }
    
    /* 다운로드 버튼 스타일 수정 (글자색 검정으로 강제 지정) */
    div.stDownloadButton > button {
        background-color: #ffffff !important; /* 버튼 배경을 흰색으로 */
        color: #000000 !important;           /* 글자색을 검정색으로 */
        border: 1px solid #ffffff;
        border-radius: 5px;
        padding: 0.5rem 1rem;
        font-weight: bold;
        width: 100%;
    }
    
    /* 버튼에 마우스를 올렸을 때(Hover) 스타일 */
    div.stDownloadButton > button:hover {
        background-color: #60a5fa !important; /* 파란색 계열로 변경 */
        color: #ffffff !important;           /* 글자색 흰색으로 변경 */
        border: 1px solid #60a5fa;
    }

    /* 상단 링크 및 타이틀 디자인 (기존과 동일) */
    .top-left-container { text-align: left; padding-top: 10px; margin-bottom: 20px; }
    .top-left-container a { 
        font-size: 1.1rem; color: #ffffff !important; 
        text-decoration: underline; font-weight: 300; 
        display: block; margin-bottom: 5px; 
    }
    .main-title { font-size: 3rem; font-weight: 800; color: #ffffff; line-height: 1.1; margin-bottom: 0.5rem; }
    .sub-title { font-size: 2.5rem; font-weight: 400; color: #60a5fa; }
    .stTable { background-color: rgba(255, 255, 255, 0.05); border-radius: 10px; }
    </style>
    """, unsafe_allow_html=True)

# --- (이하 나머지 로직은 이전과 동일하게 유지) ---
