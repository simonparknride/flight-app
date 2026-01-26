import streamlit as st
import re
import io
import pandas as pd
from datetime import datetime, timedelta
from typing import List, Dict
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# --- [이전 파싱 및 UI 설정 로직 동일 유지] ---
st.set_page_config(page_title="Flight List Factory", layout="centered")

# --- 3. DOCX 생성 (1-PAGE를 위한 2단 배열 로직 추가) ---
def build_docx(recs, is_1p=False):
    doc = Document()
    f_name = 'Air New Zealand Sans'
    sec = doc.sections[0]
    
    # 여백 극단적 최적화 (0.2인치)
    sec.top_margin = sec.bottom_margin = Inches(0.2)
    sec.left_margin = sec.right_margin = Inches(0.3)

    if is_1p:
        # --- 1-PAGE용 2단 배열 모드 ---
        # 데이터를 절반으로 나눔
        half = (len(recs) + 1) // 2
        left_side = recs[:half]
        right_side = recs[half:]
        
        # 큰 틀이 되는 테이블 생성 (좌/우 구분을 위해 2열)
        main_table = doc.add_table(rows=1, cols=2)
        main_table.width = Inches(7.5)
        
        for idx, side_data in enumerate([left_side, right_side]):
            cell = main_table.rows[0].cells[idx]
            # 셀 내부의 간격을 없애기 위해 새로 테이블 생성
            sub_table = cell.add_table(rows=0, cols=6)
            sub_table.width = Inches(3.6)
            
            last_date_str = ""
            for i, r in enumerate(side_data):
                row = sub_table.add_row()
                # 행 높이 아주 좁게 설정
                tr = row._tr
                trPr = tr.get_or_add_trPr()
                trHeight = OxmlElement('w:trHeight')
                trHeight.set(qn('w:val'), '180') # 9pt에 최적화된 높이
                trHeight.set(qn('w:hRule'), 'atLeast')
                trPr.append(trHeight)

                curr_date = r['dt'].strftime('%d %b')
                display_date = curr_date if curr_date != last_date_str else ""
                last_date_str = curr_date
                t_short = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
                
                vals = [display_date, r['flight'], t_short, r['dest'], r['type'], r['reg']]
                for j, v in enumerate(vals):
                    c = row.cells[j]
                    if i % 2 == 1: # 줄무늬 배경
                        shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'EEEEEE'); c._tc.get_or_add_tcPr().append(shd)
                    p = c.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(0)
                    run = p.add_run(str(v))
                    run.font.name = f_name
                    run.font.size = Pt(8.5) # 9pt 근처로 고정
                    if j == 0: run.bold = True
    else:
        # --- 일반 모드 (기존 14pt 단일 테이블) ---
        table = doc.add_table(rows=0, cols=6)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        last_date_str = ""
        for i, r in enumerate(recs):
            row = table.add_row()
            # ... (기존 14pt 행 설정 코드 동일) ...
            curr_date = r['dt'].strftime('%d %b')
            display_date = curr_date if curr_date != last_date_str else ""
            last_date_str = curr_date
            t_short = datetime.strptime(r['time'], '%I:%M %p').strftime('%H:%M')
            vals = [display_date, r['flight'], t_short, r['dest'], r['type'], r['reg']]
            for j, v in enumerate(vals):
                c = row.cells[j]
                if i % 2 == 1:
                    shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:fill'), 'D9D9D9'); c._tc.get_or_add_tcPr().append(shd)
                p = c.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(1)
                run = p.add_run(str(v))
                run.font.size = Pt(14.0)
                if j == 0: run.bold = True

    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf

# --- [나머지 build_labels, parse_raw_lines, UI 실행부 동일 유지] ---
