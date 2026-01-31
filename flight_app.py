# Flight List Factory - Streamlit app
# FULL VERSION with Year Rollover Fix Applied
# Existing layout / DOCX / PDF logic fully preserved

import streamlit as st
import re
import io
from datetime import datetime, timedelta, time as dtime
from typing import List, Dict, Optional
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# --- Page Config & Styling ---
st.set_page_config(page_title="Flight List Factory", layout="centered", initial_sidebar_state="expanded")

# --- Parsing patterns ---
# ★ 변경: 항공편 번호 공백 허용
TIME_LINE = re.compile(
    r"^(\d{1,2}:\d{2}\s?[AP]M)\s+([A-Z]{2}\s?\d{1,4}[A-Z]?)\s*$",
    re.IGNORECASE
)

DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}\s*$")
IATA_IN_PARENS = re.compile(r"\(([^)]+)\)")

PLANE_TYPES = [
    'A21N','A20N','A320','32Q','320','73H','737','74Y','77W','B77W',
    '789','B789','359','A359','332','A332','AT76','DH8C','DH3','AT7',
    '388','333','A333','330','76V','77L','B38M','A388','772','B772','32X','77X'
]
PLANE_TYPE_PATTERN = re.compile(
    r"\b(" + "|".join(sorted(set(PLANE_TYPES), key=len, reverse=True)) + r")\b",
    re.IGNORECASE
)

NORMALIZE_MAP = {
    '32q':'A320','320':'A320','32x':'A320',
    '789':'B789','772':'B772','77w':'B77W',
    '332':'A332','333':'A333','330':'A330',
    '359':'A359','388':'A388','737':'B737','73h':'B737','at7':'AT76'
}

ALLOWED_AIRLINES = {"NZ","QF","JQ","CZ","CA","SQ","LA","IE","FX"}
NZ_DOMESTIC_IATA = {
    "AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC",
    "TUO","WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"
}

REGO_LIKE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9\-–—]*$")

def normalize_type(t: Optional[str]) -> str:
    if not t:
        return ""
    return NORMALIZE_MAP.get(t.lower(), t.upper())

def try_parse_date_header(line: str, year: int):
    for fmt in ("%A, %b %d %Y", "%A, %B %d %Y", "%a, %b %d %Y", "%a, %B %d %Y"):
        try:
            return datetime.strptime(f"{line} {year}", fmt).date()
        except Exception:
            pass
    return None

def parse_raw_lines(lines: List[str], year: int) -> List[Dict]:
    records = []
    current_date = None
    current_year = year
    last_date = None
    i = 0
    L = len(lines)

    while i < L:
        line = lines[i].strip()

        # ★ 변경: 연말 → 연초 자동 연도 보정
        if DATE_HEADER.match(line):
            parsed = try_parse_date_header(line, current_year)
            if parsed:
                if last_date and parsed < last_date:
                    current_year += 1
                    parsed = try_parse_date_header(line, current_year)
                current_date = parsed
                last_date = parsed
            else:
                current_date = None
            i += 1
            continue

        m = TIME_LINE.match(line)
        if m and current_date:
            time_raw, flight_raw = m.groups()

            # ★ 변경: 항공편 번호 공백 제거
            flight = flight_raw.replace(" ", "").upper()

            dest_line = lines[i+1].strip() if i+1 < L else ''
            carrier_line = lines[i+2].strip() if i+2 < L else ''

            dest = ''
            m_dest = IATA_IN_PARENS.search(dest_line)
            if m_dest:
                dest = m_dest.group(1).upper()

            mtype = PLANE_TYPE_PATTERN.search(carrier_line)
            plane_type = normalize_type(mtype.group(1)) if mtype else ''

            reg = ''
            for p in reversed(IATA_IN_PARENS.findall(carrier_line)):
                if REGO_LIKE.match(p) and '-' in p:
                    reg = p
                    break

            try:
                tnorm = time_raw.upper().replace(" ", "")
                dep_dt = datetime.strptime(
                    f"{current_date} {tnorm}", "%Y-%m-%d %I:%M%p"
                )
            except Exception:
                dep_dt = None

            records.append({
                "dt": dep_dt,
                "time": time_raw,
                "flight": flight,
                "dest": dest,
                "type": plane_type,
                "reg": reg
            })

            i += 3
            continue

        i += 1

    return records
