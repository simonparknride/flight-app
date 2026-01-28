import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime, timedelta, time as dtime
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement, qn
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# =======================
# CONFIG
# =======================
ALLOWED_AIRLINES = {"NZ", "QF", "JQ", "CZ", "CA", "SQ", "LA", "IE", "FX"}
NZ_DOMESTIC_IATA = {
    "AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO",
    "WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"
}

DATE_RE = re.compile(r"^\d{1,2}\s+[A-Za-z]{3}$")

# =======================
# DATA NORMALIZATION
# =======================
def normalize_records(rows, start_year):
    records = []
    current_date = None
    year = start_year

    for r in rows:
        date_raw = str(r.get("Date", "")).strip()

        # 날짜 헤더 or 날짜 컬럼
        if DATE_RE.match(date_raw):
            d = datetime.strptime(f"{date_raw} {year}", "%d %b %Y")
            if current_date and d < current_date:
                year += 1
                d = datetime.strptime(f"{date_raw} {year}", "%d %b %Y")
            current_date = d
            continue

        if not current_date:
            continue

        try:
            dt = datetime.strptime(
                f"{current_date.strftime('%d %b %Y')} {r['Time']}",
                "%d %b %Y %H:%M"
            )

            records.append({
                "dt": dt,
                "time": r["Time"],
                "flight": r["Flight"].upper(),
                "dest": r["Dest"].upper(),
                "type": r.get("Type", ""),
                "reg": r.get("Reg", "")
            })

        except Exception as e:
            st.warning(f"파싱 실패: {r}")

    return records

# =======================
# LOADERS
# =======================
def load_txt(file):
    rows = []
    for line in file.read().decode("utf-8", errors="replace").splitlines():
        parts = [p.strip() for p in line.split(",")]
        if len(parts) >= 6:
            rows.append({
                "Date": parts[0],
                "Flight": parts[1],
                "Time": parts[2],
                "Dest": parts[3],
                "Type": parts[4],
                "Reg": parts[5],
            })
        elif DATE_RE.match(parts[0]):
            rows.append({"Date": parts[0]})
    return rows

def load_csv(file):
    return pd.read_csv(file).to_dict("records")

def load_excel(file):
    return pd.read_excel(file).to_dict("records")

# =======================
# FILTER
# =======================
def filter_records(records, s_time, e_time):
    if not records:
        return []

    base_date = records[0]["dt"].date()
    start_dt = datetime.combine(base_date, s_time)
    end_dt = datetime.combine(base_date + timedelta(days=1), e_time)

    out = [
        r for r in records
        if r["flight"][:2] in ALLOWED_AIRLINES
        and r["dest"] not in NZ_DOMESTIC_IATA
        and start_dt <= r["dt"] < end_dt
    ]
    return sorted(out, key=lambda x: x["dt"])

# =======================
# DOCX
# =======================
def build_docx(records):
    doc = Document()
    table = doc.add_table(rows=0, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for i, r in enumerate(records):
        row = table.add_row()
        for j, v in enumerate([r["flight"], r["time"], r["dest"], r["type"], r["reg"]]):
            cell = row.cells[j]
            if i % 2:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement("w:shd")
                shd.set(qn("w:fill"), "D9D9D9")
                tcPr.append(shd)
            cell.paragraphs[0].add_run(str(v)).font.size = Pt(12)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# =======================
# PDF LABELS
# =======================
def build_labels(records, start_no):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4

    for i, r in enumerate(records):
        if i and i % 10 == 0:
            c.showPage()
        c.drawString(50, h - 50 - (i % 10) * 50, f"{start_no+i} {r['flight']} {r['dest']} {r['time']}")

    c.save()
    buf.seek(0)
    return buf

# =======================
# UI
# =======================
st.set_page_config("Flight List Factory", layout="centered")
st.title("✈️ Flight List Factory")

year = st.number_input("Start Year", value=2026)
s_time = st.time_input("Start Time", dtime(5, 0))
e_time = st.time_input("End Time", dtime(4, 55))
label_start = st.number_input("Label Start", 1)

file = st.file_uploader("Upload TXT / CSV / Excel", ["txt", "csv", "xlsx"])

if file:
    if file.name.endswith(".txt"):
        raw = load_txt(file)
    elif file.name.endswith(".csv"):
        raw = load_csv(file)
    else:
        raw = load_excel(file)

    records = normalize_records(raw, year)
    filtered = filter_records(records, s_time, e_time)

    if not filtered:
        st.error("처리된 항공편이 없습니다.")
    else:
        st.success(f"{len(filtered)} flights processed")

        st.download_button("DOCX", build_docx(filtered), "flight_list.docx")
        st.download_button("PDF Labels", build_labels(filtered, label_start), "labels.pdf")

        st.dataframe(filtered)
