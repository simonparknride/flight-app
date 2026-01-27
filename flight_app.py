# =========================
# Flight List Factory 2026
# Final Stable Version
# =========================

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


# =========================
# 1. CONFIG
# =========================

YEAR = 2026
FONT_NAME = "Air New Zealand Sans"

TIME_LINE = re.compile(r"^(\d{1,2}:\d{2}\s[AP]M)\t([A-Z]{2}\d+[A-Z]?)$")
DATE_HEADER = re.compile(r"^[A-Za-z]+,\s+\w+\s+\d{1,2}$")
IATA_IN_PARENS = re.compile(r"\(([^)]+)\)")

ALLOWED_AIRLINES = {"NZ", "QF", "JQ", "CZ", "CA", "SQ", "LA", "IE", "FX"}
NZ_DOMESTIC_IATA = {
    "AKL","WLG","CHC","ZQN","TRG","NPE","PMR","NSN","NPL","DUD","IVC","TUO",
    "WRE","BHE","ROT","GIS","KKE","WHK","WAG","PPQ"
}

PLANE_TYPE_NORMALIZE = {
    "32q":"A320","320":"A320","32x":"A320",
    "789":"B789","772":"B772","77w":"B77W",
    "332":"A332","333":"A333","330":"A330",
    "359":"A359","388":"A388",
    "737":"B737","73h":"B737",
    "at7":"AT76"
}

PLANE_TYPE_PATTERN = re.compile(
    r"\b(" + "|".join(
        sorted(
            {*PLANE_TYPE_NORMALIZE.keys(), *PLANE_TYPE_NORMALIZE.values()},
            key=len, reverse=True
        )
    ) + r")\b",
    re.IGNORECASE
)


# =========================
# 2. PARSING UTILITIES
# =========================

def parse_date(line: str):
    if not DATE_HEADER.match(line):
        return None
    try:
        return datetime.strptime(f"{line} {YEAR}", "%A, %b %d %Y").date()
    except ValueError:
        return None


def normalize_plane_type(raw: str) -> str:
    return PLANE_TYPE_NORMALIZE.get(raw.lower(), raw.upper())


def extract_plane_and_reg(text: str):
    plane = ""
    reg = ""

    if m := PLANE_TYPE_PATTERN.search(text):
        plane = normalize_plane_type(m.group(1))

    for p in reversed(IATA_IN_PARENS.findall(text)):
        if "-" in p:
            reg = p.strip()
            break

    return plane, reg


def fmt_time(t: str) -> str:
    return datetime.strptime(t, "%I:%M %p").strftime("%H:%M")


def parse_raw_lines(lines: List[str]) -> List[Dict]:
    records = []
    current_date = None

    for i, line in enumerate(lines):
        line = line.strip()

        if d := parse_date(line):
            current_date = d
            continue

        if not current_date:
            continue

        if not (m := TIME_LINE.match(line)):
            continue

        try:
            time_str, flight = m.groups()
            dest = IATA_IN_PARENS.search(lines[i + 1]).group(1).upper()
            plane, reg = extract_plane_and_reg(lines[i + 2])

            dt = datetime.strptime(
                f"{current_date} {time_str}",
                "%Y-%m-%d %I:%M %p"
            )

            records.append({
                "dt": dt,
                "time": time_str,
                "flight": flight,
                "dest": dest,
                "type": plane,
                "reg": reg
            })

        except Exception:
            continue

    return records


# =========================
# 3. FILTERING
# =========================

def filter_records(records, start_hm, end_hm):
    dates = sorted({r["dt"].date() for r in records})
    if not dates:
        return [], None, None

    start_dt = datetime.combine(
        dates[0],
        datetime.strptime(start_hm, "%H:%M").time()
    )

    end_dt = datetime.combine(
        dates[1] if len(dates) > 1 else dates[0] + timedelta(days=1),
        datetime.strptime(end_hm, "%H:%M").time()
    )

    def valid(r):
        return (
            r["flight"][:2] in ALLOWED_AIRLINES and
            r["dest"] not in NZ_DOMESTIC_IATA and
            start_dt <= r["dt"] <= end_dt
        )

    return sorted(filter(valid, records), key=lambda r: r["dt"]), start_dt, end_dt


# =========================
# 4. DOCX EXPORT
# =========================

def build_docx_stream(records, start_dt, end_dt):
    doc = Document()
    section = doc.sections[0]

    section.top_margin = section.bottom_margin = Inches(0.3)
    section.left_margin = section.right_margin = Inches(0.5)

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    f_run = footer.add_run("created by Simon Park'nRide's Flight List Factory 2026")
    f_run.font.name = FONT_NAME
    f_run.font.size = Pt(10)
    f_run.font.color.rgb = RGBColor(128, 128, 128)

    head = doc.add_paragraph()
    head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h_run = head.add_run(f"{start_dt:%d}-{end_dt:%d} {start_dt:%b}")
    h_run.bold = True
    h_run.font.name = FONT_NAME
    h_run.font.size = Pt(16)

    table = doc.add_table(rows=0, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for r in records:
        cells = table.add_row().cells
        values = [
            r["flight"],
            fmt_time(r["time"]),
            r["dest"],
            r["type"],
            r["reg"]
        ]

        for cell, val in zip(cells, values):
            para = cell.paragraphs[0]
            para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            para.paragraph_format.space_before = para.paragraph_format.space_after = Pt(0)

            run = para.add_run(val)
            run.font.name = FONT_NAME
            run.font.size = Pt(14)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# =========================
# 5. PDF LABEL EXPORT
# =========================

def build_labels_stream(records, start_num):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4

    margin, gutter = 15 * mm, 6 * mm
    col_w = (w - 2 * margin - gutter) / 2
    row_h = (h - 2 * margin) / 5

    for i, r in enumerate(records):
        if i > 0 and i % 10 == 0:
            c.showPage()

        idx = i % 10
        x = margin + (idx % 2) * (col_w + gutter)
        y = h - margin - (idx // 2) * row_h

        c.rect(x, y - row_h + 2 * mm, col_w, row_h - 4 * mm)
        c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(x + 7 * mm, y - 10 * mm, str(start_num + i))
        c.setFont("Helvetica-Bold", 38)
        c.drawString(x + 15 * mm, y - 25 * mm, r["flight"])
        c.setFont("Helvetica-Bold", 29)
        c.drawString(x + 15 * mm, y - 45 * mm, fmt_time(r["time"]))

    c.save()
    buf.seek(0)
    return buf


# =========================
# 6. STREAMLIT UI
# =========================

st.set_page_config(page_title="Flight List Factory", layout="centered")

with st.sidebar:
    st.header("⚙️ Settings")

    s_time = st.text_input("Start Time (HH:MM)", value="05:00")
    e_time = st.text_input("End Time (HH:MM)", value="04:55")

    label_start = st.number_input(
        "Label Start Number",
        value=1,
        min_value=1,
        step=1
    )

def valid_hm(t: str) -> bool:
    try:
        datetime.strptime(t, "%H:%M")
        return True
    except ValueError:
        return False

if not valid_hm(s_time) or not valid_hm(e_time):
    st.error("⛔ Time must be in HH:MM format")
    st.stop()

st.title("✈️ Flight List Factory")

uploaded = st.file_uploader("Upload Raw Text File", type="txt")

if uploaded:
    lines = uploaded.read().decode("utf-8").splitlines()
    records = parse_raw_lines(lines)
    filtered, s_dt, e_dt = filter_records(records, s_time, e_time)

    if not filtered:
        st.warning("No matching flights found")
        st.stop()

    st.success(f"{len(filtered)} flights processed")

    c1, c2 = st.columns(2)
    fn = f"List_{s_dt:%d-%m}"

    c1.download_button(
        "📥 Download DOCX",
        build_docx_stream(filtered, s_dt, e_dt),
        file_name=f"{fn}.docx"
    )

    c2.download_button(
        "📥 Download PDF Labels",
        build_labels_stream(filtered, label_start),
        file_name=f"Labels_{fn}.pdf"
    )

    st.table([
        {
            "No": label_start + i,
            "Flight": r["flight"],
            "Time": r["time"],
            "Dest": r["dest"],
            "Type": r["type"]
        }
        for i, r in enumerate(filtered)
    ])
