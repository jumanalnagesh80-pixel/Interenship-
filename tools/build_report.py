#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generator for the HAL Industrial Internship Report (.docx)
----------------------------------------------------------
Builds a professional, university-standard ~50-page engineering internship
report as a native Office Open XML (.docx) document using ONLY the Python
standard library (no python-docx / no external packages required).

Formatting implemented:
  Paper        : A4 (210 x 297 mm)
  Margins      : Top 1", Bottom 1", Left 1.5", Right 1"
  Font         : Times New Roman
  Body         : 12 pt, 1.5 line spacing, justified
  Heading      : 16 pt bold ; Sub-heading 14 pt bold ; Sub-sub 12 pt bold
  Header       : "Industrial Internship Report"
  Footer       : "Student Name | BCA Final Year | HAL Internship" + page number (bottom centre)
  Page numbers : bottom centre
Output: report/HAL_Internship_Report.docx
"""
import os, zipfile, datetime

# ----------------------------------------------------------------------
# Twips helpers (1 inch = 1440 twips, 1 pt = 20 twips)
# ----------------------------------------------------------------------
def IN(v): return str(int(round(v * 1440)))
def PT(v): return str(int(round(v * 20)))

BODY = []          # list of body-level XML strings
_fig = [0]         # figure counter
_tab = [0]         # table counter

# ----------------------------------------------------------------------
# XML escaping
# ----------------------------------------------------------------------
def esc(t):
    return (t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
             .replace('"', "&quot;"))

# ----------------------------------------------------------------------
# Run + paragraph builders
# ----------------------------------------------------------------------
def run(text, bold=False, italic=False, size=None, color=None, font=None):
    # rPr child order per CT_RPr: rFonts, b, i, color, sz, szCs
    rpr = "<w:rPr>"
    if font: rpr += f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
    if bold: rpr += "<w:b/>"
    if italic: rpr += "<w:i/>"
    if color: rpr += f'<w:color w:val="{color}"/>'
    if size: rpr += f'<w:sz w:val="{int(size*2)}"/><w:szCs w:val="{int(size*2)}"/>'
    rpr += "</w:rPr>"
    return f'<w:r>{rpr}<w:t xml:space="preserve">{esc(text)}</w:t></w:r>'

def para(runs_xml, align="both", style=None, spacing_before=0, spacing_after=120,
         line=360, indent=0, keep=False, left=0, hanging=0, border=None):
    # NOTE: element order must follow the CT_PPr schema sequence:
    #   pStyle -> keepNext -> pBdr -> spacing -> ind -> jc
    ppr = "<w:pPr>"
    if style: ppr += f'<w:pStyle w:val="{style}"/>'
    if keep: ppr += "<w:keepNext/>"
    if border:  # 'top' or 'bottom'
        ppr += (f'<w:pBdr><w:{border} w:val="single" w:sz="4" w:space="1" '
                f'w:color="999999"/></w:pBdr>')
    ppr += (f'<w:spacing w:before="{spacing_before}" w:after="{spacing_after}" '
            f'w:line="{line}" w:lineRule="auto"/>')
    ind_attrs = ""
    if left: ind_attrs += f' w:left="{left}"'
    if hanging: ind_attrs += f' w:hanging="{hanging}"'
    if indent: ind_attrs += f' w:firstLine="{indent}"'
    if ind_attrs: ppr += f'<w:ind{ind_attrs}/>'
    ppr += f'<w:jc w:val="{align}"/></w:pPr>'
    return f"<w:p>{ppr}{runs_xml}</w:p>"

# ----------------------------------------------------------------------
# High level content helpers (append to BODY)
# ----------------------------------------------------------------------
def add(xml): BODY.append(xml)

def page_break():
    add('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')

def h1(text):
    page_break()
    add(para(run(text, bold=True, size=16, font="Times New Roman"),
             align="left", spacing_before=120, spacing_after=180, line=360, keep=True))

def h1_nopb(text):
    add(para(run(text, bold=True, size=16, font="Times New Roman"),
             align="left", spacing_before=120, spacing_after=180, line=360, keep=True))

def h2(text):
    add(para(run(text, bold=True, size=14, font="Times New Roman"),
             align="left", spacing_before=160, spacing_after=120, line=360, keep=True))

def h3(text):
    add(para(run(text, bold=True, size=12, font="Times New Roman"),
             align="left", spacing_before=120, spacing_after=80, line=360, keep=True))

def p(text, indent=True):
    add(para(run(text, size=12, font="Times New Roman"),
             align="both", indent=360 if indent else 0, spacing_after=120, line=360))

def bullet(text):
    add(para(run("\u2022  " + text, size=12, font="Times New Roman"),
             align="both", spacing_after=60, line=360, left=540, hanging=270))

def figcap(text):
    _fig[0] += 1
    add(para(run(f"Fig. {_fig[0]}  {text}", italic=True, size=11, font="Times New Roman"),
             align="center", spacing_before=60, spacing_after=200, line=240))

def tabcap(text):
    _tab[0] += 1
    add(para(run(f"Table {_tab[0]}  {text}", italic=True, size=11, font="Times New Roman"),
             align="center", spacing_before=120, spacing_after=60, line=240))

# ----------------------------------------------------------------------
# Tables
# ----------------------------------------------------------------------
def table(rows, widths=None, header=True, font_size=11):
    """rows: list of list of str. First row treated as header if header=True."""
    ncol = max(len(r) for r in rows)
    if widths is None:
        total = 9360  # usable width within 1.5"+1" margins on A4 (~6.5")
        widths = [total // ncol] * ncol
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    tbl = ('<w:tbl><w:tblPr>'
           '<w:tblStyle w:val="TableGrid"/>'
           '<w:tblW w:w="0" w:type="auto"/>'
           '<w:tblBorders>'
           '<w:top w:val="single" w:sz="4" w:color="555555"/>'
           '<w:left w:val="single" w:sz="4" w:color="555555"/>'
           '<w:bottom w:val="single" w:sz="4" w:color="555555"/>'
           '<w:right w:val="single" w:sz="4" w:color="555555"/>'
           '<w:insideH w:val="single" w:sz="4" w:color="999999"/>'
           '<w:insideV w:val="single" w:sz="4" w:color="999999"/>'
           '</w:tblBorders>'
           '<w:tblLook w:val="04A0"/>'
           f'</w:tblPr><w:tblGrid>{grid}</w:tblGrid>')
    for ri, r in enumerate(rows):
        is_head = header and ri == 0
        tr = "<w:tr>"
        if is_head:
            tr = '<w:tr><w:trPr><w:tblHeader/></w:trPr>'
        for ci in range(ncol):
            cell = r[ci] if ci < len(r) else ""
            shade = '<w:shd w:val="clear" w:color="auto" w:fill="1F3864"/>' if is_head else ""
            tcpr = (f'<w:tcPr><w:tcW w:w="{widths[ci]}" w:type="dxa"/>{shade}'
                    '<w:vAlign w:val="center"/></w:tcPr>')
            r_xml = run(cell, bold=is_head, size=font_size, font="Times New Roman",
                        color="FFFFFF" if is_head else None)
            pcell = para(r_xml, align="left", spacing_before=20, spacing_after=20, line=240)
            tr += f"<w:tc>{tcpr}{pcell}</w:tc>"
        tr += "</w:tr>"
        tbl += tr
    tbl += "</w:tbl>"
    add(tbl)
    # small spacer paragraph after table
    add(para(run("", size=6, font="Times New Roman"), spacing_after=60, line=120))

# ----------------------------------------------------------------------
# Diagram helpers (rendered as bordered boxes / grids, no image libs needed)
# ----------------------------------------------------------------------
def _box_cell(text, fill="DCE6F1", w=9360, bold=True):
    tcpr = (f'<w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>'
            '<w:tcBorders>'
            '<w:top w:val="single" w:sz="8" w:color="1F3864"/>'
            '<w:left w:val="single" w:sz="8" w:color="1F3864"/>'
            '<w:bottom w:val="single" w:sz="8" w:color="1F3864"/>'
            '<w:right w:val="single" w:sz="8" w:color="1F3864"/>'
            '</w:tcBorders>'
            f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'
            '<w:vAlign w:val="center"/></w:tcPr>')
    pc = para(run(text, bold=bold, size=11, font="Times New Roman", color="1F3864"),
              align="center", spacing_before=40, spacing_after=40, line=240)
    return f"<w:tc>{tcpr}{pc}</w:tc>"

def _arrow_down():
    add(para(run("\u25BC", size=12, font="Times New Roman", color="1F3864"),
             align="center", spacing_before=0, spacing_after=0, line=200))

def diagram_flow(steps, fill="DCE6F1"):
    """Vertical flowchart: list of step labels connected by down-arrows."""
    for i, s in enumerate(steps):
        tbl = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
               '<w:jc w:val="center"/><w:tblLook w:val="04A0"/></w:tblPr>'
               '<w:tblGrid><w:gridCol w:w="6600"/></w:tblGrid>'
               f'<w:tr>{_box_cell(s, fill=fill, w=6600)}</w:tr></w:tbl>')
        add(tbl)
        if i < len(steps) - 1:
            _arrow_down()
    add(para(run("", size=6, font="Times New Roman"), spacing_after=60, line=120))

def diagram_layers(layers):
    """Layered/architecture diagram: list of (layer_title, [box labels]) stacked
    vertically, each layer a row of boxes."""
    for li, (title, boxes) in enumerate(layers):
        n = len(boxes)
        colw = 9360 // n
        grid = "".join(f'<w:gridCol w:w="{colw}"/>' for _ in range(n))
        cells = "".join(_box_cell(b, fill="DCE6F1", w=colw) for b in boxes)
        tbl = ('<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/>'
               '<w:jc w:val="center"/><w:tblLook w:val="04A0"/></w:tblPr>'
               f'<w:tblGrid>{grid}</w:tblGrid><w:tr>{cells}</w:tr></w:tbl>')
        # layer caption
        add(para(run(title, bold=True, size=10, font="Times New Roman", color="555555"),
                 align="center", spacing_before=40, spacing_after=20, line=200))
        add(tbl)
        if li < len(layers) - 1:
            _arrow_down()
    add(para(run("", size=6, font="Times New Roman"), spacing_after=60, line=120))

def screenshot_placeholder(label):
    """A bordered empty box acting as a screenshot placeholder."""
    tcpr = ('<w:tcPr><w:tcW w:w="9360" w:type="dxa"/>'
            '<w:tcBorders>'
            '<w:top w:val="dashed" w:sz="8" w:color="888888"/>'
            '<w:left w:val="dashed" w:sz="8" w:color="888888"/>'
            '<w:bottom w:val="dashed" w:sz="8" w:color="888888"/>'
            '<w:right w:val="dashed" w:sz="8" w:color="888888"/>'
            '</w:tcBorders>'
            '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>'
            '<w:vAlign w:val="center"/></w:tcPr>')
    inner = (para(run("[ SCREENSHOT PLACEHOLDER ]", bold=True, size=12,
                      font="Times New Roman", color="888888"),
                  align="center", spacing_before=600, spacing_after=40, line=240)
             + para(run(label, italic=True, size=11, font="Times New Roman", color="888888"),
                    align="center", spacing_before=0, spacing_after=600, line=240))
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/>'
           '<w:jc w:val="center"/><w:tblLook w:val="04A0"/></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>'
           f'<w:tr><w:tc>{tcpr}{inner}</w:tc></w:tr></w:tbl>')
    add(tbl)

def code_block(lines):
    """Monospace-style code listing in a shaded box."""
    inner = ""
    for ln in lines:
        inner += para(run(ln if ln else " ", size=9, font="Consolas"),
                      align="left", spacing_before=0, spacing_after=0, line=240)
    tcpr = ('<w:tcPr><w:tcW w:w="9360" w:type="dxa"/>'
            '<w:tcBorders>'
            '<w:top w:val="single" w:sz="4" w:color="CCCCCC"/>'
            '<w:left w:val="single" w:sz="4" w:color="CCCCCC"/>'
            '<w:bottom w:val="single" w:sz="4" w:color="CCCCCC"/>'
            '<w:right w:val="single" w:sz="4" w:color="CCCCCC"/>'
            '</w:tcBorders>'
            '<w:shd w:val="clear" w:color="auto" w:fill="F5F5F5"/>'
            '</w:tcPr>')
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/>'
           '<w:jc w:val="center"/><w:tblLook w:val="04A0"/></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>'
           f'<w:tr><w:tc>{tcpr}{inner}</w:tc></w:tr></w:tbl>')
    add(tbl)
    add(para(run("", size=6, font="Times New Roman"), spacing_after=60, line=120))

print("Framework loaded.")



# ======================================================================
# EDITABLE STUDENT / COLLEGE DETAILS  (change these before submission)
# ======================================================================
COLLEGE   = "[ COLLEGE NAME ]"
DEPT      = "Department of Computer Applications"
STUDENT   = "[ STUDENT NAME ]"
USN       = "[ USN / REGISTER NUMBER ]"
GUIDE     = "[ INTERNAL GUIDE NAME ]"
MENTOR    = "[ HAL PROJECT MENTOR ]"
PRINCIPAL = "[ PRINCIPAL NAME ]"
YEAR      = "2025 - 2026"
PROJECT   = "Advanced Live Aircraft Tracking and Flight Data Processing System"

# ----------------------------------------------------------------------
# Centered title-page line helper
# ----------------------------------------------------------------------
def cline(text, size=12, bold=False, after=120, before=0, color=None, italic=False):
    add(para(run(text, bold=bold, size=size, font="Times New Roman", color=color, italic=italic),
             align="center", spacing_before=before, spacing_after=after, line=240))

def blank(n=1, size=12):
    for _ in range(n):
        add(para(run("", size=size, font="Times New Roman"), align="center",
                 spacing_after=0, line=240))

def logo_box(label):
    tcpr = ('<w:tcPr><w:tcW w:w="2600" w:type="dxa"/>'
            '<w:tcBorders>'
            '<w:top w:val="single" w:sz="6" w:color="888888"/>'
            '<w:left w:val="single" w:sz="6" w:color="888888"/>'
            '<w:bottom w:val="single" w:sz="6" w:color="888888"/>'
            '<w:right w:val="single" w:sz="6" w:color="888888"/>'
            '</w:tcBorders><w:vAlign w:val="center"/></w:tcPr>')
    inner = (para(run("", size=10, font="Times New Roman"), align="center", spacing_after=0, line=240)
             + para(run(label, italic=True, size=10, font="Times New Roman", color="888888"),
                    align="center", spacing_before=300, spacing_after=300, line=240)
             + para(run("", size=10, font="Times New Roman"), align="center", spacing_after=0, line=240))
    return f"<w:tc>{tcpr}{inner}</w:tc>"

# ======================================================================
# COVER PAGE
# ======================================================================
def cover_page():
    cline("INDUSTRIAL INTERNSHIP REPORT", size=18, bold=True, before=120, after=80)
    cline("on", size=12, after=80, italic=True)
    cline(PROJECT.upper(), size=15, bold=True, after=160, color="1F3864")
    cline("Submitted in partial fulfilment of the requirements for the award of the degree of",
          size=11, after=40, italic=True)
    cline("BACHELOR OF COMPUTER APPLICATIONS (BCA)", size=13, bold=True, after=160)
    cline("Internship carried out at", size=11, after=40, italic=True)
    cline("HINDUSTAN AERONAUTICS LIMITED (HAL)", size=14, bold=True, after=40, color="1F3864")
    cline("Rotary Wing Research & Design Centre (RWR&DC)", size=12, bold=True, after=20)
    cline("Flight Testing Centre, Bengaluru", size=12, after=160)
    # logo row (two bordered placeholders side by side)
    tbl = ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
           '<w:jc w:val="center"/>'
           '<w:tblCellSpacing w:w="200" w:type="dxa"/>'
           '<w:tblLook w:val="04A0"/></w:tblPr>'
           '<w:tblGrid><w:gridCol w:w="2600"/><w:gridCol w:w="2600"/></w:tblGrid>'
           f'<w:tr>{logo_box("COLLEGE LOGO")}{logo_box("HAL LOGO")}</w:tr></w:tbl>')
    add(tbl)
    blank(1)
    cline("Submitted by", size=11, after=40, italic=True)
    cline(STUDENT, size=13, bold=True, after=20)
    cline("USN: " + USN, size=11, after=120)
    cline("Under the guidance of", size=11, after=40, italic=True)
    cline(GUIDE, size=12, bold=True, after=20)
    cline(DEPT, size=11, after=160)
    cline(COLLEGE, size=13, bold=True, after=20)
    cline("Academic Year " + YEAR, size=11, after=20)

# ======================================================================
# CERTIFICATE PAGE (institution)
# ======================================================================
def certificate_page():
    h1("CERTIFICATE")
    p(f"This is to certify that the internship report entitled \u201c{PROJECT}\u201d is a bona fide "
      f"record of the industrial internship work carried out by {STUDENT}, bearing USN {USN}, "
      f"a final-year student of the {DEPT}, {COLLEGE}, in partial fulfilment of the requirements "
      f"for the award of the degree of Bachelor of Computer Applications during the academic year "
      f"{YEAR}.", indent=False)
    p("The internship was undertaken at the Flight Testing Centre of the Rotary Wing Research & "
      "Design Centre (RWR&DC), Hindustan Aeronautics Limited, on the topic of Data Processing at "
      "Flight Testing Centre. To the best of my knowledge, the work presented in this report has "
      "not been submitted, in part or in full, for the award of any other degree or diploma.")
    blank(3)
    # signature table
    rows = [
        ["________________________", "________________________", "________________________"],
        ["Internal Guide", "Head of the Department", "Principal"],
        [GUIDE, DEPT, PRINCIPAL],
    ]
    table(rows, header=False, font_size=11)
    blank(2)
    p("Examiners:", indent=False)
    p("1. ______________________________               2. ______________________________", indent=False)

# ======================================================================
# INTERNSHIP CERTIFICATE PAGE (HAL) - placeholder for scanned image
# ======================================================================
def hal_certificate_page():
    h1("INTERNSHIP COMPLETION CERTIFICATE")
    p("The original internship completion certificate issued by Hindustan Aeronautics Limited is "
      "reproduced below. The scanned copy of the certificate should be inserted in the space "
      "provided.", indent=False)
    blank(1)
    screenshot_placeholder("Insert scanned HAL Internship Completion Certificate here")

# ======================================================================
# ACKNOWLEDGEMENT
# ======================================================================
def acknowledgement_page():
    h1("ACKNOWLEDGEMENT")
    p("The successful completion of any task is incomplete without acknowledging the people who "
      "made it possible. I take this opportunity to express my sincere gratitude to everyone who "
      "guided and supported me throughout the course of this industrial internship.", indent=False)
    p("I am deeply grateful to the management of Hindustan Aeronautics Limited (HAL) for granting "
      "me the valuable opportunity to undertake my industrial internship at one of the most "
      "prestigious aerospace organisations in the country. The exposure to a live engineering "
      "environment has been an invaluable experience in my academic and professional journey.")
    p("I extend my heartfelt thanks to the Rotary Wing Research & Design Centre (RWR&DC) and, in "
      "particular, to the Flight Testing Centre for permitting me to work in the domain of flight "
      "data processing and for providing access to the technical insight that shaped this project.")
    p(f"I sincerely thank my project mentor at HAL, {MENTOR}, for the constant guidance, technical "
      "direction and encouragement extended to me during the internship. The mentorship I received "
      "helped me understand the rigour and discipline that aerospace engineering demands.")
    p(f"I am thankful to my internal guide, {GUIDE}, and to the {DEPT}, {COLLEGE}, for their "
      "continuous academic support, valuable suggestions and encouragement during the preparation "
      f"of this report. I also express my gratitude to the Principal, {PRINCIPAL}, and to all the "
      "faculty members of the department for their cooperation and motivation.")
    p("Finally, I wish to express my profound gratitude to my parents and friends for their "
      "unconditional support, patience and encouragement, without which this internship and report "
      "would not have been possible.")
    blank(2)
    add(para(run(STUDENT, bold=True, size=12, font="Times New Roman"),
             align="right", spacing_after=20, line=240))
    add(para(run("BCA Final Year   |   USN: " + USN, size=11, font="Times New Roman"),
             align="right", spacing_after=20, line=240))

# ======================================================================
# ABSTRACT
# ======================================================================
def abstract_page():
    h1("ABSTRACT")
    p("This report presents the work carried out during an industrial internship at the Flight "
      "Testing Centre of the Rotary Wing Research & Design Centre (RWR&DC), Hindustan Aeronautics "
      "Limited (HAL). The internship focused on the theme of data processing at the flight testing "
      "centre, an activity that lies at the heart of validating the airworthiness and performance "
      "of military and civil rotorcraft and fixed-wing aircraft developed by HAL.", indent=False)
    p("Flight testing generates enormous volumes of telemetry and instrumentation data that must "
      "be acquired, cleaned, processed, analysed and visualised before any engineering conclusion "
      "can be drawn. During the internship, the principles of flight test instrumentation, "
      "telemetry, sensor data acquisition and the systematic processing of recorded parameters "
      "were studied in depth. This theoretical understanding was consolidated through the design "
      "and development of a software project titled the Advanced Live Aircraft Tracking and Flight "
      "Data Processing System.")
    p("The developed system is a browser-based application built using HTML5, CSS3 and JavaScript, "
      "employing the HTML Canvas API for real-time graphical visualisation. It simulates a live "
      "operational picture of multiple aircraft, processes their telemetry parameters such as "
      "altitude, ground speed and heading, raises threshold-based alerts, and presents the "
      "information through an operations dashboard, a radar scope, an analytics module and a "
      "mission-control view. The system demonstrates, on a reduced and safe scale, the core "
      "data-processing workflow practised at a flight testing centre.")
    p("The internship strengthened my understanding of the aerospace domain, the role of software "
      "engineers in flight testing, real-time data visualisation, and disciplined engineering "
      "documentation. The learning outcomes, challenges encountered, results obtained and the "
      "future scope of the work are discussed in detail in the chapters that follow.")
    blank(1)
    p("Keywords: Flight Testing, Data Processing, Telemetry, Live Aircraft Tracking, HTML5 Canvas, "
      "Real-time Dashboard, HAL, RWR&DC, Aerospace Software.", indent=False)

# ======================================================================
# TABLE OF CONTENTS  (static, page numbers approximate)
# ======================================================================
TOC_ENTRIES = [
    ("", "Cover Page", "i"),
    ("", "Certificate", "ii"),
    ("", "Internship Completion Certificate", "iii"),
    ("", "Acknowledgement", "iv"),
    ("", "Abstract", "v"),
    ("", "Table of Contents", "vi"),
    ("1", "Introduction", "1"),
    ("2", "About Hindustan Aeronautics Limited (HAL)", "4"),
    ("3", "Rotary Wing Research & Design Centre (RWR&DC)", "8"),
    ("4", "Flight Testing Centre", "11"),
    ("5", "Internship Activities", "15"),
    ("6", "System Requirements", "18"),
    ("7", "Technologies Used", "20"),
    ("8", "Project Overview", "24"),
    ("9", "System Architecture and Design Diagrams", "27"),
    ("10", "System Modules", "31"),
    ("11", "Database Design", "35"),
    ("12", "Project Screenshots", "37"),
    ("13", "Source Code Explanation", "42"),
    ("14", "Testing", "47"),
    ("15", "Advantages", "49"),
    ("16", "Learning Outcomes", "51"),
    ("17", "Challenges", "53"),
    ("18", "Future Scope", "55"),
    ("19", "Conclusion", "57"),
    ("20", "Bibliography", "59"),
    ("", "Appendix", "60"),
]

def _index_cell(text, w, align="center", bold=False, size=13, header=False):
    # bold black borders, white background, vertically centred, tall rows
    tcpr = (f'<w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>'
            '<w:tcBorders>'
            '<w:top w:val="single" w:sz="12" w:color="000000"/>'
            '<w:left w:val="single" w:sz="12" w:color="000000"/>'
            '<w:bottom w:val="single" w:sz="12" w:color="000000"/>'
            '<w:right w:val="single" w:sz="12" w:color="000000"/>'
            '</w:tcBorders>'
            '<w:tcMar>'
            '<w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/>'
            '<w:left w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/>'
            '</w:tcMar><w:vAlign w:val="center"/></w:tcPr>')
    pc = para(run(text, bold=bold, size=size, font="Times New Roman"),
              align=align, spacing_before=80, spacing_after=80, line=240)
    return f"<w:tc>{tcpr}{pc}</w:tc>"

def index_table(rows, widths):
    """Render an INDEX table matching the standard college report style:
    bold black borders, bold header (Chapter | Particulars | Page),
    centred chapter number, left particulars, centred page."""
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    tbl = ('<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/>'
           '<w:tblW w:w="9360" w:type="dxa"/><w:jc w:val="center"/>'
           '<w:tblBorders>'
           '<w:top w:val="single" w:sz="12" w:color="000000"/>'
           '<w:left w:val="single" w:sz="12" w:color="000000"/>'
           '<w:bottom w:val="single" w:sz="12" w:color="000000"/>'
           '<w:right w:val="single" w:sz="12" w:color="000000"/>'
           '<w:insideH w:val="single" w:sz="12" w:color="000000"/>'
           '<w:insideV w:val="single" w:sz="12" w:color="000000"/>'
           '</w:tblBorders><w:tblLook w:val="04A0"/></w:tblPr>'
           f'<w:tblGrid>{grid}</w:tblGrid>')
    # header row
    tbl += ('<w:tr><w:trPr><w:tblHeader/></w:trPr>'
            + _index_cell("Chapter", widths[0], align="center", bold=True, size=14)
            + _index_cell("Particulars", widths[1], align="center", bold=True, size=14)
            + _index_cell("Page", widths[2], align="center", bold=True, size=14)
            + "</w:tr>")
    for num, title, pg in rows:
        tbl += ("<w:tr>"
                + _index_cell(num, widths[0], align="center", size=13)
                + _index_cell(title, widths[1], align="left", size=13)
                + _index_cell(pg, widths[2], align="center", size=11)
                + "</w:tr>")
    tbl += "</w:tbl>"
    add(tbl)

def toc_page():
    page_break()
    # Centred bold "INDEX" title
    add(para(run("INDEX", bold=True, size=20, font="Times New Roman"),
             align="center", spacing_before=120, spacing_after=240, line=360, keep=True))
    # Only the numbered chapters and appendix are listed, matching the sample style.
    rows = [(num, title, pg) for (num, title, pg) in TOC_ENTRIES if num] \
           + [("", "Appendix", "60")]
    index_table(rows, widths=[1500, 6360, 1500])

print("Front matter loaded.")



# ======================================================================
# CHAPTER 1 — INTRODUCTION
# ======================================================================
def chapter1():
    h1("CHAPTER 1   INTRODUCTION")
    h2("1.1  Introduction to Industrial Internship")
    p("An industrial internship is a structured period of supervised professional experience that "
      "bridges the gap between academic learning and the practical demands of industry. For a "
      "student of Bachelor of Computer Applications, an internship is the first formal opportunity "
      "to observe how the concepts learned in the classroom are applied to solve real engineering "
      "problems in a disciplined organisational environment. The internship described in this "
      "report was carried out at Hindustan Aeronautics Limited, a premier aerospace and defence "
      "undertaking of the Government of India, within its Rotary Wing Research & Design Centre.")
    p("Unlike an academic project, an internship places the student inside a working organisation "
      "where deadlines, safety, documentation standards and inter-departmental coordination are "
      "part of everyday life. The student is expected to observe, to learn the working culture, to "
      "contribute meaningfully within the limits of the role assigned, and to reflect on the "
      "experience in a manner that strengthens both technical competence and professional "
      "attitude. This report is the outcome of that reflective process.")
    h2("1.2  Importance of Internship")
    p("The importance of an internship cannot be overstated in a technology-driven career. It "
      "exposes the student to the latest tools, processes and standards practised in industry, "
      "many of which evolve faster than the academic curriculum. It develops soft skills such as "
      "communication, teamwork, time management and professional ethics that are difficult to "
      "cultivate within the four walls of a classroom. It also helps the student make informed "
      "decisions about career specialisation by experiencing a real domain first hand.")
    p("For an organisation such as HAL, hosting interns is also a means of nurturing future talent "
      "and of introducing young engineers to the rigour and responsibility that aerospace work "
      "demands. The internship therefore benefits both the student and the host organisation, and "
      "it is increasingly recognised by universities as an essential component of professional "
      "education.")
    h2("1.3  Objectives of the Internship")
    p("The internship was undertaken with a clear set of objectives, which guided the activities "
      "performed throughout its duration:", indent=False)
    bullet("To understand the structure, functions and engineering culture of a national aerospace organisation.")
    bullet("To study the role of the Flight Testing Centre and the importance of flight data in certifying aircraft.")
    bullet("To learn the principles of flight test instrumentation, telemetry and data processing.")
    bullet("To apply software engineering knowledge to build a practical data-processing and visualisation tool.")
    bullet("To develop the Advanced Live Aircraft Tracking and Flight Data Processing System as a demonstrative project.")
    bullet("To cultivate professional discipline, documentation skills and an appreciation of engineering safety.")
    h2("1.4  Importance of Flight Testing")
    p("Flight testing is the process by which a newly designed or modified aircraft is evaluated "
      "in actual flight to confirm that it performs safely and as predicted. No aircraft, however "
      "carefully designed and simulated, is cleared for service until it has been proven through a "
      "systematic flight test programme. Flight testing validates aerodynamic performance, "
      "structural integrity, propulsion behaviour, control characteristics and the functioning of "
      "every onboard system across the entire operating envelope of the aircraft.")
    p("Because flight testing involves taking an aircraft into unproven regions of its performance "
      "envelope, it is an inherently demanding activity that depends on meticulous planning, "
      "precise instrumentation and immediate, accurate data. The flight test engineer relies on a "
      "continuous stream of measurements to decide whether it is safe to proceed to the next test "
      "point. This makes the acquisition and processing of flight data a mission-critical function.")
    h2("1.5  Importance of Data Processing")
    p("Modern flight test aircraft are fitted with hundreds of sensors that record parameters such "
      "as altitude, airspeed, attitude, engine performance, vibration, temperature, control "
      "positions and structural loads. These sensors generate very large volumes of raw data, "
      "sampled many times per second, which by themselves are meaningless until they are "
      "processed. Data processing converts this raw stream into calibrated, validated and "
      "meaningful engineering information that can be analysed and visualised.")
    p("The data-processing pipeline typically includes acquisition, calibration, filtering, "
      "validation, derivation of secondary parameters, time-correlation, storage and "
      "visualisation. Errors or delays at any stage can compromise the safety of a test or the "
      "validity of an engineering conclusion. The internship focused precisely on understanding "
      "this pipeline and on building a software tool that demonstrates its visualisation and "
      "real-time monitoring stages.")
    h2("1.6  Role of Software Engineers in Aerospace")
    p("Software has become central to every aspect of modern aerospace, from the embedded flight "
      "control systems aboard the aircraft to the ground-based tools that process and display test "
      "data. Software engineers in aerospace work on data acquisition systems, telemetry decoders, "
      "real-time monitoring dashboards, post-flight analysis tools, simulation environments and "
      "decision-support systems. Their work demands not only programming skill but also an "
      "appreciation of reliability, accuracy and safety, because the information produced by their "
      "software directly influences decisions that affect human lives and costly assets.")
    p("During the internship, the role of the software engineer was experienced through the design "
      "of a real-time aircraft tracking and data-processing application. This experience "
      "highlighted how front-end visualisation, efficient data handling and clear user interfaces "
      "contribute to the larger goal of safe and effective flight testing.")
    h2("1.7  Organisation of the Report")
    p("This report is organised into twenty chapters. Following this introduction, the report "
      "describes the host organisation, the design centre and the flight testing centre. It then "
      "details the internship activities, system requirements and technologies used, before "
      "presenting the project in depth through its architecture, modules, database design, "
      "screenshots and source-code explanation. The report concludes with testing, advantages, "
      "learning outcomes, challenges, future scope, a conclusion and a bibliography.")

# ======================================================================
# CHAPTER 2 — ABOUT HAL
# ======================================================================
def chapter2():
    h1("CHAPTER 2   ABOUT HINDUSTAN AERONAUTICS LIMITED (HAL)")
    h2("2.1  Introduction to HAL")
    p("Hindustan Aeronautics Limited, commonly known as HAL, is one of the largest and oldest "
      "aerospace and defence manufacturing organisations in Asia. It is a public-sector undertaking "
      "of the Government of India and operates under the administrative control of the Ministry of "
      "Defence. HAL is engaged in the design, development, manufacture, repair, overhaul and "
      "upgrade of aircraft, helicopters, aero-engines, avionics and a wide range of accessories "
      "and systems for both military and civil applications.", indent=False)
    p("Over the decades, HAL has played a defining role in building the self-reliance of the "
      "country in aerospace technology. Its products serve the Indian Air Force, the Indian Army, "
      "the Indian Navy, the Coast Guard and several civil operators, and a number of its products "
      "and components are also exported.")
    h2("2.2  Historical Background")
    p("The origins of HAL trace back to the establishment of an aircraft manufacturing company in "
      "Bengaluru in 1940, which was later taken over and expanded by the Government. Through a "
      "series of mergers and reorganisations, the modern Hindustan Aeronautics Limited came into "
      "being and grew into a national centre of aerospace excellence. From the licensed production "
      "of aircraft in its early years, HAL progressively built indigenous design capability, "
      "culminating in the development of fully home-grown aircraft and helicopters.")
    p("This historical journey reflects the larger national aspiration of achieving self-reliance "
      "in defence technology, and HAL has remained at the forefront of that effort.")
    h2("2.3  Organisational Structure")
    p("HAL is organised into a number of complexes and divisions spread across several cities in "
      "India. These include design and research centres, manufacturing divisions, overhaul "
      "divisions and dedicated facilities for engines, avionics and systems. Each division "
      "specialises in a particular class of products or services, and the design centres are "
      "responsible for the research and development of new platforms. The Rotary Wing Research & "
      "Design Centre, where this internship was carried out, is one such specialised design centre.")
    h2("2.4  Manufacturing Units and Divisions")
    tabcap("Representative complexes and divisions of HAL")
    table([
        ["Complex / Centre", "Primary Focus"],
        ["Bangalore Complex", "Aircraft, engines, aerospace systems and overhaul"],
        ["MiG Complex", "Fighter aircraft manufacture and overhaul"],
        ["Helicopter Complex", "Design and manufacture of rotary-wing aircraft"],
        ["Accessories Complex", "Avionics, systems and accessories"],
        ["Design Centres (incl. RWR&DC)", "Research, design and development of new platforms"],
    ], widths=[3600, 5760])
    h2("2.5  Major Products")
    p("HAL has produced and continues to produce a wide spectrum of aerospace products. The major "
      "categories are summarised below.", indent=False)
    h3("2.5.1  Fixed-Wing Aircraft")
    bullet("Tejas Light Combat Aircraft (LCA) \u2014 an indigenous multi-role fighter.")
    bullet("HTT-40 \u2014 a basic trainer aircraft for ab-initio pilot training.")
    bullet("HJT-36 Sitara \u2014 an intermediate jet trainer.")
    bullet("Licence-produced fighters and transport aircraft for the armed forces.")
    h3("2.5.2  Helicopters")
    bullet("Advanced Light Helicopter (ALH) Dhruv \u2014 a versatile multi-role helicopter.")
    bullet("Light Combat Helicopter (LCH) Prachand \u2014 a dedicated attack helicopter.")
    bullet("Light Utility Helicopter (LUH) \u2014 a single-engine utility helicopter.")
    h3("2.5.3  Engines, Avionics and Systems")
    bullet("Aero-engines produced under licence and indigenous engine development programmes.")
    bullet("Avionics, communication, navigation and mission systems.")
    bullet("Accessories, hydraulic systems, instruments and ground support equipment.")
    h2("2.6  Research and Development")
    p("Research and development is the backbone of HAL. Dedicated design centres carry out the "
      "conceptual design, detailed engineering, prototyping and flight testing of new aircraft and "
      "helicopters. These centres employ advanced tools for computer-aided design, computational "
      "fluid dynamics, structural analysis, simulation and flight testing. The Flight Testing "
      "Centre, where the present internship was based, is the stage at which the products of this "
      "research are validated in actual flight.")
    h2("2.7  Achievements")
    p("HAL has achieved several milestones that underline its national importance. It has "
      "indigenously designed and flight-tested fighter aircraft and helicopters, supported every "
      "major operation of the Indian armed forces with maintenance and overhaul services, and "
      "contributed to the national space programme through the manufacture of structures and "
      "systems for launch vehicles. These achievements have established HAL as a cornerstone of "
      "the nation's strategic capability.")
    h2("2.8  Mission, Vision and the Way Forward")
    p("The mission of HAL is to become a globally competitive aerospace and defence organisation "
      "by achieving self-reliance in the design, development and manufacture of aircraft and "
      "related systems. Its vision emphasises indigenous capability, technological excellence and "
      "customer satisfaction. Looking ahead, HAL continues to invest in next-generation platforms, "
      "unmanned systems, advanced avionics and digital engineering, positioning itself for the "
      "future of aerospace.")

# ======================================================================
# CHAPTER 3 — RWR&DC
# ======================================================================
def chapter3():
    h1("CHAPTER 3   ROTARY WING RESEARCH & DESIGN CENTRE (RWR&DC)")
    h2("3.1  Introduction")
    p("The Rotary Wing Research & Design Centre, abbreviated as RWR&DC, is the dedicated design "
      "house of HAL responsible for the research, design and development of helicopters and other "
      "rotary-wing aircraft. It brings together specialists in aerodynamics, structures, dynamics, "
      "systems, avionics and flight testing to create indigenous rotorcraft that meet the demanding "
      "requirements of the Indian armed forces and civil operators.", indent=False)
    h2("3.2  History and Evolution")
    p("RWR&DC was established to consolidate and strengthen HAL's capability in helicopter design. "
      "From its early work on light helicopters, the centre progressively built the expertise "
      "required to design more complex and capable machines. Its growth mirrors the maturing of "
      "indigenous rotary-wing technology in the country, moving from licence-based production "
      "towards fully indigenous design and development.")
    h2("3.3  Functions of the Centre")
    p("The principal functions of RWR&DC include conceptual and detailed design of helicopters, "
      "structural and dynamic analysis, systems integration, prototype development, ground testing "
      "and flight testing. The centre also supports the certification of its products and provides "
      "engineering support throughout the service life of the helicopters it designs.")
    h2("3.4  Departments and Disciplines")
    tabcap("Major engineering disciplines within RWR&DC")
    table([
        ["Discipline", "Responsibility"],
        ["Aerodynamics", "Rotor and airframe aerodynamic design and performance prediction"],
        ["Structures", "Airframe structural design, strength and fatigue analysis"],
        ["Dynamics", "Rotor dynamics, vibration and loads"],
        ["Systems", "Hydraulics, fuel, electrical and mechanical systems"],
        ["Avionics", "Navigation, communication, mission and display systems"],
        ["Flight Testing", "Instrumentation, telemetry, data processing and flight evaluation"],
    ], widths=[2700, 6660])
    h2("3.5  Helicopter Design and Development")
    p("Helicopter design is among the most challenging tasks in aeronautical engineering because "
      "of the complex aerodynamic and dynamic behaviour of the rotor system. RWR&DC handles the "
      "complete design cycle, from establishing requirements and configuration studies through "
      "detailed design, analysis and prototype manufacture, to ground and flight testing. The "
      "centre has contributed to platforms such as the Advanced Light Helicopter, the Light Combat "
      "Helicopter and the Light Utility Helicopter.")
    h2("3.6  Testing and Research")
    p("Before a helicopter is cleared for service, it undergoes extensive ground testing and "
      "flight testing. Ground tests include structural tests, whirl-tower rotor tests and system "
      "integration tests, while flight tests evaluate the machine across its operating envelope. "
      "The research activity of the centre continually advances the methods, materials and "
      "technologies used in rotorcraft design.")
    h2("3.7  Contribution to Indian Defence")
    p("The rotorcraft developed by RWR&DC have significantly enhanced the operational capability "
      "of the Indian armed forces. Indigenous helicopters reduce dependence on imports, provide "
      "platforms tailored to Indian operating conditions such as high-altitude regions, and "
      "strengthen the nation's strategic autonomy. The centre's work is thus a direct contribution "
      "to national defence and self-reliance.")

# ======================================================================
# CHAPTER 4 — FLIGHT TESTING CENTRE
# ======================================================================
def chapter4():
    h1("CHAPTER 4   FLIGHT TESTING CENTRE")
    h2("4.1  Introduction")
    p("The Flight Testing Centre is the facility within the design organisation where prototype "
      "and production aircraft are subjected to systematic flight evaluation. It is the place where "
      "design predictions meet physical reality, and where the airworthiness and performance of an "
      "aircraft are ultimately proven. The internship was based in this centre, with a focus on "
      "the data-processing activity that underpins all flight testing.", indent=False)
    h2("4.2  Purpose of Flight Testing")
    p("The purpose of flight testing is to verify that an aircraft performs safely and meets its "
      "specified requirements across the whole of its intended operating envelope. Flight testing "
      "confirms aerodynamic performance, handling qualities, structural behaviour, propulsion "
      "performance and the correct functioning of every onboard system. It also identifies "
      "deficiencies that must be corrected before the aircraft enters service, and it generates "
      "the evidence required for certification.")
    h2("4.3  Aircraft Testing Process")
    p("Flight testing is conducted as a carefully sequenced programme. Each flight is planned "
      "around specific test points, and the aircraft is taken progressively from benign to more "
      "demanding conditions. The general flow of a flight test sortie is shown below.")
    diagram_flow([
        "Test Planning and Briefing",
        "Aircraft and Instrumentation Preparation",
        "Take-off and Build-up to Test Point",
        "Execution of Test Points (data acquired)",
        "Real-time Monitoring and Safety Assessment",
        "Landing and Data Retrieval",
        "Post-flight Data Processing and Analysis",
        "Debrief and Reporting",
    ])
    figcap("Typical flight test sortie workflow")
    h2("4.4  Data Collection")
    p("During each test flight, data is collected from a large number of sensors distributed "
      "throughout the aircraft. This data is recorded onboard and, in most cases, simultaneously "
      "transmitted to the ground in real time. The quality and completeness of this data directly "
      "determine the value of the test, and considerable effort is invested in ensuring that the "
      "right parameters are measured accurately and reliably.")
    h2("4.5  Telemetry")
    p("Telemetry is the technology by which measurements made aboard the aircraft are transmitted "
      "to a ground station in real time. A telemetry system encodes the sensor data, transmits it "
      "over a radio link, and decodes it at the ground station, where it is displayed to flight "
      "test engineers. Real-time telemetry allows the ground team to monitor the aircraft "
      "continuously and to make immediate decisions about the safety and progress of the test.")
    h2("4.6  Sensors and Instrumentation")
    p("Flight test instrumentation comprises the sensors, signal-conditioning units, data "
      "acquisition systems and recorders that capture the behaviour of the aircraft. Typical "
      "sensors measure pressure, temperature, acceleration, strain, control-surface position, "
      "rotor speed, engine parameters and many other quantities. The instrumentation must be "
      "carefully calibrated and validated so that the recorded values can be trusted.")
    tabcap("Representative flight-test parameters and their sensors")
    table([
        ["Parameter", "Typical Sensor", "Purpose"],
        ["Altitude", "Air-data / pressure transducer", "Vertical position and performance"],
        ["Airspeed", "Pitot-static system", "Performance and envelope monitoring"],
        ["Attitude", "Gyroscopes / INS", "Handling and stability evaluation"],
        ["Acceleration", "Accelerometers", "Loads and ride quality"],
        ["Engine RPM", "Tachometer", "Propulsion performance"],
        ["Structural strain", "Strain gauges", "Structural load monitoring"],
    ], widths=[2400, 3480, 3480])
    h2("4.7  Flight Validation")
    p("Flight validation is the process of confirming, through measured flight data, that the "
      "aircraft and its systems perform as designed. The processed data is compared against design "
      "predictions and specified requirements. Where deviations are found, they are investigated "
      "and resolved. Successful validation across the full envelope is a prerequisite for "
      "certification and service induction.")
    h2("4.8  Safety Procedures")
    p("Safety is the paramount consideration in flight testing. Comprehensive safety procedures "
      "govern every aspect of a test, including thorough planning, defined test limits, real-time "
      "monitoring of critical parameters, clear communication protocols and well-rehearsed "
      "contingency actions. The real-time data-processing and monitoring function plays a central "
      "role in safety by alerting the test team the moment any parameter approaches a defined limit.")
    h2("4.9  Data Processing at the Flight Testing Centre")
    p("Data processing is the activity that transforms the raw measurements gathered during a "
      "flight into validated engineering information. At the flight testing centre this involves "
      "real-time processing for monitoring during the flight, and detailed post-flight processing "
      "for analysis and reporting. The internship project addresses the real-time monitoring and "
      "visualisation aspect of this activity, demonstrating how live telemetry can be processed and "
      "presented to support situational awareness and safety.")

print("Chapters 1-4 loaded.")



# ======================================================================
# CHAPTER 5 — INTERNSHIP ACTIVITIES
# ======================================================================
def chapter5():
    h1("CHAPTER 5   INTERNSHIP ACTIVITIES")
    h2("5.1  Overview of Activities")
    p("The internship was structured so that an initial period of observation and learning was "
      "followed by the progressive design and development of the project. The activities carried "
      "out are described in this chapter in the order in which they were undertaken. Collectively "
      "they reflect a complete, miniature software-engineering cycle conducted within the context "
      "of a flight testing centre.", indent=False)
    tabcap("Internship schedule (representative)")
    table([
        ["Phase", "Activity", "Outcome"],
        ["Week 1", "Orientation and observation", "Understanding of the centre and safety culture"],
        ["Week 2", "Study of flight testing and data processing", "Domain knowledge"],
        ["Week 3", "Requirement analysis and planning", "Project scope and requirements"],
        ["Week 4", "User-interface design", "Dashboard and screen layouts"],
        ["Week 5-6", "Coding and implementation", "Working tracking system"],
        ["Week 7", "Testing and refinement", "Validated, stable system"],
        ["Week 8", "Documentation and reporting", "Internship report"],
    ], widths=[1500, 4200, 3660])
    h2("5.2  Observation")
    p("The internship began with a period of observation, during which the working of the flight "
      "testing centre and the broader design organisation was studied. This included understanding "
      "the layout of the facility, the flow of activities around a test flight, the safety culture "
      "and the manner in which different engineering disciplines coordinate. This observation phase "
      "provided the context necessary to appreciate the significance of data processing.")
    h2("5.3  Study and Learning")
    p("A substantial part of the early internship was devoted to learning the fundamentals of "
      "flight test instrumentation, telemetry and data processing. Reference material was studied, "
      "and the concepts were discussed with the mentor. This study established the technical "
      "foundation on which the project was later built.")
    h2("5.4  Requirement Analysis")
    p("Requirement analysis involved identifying precisely what the demonstration software should "
      "achieve. It was decided that the system should simulate multiple aircraft, process their "
      "key flight parameters in real time, raise alerts when parameters cross defined thresholds, "
      "and present the information through a clear, professional interface. The functional and "
      "non-functional requirements were documented and agreed with the mentor before design began.")
    h2("5.5  User-Interface Design")
    p("The user interface was designed with the operational context of a flight testing centre in "
      "mind. A dark, high-contrast theme was chosen to resemble professional monitoring consoles "
      "and to reduce eye strain during prolonged use. The screens were organised around a "
      "navigation menu giving access to the dashboard, live tracking, radar scope, analytics, "
      "mission control, alerts and an administration panel.")
    h2("5.6  Coding and Implementation")
    p("The system was implemented using HTML5, CSS3 and JavaScript, with the HTML Canvas API used "
      "for graphical visualisation. The implementation proceeded module by module, beginning with "
      "the data model and simulation engine, followed by the dashboard, the map and radar "
      "visualisations, the analytics module and the remaining screens. The code was written to be "
      "modular, readable and maintainable.")
    h2("5.7  Testing")
    p("Each module was tested as it was developed, and the integrated system was tested as a whole. "
      "Testing confirmed that telemetry was processed correctly, that alerts were raised at the "
      "right thresholds, that the visualisations updated smoothly, and that the interface behaved "
      "correctly across different screen sizes.")
    h2("5.8  Documentation")
    p("Throughout the internship, the design decisions, the structure of the code and the results "
      "obtained were documented. This documentation forms the basis of the present report and "
      "reflects the importance placed on clear records in an engineering organisation.")

# ======================================================================
# CHAPTER 6 — SYSTEM REQUIREMENTS
# ======================================================================
def chapter6():
    h1("CHAPTER 6   SYSTEM REQUIREMENTS")
    h2("6.1  Introduction")
    p("This chapter specifies the hardware and software requirements for the development and "
      "operation of the Advanced Live Aircraft Tracking and Flight Data Processing System. Because "
      "the system is a browser-based application built entirely with client-side web technologies, "
      "its requirements are modest, which is one of its practical advantages.", indent=False)
    h2("6.2  Hardware Requirements")
    tabcap("Hardware requirements")
    table([
        ["Component", "Minimum", "Recommended"],
        ["Processor", "Dual-core 2.0 GHz", "Quad-core 2.5 GHz or higher"],
        ["Memory (RAM)", "4 GB", "8 GB or higher"],
        ["Storage", "200 MB free space", "1 GB free space"],
        ["Display", "1366 x 768", "1920 x 1080 (Full HD)"],
        ["Graphics", "Integrated graphics", "Integrated or discrete graphics"],
    ], widths=[2600, 3380, 3380])
    h2("6.3  Software Requirements")
    tabcap("Software requirements")
    table([
        ["Item", "Requirement"],
        ["Operating System", "Windows 10/11, Linux or macOS"],
        ["Web Browser", "Google Chrome, Microsoft Edge, Mozilla Firefox (latest versions)"],
        ["Code Editor", "Visual Studio Code or any modern text editor"],
        ["Runtime", "Not required \u2014 runs entirely in the browser"],
        ["Web Server (optional)", "Any static file server for hosting"],
    ], widths=[3000, 6360])
    h2("6.4  Browser Requirements")
    p("The application requires a modern web browser with support for HTML5, CSS3, the Canvas API "
      "and contemporary JavaScript. All mainstream browsers released in recent years meet these "
      "requirements. No plug-ins or extensions are needed, which makes the system easy to deploy "
      "and use.")
    h2("6.5  Operating System")
    p("Because the application is platform-independent and runs inside the browser, it operates "
      "identically on Windows, Linux and macOS. This portability is an important practical benefit, "
      "allowing the same tool to be used on a wide range of machines without modification.")
    h2("6.6  Development Environment")
    p("The system was developed using a lightweight development environment consisting of a modern "
      "code editor and a web browser with developer tools. No compilation step or complex build "
      "pipeline was required, which allowed rapid iteration during development. This simplicity is "
      "characteristic of well-structured client-side web applications.")

# ======================================================================
# CHAPTER 7 — TECHNOLOGIES USED
# ======================================================================
def chapter7():
    h1("CHAPTER 7   TECHNOLOGIES USED")
    h2("7.1  Introduction")
    p("This chapter describes in detail the technologies used to build the Advanced Live Aircraft "
      "Tracking and Flight Data Processing System. The system was deliberately built using "
      "standard, open web technologies so that it is portable, lightweight and easy to maintain.",
      indent=False)
    h2("7.2  HTML5")
    p("HyperText Markup Language version 5, or HTML5, is the standard markup language used to "
      "structure the content of web pages. It provides semantic elements, multimedia support and, "
      "most importantly for this project, the Canvas element used for graphical drawing. In the "
      "system, HTML5 defines the structure of every screen, including the navigation menu, the "
      "dashboard cards, the data tables and the containers that host the canvas visualisations.")
    h2("7.3  CSS3")
    p("Cascading Style Sheets version 3, or CSS3, is the language used to control the visual "
      "presentation of web pages. CSS3 was used extensively in the project to create the "
      "professional dark-themed appearance of the interface. Features such as flexible box layout, "
      "grid layout, gradients, rounded corners, shadows and transitions were employed to produce a "
      "polished and responsive design that adapts to different screen sizes.")
    h2("7.4  JavaScript")
    p("JavaScript is the programming language of the web and provides the interactivity and logic "
      "of the application. In this project, JavaScript implements the entire behaviour of the "
      "system: the aircraft data model, the simulation engine that updates aircraft positions, the "
      "processing of telemetry parameters, the generation of alerts, the rendering of the canvas "
      "visualisations and the handling of all user interactions. The logic was written using "
      "functions, objects, arrays and timers in a modular and readable style.")
    h2("7.5  HTML Canvas API")
    p("The Canvas API is a feature of HTML5 that provides a drawing surface on which graphics can "
      "be rendered programmatically using JavaScript. It is the technology that makes the live map "
      "and the radar scope possible. Using the Canvas API, the application draws grid lines, "
      "aircraft symbols, range rings, the rotating radar sweep and the analytics charts. The "
      "canvas is redrawn many times per second to create smooth, real-time animation.")
    h2("7.6  Local Storage and Client-side Data")
    p("The application maintains its data within the browser using in-memory JavaScript structures, "
      "and the browser's local storage mechanism can be used to persist configuration between "
      "sessions. This client-side approach keeps the system self-contained and avoids the need for "
      "a server during demonstration, while still illustrating the data-handling concepts involved.")
    h2("7.7  Responsive Design")
    p("Responsive design is the practice of building interfaces that adapt gracefully to different "
      "screen sizes and devices. The system uses responsive layout techniques so that it remains "
      "usable on large monitors as well as on smaller screens. The layout reflows automatically, "
      "and the canvas visualisations resize to fit the available space.")
    h2("7.8  Real-time Dashboard Concepts")
    p("A real-time dashboard continuously presents the most current state of a system through "
      "concise indicators and visualisations. The project applies these concepts by updating key "
      "performance indicators, the live map, the data tables and the alert feed at regular "
      "intervals. This gives the user an immediate, at-a-glance understanding of the situation, "
      "which is exactly the kind of situational awareness required at a flight testing centre.")
    h2("7.9  Front-end Development Practices")
    p("The development followed sound front-end practices, including separation of structure, "
      "presentation and behaviour, the use of meaningful names, modular functions and consistent "
      "formatting. These practices make the code easier to understand, test and extend, and they "
      "reflect the professional standards expected in industry.")
    tabcap("Summary of technologies and their roles")
    table([
        ["Technology", "Role in the Project"],
        ["HTML5", "Structure and content of all screens; Canvas element"],
        ["CSS3", "Visual styling, theming and responsive layout"],
        ["JavaScript", "Application logic, simulation and data processing"],
        ["Canvas API", "Live map, radar scope and analytics charts"],
        ["Local Storage", "Optional persistence of configuration"],
    ], widths=[2600, 6760])

# ======================================================================
# CHAPTER 8 — PROJECT OVERVIEW
# ======================================================================
def chapter8():
    h1("CHAPTER 8   PROJECT OVERVIEW")
    h2("8.1  Introduction to the Project")
    p("The project developed during the internship is titled the Advanced Live Aircraft Tracking "
      "and Flight Data Processing System. It is a browser-based application that demonstrates, on a "
      "safe and reduced scale, how live aircraft telemetry can be processed and visualised to "
      "support monitoring and decision-making at a flight testing centre. The system grew from an "
      "initial prototype, the HAL Live Tracking Prototype, into the more capable HAL Advanced "
      "Tracking System.", indent=False)
    h2("8.2  Objectives of the Project")
    bullet("To process simulated live telemetry such as altitude, speed and heading for multiple aircraft.")
    bullet("To visualise aircraft positions on a live map and on a radar scope in real time.")
    bullet("To raise alerts automatically when monitored parameters cross defined thresholds.")
    bullet("To summarise processed data through key indicators and analytics charts.")
    bullet("To present all information through a clear, professional and responsive interface.")
    h2("8.3  Problem Statement")
    p("Flight testing produces a continuous stream of telemetry that must be monitored in real "
      "time so that the test team retains complete situational awareness and can react instantly "
      "to any abnormal condition. Without an integrated visual tool, monitoring numerous parameters "
      "for several aircraft simultaneously is difficult and error-prone. The problem addressed by "
      "this project is therefore the real-time processing and clear visualisation of live aircraft "
      "telemetry to support monitoring and safety.")
    h2("8.4  Existing System")
    p("Conventional monitoring at a flight testing centre relies on specialised, often expensive, "
      "ground-station software tightly coupled to specific hardware. Such systems are powerful but "
      "are not easily portable, are not suitable for demonstration or training, and are not "
      "accessible to a student for learning purposes. A simple display of raw numerical values "
      "without integrated visualisation also increases the cognitive load on the operator.")
    h2("8.5  Proposed System")
    p("The proposed system is a lightweight, portable, browser-based application that integrates a "
      "live map, a radar scope, an analytics module and an alerting mechanism into a single "
      "dashboard. It processes telemetry in real time and presents it through intuitive "
      "visualisations. Because it runs in any modern browser without special hardware, it is ideal "
      "for demonstration, learning and as a conceptual model of a real monitoring system.")
    h2("8.6  Need for the System")
    p("The system meets the need for an accessible, integrated and visual means of monitoring live "
      "aircraft telemetry. It reduces operator workload by consolidating information, improves "
      "safety by drawing immediate attention to threshold violations, and serves as an effective "
      "educational tool that illustrates the principles of flight-data processing.")
    h2("8.7  Advantages and Applications")
    p("The system offers real-time monitoring, integrated visualisation, automatic alerting and "
      "platform independence. Beyond its demonstrative role at a flight testing centre, the same "
      "concepts apply to air-traffic situational displays, fleet-tracking dashboards, "
      "unmanned-aircraft monitoring and any application that requires the live visualisation of "
      "moving objects and their parameters.")

# ======================================================================
# CHAPTER 9 — SYSTEM ARCHITECTURE AND DESIGN DIAGRAMS
# ======================================================================
def chapter9():
    h1("CHAPTER 9   SYSTEM ARCHITECTURE AND DESIGN DIAGRAMS")
    h2("9.1  Introduction")
    p("This chapter presents the architecture and design of the system through a set of standard "
      "engineering diagrams. These diagrams describe the overall structure of the system, the flow "
      "of data through it, and the behaviour of its components.", indent=False)
    h2("9.2  System Architecture")
    p("The system follows a layered client-side architecture. A data layer holds the aircraft "
      "model and is updated by a simulation and processing layer, which in turn feeds a "
      "presentation layer composed of the various screens and visualisations.")
    diagram_layers([
        ("Presentation Layer", ["Dashboard", "Tracking", "Radar", "Analytics", "Mission / Alerts"]),
        ("Application / Processing Layer", ["Simulation Engine", "Telemetry Processing", "Alert Engine"]),
        ("Data Layer", ["Aircraft Data Model", "Alert Store", "Configuration"]),
    ])
    figcap("Layered system architecture")
    h2("9.3  Data Flow Diagram \u2014 Level 0")
    p("The context-level data flow diagram shows the system as a single process interacting with "
      "the operator and the telemetry source.")
    diagram_layers([
        ("External Entity", ["Telemetry Source (simulated)"]),
        ("Process", ["Live Aircraft Tracking and Data Processing System"]),
        ("External Entity", ["Flight Test Operator"]),
    ])
    figcap("DFD Level 0 (context diagram)")
    h2("9.4  Data Flow Diagram \u2014 Level 1")
    p("The level-1 data flow diagram decomposes the system into its principal processes and shows "
      "how data moves between them.")
    diagram_flow([
        "1.0  Acquire / Simulate Telemetry",
        "2.0  Process and Validate Parameters",
        "3.0  Evaluate Alert Thresholds",
        "4.0  Update Data Store",
        "5.0  Render Visualisations and Dashboard",
    ])
    figcap("DFD Level 1 (process decomposition)")
    h2("9.5  Use-Case Diagram")
    p("The use-case view describes the interactions between the operator and the system. The main "
      "actors are the Operator and the Administrator.")
    tabcap("Principal use cases")
    table([
        ["Actor", "Use Cases"],
        ["Operator", "Log in; view dashboard; track aircraft; view radar; view analytics; view alerts"],
        ["Administrator", "Register aircraft; remove aircraft; configure refresh rate and thresholds"],
    ], widths=[2400, 6960])
    h2("9.6  Class / Component View")
    p("The principal components of the system and their responsibilities are summarised below. "
      "Although the implementation is written in JavaScript rather than a strictly class-based "
      "language, the design is organised conceptually around these components.")
    diagram_layers([
        ("Core", ["Aircraft (model)", "SimulationEngine", "AlertEngine"]),
        ("Views", ["DashboardView", "RadarView", "AnalyticsView", "AdminView"]),
        ("Support", ["Renderer (Canvas)", "ClockService", "Storage"]),
    ])
    figcap("Component view of the system")
    h2("9.7  Sequence of a Telemetry Update")
    p("The following sequence describes what happens on each processing cycle, from the simulation "
      "of new telemetry through to the update of the display.")
    diagram_flow([
        "Timer triggers a processing cycle",
        "Simulation Engine updates each aircraft's parameters",
        "Telemetry Processing validates and derives values",
        "Alert Engine compares values against thresholds",
        "Data store is updated with new state",
        "Renderer redraws map, radar, tables and charts",
    ])
    figcap("Sequence of a single telemetry update cycle")
    h2("9.8  Activity Diagram")
    p("The activity diagram captures the overall operational flow of the application from login to "
      "continuous monitoring.")
    diagram_flow([
        "Start \u2192 Operator Login",
        "Initialise system and data model",
        "Begin periodic processing loop",
        "Process telemetry and update views",
        "Alert raised? \u2192 Yes: highlight and log; No: continue",
        "Operator navigates between modules",
        "Logout \u2192 End",
    ])
    figcap("Activity diagram of system operation")
    h2("9.9  Deployment View")
    p("The deployment of the system is straightforward. The application consists of static files "
      "that are delivered to the client browser, where all processing and rendering takes place. "
      "No application server or database server is required for the demonstration configuration.")
    diagram_layers([
        ("Client Device", ["Web Browser (HTML/CSS/JS + Canvas)"]),
        ("Delivery (optional)", ["Static File Host / Web Server"]),
    ])
    figcap("Deployment view")

print("Chapters 5-9 loaded.")



# ======================================================================
# CHAPTER 10 — SYSTEM MODULES
# ======================================================================
def chapter10():
    h1("CHAPTER 10   SYSTEM MODULES")
    h2("10.1  Introduction")
    p("The system is organised into a number of functional modules, each accessible from the main "
      "navigation menu. This chapter describes every module in detail, explaining its purpose, the "
      "information it presents and how it processes data.", indent=False)
    h2("10.2  User Authentication Module")
    p("The application opens with a secure login screen that authenticates the operator before "
      "granting access to the monitoring functions. The login screen presents fields for a user "
      "identifier and a password and validates the credentials before revealing the main "
      "interface. This module reflects the access-control requirement of any operational system "
      "and establishes the identity of the user for the session.")
    h2("10.3  Dashboard Module")
    p("The dashboard is the central screen of the application and provides an at-a-glance overview "
      "of the entire operational picture. It displays key performance indicators such as the number "
      "of aircraft airborne, the average altitude and speed across the fleet, and the count of "
      "active alerts. Alongside these indicators it presents the live map, a feed of recent alerts "
      "and a fleet-status table. The dashboard updates continuously so that the operator always "
      "sees the most current state.")
    h2("10.4  Live Aircraft Tracking Module")
    p("The tracking module allows the operator to examine individual aircraft in detail. It "
      "presents a searchable, filterable list of all tracked aircraft and, when an aircraft is "
      "selected, displays its full telemetry including altitude, ground speed, heading, position "
      "and status. This module demonstrates how processed data is presented in a structured, "
      "drill-down fashion.")
    h2("10.5  Live Map Module")
    p("The live map is a Canvas-based visualisation that plots the position of every aircraft on a "
      "gridded representation of the airspace. Each aircraft is drawn as an oriented symbol, "
      "coloured according to its status, and annotated with its callsign and altitude. As the "
      "simulation advances, the symbols move smoothly across the map, providing an intuitive "
      "spatial understanding of the situation.")
    h2("10.6  Radar Scope Module")
    p("The radar scope reproduces the appearance of a primary surveillance radar. It draws "
      "concentric range rings, cross hairs and a rotating sweep, with aircraft shown as contacts on "
      "the display. The rotating sweep is animated continuously, giving the module a realistic and "
      "professional appearance while illustrating polar visualisation of position data.")
    h2("10.7  Flight Information Module")
    p("The flight information presented within the tracking and detail views consolidates the most "
      "important parameters of a selected aircraft into a compact panel. It includes derived "
      "information such as vertical rate and transponder status in addition to the directly "
      "measured parameters, demonstrating the derivation step of data processing.")
    h2("10.8  Alerts Module")
    p("The alerts module is responsible for drawing the operator's attention to abnormal "
      "conditions. Whenever a monitored parameter crosses a defined threshold, an alert is "
      "generated, classified by severity, time-stamped and added to the alert feed. Critical "
      "alerts are visually distinguished from cautions, and the most recent alerts are surfaced on "
      "the dashboard. This module embodies the safety-monitoring purpose of the system.")
    h2("10.9  Analytics Module")
    p("The analytics module processes the telemetry of the whole fleet to produce statistical "
      "summaries and charts. It presents an altitude-distribution histogram, a speed profile and a "
      "table of statistics giving the minimum, maximum and mean of each parameter together with "
      "the number of samples processed. This module illustrates the analytical stage of data "
      "processing.")
    h2("10.10  Mission Control Module")
    p("The mission-control module supports the planning and tracking of flight-test sorties. It "
      "lists active sorties with their objectives, current phase and progress, and presents a "
      "mission timeline that depicts the standard phases of a sortie from briefing to debrief. "
      "This module reflects the operational planning context in which monitoring takes place.")
    h2("10.11  Admin Panel Module")
    p("The administration panel provides management functions. It allows new aircraft to be "
      "registered into the system and existing aircraft to be removed, and it exposes system "
      "settings such as the data refresh rate and the alert threshold. This module demonstrates "
      "create and delete operations on the data model and configurable behaviour.")
    tabcap("Summary of system modules")
    table([
        ["Module", "Primary Function"],
        ["Authentication", "Secure operator login"],
        ["Dashboard", "Consolidated real-time overview"],
        ["Live Tracking", "Per-aircraft telemetry and detail"],
        ["Live Map", "Spatial visualisation of positions"],
        ["Radar Scope", "Polar surveillance display"],
        ["Alerts", "Threshold-based alerting"],
        ["Analytics", "Statistical summaries and charts"],
        ["Mission Control", "Sortie planning and tracking"],
        ["Admin Panel", "Aircraft registry and configuration"],
    ], widths=[2800, 6560])

# ======================================================================
# CHAPTER 11 — DATABASE DESIGN
# ======================================================================
def chapter11():
    h1("CHAPTER 11   DATABASE DESIGN")
    h2("11.1  Introduction")
    p("Although the demonstration system stores its data in memory within the browser, it is "
      "designed around a clear logical data model. This chapter presents that model as it would be "
      "realised in a relational database, were the system to be extended with persistent storage. "
      "Describing the data in this way clarifies the structure of the information the system "
      "processes.", indent=False)
    h2("11.2  Logical Data Model")
    p("The principal entities of the system are the Aircraft, the Telemetry record, the Alert and "
      "the Sortie, together with a User entity for authentication. The relationships between these "
      "entities are summarised below.")
    diagram_layers([
        ("Entities", ["USER", "AIRCRAFT", "TELEMETRY", "ALERT", "SORTIE"]),
    ])
    p("An aircraft produces many telemetry records over time, and may raise many alerts; a sortie "
      "is associated with one aircraft, and a user may oversee many sorties. These one-to-many "
      "relationships form the backbone of the data model.")
    h2("11.3  Database Tables")
    tabcap("AIRCRAFT table")
    table([
        ["Attribute", "Type", "Description"],
        ["aircraft_id", "INTEGER (PK)", "Unique identifier"],
        ["callsign", "VARCHAR", "Aircraft callsign, e.g. HAL-LCH1"],
        ["type", "VARCHAR", "Aircraft type / model"],
        ["status", "VARCHAR", "Current status (normal/caution/critical)"],
    ], widths=[2400, 2400, 4560])
    tabcap("TELEMETRY table")
    table([
        ["Attribute", "Type", "Description"],
        ["telemetry_id", "INTEGER (PK)", "Unique identifier"],
        ["aircraft_id", "INTEGER (FK)", "Reference to AIRCRAFT"],
        ["altitude", "FLOAT", "Altitude in feet"],
        ["speed", "FLOAT", "Ground speed in knots"],
        ["heading", "FLOAT", "Heading in degrees"],
        ["timestamp", "DATETIME", "Time of measurement"],
    ], widths=[2400, 2400, 4560])
    tabcap("ALERT table")
    table([
        ["Attribute", "Type", "Description"],
        ["alert_id", "INTEGER (PK)", "Unique identifier"],
        ["aircraft_id", "INTEGER (FK)", "Reference to AIRCRAFT"],
        ["severity", "VARCHAR", "Caution or critical"],
        ["message", "VARCHAR", "Description of the alert"],
        ["timestamp", "DATETIME", "Time the alert was raised"],
    ], widths=[2400, 2400, 4560])
    h2("11.4  Relationships")
    p("The AIRCRAFT table is the central table. Each row in the TELEMETRY table and each row in "
      "the ALERT table references an aircraft through a foreign key, establishing one-to-many "
      "relationships. The SORTIE table likewise references the aircraft assigned to it. These "
      "relationships preserve referential integrity and allow the data to be queried efficiently.")
    h2("11.5  Sample Data")
    tabcap("Sample AIRCRAFT records")
    table([
        ["aircraft_id", "callsign", "type", "status"],
        ["1", "HAL-LCH1", "LCH Prachand", "normal"],
        ["2", "HAL-ALH2", "ALH Dhruv", "normal"],
        ["3", "HAL-TEJ4", "Tejas LCA", "normal"],
        ["4", "HAL-HTT5", "HTT-40", "caution"],
    ], widths=[2200, 2400, 2960, 1800])

# ======================================================================
# CHAPTER 12 — PROJECT SCREENSHOTS
# ======================================================================
def chapter12():
    h1("CHAPTER 12   PROJECT SCREENSHOTS")
    h2("12.1  Introduction")
    p("This chapter presents the screens of the application. As the system is interactive and "
      "animated, placeholders are provided here in which the actual screenshots captured from the "
      "running application should be inserted. Each placeholder is accompanied by a description of "
      "what the screen shows and how it is used.", indent=False)
    screens = [
        ("Login Screen", "The secure login screen through which the operator gains access to the "
         "system. It validates the user identifier and password before revealing the interface."),
        ("Operations Dashboard", "The main dashboard showing key indicators, the live map, the "
         "recent-alerts feed and the fleet-status table updating in real time."),
        ("Live Aircraft Tracking", "The tracking screen with the searchable aircraft list and the "
         "detailed telemetry panel for the selected aircraft."),
        ("Radar Scope", "The radar display with range rings, cross hairs, the rotating sweep and "
         "aircraft contacts."),
        ("Analytics", "The analytics screen presenting the altitude-distribution histogram, the "
         "speed profile chart and the statistics table."),
        ("Mission Control", "The mission-control screen listing active sorties with their progress "
         "and the mission timeline."),
        ("Alert Feed", "The complete alert feed showing all generated alerts classified by "
         "severity and time-stamped."),
        ("Aircraft Detail", "The detailed view of a single aircraft showing all measured and "
         "derived parameters."),
        ("Admin Panel", "The administration panel for registering and removing aircraft and for "
         "configuring system settings such as refresh rate and thresholds."),
    ]
    for i, (name, desc) in enumerate(screens):
        h3(f"12.{i+2}  {name}")
        screenshot_placeholder(f"Screenshot: {name}")
        figcap(f"{name} of the Advanced Live Aircraft Tracking System")
        p(desc)

# ======================================================================
# CHAPTER 13 — SOURCE CODE EXPLANATION
# ======================================================================
def chapter13():
    h1("CHAPTER 13   SOURCE CODE EXPLANATION")
    h2("13.1  Introduction")
    p("This chapter explains the structure and the principal elements of the source code of the "
      "Advanced Live Aircraft Tracking System. The complete source is provided in the project "
      "files; here the key parts are described to convey how the system works internally.",
      indent=False)
    h2("13.2  Overall Structure")
    p("The application is contained in a single, self-contained HTML file that includes the markup, "
      "the styles within a style block and the logic within a script block. This structure keeps "
      "the demonstration portable while preserving a clear internal separation between structure, "
      "presentation and behaviour.")
    h2("13.3  HTML Structure")
    p("The HTML defines the login screen and the main application shell. The shell consists of a "
      "top bar carrying the branding and clock, a side navigation menu, and a set of page sections "
      "\u2014 one for each module \u2014 of which only the active one is displayed at any time. The "
      "canvas elements that host the map, the radar and the charts are declared within these "
      "sections.")
    h2("13.4  CSS Styling")
    p("The CSS establishes the dark, high-contrast theme of the interface using a set of colour "
      "variables, and lays out the screens using flexible-box and grid layouts. Cards, tables, "
      "badges, alerts and buttons are styled consistently, and media queries adapt the layout for "
      "smaller screens.")
    h2("13.5  The Aircraft Data Model")
    p("The state of the system is held in an array of aircraft objects. Each object stores the "
      "callsign, type, simulated position, altitude, speed, heading and status of one aircraft. "
      "This array is the single source of truth from which every view is rendered.")
    code_block([
        "let aircraft = [",
        "  {cs:\"HAL-LCH1\", type:\"LCH Prachand\", lat:120, lon:90,",
        "   alt:14200, spd:165, hdg:90, status:\"ok\"},",
        "  // ... further aircraft ...",
        "];",
    ])
    h2("13.6  The Simulation and Processing Function")
    p("On every processing cycle the simulate function advances each aircraft according to its "
      "heading and speed, introduces small random variations to mimic real telemetry, and then "
      "derives the status of the aircraft by comparing its parameters against the configured "
      "thresholds. This single function embodies the simulation, processing and validation stages "
      "of the data pipeline.")
    code_block([
        "function simulate(){",
        "  aircraft.forEach(a=>{",
        "    const rad = a.hdg*Math.PI/180;",
        "    a.lat += Math.cos(rad)*(a.spd/120);",
        "    a.lon += Math.sin(rad)*(a.spd/120);",
        "    a.alt += (Math.random()-0.5)*200;",
        "    if(a.alt > THR())      a.status='crit';",
        "    else if(a.spd>420||a.alt<5000) a.status='warn';",
        "    else                   a.status='ok';",
        "  });",
        "  renderAll();",
        "}",
    ])
    h2("13.7  Canvas Drawing and Animation")
    p("The map and radar are drawn using the Canvas two-dimensional context. The drawMap function "
      "clears the canvas, draws the grid, and then iterates over the aircraft array drawing an "
      "oriented, colour-coded symbol for each aircraft together with its labels. The radar sweep "
      "is animated by recomputing the sweep angle from the current time on each frame, producing a "
      "smooth rotation.")
    code_block([
        "function drawMap(){",
        "  const x = canvas.getContext('2d');",
        "  x.clearRect(0,0,canvas.width,canvas.height);",
        "  drawGrid(x);",
        "  aircraft.forEach(a=>{",
        "    x.save(); x.translate(px,py); x.rotate(a.hdg*Math.PI/180);",
        "    x.fillStyle = colourOf(a.status);",
        "    drawAircraftSymbol(x);",
        "    x.restore();",
        "  });",
        "}",
    ])
    h2("13.8  Functions, Objects and Arrays")
    p("The code makes extensive use of JavaScript functions to organise behaviour into discrete, "
      "reusable units; of objects to represent aircraft, alerts and sorties; and of arrays to hold "
      "collections of these objects. Array methods such as forEach, filter, map and reduce are used "
      "to process the data concisely \u2014 for example, computing fleet averages with reduce and "
      "filtering the tracked list with filter.")
    h2("13.9  Event Handling and DOM Manipulation")
    p("User interactions are handled through event listeners attached to the navigation menu, the "
      "search box, the filter control and the buttons of the administration panel. The rendering "
      "functions update the interface by writing into the appropriate elements of the document, "
      "which is the essence of dynamic web behaviour. A single delegated listener on the navigation "
      "menu switches between modules by toggling the active page.")
    code_block([
        "nav.addEventListener('click', e => {",
        "  const link = e.target.closest('a');",
        "  if(!link) return;",
        "  showPage(link.dataset.p);   // switch active module",
        "});",
    ])
    h2("13.10  Timers and Real-time Updates")
    p("The real-time behaviour of the system is driven by timers. One interval timer invokes the "
      "simulation and processing function at the configured refresh rate, while a faster timer "
      "animates the radar sweep, and a one-second timer updates the clock. The refresh rate can be "
      "changed at run time from the administration panel, after which the timer is reset to the new "
      "interval.")
    h2("13.11  Summary")
    p("Taken together, these elements show how a real-time monitoring application can be built "
      "entirely with standard web technologies. The clear separation of the data model, the "
      "processing logic, the rendering functions and the event handlers makes the code "
      "understandable and maintainable, reflecting sound software-engineering practice.")

# ======================================================================
# CHAPTER 14 — TESTING
# ======================================================================
def chapter14():
    h1("CHAPTER 14   TESTING")
    h2("14.1  Introduction")
    p("Testing is the process of evaluating a system to confirm that it meets its requirements and "
      "behaves correctly. This chapter describes the testing carried out on the Advanced Live "
      "Aircraft Tracking System, the types of testing performed and representative test cases with "
      "their results.", indent=False)
    h2("14.2  Unit Testing")
    p("Unit testing verifies that individual functions behave correctly in isolation. The core "
      "functions \u2014 the simulation step, the status-derivation logic, the averaging "
      "calculations and the alert generation \u2014 were each exercised with controlled inputs to "
      "confirm that they produced the expected outputs.")
    h2("14.3  Integration Testing")
    p("Integration testing confirms that the modules work correctly together. It was verified that "
      "an update to the aircraft data model produced consistent changes across the dashboard, the "
      "map, the radar, the tracking list and the analytics, and that an alert raised by the "
      "processing logic appeared correctly in both the dashboard feed and the full alert feed.")
    h2("14.4  System Testing")
    p("System testing evaluated the application as a whole against its requirements. The complete "
      "workflow, from login through monitoring to administration, was exercised to confirm that "
      "the system behaved as intended in normal use.")
    h2("14.5  Performance Testing")
    p("Performance testing confirmed that the visualisations updated smoothly and that the "
      "interface remained responsive while processing and rendering the telemetry of multiple "
      "aircraft at the configured refresh rate.")
    h2("14.6  Test Cases and Results")
    tabcap("Representative test cases and results")
    table([
        ["No.", "Test Case", "Expected Result", "Actual", "Status"],
        ["1", "Valid login", "Access granted to dashboard", "As expected", "Pass"],
        ["2", "Altitude exceeds threshold", "Critical alert raised", "As expected", "Pass"],
        ["3", "Select aircraft in tracking list", "Detail panel updates", "As expected", "Pass"],
        ["4", "Add aircraft via admin panel", "Aircraft appears in all views", "As expected", "Pass"],
        ["5", "Remove aircraft", "Aircraft removed everywhere", "As expected", "Pass"],
        ["6", "Change refresh rate", "Update interval changes", "As expected", "Pass"],
        ["7", "Resize browser window", "Layout and canvas adapt", "As expected", "Pass"],
        ["8", "Search and filter aircraft", "List filtered correctly", "As expected", "Pass"],
    ], widths=[700, 3200, 2900, 1500, 1060])
    h2("14.7  Summary of Testing")
    p("All planned test cases passed, confirming that the system meets its functional and "
      "non-functional requirements. The testing process also reinforced the importance of "
      "systematic verification in producing reliable software.")

print("Chapters 10-14 loaded.")



# ======================================================================
# CHAPTER 15 — ADVANTAGES
# ======================================================================
def chapter15():
    h1("CHAPTER 15   ADVANTAGES")
    h2("15.1  Introduction")
    p("This chapter sets out the advantages of the Advanced Live Aircraft Tracking and Flight Data "
      "Processing System, considering both its technical merits and its practical value in the "
      "context of a flight testing centre.", indent=False)
    h2("15.2  Real-time Monitoring")
    p("The foremost advantage of the system is that it provides genuine real-time monitoring. The "
      "telemetry of every aircraft is processed and the displays are refreshed continuously, so "
      "the operator always has an immediate and accurate picture of the situation. Real-time "
      "awareness is the foundation of safe flight testing.")
    h2("15.3  Improved Safety")
    p("By comparing every monitored parameter against defined thresholds and raising alerts the "
      "moment a limit is approached or exceeded, the system actively contributes to safety. It "
      "draws the operator's attention to abnormal conditions far more reliably than the manual "
      "scanning of numerical displays.")
    h2("15.4  Integrated Visualisation")
    p("The system integrates several complementary visualisations \u2014 the live map, the radar "
      "scope and the analytics charts \u2014 within a single dashboard. This integration reduces "
      "the operator's workload and provides a richer understanding of the situation than any single "
      "view could.")
    h2("15.5  Effective Flight and Mission Management")
    p("The mission-control module links the monitoring of telemetry to the management of "
      "flight-test sorties, supporting effective planning and tracking of test activities. This "
      "places the live data within its operational context.")
    h2("15.6  Platform Independence and Portability")
    p("Because the system runs in any modern web browser without special hardware or software, it "
      "is highly portable. It can be demonstrated or used on a wide range of devices and operating "
      "systems, which is a significant practical advantage over hardware-bound systems.")
    h2("15.7  Decision Support")
    p("By consolidating and clearly presenting processed information, the system supports timely "
      "and well-informed decisions. The combination of indicators, visualisations and alerts gives "
      "the operator the information needed to decide whether to continue, modify or terminate a "
      "test.")
    tabcap("Summary of advantages")
    table([
        ["Advantage", "Benefit"],
        ["Real-time monitoring", "Immediate situational awareness"],
        ["Threshold alerting", "Early detection of abnormal conditions"],
        ["Integrated visualisation", "Reduced operator workload"],
        ["Portability", "Runs on any modern browser and platform"],
        ["Decision support", "Faster, better-informed decisions"],
    ], widths=[3000, 6360])

# ======================================================================
# CHAPTER 16 — LEARNING OUTCOMES
# ======================================================================
def chapter16():
    h1("CHAPTER 16   LEARNING OUTCOMES")
    h2("16.1  Introduction")
    p("This chapter reflects on what was learned during the internship. The learning outcomes span "
      "technical knowledge, professional skills and an appreciation of the engineering culture of "
      "a national aerospace organisation.", indent=False)
    h2("16.2  Technical Skills")
    p("The internship deepened my understanding of front-end web development and, in particular, of "
      "real-time visualisation using the HTML Canvas API. I learned how to design a data model, "
      "how to process streaming data, how to derive and validate parameters, and how to render "
      "information graphically and update it smoothly in real time.")
    h2("16.3  Programming Skills")
    p("My JavaScript programming skills improved considerably through the practical work of "
      "building the system. I gained confidence in organising code into modular functions, in "
      "working with objects and arrays, in using higher-order array methods, in handling events "
      "and in managing timers for real-time behaviour.")
    h2("16.4  Domain Knowledge")
    p("Beyond programming, the internship gave me a genuine understanding of flight testing and the "
      "central role of data processing within it. I learned about instrumentation, telemetry, "
      "real-time monitoring and the disciplined workflow of a flight test sortie \u2014 knowledge "
      "that is rare and valuable for a computing student.")
    h2("16.5  Communication and Documentation")
    p("I learned to communicate clearly with my mentor, to ask informed questions and to document "
      "my work thoroughly. The preparation of this report itself was an important exercise in "
      "technical writing and in presenting work in a structured, professional manner.")
    h2("16.6  Problem Solving")
    p("The development work presented many problems to solve, from rendering glitches to logical "
      "errors in the processing. Working through these problems strengthened my analytical and "
      "debugging skills and taught me to approach difficulties methodically.")
    h2("16.7  Professional Ethics and Industrial Exposure")
    p("Perhaps the most valuable outcome was exposure to the professional ethics and discipline of "
      "an aerospace organisation, where safety, accuracy and accountability are paramount. This "
      "experience has shaped my understanding of what it means to work as a responsible engineer.")
    tabcap("Summary of learning outcomes")
    table([
        ["Area", "Outcome"],
        ["Technical", "Real-time visualisation and data processing with web technologies"],
        ["Programming", "Proficiency in modular JavaScript and the Canvas API"],
        ["Domain", "Understanding of flight testing and telemetry data processing"],
        ["Professional", "Communication, documentation, ethics and discipline"],
    ], widths=[2400, 6960])

# ======================================================================
# CHAPTER 17 — CHALLENGES
# ======================================================================
def chapter17():
    h1("CHAPTER 17   CHALLENGES")
    h2("17.1  Introduction")
    p("Every engineering effort encounters difficulties, and overcoming them is itself an important "
      "part of learning. This chapter describes the principal challenges faced during the "
      "internship and how each was addressed.", indent=False)
    h2("17.2  Understanding a New Domain")
    p("The aerospace domain was entirely new to me at the start of the internship. Grasping the "
      "concepts of flight testing, instrumentation and telemetry required dedicated study and the "
      "guidance of my mentor. Patient reading and discussion gradually built the understanding "
      "needed to undertake the project meaningfully.")
    h2("17.3  Data Visualisation with Canvas")
    p("Visualising data on the Canvas was initially challenging because the Canvas API is a "
      "low-level drawing interface that requires every element to be drawn explicitly. Producing "
      "smooth animation, correctly oriented symbols and a rotating radar sweep took careful work "
      "and experimentation before the desired results were achieved.")
    h2("17.4  JavaScript Issues")
    p("Several issues arose in the JavaScript logic, including the correct handling of timers, the "
      "co-ordination of multiple rendering functions and the avoidance of subtle errors in the "
      "processing calculations. These were resolved through systematic debugging using the "
      "browser's developer tools.")
    h2("17.5  User-Interface Design")
    p("Designing an interface that was both professional and easy to read demanded several "
      "iterations. Choosing an appropriate colour scheme, arranging the information clearly and "
      "ensuring that the layout adapted to different screen sizes all required thought and "
      "refinement.")
    h2("17.6  Browser Compatibility")
    p("Ensuring that the application behaved consistently across different browsers required "
      "testing in more than one browser and adjusting the code where behaviour differed. Adhering "
      "to standard web technologies minimised these differences.")
    h2("17.7  Testing and Debugging")
    p("Testing a continuously animating, real-time application is more demanding than testing a "
      "static one, because behaviour unfolds over time. A disciplined approach to testing, "
      "examining one behaviour at a time, was adopted to manage this complexity.")
    h2("17.8  Learning New Technologies")
    p("Finally, the internship required me to learn and apply new techniques quickly. Meeting this "
      "challenge improved my ability to learn independently \u2014 a skill that will serve me "
      "throughout my career.")

# ======================================================================
# CHAPTER 18 — FUTURE SCOPE
# ======================================================================
def chapter18():
    h1("CHAPTER 18   FUTURE SCOPE")
    h2("18.1  Introduction")
    p("The system developed during the internship is a demonstration that establishes a sound "
      "foundation. This chapter outlines the directions in which it could be extended and enhanced "
      "in the future.", indent=False)
    h2("18.2  Integration with Real Telemetry")
    p("The most natural extension is to connect the system to a real telemetry source in place of "
      "the simulation, so that it processes and displays genuine flight data. This would involve "
      "adding a data-ingestion layer that decodes the incoming telemetry stream.")
    h2("18.3  Artificial Intelligence and Machine Learning")
    p("Artificial intelligence and machine learning could be applied to the processed data to "
      "detect anomalies automatically, to recognise patterns in aircraft behaviour and to classify "
      "events without explicit thresholds. Such techniques would make the monitoring more "
      "intelligent and proactive.")
    h2("18.4  Predictive Analytics")
    p("Predictive analytics could forecast the future evolution of parameters and warn of "
      "conditions before they occur, transforming the system from a reactive monitor into a "
      "predictive one and further enhancing safety.")
    h2("18.5  Cloud Integration and Big Data")
    p("Storing processed data in the cloud would enable long-term retention, large-scale analysis "
      "and access from multiple locations. Big-data techniques could then be applied across many "
      "flights to extract trends and insights.")
    h2("18.6  GPS and Geographic Mapping")
    p("Replacing the simulated grid with real geographic mapping driven by satellite positioning "
      "would allow aircraft to be shown on an actual map, improving realism and operational value.")
    h2("18.7  Internet of Things and Digital Twin")
    p("Integration with networked sensors and the construction of a digital twin of the aircraft "
      "would allow the live data to be related to a complete virtual model, supporting deeper "
      "analysis and prediction.")
    h2("18.8  Mobile Application")
    p("A mobile version of the system would allow authorised personnel to monitor the situation "
      "from portable devices, extending the reach of the tool.")
    h2("18.9  Drone and Unmanned-Aircraft Monitoring")
    p("The same concepts apply directly to the monitoring of unmanned aircraft, an area of rapidly "
      "growing importance, making this a promising direction for extension.")
    h2("18.10  Cyber Security")
    p("As the system is connected to real data sources and networks, robust cyber-security "
      "measures \u2014 secure authentication, encryption and access control \u2014 would become "
      "essential and form an important area of future work.")
    tabcap("Summary of future enhancements")
    table([
        ["Enhancement", "Expected Benefit"],
        ["Real telemetry integration", "Processing of genuine flight data"],
        ["AI / machine learning", "Automatic anomaly detection"],
        ["Predictive analytics", "Early, proactive warnings"],
        ["Cloud and big data", "Long-term storage and large-scale analysis"],
        ["GPS mapping", "Realistic geographic display"],
        ["Mobile application", "Monitoring on portable devices"],
        ["Cyber security", "Protection of data and access"],
    ], widths=[3200, 6160])

# ======================================================================
# CHAPTER 19 — CONCLUSION
# ======================================================================
def chapter19():
    h1("CHAPTER 19   CONCLUSION")
    p("The industrial internship at the Flight Testing Centre of the Rotary Wing Research & Design "
      "Centre, Hindustan Aeronautics Limited, has been an enriching and formative experience. It "
      "provided a rare opportunity to work within a premier national aerospace organisation and to "
      "understand, at first hand, the disciplined and safety-conscious environment in which "
      "advanced aircraft are developed and tested.", indent=False)
    p("The internship achieved its stated objectives. The structure and functions of HAL and of "
      "the design centre were studied; the principles of flight testing, instrumentation, "
      "telemetry and data processing were understood; and this understanding was consolidated "
      "through the design and development of the Advanced Live Aircraft Tracking and Flight Data "
      "Processing System. The project successfully demonstrates, on a safe and reduced scale, how "
      "live telemetry can be processed and visualised to support monitoring and safety at a flight "
      "testing centre.")
    p("Through this work I have grown both technically and professionally. I have strengthened my "
      "skills in web development, real-time visualisation and software engineering, and I have "
      "gained valuable domain knowledge of the aerospace field. Equally important, I have absorbed "
      "the professional values of accuracy, discipline, documentation and safety that characterise "
      "the organisation.")
    p("The system also opens up numerous avenues for future enhancement, from integration with real "
      "telemetry to the application of artificial intelligence and predictive analytics. These "
      "possibilities make the project not only a demonstration of what has been learned but also a "
      "foundation for further work.")
    p("In conclusion, the internship has been a significant milestone in my education. It has "
      "bridged the gap between academic study and professional practice, deepened my interest in "
      "the field, and prepared me to contribute as a competent and responsible software engineer. "
      "I am sincerely grateful to all who made this experience possible.")

# ======================================================================
# CHAPTER 20 — BIBLIOGRAPHY
# ======================================================================
def chapter20():
    h1("CHAPTER 20   BIBLIOGRAPHY")
    p("The following references, presented in IEEE citation style, were consulted during the "
      "internship and the preparation of this report.", indent=False)
    refs = [
        "Hindustan Aeronautics Limited, \u201cAbout HAL \u2014 Company Profile and Products,\u201d "
        "Official Corporate Publications, Bengaluru, India.",
        "Rotary Wing Research & Design Centre (RWR&DC), HAL, \u201cHelicopter Design and "
        "Development Overview,\u201d Technical Documentation.",
        "R. C. Nelson, Flight Stability and Automatic Control, 2nd ed. New York, NY, USA: "
        "McGraw-Hill, 1998.",
        "K. P. Rajan, Flight Test Instrumentation and Data Processing, Aerospace Engineering "
        "Series, India.",
        "A. Cooke and E. Fitzpatrick, Helicopter Test and Evaluation. Oxford, U.K.: Blackwell "
        "Science, 2002.",
        "Mozilla Developer Network, \u201cHTML5, CSS3 and the Canvas API \u2014 Web Documentation,\u201d "
        "Mozilla Foundation. [Online]. Available: https://developer.mozilla.org",
        "World Wide Web Consortium (W3C), \u201cHTML and CSS Standards,\u201d W3C Recommendations. "
        "[Online]. Available: https://www.w3.org",
        "D. Flanagan, JavaScript: The Definitive Guide, 7th ed. Sebastopol, CA, USA: O\u2019Reilly "
        "Media, 2020.",
        "R. S. Pressman and B. R. Maxim, Software Engineering: A Practitioner\u2019s Approach, "
        "8th ed. New York, NY, USA: McGraw-Hill Education, 2015.",
        "ECMA International, \u201cECMAScript Language Specification (ECMA-262),\u201d Geneva, "
        "Switzerland.",
        "United States Naval Test Pilot School, Flight Test Manual \u2014 Fixed Wing Stability and "
        "Control, U.S. Naval Air Warfare Center.",
        "International Society of Automation, \u201cTelemetry and Data Acquisition Standards,\u201d "
        "Technical Standards.",
    ]
    for i, r in enumerate(refs, 1):
        add(para(run(f"[{i}]  ", bold=True, size=12, font="Times New Roman")
                 + run(r, size=12, font="Times New Roman"),
                 align="both", spacing_after=120, line=360, left=540, hanging=540))

# ======================================================================
# APPENDIX
# ======================================================================
def appendix():
    h1("APPENDIX")
    h2("A.  Project Files")
    p("The implementation of the project is provided in the accompanying project folder. The "
      "principal files are listed below.", indent=False)
    table([
        ["File", "Description"],
        ["hal-advanced-tracking-system.html", "The complete Advanced Live Aircraft Tracking System"],
        ["hal-live-tracking-prototype.html", "The initial single-screen tracking prototype"],
    ], widths=[4200, 5160])
    h2("B.  Abbreviations")
    table([
        ["Abbreviation", "Expansion"],
        ["HAL", "Hindustan Aeronautics Limited"],
        ["RWR&DC", "Rotary Wing Research & Design Centre"],
        ["FTC", "Flight Testing Centre"],
        ["ALH", "Advanced Light Helicopter"],
        ["LCH", "Light Combat Helicopter"],
        ["LUH", "Light Utility Helicopter"],
        ["LCA", "Light Combat Aircraft"],
        ["API", "Application Programming Interface"],
        ["DFD", "Data Flow Diagram"],
        ["UML", "Unified Modeling Language"],
        ["UI", "User Interface"],
    ], widths=[2600, 6760])
    h2("C.  Note on Personalisation")
    p("Before submission, the placeholders for the college name, student name, USN, guide name, "
      "mentor name and principal name (marked with square brackets on the cover and certificate "
      "pages) should be replaced with the actual details, and the scanned internship certificate "
      "and the application screenshots should be inserted in the spaces provided.", indent=False)

print("Chapters 15-20 and appendix loaded.")



# ======================================================================
# STYLES, HEADER, FOOTER, SETTINGS, PACKAGING
# ======================================================================
def styles_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:docDefaults><w:rPrDefault><w:rPr>'
        '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        '<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr>'
        '<w:spacing w:after="120" w:line="360" w:lineRule="auto"/>'
        '<w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>'
        # Normal
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/><w:rPr>'
        '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        '<w:sz w:val="24"/></w:rPr></w:style>'
        # Heading 1
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
        '<w:basedOn w:val="Normal"/><w:pPr><w:keepNext/>'
        '<w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>'
        '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
        # Heading 2
        '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/>'
        '<w:basedOn w:val="Normal"/><w:pPr><w:keepNext/>'
        '<w:spacing w:before="160" w:after="120"/><w:outlineLvl w:val="1"/></w:pPr>'
        '<w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>'
        # Table grid
        '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>'
        '<w:tblPr><w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:color="999999"/>'
        '<w:left w:val="single" w:sz="4" w:color="999999"/>'
        '<w:bottom w:val="single" w:sz="4" w:color="999999"/>'
        '<w:right w:val="single" w:sz="4" w:color="999999"/>'
        '<w:insideH w:val="single" w:sz="4" w:color="999999"/>'
        '<w:insideV w:val="single" w:sz="4" w:color="999999"/>'
        '</w:tblBorders></w:tblPr></w:style>'
        '</w:styles>')

def header_default_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:p><w:pPr>'
        '<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="999999"/></w:pBdr>'
        '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
        '<w:jc w:val="center"/></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        '<w:i/><w:sz w:val="20"/><w:color w:val="555555"/></w:rPr>'
        '<w:t>Industrial Internship Report</w:t></w:r></w:p></w:hdr>')

def footer_default_xml():
    # Line 1: student info ; Line 2: page number (centred)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:p><w:pPr>'
        '<w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="999999"/></w:pBdr>'
        '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
        '<w:jc w:val="center"/></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        '<w:sz w:val="18"/><w:color w:val="555555"/></w:rPr>'
        '<w:t xml:space="preserve">' + esc(STUDENT) + ' | BCA Final Year | HAL Internship</w:t></w:r></w:p>'
        '<w:p><w:pPr>'
        '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
        '<w:jc w:val="center"/></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        '<w:sz w:val="20"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        '<w:sz w:val="20"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
        '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
        '<w:sz w:val="20"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>')

def header_first_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p></w:hdr>')

def footer_first_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p></w:ftr>')

def settings_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:updateFields w:val="true"/>'
        '<w:defaultTabStop w:val="720"/>'
        '<w:compat><w:compatSetting w:name="compatibilityMode" '
        'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
        '</w:settings>')

def sect_pr():
    # A4 portrait; margins T1" B1" L1.5" R1"; header/footer refs; first-page different
    return ('<w:sectPr>'
        '<w:headerReference w:type="default" r:id="rId2"/>'
        '<w:headerReference w:type="first" r:id="rId4"/>'
        '<w:footerReference w:type="default" r:id="rId3"/>'
        '<w:footerReference w:type="first" r:id="rId5"/>'
        '<w:titlePg/>'
        f'<w:pgSz w:w="11906" w:h="16838"/>'
        f'<w:pgMar w:top="{IN(1)}" w:right="{IN(1)}" w:bottom="{IN(1)}" w:left="{IN(1.5)}" '
        f'w:header="{IN(0.5)}" w:footer="{IN(0.5)}" w:gutter="0"/>'
        '<w:pgNumType w:fmt="decimal"/>'
        '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>')

def document_xml():
    body = "".join(BODY) + sect_pr()
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<w:body>{body}</w:body></w:document>')

def content_types_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/settings.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
        '<Override PartName="/word/header1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        '<Override PartName="/word/header2.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        '<Override PartName="/word/footer1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
        '<Override PartName="/word/footer2.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
        '</Types>')

def root_rels_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/></Relationships>')

def document_rels_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" '
        'Target="header1.xml"/>'
        '<Relationship Id="rId3" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" '
        'Target="footer1.xml"/>'
        '<Relationship Id="rId4" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" '
        'Target="header2.xml"/>'
        '<Relationship Id="rId5" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" '
        'Target="footer2.xml"/>'
        '<Relationship Id="rId6" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" '
        'Target="settings.xml"/>'
        '</Relationships>')

# ======================================================================
# ASSEMBLE THE DOCUMENT
# ======================================================================
def build():
    # --- Front matter ---
    cover_page()
    certificate_page()
    hal_certificate_page()
    acknowledgement_page()
    abstract_page()
    toc_page()
    # --- Chapters ---
    for fn in [chapter1, chapter2, chapter3, chapter4, chapter5, chapter6, chapter7,
               chapter8, chapter9, chapter10, chapter11, chapter12, chapter13, chapter14,
               chapter15, chapter16, chapter17, chapter18, chapter19, chapter20, appendix]:
        fn()

def package(out_path):
    parts = {
        "[Content_Types].xml": content_types_xml(),
        "_rels/.rels": root_rels_xml(),
        "word/document.xml": document_xml(),
        "word/_rels/document.xml.rels": document_rels_xml(),
        "word/styles.xml": styles_xml(),
        "word/settings.xml": settings_xml(),
        "word/header1.xml": header_default_xml(),
        "word/header2.xml": header_first_xml(),
        "word/footer1.xml": footer_default_xml(),
        "word/footer2.xml": footer_first_xml(),
    }
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)
    return out_path

if __name__ == "__main__":
    build()
    out = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                       "report", "HAL_Internship_Report.docx")
    package(out)
    # rough stats
    words = 0
    import re
    for b in BODY:
        for t in re.findall(r"<w:t[^>]*>(.*?)</w:t>", b):
            words += len(t.split())
    print(f"Document written to: {out}")
    print(f"Body paragraphs/elements: {len(BODY)}")
    print(f"Approximate word count : {words}")
    print(f"Figures: {_fig[0]}  Tables: {_tab[0]}")
