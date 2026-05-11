from pathlib import Path
from datetime import datetime
import re

from docx import Document
from docx.shared import RGBColor, Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter


BLUE = "0000FF"
HEADER_BLUE = "0070C0"
LIGHT_BLUE = "D9EAF7"
LIGHT_GRAY = "D9D9D9"
YELLOW = "FFFF00"
RED = "FF0000"
WARNING_YELLOW = "FFFF00"
OK_GREEN = "92D050"
WHITE = "FFFFFF"
BLACK = "000000"


# =========================================================
# Basic Word helpers
# =========================================================

def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_text(cell, text, bold=False, color=BLACK, size=8, align="left"):
    cell.text = ""
    p = cell.paragraphs[0]

    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "right":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    run = p.add_run(str(text))
    run.font.name = "Arial"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)

    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def set_table_borders(table, color="000000", size="4"):
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")

    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)

    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = "w:" + edge
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)

        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), color)


def add_monospaced_box(doc, text, font_size=7):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    set_table_borders(table, color="000000", size="6")

    cell = table.cell(0, 0)
    cell.text = ""
    for index, line in enumerate(str(text).splitlines()):
        paragraph = cell.paragraphs[0] if index == 0 else cell.add_paragraph()
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        run = paragraph.add_run(line)
        run.font.name = "Courier New"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Courier New")
        run.font.size = Pt(font_size)
        run.font.color.rgb = RGBColor.from_string(BLACK)

    return table


def set_doc_layout(doc):
    section = doc.sections[0]
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.55)
    section.left_margin = Inches(0.55)
    section.right_margin = Inches(0.55)

    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(8)


def add_heading1(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = "Arial"
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.color.rgb = RGBColor.from_string(BLUE)

    p_format = p.paragraph_format
    p_format.space_before = Pt(6)
    p_format.space_after = Pt(2)

    # blue underline
    border_p = p._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), BLUE)
    p_bdr.append(bottom)
    border_p.append(p_bdr)

    return p


def add_heading2(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = "Arial"
    run.font.size = Pt(10)
    run.font.bold = True
    run.font.italic = True
    run.font.color.rgb = RGBColor.from_string(BLUE)

    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(2)

    return p


def add_heading3(doc, text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = "Arial"
    run.font.size = Pt(9)
    run.font.bold = True
    run.font.color.rgb = RGBColor.from_string(BLACK)

    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)

    return p


def add_normal(doc, text, size=8, bold=False):
    p = doc.add_paragraph()
    run = p.add_run(str(text))
    run.font.name = "Arial"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(BLACK)
    return p


def add_small_table(doc, headers, rows, header_fill=LIGHT_GRAY, highlight_rule=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.style = "Table Grid"
    set_table_borders(table)

    for i, h in enumerate(headers):
        set_cell_shading(table.rows[0].cells[i], header_fill)
        set_cell_text(table.rows[0].cells[i], h, bold=True, size=8)

    for row in rows:
        cells = table.add_row().cells
        for i, h in enumerate(headers):
            value = row.get(h, "")
            fill = WHITE

            if highlight_rule and highlight_rule(row, h, value):
                fill = YELLOW

            set_cell_shading(cells[i], fill)
            set_cell_text(cells[i], value, size=8)

    return table


# =========================================================
# 0. Manual TOC
# =========================================================

def add_table_of_contents(doc):
    """Generate a professional TOC with proper tab stops, dot leaders, and color-coded levels."""

    # --- Title ---
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(18)
    run = p.add_run("Table of Contents")
    run.font.name = "Arial"
    run.font.size = Pt(18)
    run.font.bold = True
    run.font.color.rgb = RGBColor.from_string(BLUE)

    # Define TOC entries:  (title, page_number, level)
    # level 0 = main section,  level 1 = sub-section,  level 2 = sub-sub-section
    toc_rows = [
        ("Defect Classification", 4, 0),
        ("Suggestion Summary", 5, 0),
        ("1. General Project Information", 6, 0),
        ("1.1 Project Information", 6, 1),
        ("1.2 Customer Contact Information", 6, 1),
        ("1.3 MFEC Engineers' Information", 6, 1),
        ("1.4 Change Record", 6, 1),
        ("1.5 Reviewers", 6, 1),
        ("2. Oracle minimum requirement", 7, 0),
        ("3. System Checklist", 10, 0),
        ("4. Database Information", 11, 0),
        ("4.1 Database Configuration.", 11, 1),
        ("4.2 Database Parameter", 11, 1),
        ("4.3 Major Security Initialization Parameters", 13, 1),
        ("4.4 Database Files", 13, 1),
        ("4.5 Temporary Files", 14, 1),
        ("4.6 Redo Log File", 14, 1),
        ("4.7 Control File", 14, 1),
        ("4.8 Database Patch", 15, 1),
        ("5. RDBMS Performance", 16, 0),
        ("5.1 Performance Review", 16, 1),
        ("5.2 Database Growth Rate", 18, 1),
        ("5.3 Performance Analysis", 19, 1),
        ("6. Tablespace Free Space", 23, 0),
        ("7. Default tablespace and temporary tablespace", 24, 0),
        ("8. Database Registry", 26, 0),
        ("APPENDIX A – Invalid Object", 27, 0),
        ("APPENDIX B – Information from alert log", 36, 0),
        ("APPENDIX C – SQL Statement should be to investigate", 37, 0),
        ("APPENDIX D – Operating system log", 38, 0),
        ("APPENDIX E – Backup Configuration", 39, 0),
        ("APPENDIX F – OS Performance Summary", 40, 0),
    ]

    # Style settings per level
    level_config = {
        0: {"font_size": Pt(10), "bold": True,  "color": RGBColor(0x00, 0x00, 0xCC), "indent_cm": 0},
        1: {"font_size": Pt(9),  "bold": False, "color": RGBColor(0x00, 0x70, 0xC0), "indent_cm": 0.8},
        2: {"font_size": Pt(8),  "bold": False, "color": RGBColor(0x33, 0x33, 0x33), "indent_cm": 1.6},
    }

    right_tab_pos = 9000  # EMU-like position in twentieths of a point (~16 cm)

    for title, page, level in toc_rows:
        cfg = level_config.get(level, level_config[0])
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)

        # Set left indent
        indent_twips = int(cfg["indent_cm"] * 567)  # 1 cm = 567 twips
        p.paragraph_format.left_indent = Pt(cfg["indent_cm"] * 28.35)

        # Add a right-aligned tab stop with dot leader via XML
        pPr = p._p.get_or_add_pPr()
        tabs_elem = OxmlElement('w:tabs')
        tab_elem = OxmlElement('w:tab')
        tab_elem.set(qn('w:val'), 'right')
        tab_elem.set(qn('w:leader'), 'dot')
        tab_elem.set(qn('w:pos'), str(right_tab_pos))
        tabs_elem.append(tab_elem)
        pPr.append(tabs_elem)

        # Add the title text run
        r_title = p.add_run(title)
        r_title.font.name = "Arial"
        r_title.font.size = cfg["font_size"]
        r_title.font.bold = cfg["bold"]
        r_title.font.color.rgb = cfg["color"]

        # Add a tab character (this triggers the dot leader)
        r_tab = p.add_run("\t")
        r_tab.font.name = "Arial"
        r_tab.font.size = cfg["font_size"]

        # Add the page number
        r_page = p.add_run(str(page))
        r_page.font.name = "Arial"
        r_page.font.size = cfg["font_size"]
        r_page.font.bold = cfg["bold"]
        r_page.font.color.rgb = cfg["color"]

    doc.add_page_break()


# =========================================================
# Defect / Summary
# =========================================================

def add_defect_classification(doc):
    add_heading1(doc, "Defect Classification")
    add_normal(
        doc,
        "Each test will be either deemed as a pass or a severity level will be given using the following severity level classification.",
        size=8,
    )

    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    set_table_borders(table)

    headers = ["Severity\nLevel", "Definition", "Actions"]
    for i, h in enumerate(headers):
        set_cell_shading(table.rows[0].cells[i], LIGHT_GRAY)
        set_cell_text(table.rows[0].cells[i], h, bold=True, align="center")

    rows = [
        ("1", "Critical", "Chance of Database down, need immediate action", RED),
        ("2", "Warning", "Change of hindering some database function, doesn’t require immediate action", WARNING_YELLOW),
        ("3", "OK (Stable)", "Database can work normally. No action required.", OK_GREEN),
    ]

    for sev, definition, action, fill in rows:
        cells = table.add_row().cells
        for c in cells:
            set_cell_shading(c, fill)

        set_cell_text(cells[0], sev, align="center", bold=True)
        set_cell_text(cells[1], definition, align="center")
        set_cell_text(cells[2], action, align="left")

    doc.add_page_break()


def build_suggestion_summary(data):
    invalid_count = len(data.get("invalid_objects_list", []))
    patch_info = data.get("patch_info", [])
    tablespace_free = data.get("tablespace_free", {})
    sec53 = data.get("section53", {})
    os_perf = data.get("os_perf", {})
    db_params = data.get("db_parameters", {})

    # --- Patch status ---
    patch_status = "OK"
    if patch_info:
        for p in patch_info:
            if "recommend" in str(p.get("suggestion", "")).lower():
                patch_status = "Critical"

    # --- Tablespace status ---
    tbs_status = "OK"
    for _, info in tablespace_free.items():
        try:
            pct_free = float(info.get("pct_free", 0))
        except Exception:
            pct_free = 100
        if pct_free < 10:
            tbs_status = "Critical"
            break
        elif pct_free < 20:
            tbs_status = "Warning"

    # --- Invalid objects ---
    invalid_status = "Warning" if invalid_count > 0 else "OK"

    # --- 4.2/4.3 audit_trail check ---
    audit_trail = str(db_params.get("audit_trail", "")).upper()
    param_status = "Warning" if "EXTENDED" in audit_trail or "DB" in audit_trail else "OK"

    # --- 4.4 Datafile inc_mb check ---
    df_status = "OK"
    for f in data.get("datafiles", []):
        try:
            inc = float(f.get("inc_mb", "0"))
            if 0 < inc < 5:
                df_status = "Warning"
                break
        except (ValueError, TypeError):
            pass

    # --- 5.1 Archive mode ---
    archive_mode = data.get("archive_mode", "")
    review_status = "OK" if archive_mode == "ARCHIVELOG" else "Warning"

    # --- 5.3 CPU Section (Parse CPU to Parse Elapsd) ---
    cpu_status = "OK"
    inst_eff = sec53.get("instance_efficiency", {})
    try:
        parse_val = float(inst_eff.get("Parse CPU to Parse Elapsd %", "100"))
        if parse_val < 50.0:
            cpu_status = "Warning"
    except ValueError:
        pass

    # --- 5.3 SQL Section (open cursors / dblink) ---
    sql_status = "OK"
    cursors = sec53.get("open_cursors", {})
    try:
        pct_cursor = float(str(cursors.get("pct", "0")).replace("%", ""))
        if pct_cursor >= 95:
            sql_status = "Warning"
    except ValueError:
        pass
    top5 = sec53.get("top5_events", [])
    if any("dblink" in ev.lower() for ev in top5):
        sql_status = "Warning"

    # --- 5.3 Memory Section ---
    mem_status = "OK"

    # --- APPENDIX F OS Performance ---
    os_status = "OK"
    if os_perf.get("found"):
        cpu_trend = os_perf.get("cpu_trend", [])
        if cpu_trend:
            max_cpu = max([r.get("usr", 0) + r.get("sys", 0) for r in cpu_trend], default=0)
            if max_cpu >= 95:
                os_status = "Warning"
        runq = os_perf.get("runq_trend", [])
        if runq:
            max_runq = max((r.get("max", 0) for r in runq), default=0)
            try:
                cpu_count = int(data.get("os_info", {}).get("cpu_cores", "2"))
            except ValueError:
                cpu_count = 2
            if max_runq > cpu_count * 2:
                os_status = "Warning"

    return [
        ("2. Oracle Minimum Requirement", "OK", ""),
        ("3. System Checklist", "OK", ""),
        ("4. Database Information", "", ""),
        ("   4.1 Database Configuration", "OK", ""),
        ("   4.2 Database Parameter", param_status, "See more detail in 4.3" if param_status != "OK" else ""),
        ("   4.3 Major Security Initialization Parameter", param_status, "See more detail in 4.3" if param_status != "OK" else ""),
        ("   4.4 Database Files", df_status, "See more detail in 4.4" if df_status != "OK" else ""),
        ("   4.5 Temporary Files", "OK", ""),
        ("   4.6 Redo Log File", "OK", ""),
        ("   4.7 Control File", "OK", ""),
        ("   4.8 Database Patch", patch_status, "See more detail in 4.8" if patch_status != "OK" else ""),
        ("5. RDBMS Performance", "", ""),
        ("   5.1 Performance Review", review_status, "See more detail in 5.1" if review_status != "OK" else ""),
        ("   5.2 Database Growth Rate", "OK", ""),
        ("   5.3 Performance Analysis (CPU Section)", cpu_status, "See more detail in 5.3 (CPU Section)" if cpu_status != "OK" else ""),
        ("   5.3 Performance Analysis (SQL Section)", sql_status, "See more detail in 5.3 (SQL Section)" if sql_status != "OK" else ""),
        ("   5.3 Performance Analysis (Memory Section)", mem_status, ""),
        ("6. Tablespace Free Space", tbs_status, "See more detail in 6.1" if tbs_status != "OK" else ""),
        ("7. Default tablespace and temporary tablespace", "OK", ""),
        ("8. Database Registry", "OK", ""),
        ("APPENDIX A. Invalid Object", invalid_status, "See more detail in APPENDIX A" if invalid_status != "OK" else ""),
        ("APPENDIX B. Information from alert log", "OK", ""),
        ("APPENDIX C. SQL Statement should be to investigate", "OK", ""),
        ("APPENDIX D. Operating System Log", "OK", ""),
        ("APPENDIX E. Backup Configuration", "OK", ""),
        ("APPENDIX F. OS Performance Summary", os_status, "See more detail in APPENDIX F" if os_status != "OK" else ""),
    ]



def add_suggestion_summary(doc, data):
    add_heading1(doc, "Suggestion Summary")

    rows = build_suggestion_summary(data)

    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    set_table_borders(table)

    headers = ["Topic", "Critical", "Warning", "OK", "Remark"]
    for i, h in enumerate(headers):
        set_cell_shading(table.rows[0].cells[i], HEADER_BLUE)
        set_cell_text(table.rows[0].cells[i], h, bold=True, color=WHITE, align="center")

    for topic, status, remark in rows:
        cells = table.add_row().cells

        for c in cells:
            set_cell_shading(c, WHITE)

        set_cell_text(cells[0], topic, size=7)
        set_cell_text(cells[1], "x" if status == "Critical" else "", align="center", size=7)
        set_cell_text(cells[2], "x" if status == "Warning" else "", align="center", size=7)
        set_cell_text(cells[3], "x" if status == "OK" else "", align="center", size=7)
        set_cell_text(cells[4], remark, size=7)

    doc.add_page_break()


# =========================================================
# Project Information
# =========================================================

def add_general_project_information(doc, data, source_zip):
    add_heading1(doc, "1. General Project Information")

    add_heading2(doc, "1.1 Project Information")

    gui_meta = data.get("gui_metadata", {})
    db_name = gui_meta.get("db_name") or data.get("instance_name") or data.get("db_parameters", {}).get("db_name", "")
    
    project_rows = [
        {"Field": "Project name", "Value": gui_meta.get("project_name", "MA Oracle 5 Years")},
        {"Field": "MFEC project no.", "Value": gui_meta.get("project_no", "")},
        {"Field": "MFEC sales name", "Value": gui_meta.get("sale", "")},
        {"Field": "Phone no.", "Value": ""}, # Can be derived if sale has contact
        {"Field": "Email address", "Value": ""},
        {"Field": "Database name", "Value": db_name},
        {"Field": "Source ZIP", "Value": source_zip.name},
    ]
    add_small_table(doc, ["Field", "Value"], project_rows)

    add_heading2(doc, "1.2 Customer Contact Information")
    cust_list = gui_meta.get("customers", [])
    if not cust_list:
        cust_list = [{"name": "", "phone": "", "email": ""}]
        
    cust_rows = [{"Name": c.get("name",""), "Phone no.": c.get("phone",""), "Email address": c.get("email","")} for c in cust_list]
    add_small_table(doc, ["Name", "Phone no.", "Email address"], cust_rows)

    add_heading2(doc, "1.3 MFEC Engineers' Information")
    eng_list = gui_meta.get("engineers", [])
    if not eng_list:
        eng_list = [{"name": "", "phone": "", "email": ""}]
        
    eng_rows = [{"Name": e.get("name",""), "Phone no.": e.get("phone",""), "Email address": e.get("email","")} for e in eng_list]
    add_small_table(doc, ["Name", "Phone no.", "Email address"], eng_rows)

    add_heading2(doc, "1.4 Change Record")
    cr_list = gui_meta.get("change_records", [])
    if not cr_list:
        cr_list = [{"date": datetime.now().strftime("%d-%b-%Y"), "author": "", "version": "1.0", "ref": "Initial Document"}]
        
    cr_rows = [{"Date": r.get("date",""), "Author": r.get("author",""), "Version": r.get("version",""), "Change Reference": r.get("ref","")} for r in cr_list]
    add_small_table(doc, ["Date", "Author", "Version", "Change Reference"], cr_rows)

    add_heading2(doc, "1.5 Reviewers")
    rev_list = gui_meta.get("reviewers", [])
    if not rev_list:
        rev_list = [{"date": "", "name": "", "position": ""}]
        
    rev_rows = [{"Date": r.get("date",""), "Name": r.get("name",""), "Position": r.get("position","")} for r in rev_list]
    add_small_table(doc, ["Date", "Name", "Position"], rev_rows)

    doc.add_page_break()


# =========================================================
# Section 2: Oracle Minimum Requirement
# =========================================================

def add_oracle_minimum_requirement(doc, data):
    os_info = data.get("os_info", {})

    add_heading1(doc, "2. Oracle minimum requirement")

    # ----- 2.1 Checking server machine specification -----
    add_heading2(doc, "2.1 Checking server machine specification")
    add_normal(doc, "Server machine specification collected from the operating system:")

    spec_rows = [
        {"Item": "Hostname", "Detail": os_info.get("hostname", "")},
        {"Item": "IP Address", "Detail": os_info.get("ip_address", "")},
        {"Item": "Subnet Mask", "Detail": os_info.get("subnet_mask", "")},
        {"Item": "OS Version", "Detail": os_info.get("os_version", "")},
        {"Item": "Kernel Version", "Detail": os_info.get("kernel_version", "")},
        {"Item": "Architecture", "Detail": os_info.get("architecture", "")},
        {"Item": "CPU Model", "Detail": os_info.get("cpu_model", "")},
        {"Item": "Number of CPU(s)", "Detail": os_info.get("cpu_count", "")},
        {"Item": "Core(s) per socket", "Detail": os_info.get("cores_per_socket", "")},
        {"Item": "Socket(s)", "Detail": os_info.get("sockets", "")},
        {"Item": "CPU MHz", "Detail": os_info.get("cpu_mhz", "")},
        {"Item": "CPU Cache Size", "Detail": os_info.get("cache_size", "")},
        {"Item": "Total RAM (GB)", "Detail": os_info.get("total_ram_gb", "")},
        {"Item": "Total Swap (MB)", "Detail": os_info.get("swap_total_mb", "")},
        {"Item": "Swap Used (MB)", "Detail": os_info.get("swap_used_mb", "")},
        {"Item": "/tmp Total Size", "Detail": os_info.get("tmp_size", "")},
        {"Item": "/tmp Available", "Detail": os_info.get("tmp_avail", "")},
        {"Item": "JDK/JRE Version", "Detail": os_info.get("jdk_version", "")},
    ]
    add_small_table(doc, ["Item", "Detail"], spec_rows)

    # ----- 2.2 Compare to Oracle requirement -----
    add_heading2(doc, "2.2 Compare to Oracle requirement")
    add_normal(doc, "Compare current server specification with Oracle Database 19c minimum requirements:")

    # Calculate comparison values
    try:
        ram_gb = float(os_info.get("total_ram_gb", "0") or "0")
    except ValueError:
        ram_gb = 0.0

    try:
        swap_mb = float(os_info.get("swap_total_mb", "0") or "0")
        swap_gb = swap_mb / 1024
    except ValueError:
        swap_gb = 0.0

    # Oracle 19c swap requirement based on RAM
    if ram_gb <= 2:
        swap_req = "1.5x RAM"
        swap_req_gb = ram_gb * 1.5
    elif ram_gb <= 16:
        swap_req = "Equal to RAM"
        swap_req_gb = ram_gb
    else:
        swap_req = "16 GB"
        swap_req_gb = 16.0

    tmp_avail_str = os_info.get("tmp_avail", "0")
    try:
        if "G" in tmp_avail_str.upper():
            tmp_avail_gb = float(tmp_avail_str.upper().replace("G", "").strip())
        elif "M" in tmp_avail_str.upper():
            tmp_avail_gb = float(tmp_avail_str.upper().replace("M", "").strip()) / 1024
        else:
            tmp_avail_gb = float(tmp_avail_str)
    except (ValueError, AttributeError):
        tmp_avail_gb = 0.0

    arch = os_info.get("architecture", "")
    arch_ok = "x86_64" in arch or "aarch64" in arch

    req_rows = [
        {
            "Requirement": "RAM (minimum)",
            "Oracle 19c Minimum": "1 GB",
            "Actual": f"{ram_gb:.1f} GB",
            "Status": "Pass" if ram_gb >= 1 else "Fail",
        },
        {
            "Requirement": "Swap Space",
            "Oracle 19c Minimum": f"{swap_req} ({swap_req_gb:.1f} GB)",
            "Actual": f"{swap_gb:.1f} GB",
            "Status": "Pass" if swap_gb >= swap_req_gb else "Fail",
        },
        {
            "Requirement": "/tmp space",
            "Oracle 19c Minimum": "1 GB",
            "Actual": f"{tmp_avail_gb:.1f} GB available",
            "Status": "Pass" if tmp_avail_gb >= 1 else "Fail",
        },
        {
            "Requirement": "Architecture",
            "Oracle 19c Minimum": "x86_64 / aarch64",
            "Actual": arch,
            "Status": "Pass" if arch_ok else "Fail",
        },
    ]

    def highlight_req(row, h, value):
        return row.get("Status") == "Fail"

    add_small_table(doc, ["Requirement", "Oracle 19c Minimum", "Actual", "Status"], req_rows,
                    header_fill=HEADER_BLUE, highlight_rule=highlight_req)

    # ----- 2.3 User's environment -----
    add_heading2(doc, "2.3 User's environment")
    add_normal(doc, "Oracle user environment configuration:")

    env_rows = [
        {"Item": "Oracle Owner", "Value": os_info.get("oracle_owner", "")},
        {"Item": "Oracle Home", "Value": os_info.get("oracle_home", "")},
        {"Item": "Oracle SID", "Value": os_info.get("oracle_sid", "")},
        {"Item": "Oracle Base", "Value": os_info.get("oracle_base", "")},
        {"Item": "Shell", "Value": os_info.get("shell", "")},
        {"Item": "Groups", "Value": os_info.get("oracle_groups", "")},
        {"Item": "Semaphores", "Value": os_info.get("semaphores", "")},
    ]
    add_small_table(doc, ["Item", "Value"], env_rows)

    # Kernel parameters sub-table
    kernel_params = os_info.get("kernel_params", {})
    if kernel_params:
        add_heading2(doc, "2.3.1 Kernel Parameters")
        kp_rows = [{"Parameter": k, "Value": v} for k, v in kernel_params.items()]
        add_small_table(doc, ["Parameter", "Value"], kp_rows)

    # Disk space sub-table
    disk_info = os_info.get("disk_info", [])
    if disk_info:
        add_heading2(doc, "2.3.2 Disk Space")
        disk_rows = []
        for d in disk_info:
            disk_rows.append({
                "Filesystem": d.get("filesystem", ""),
                "Size": d.get("size", ""),
                "Used": d.get("used", ""),
                "Avail": d.get("avail", ""),
                "Use%": d.get("use_pct", ""),
                "Mounted on": d.get("mounted_on", ""),
            })

        def highlight_disk(row, h, value):
            if h == "Use%":
                try:
                    pct = int(str(value).replace("%", ""))
                    return pct >= 85
                except (ValueError, TypeError):
                    pass
            return False

        add_small_table(doc, ["Filesystem", "Size", "Used", "Avail", "Use%", "Mounted on"],
                        disk_rows, highlight_rule=highlight_disk)

    # User limits sub-table
    user_limits = os_info.get("user_limits", {})
    if user_limits:
        add_heading2(doc, "2.3.3 User Limits (ulimit)")
        ul_rows = [{"Limit": k, "Value": v} for k, v in user_limits.items()]
        add_small_table(doc, ["Limit", "Value"], ul_rows)

    doc.add_page_break()


# =========================================================
# Section 3: System Checklist
# =========================================================

def add_system_checklist(doc, data):
    os_info = data.get("os_info", {})

    add_heading1(doc, "3. System Checklist")

    # ----- 3.1 Hardware configuration -----
    add_heading2(doc, "3.1 Hardware configuration")
    add_normal(doc, "Hardware and system configuration details:")

    hw_rows = [
        {"Item": "Machine Name (Hostname)", "Detail": os_info.get("hostname", "")},
        {"Item": "Platform / OS", "Detail": os_info.get("os_version", "")},
        {"Item": "Architecture", "Detail": os_info.get("architecture", "")},
        {"Item": "Kernel Version", "Detail": os_info.get("kernel_version", "")},
        {"Item": "CPU Type", "Detail": os_info.get("cpu_model", "")},
        {"Item": "Number of CPU(s)", "Detail": os_info.get("cpu_count", "")},
        {"Item": "Socket(s)", "Detail": os_info.get("sockets", "")},
        {"Item": "Core(s) per socket", "Detail": os_info.get("cores_per_socket", "")},
        {"Item": "CPU Clock Frequency (MHz)", "Detail": os_info.get("cpu_mhz", "")},
        {"Item": "CPU Cache Size", "Detail": os_info.get("cache_size", "")},
        {"Item": "Physical Memory (RAM)", "Detail": f"{os_info.get('total_ram_gb', '')} GB ({os_info.get('total_ram_mb', '')} MB)" if os_info.get("total_ram_gb") else ""},
        {"Item": "Swap Total", "Detail": f"{os_info.get('swap_total_mb', '')} MB" if os_info.get("swap_total_mb") else ""},
        {"Item": "Swap Used", "Detail": f"{os_info.get('swap_used_mb', '')} MB" if os_info.get("swap_used_mb") else ""},
    ]
    add_small_table(doc, ["Item", "Detail"], hw_rows)

    # ----- 3.2 Network configuration -----
    add_heading2(doc, "3.2 Network configuration")
    add_normal(doc, "Network interface and hosts configuration:")

    net_rows = [
        {"Item": "IP Address", "Detail": os_info.get("ip_address", "")},
        {"Item": "Subnet Mask", "Detail": os_info.get("subnet_mask", "")},
        {"Item": "Hostname", "Detail": os_info.get("hostname", "")},
    ]
    add_small_table(doc, ["Item", "Detail"], net_rows)

    # Hosts file
    hosts_entries = os_info.get("hosts_entries", [])
    if hosts_entries:
        add_heading2(doc, "3.2.1 Hosts File (/etc/hosts)")
        hosts_rows = [{"IP Address": h.get("ip", ""), "Hostname(s)": h.get("hostnames", "")} for h in hosts_entries]
        add_small_table(doc, ["IP Address", "Hostname(s)"], hosts_rows)

    # ----- 3.3 Crontab information -----
    add_heading2(doc, "3.3 Crontab information")

    crontab_entries = os_info.get("crontab_entries", [])
    if crontab_entries:
        add_normal(doc, f"Total crontab entries: {len(crontab_entries)} (Active: {sum(1 for c in crontab_entries if c.get('active'))}, Inactive: {sum(1 for c in crontab_entries if not c.get('active'))})")

        cron_rows = []
        for c in crontab_entries:
            cron_rows.append({
                "Schedule": c.get("schedule", ""),
                "Command": c.get("command", ""),
                "Status": "Active" if c.get("active") else "Inactive",
            })

        def highlight_cron(row, h, value):
            return row.get("Status") == "Inactive"

        add_small_table(doc, ["Schedule", "Command", "Status"], cron_rows, highlight_rule=highlight_cron)
    else:
        add_normal(doc, "No crontab entries found.")

    doc.add_page_break()


# =========================================================
# Database Information
# =========================================================

def add_database_information(doc, data):
    add_heading1(doc, "4. Database Information")

    add_heading2(doc, "4.1 Database Configuration.")

    db_name = data.get("instance_name") or data.get("db_parameters", {}).get("db_name", "")
    archive_enabled = "Yes" if data.get("archive_mode") == "ARCHIVELOG" else "No"

    rows = [
        {"What?": "Instance name", "Values?": db_name},
        {"What?": "RDBMS Version/Release?", "Values?": data.get("rdbms_version", "")},
        {"What?": "Number of datafiles?", "Values?": str(len(data.get("datafiles", [])))},
        {"What?": "Number of tempfiles?", "Values?": str(len(data.get("tempfiles", [])))},
        {"What?": "Is Archiving enabled?", "Values?": archive_enabled},
        {"What?": "Disk Space(DataFile + TempFile + RedoLogFile)", "Values?": data.get("disk_space", "")},
        {"What?": "Optimizer mode", "Values?": data.get("optimizer_mode", "")},
    ]

    add_small_table(doc, ["What?", "Values?"], rows)

    add_heading2(doc, "4.2 Database Parameter")

    param_rows = []
    for k, v in data.get("db_parameters", {}).items():
        param_rows.append({"Name": k, "Values": v})

    def highlight_param(row, h, value):
        return row.get("Name", "").lower() in ["audit_trail"] or row.get("Name", "").lower() == "db_file_multiblock_read_count"

    add_small_table(doc, ["Name", "Values"], param_rows, highlight_rule=highlight_param)

    add_heading2(doc, "4.3 Major Security Initialization Parameters")
    sec_params = []
    for k in ["O7_DICTIONARY_ACCESSIBILITY", "audit_trail", "remote_login_passwordfile", "remote_os_authent"]:
        sec_params.append({"Name": k, "Values": data.get("db_parameters", {}).get(k, "")})

    add_small_table(doc, ["Name", "Values"], sec_params, highlight_rule=highlight_param)

    add_normal(doc, "Additional Suggestion:", bold=True)
    add_normal(
        doc,
        "- audit_trail parameter should be reviewed. If audit data is stored in database tablespace, monitor usage and purge data according to retention policy.",
    )

    add_heading2(doc, "4.4 Database Files")
    df_rows = []
    for r in data.get("datafiles", []):
        df_rows.append({
            "Tbs Name": r.get("tbs_name", ""),
            "File Name": r.get("file_name", ""),
            "Size(MB)": r.get("size_mb", ""),
            "Max(MB)": r.get("max_mb", ""),
            "Aut": r.get("auto_ext", ""),
            "Inc.(MB)": r.get("inc_mb", ""),
        })

    def highlight_datafile(row, h, value):
        return row.get("Tbs Name", "") == "USERS" or str(row.get("Inc.(MB)", "")).strip() in ["1", "1.25"]

    add_small_table(doc, ["Tbs Name", "File Name", "Size(MB)", "Max(MB)", "Aut", "Inc.(MB)"], df_rows, highlight_rule=highlight_datafile)

    # Additional Suggestion for low increment datafiles
    low_inc_files = []
    for r in data.get("datafiles", []):
        try:
            inc = float(r.get("inc_mb", "0"))
            if 0 < inc < 5:
                low_inc_files.append((r.get("tbs_name", ""), inc))
        except (ValueError, TypeError):
            pass
    if low_inc_files:
        add_normal(doc, "Additional Suggestion:", bold=True)
        tbs_names = list(set(t[0] for t in low_inc_files))
        for tbs in tbs_names:
            inc_val = [v for t, v in low_inc_files if t == tbs][0]
            add_normal(doc, f"- Refer database file table, the data file of {tbs} tablespace is auto extend by increase {inc_val} MB, it's frequently auto extend size of data file. We recommend you configure auto extend by increase to 50 MB for reduce time to amount of extend size of data file.")

    add_heading2(doc, "4.5 Temporary Files")
    tmp_rows = []
    for r in data.get("tempfiles", []):
        tmp_rows.append({
            "Tbs Name": r.get("tbs_name", ""),
            "File Name": r.get("file_name", ""),
            "Size(MB)": r.get("size_mb", ""),
            "Max(MB)": r.get("max_mb", ""),
            "Aut": r.get("auto_ext", ""),
            "Inc.(MB)": r.get("inc_mb", ""),
        })
    add_small_table(doc, ["Tbs Name", "File Name", "Size(MB)", "Max(MB)", "Aut", "Inc.(MB)"], tmp_rows)

    add_heading2(doc, "4.6 Redo Log File")
    redo_rows = []
    for r in data.get("redo_logs", []):
        redo_rows.append({
            "Group#": r.get("group_no", ""),
            "Member": r.get("member", ""),
            "Size(MB)": r.get("size_mb", ""),
        })
    add_small_table(doc, ["Group#", "Member", "Size(MB)"], redo_rows)

    add_heading2(doc, "4.7 Control File")
    control_rows = [{"Control File Name#": r.get("file_name", "")} for r in data.get("control_files", [])]
    add_small_table(doc, ["Control File Name#"], control_rows)

    add_heading2(doc, "4.8 Database Patch")
    patch_rows = []
    for r in data.get("patch_info", []):
        patch_rows.append({
            "Component": r.get("component", ""),
            "Current Version": r.get("current_version", ""),
            "Recommended version": r.get("recommended_version", ""),
        })

    add_small_table(doc, ["Component", "Current Version", "Recommended version"], patch_rows, header_fill=LIGHT_GRAY)

    add_normal(doc, "Additional suggestion:", bold=True)
    add_normal(doc, "- Recommend applying the latest patch to fix bugs and update database security.")

    doc.add_page_break()


# =========================================================
# Growth graph
# =========================================================

def create_growth_chart(growth_rows, output_dir):
    if not growth_rows:
        return None

    months = [str(r.get("month", "")) for r in growth_rows]
    used = [float(r.get("used_gb", 0) or 0) for r in growth_rows]
    allocated = [float(r.get("allocated_gb", 0) or 0) for r in growth_rows]

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    chart_path = output_dir / "database_growth_rate.png"

    plt.figure(figsize=(6.2, 3.0), facecolor="#333333")
    ax = plt.gca()
    ax.set_facecolor("#333333")

    ax.plot(months, used, marker="o", label="Total use per month(GB)", linewidth=2, color="#5B9BD5")
    ax.plot(months, allocated, marker="o", label="Current Allocate (GB)", linewidth=2, color="#ED7D31")

    ax.set_title("Database Growth Rate", color="white", fontsize=11, weight="bold")
    ax.tick_params(colors="white", labelsize=8)
    ax.grid(True, color="#555555", linewidth=0.5)

    for spine in ax.spines.values():
        spine.set_color("#666666")

    legend = ax.legend(fontsize=6, loc="upper center", ncol=2, bbox_to_anchor=(0.5, 1.0))
    legend.get_frame().set_facecolor("#333333")
    legend.get_frame().set_edgecolor("#333333")
    for text in legend.get_texts():
        text.set_color("white")

    for x, y in zip(months, used):
        ax.text(x, y, f"{y:.2f}", color="white", fontsize=7, ha="center", va="bottom")

    for x, y in zip(months, allocated):
        ax.text(x, y, f"{y:.2f}", color="white", fontsize=7, ha="center", va="bottom")

    plt.tight_layout()
    plt.savefig(chart_path, dpi=180, facecolor="#333333")
    plt.close()

    return chart_path


def add_performance_review(doc, data):
    """Section 5.1 Performance Review table"""
    perf = data.get("perf_stats", {})
    db_params = data.get("db_parameters", {})
    oracle_stats = data.get("oracle_stats", {})

    add_heading2(doc, "5.1 Performance Review")
    add_normal(doc, "Performance statistics from Oracle database instance:")

    def _fmt_pct(val):
        """Format a ratio or percentage value nicely"""
        try:
            f = float(val)
            if f < 1 and f > 0:
                return f"{f * 100:.4f}%"
            return f"{f:.4f}%" if f < 10 else f"{f:.2f}%"
        except (ValueError, TypeError):
            return str(val)

    def _status_pct(val, good_above=90):
        """Determine OK/Warning/Critical based on percentage threshold"""
        try:
            f = float(val)
            pct = f * 100 if f < 1 else f
            if pct >= good_above:
                return "OK"
            elif pct >= good_above - 10:
                return "Warning"
            else:
                return "Critical"
        except (ValueError, TypeError):
            return ""

    # Build the performance review rows
    review_rows = []

    # Library Cache Hit Ratio
    lib_hit = perf.get("LIBRARY_CACHE_HITRATIO", "")
    if lib_hit:
        review_rows.append({
            "Metric": "Library cache hit ratio",
            "Value": _fmt_pct(lib_hit),
            "Threshold": "> 90%",
            "Status": _status_pct(lib_hit, 90),
            "Remark": "Should be > 90%, else increase SHARED_POOL_SIZE",
        })

    # PIN / RELOAD ratio
    pin_reload = perf.get("LIBRARY_CACHE_PIN_RELOAD", "")
    if pin_reload:
        review_rows.append({
            "Metric": "PIN / RELOAD ratio",
            "Value": _fmt_pct(pin_reload),
            "Threshold": "> 99%",
            "Status": _status_pct(pin_reload, 99),
            "Remark": "",
        })

    # Dictionary cache miss ratio
    dict_hit = perf.get("DICT_CACHE_HITRATIO", "")
    if dict_hit:
        review_rows.append({
            "Metric": "Dictionary cache miss ratio",
            "Value": _fmt_pct(dict_hit),
            "Threshold": "> 85%",
            "Status": _status_pct(dict_hit, 85),
            "Remark": "Should be > 85%, else increase SHARED_POOL_SIZE",
        })

    # Shared pool size
    shared_pool = db_params.get("shared_pool_size", "")
    review_rows.append({
        "Metric": "Shared pool size",
        "Value": shared_pool if shared_pool else "Auto (SGA_TARGET)",
        "Threshold": "",
        "Status": "OK",
        "Remark": "",
    })

    # Redo log space request
    redo_req = perf.get("REDO_LOG_SPACE_REQUEST", "")
    if redo_req:
        try:
            redo_val = int(float(redo_req))
            redo_status = "OK" if redo_val < 100 else "Warning" if redo_val < 1000 else "Critical"
        except ValueError:
            redo_status = ""
        review_rows.append({
            "Metric": "Redo log space request",
            "Value": redo_req,
            "Threshold": "< 100",
            "Status": redo_status,
            "Remark": "High value may indicate redo log too small" if redo_status != "OK" else "",
        })

    # DB block buffer cache hit ratio
    buf_hit = perf.get("BUFFER_CACHE_HIT_RATIO", "")
    if buf_hit:
        review_rows.append({
            "Metric": "DB block buffer cache hit ratio",
            "Value": _fmt_pct(buf_hit),
            "Threshold": "> 90%",
            "Status": _status_pct(buf_hit, 90),
            "Remark": "Should be > 90%, else increase DB_CACHE_SIZE",
        })

    # Latch hit ratio
    latch_hit = perf.get("LATCH_HIT_RATIO", "")
    if latch_hit:
        review_rows.append({
            "Metric": "Latch hit ratio",
            "Value": _fmt_pct(latch_hit),
            "Threshold": "> 99%",
            "Status": _status_pct(latch_hit, 99),
            "Remark": "",
        })

    # Disk sort ratio
    disk_sort = perf.get("DISK_SORT_RATIO", "")
    if disk_sort:
        try:
            ds_val = float(disk_sort)
            ds_pct = ds_val if ds_val > 1 else ds_val * 100
            ds_status = "OK" if ds_pct < 5 else "Warning"
        except ValueError:
            ds_status = ""
        review_rows.append({
            "Metric": "Disk sort ratio",
            "Value": _fmt_pct(disk_sort),
            "Threshold": "< 5%",
            "Status": ds_status,
            "Remark": "Should be < 5%, else increase SORT_AREA_SIZE" if ds_status != "OK" else "",
        })

    # Rollback segment waits
    rb_waits = perf.get("ROLLBACK_SEGMENT_WAITS", "")
    if rb_waits:
        try:
            rb_val = float(rb_waits)
            rb_pct = rb_val if rb_val > 1 else rb_val * 100
            rb_status = "OK" if rb_pct < 1 else "Warning"
        except ValueError:
            rb_status = ""
        review_rows.append({
            "Metric": "Rollback segment waits",
            "Value": _fmt_pct(rb_waits),
            "Threshold": "< 1%",
            "Status": rb_status,
            "Remark": "",
        })

    # Dispatcher workload
    disp = perf.get("DISPATCHER_WORKLOAD", "")
    if disp:
        try:
            disp_val = float(disp)
            disp_pct = disp_val if disp_val > 1 else disp_val * 100
            disp_status = "OK" if disp_pct < 50 else "Warning"
        except ValueError:
            disp_status = ""
        review_rows.append({
            "Metric": "Dispatcher workload",
            "Value": _fmt_pct(disp),
            "Threshold": "< 50%",
            "Status": disp_status,
            "Remark": "",
        })

    # PGA cache hit percentage
    pga_hit = perf.get("PGA_CACHE_HIT_PCT", oracle_stats.get("PGA_CACHE_HIT_PCT", ""))
    if pga_hit:
        review_rows.append({
            "Metric": "PGA cache hit percentage",
            "Value": f"{pga_hit}%",
            "Threshold": "> 70%",
            "Status": _status_pct(pga_hit, 70),
            "Remark": "",
        })

    # Undo tablespace
    undo_tbs = db_params.get("undo_tablespace", "")
    review_rows.append({
        "Metric": "Undo tablespace",
        "Value": undo_tbs,
        "Threshold": "",
        "Status": "OK" if undo_tbs else "",
        "Remark": "",
    })

    # Archive mode
    archive_mode = data.get("archive_mode", "")
    review_rows.append({
        "Metric": "Archive mode status",
        "Value": archive_mode,
        "Threshold": "",
        "Status": "OK" if archive_mode == "ARCHIVELOG" else "Warning",
        "Remark": "NOARCHIVELOG — data recovery limited" if archive_mode != "ARCHIVELOG" else "",
    })

    def highlight_perf(row, h, value):
        return row.get("Status") in ("Critical", "Warning")

    if review_rows:
        add_small_table(doc, ["Metric", "Value", "Threshold", "Status", "Remark"],
                        review_rows, header_fill=HEADER_BLUE, highlight_rule=highlight_perf)
    else:
        add_normal(doc, "No performance statistics found.")

    # Additional Suggestion for archive mode
    archive_mode = data.get("archive_mode", "")
    if archive_mode != "ARCHIVELOG":
        add_normal(doc, "Additional Suggestion:", bold=True)
        add_normal(doc, "- An archived redo log file is a copy of one of the filled members of a redo log group. This database is in NOARCHIVELOG indicates you disable the archiving of the redo log. If a media failure occurs while the database is in NOARCHIVELOG mode, you can only restore the database to the point of the most recent full database backup. You cannot recover transactions subsequent to that backup.")
        add_normal(doc, "- We recommend you run a database in ARCHIVELOG mode indicates you enable the archiving of the redo log. A database backup, together with online and archived redo log files, guarantees that you can recover all committed transactions in the event of an operating system or disk failure. When your database is in archivelog mode, you have to have more disk space to store archive log file and it will decrease your database performance because Archiver Processes (ARCn) copy redo log files to a storage device.")


def add_database_growth_rate(doc, data, output_dir):
    add_heading1(doc, "5. RDBMS Performance")
    add_performance_review(doc, data)
    add_heading2(doc, "5.2 Database Growth Rate")

    growth_rows = data.get("growth_rows_raw", [])
    db_name = data.get("instance_name") or data.get("db_parameters", {}).get("db_name", "")

    if not growth_rows:
        add_normal(doc, "No database growth data found.")
        return

    display_rows = growth_rows[-4:] if len(growth_rows) > 4 else growth_rows
    current = display_rows[-1]
    current_used = current.get("used_gb", "")
    current_alloc = current.get("allocated_gb", "")
    current_pct = current.get("pct_used", "")

    add_normal(doc, f"Current data information about space usage of database {db_name} as follows")

    add_small_table(doc, ["Name", "Value"], [
        {"Name": "Database Name", "Value": db_name},
        {"Name": "Current Allocated (GB)", "Value": current_alloc},
        {"Name": "Current Used data (GB)", "Value": current_used},
        {"Name": "Percent Used", "Value": current_pct},
    ], header_fill=HEADER_BLUE)

    chart_path = create_growth_chart(display_rows, output_dir)
    if chart_path:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(chart_path), width=Inches(6.2))

    add_normal(doc, "The result of total database usage size in this quarter as the below table.")

    table_rows = []
    for r in display_rows:
        table_rows.append({
            "Month": r.get("month", ""),
            "Current Used (GB)": r.get("used_gb", ""),
            "Current Allocate (GB)": r.get("allocated_gb", ""),
            "% Used": r.get("pct_used", ""),
        })

    add_small_table(doc, ["Month", "Current Used (GB)", "Current Allocate (GB)", "% Used"], table_rows, header_fill=HEADER_BLUE)

    add_normal(
        doc,
        f"The current size of the database is {current_used} GB. ",
    )

    # Dynamic growth rate analysis
    if len(display_rows) >= 2:
        first = display_rows[0]
        last = display_rows[-1]
        used_diff = last.get("used_gb", 0) - first.get("used_gb", 0)
        n_months = max(len(display_rows) - 1, 1)
        avg_growth_per_month = round(used_diff / n_months, 2)

        alloc_diff = last.get("allocated_gb", 0) - first.get("allocated_gb", 0)
        avg_alloc_per_month = round(alloc_diff / n_months, 2)

        if avg_growth_per_month > 0:
            add_normal(
                doc,
                f"We compare data growth rate between current allocated data which is {current_alloc} GB "
                f"and the allocated size increasing around {avg_alloc_per_month} GB per month. "
                f"The database growth rate is around {avg_growth_per_month} GB per month. "
                f"As a result of database size, the current allocated size of the database is sufficient to support database growth rate in this current situation.",
            )
        else:
            add_normal(
                doc,
                f"The current allocated size is {current_alloc} GB. "
                f"The database growth rate decreased by around {abs(avg_growth_per_month)} GB per month. "
                f"As a result of database size, the current allocated size of the database is sufficient to support database growth rate in this current situation.",
            )
    else:
        add_normal(
            doc,
            f"The current allocated size is {current_alloc} GB. "
            f"As a result of database size, the current allocated size of the database is sufficient to support database growth rate in this current situation.",
        )

    doc.add_page_break()


# =========================================================
# Section 5.3: Performance Analysis
# =========================================================

def _fmt_number(value, decimals=0):
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value)
    if decimals:
        return f"{number:,.{decimals}f}"
    return f"{number:,.0f}"


def _current_index(rows):
    for index, row in enumerate(rows):
        if row.get("is_current"):
            return index
    factors = []
    for index, row in enumerate(rows):
        factor = row.get("pga_target_factor", row.get("sga_size_factor", None))
        try:
            factors.append((abs(float(factor) - 1.0), index))
        except (TypeError, ValueError):
            pass
    return min(factors)[1] if factors else 0


def _advisory_window(rows, before=1, after=3):
    if not rows:
        return []
    current = _current_index(rows)
    start = max(0, current - before)
    end = min(len(rows), current + after + 1)
    if end - start < before + after + 1:
        start = max(0, end - (before + after + 1))
    return rows[start:end]


def _dark_advisory_chart(rows, x_labels, y_values, title, x_label, y_label, legend_label, output_path):
    fig, ax = plt.subplots(figsize=(7.1, 3.5), dpi=150, facecolor="#2F2F2F")
    ax.set_facecolor("#3B3B3B")

    ax.plot(x_labels, y_values, marker="o", color="#ED7D31", linewidth=2.2, markersize=4, label=legend_label)
    ax.set_title(title, color="#EDEDED", fontsize=13, fontweight="bold", pad=12)
    ax.set_xlabel(x_label, color="#D9D9D9", fontsize=8, fontweight="bold")
    ax.set_ylabel(y_label, color="#D9D9D9", fontsize=8, fontweight="bold")
    ax.tick_params(colors="#D9D9D9", labelsize=7)
    ax.yaxis.set_major_formatter(FuncFormatter(lambda value, _: f"{value:,.0f}"))
    ax.grid(True, axis="y", color="#5A5A5A", linewidth=0.6)

    for spine in ax.spines.values():
        spine.set_visible(False)

    for x, y in zip(x_labels, y_values):
        ax.text(x, y, _fmt_number(y), color="#CFCFCF", fontsize=7, ha="center", va="bottom")

    current = _current_index(rows)
    if 0 <= current < len(rows):
        current_x = x_labels[current]
        current_y = y_values[current]
        ax.scatter([current_x], [current_y], s=90, facecolors="none", edgecolors="yellow", linewidths=2.2, zorder=4)
        ax.annotate(
            "Current",
            xy=(current_x, current_y),
            xytext=(-55, -25),
            textcoords="offset points",
            bbox={"boxstyle": "square,pad=0.35", "fc": "yellow", "ec": "black", "lw": 0.8},
            color="black",
            fontsize=8,
            arrowprops={"arrowstyle": "-", "color": "yellow", "lw": 1.2},
        )

    legend = ax.legend(loc="lower center", bbox_to_anchor=(0.5, -0.28), fontsize=7, frameon=False)
    for text in legend.get_texts():
        text.set_color("#D9D9D9")

    fig.tight_layout(pad=2.0)
    fig.savefig(output_path, facecolor="#2F2F2F")
    plt.close(fig)


def create_pga_advisory_chart(pga_rows, output_dir):
    if not pga_rows:
        return None

    rows = _advisory_window(pga_rows)
    labels = []
    values = []
    for row in rows:
        size_mb = int(float(row.get("pga_size_bytes", 0) or 0) / (1024 * 1024))
        label = _fmt_number(size_mb)
        if row.get("is_current"):
            label = f"{label} (Current)"
        labels.append(label)
        values.append(float(row.get("estd_extra_mb_rw", 0) or 0))

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    chart_path = output_dir / "database_pga_advisory.png"

    _dark_advisory_chart(
        rows,
        labels,
        values,
        title="PGA Memory Advisory",
        x_label="PGA TARGET EST (MB)",
        y_label="ESTD EXTRA READ/WRITTEN TO DISK",
        legend_label="Estd Extra W/A MB Read/Written to Disk",
        output_path=chart_path,
    )
    
    return chart_path
    
def create_sga_advisory_chart(sga_rows, output_dir):
    if not sga_rows:
        return None

    rows = _advisory_window(sga_rows)
    labels = []
    values = []
    for row in rows:
        size_mb = int(float(row.get("sga_size_mb", 0) or 0))
        label = _fmt_number(size_mb)
        if row.get("is_current"):
            label = f"{label} (Current)"
        labels.append(label)
        values.append(float(row.get("estd_physical_reads", 0) or 0))

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    chart_path = output_dir / "database_sga_advisory.png"

    _dark_advisory_chart(
        rows,
        labels,
        values,
        title="SGA Memory Advisory",
        x_label="SGA TARGET SIZE (MB)",
        y_label="EST PHYSICAL READS",
        legend_label="Est Physical Reads",
        output_path=chart_path,
    )
    
    return chart_path


def _sample_tick_labels(samples, max_ticks=10):
    if not samples:
        return [], []

    step = max(1, len(samples) // max_ticks)
    positions = list(range(0, len(samples), step))
    if positions[-1] != len(samples) - 1:
        positions.append(len(samples) - 1)

    labels = [samples[pos].get("day", "") for pos in positions]
    return positions, labels


def _create_oswatcher_chart(samples, value_key, chart_title, output_path, y_limit=None):
    if not samples:
        return None

    output_path = Path(output_path)
    values = [float(row.get(value_key, 0) or 0) for row in samples]
    x_values = list(range(len(values)))
    tick_positions, tick_labels = _sample_tick_labels(samples)

    fig, ax = plt.subplots(figsize=(8.2, 2.0), dpi=140)
    fig.patch.set_facecolor("#C9C7EF")
    ax.set_facecolor("#BFBFBF")
    ax.plot(x_values, values, color="#002BFF", linewidth=0.8)

    ax.set_title(chart_title, fontsize=7, fontweight="bold", pad=3)
    ax.grid(True, color="#777777", linewidth=0.35, alpha=0.65)
    ax.tick_params(axis="both", labelsize=5, pad=1)
    ax.set_xlim(0, max(len(values) - 1, 1))
    if y_limit:
        ax.set_ylim(0, y_limit)

    ax.set_xticks(tick_positions)
    ax.set_xticklabels(tick_labels)
    for spine in ax.spines.values():
        spine.set_color("#666666")
        spine.set_linewidth(0.6)

    fig.text(0.02, 0.02, "Host=MotifPRDDBN", fontsize=5, color="#333333")
    plt.tight_layout(pad=0.8)
    plt.savefig(output_path, facecolor=fig.get_facecolor())
    plt.close(fig)
    return output_path


def _create_oswatcher_multi_chart(samples, series, chart_title, output_path, y_limit=None):
    if not samples:
        return None

    output_path = Path(output_path)
    x_values = list(range(len(samples)))
    tick_positions, tick_labels = _sample_tick_labels(samples)

    fig, ax = plt.subplots(figsize=(8.2, 2.0), dpi=140)
    fig.patch.set_facecolor("#C9C7EF")
    ax.set_facecolor("#BFBFBF")

    for key, label, color in series:
        values = [float(row.get(key, 0) or 0) for row in samples]
        ax.plot(x_values, values, color=color, linewidth=0.8, label=label)

    ax.set_title(chart_title, fontsize=7, fontweight="bold", pad=3)
    ax.grid(True, color="#777777", linewidth=0.35, alpha=0.65)
    ax.tick_params(axis="both", labelsize=5, pad=1)
    ax.set_xlim(0, max(len(samples) - 1, 1))
    if y_limit:
        ax.set_ylim(0, y_limit)
    ax.set_xticks(tick_positions)
    ax.set_xticklabels(tick_labels)
    if len(series) > 1:
        ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.18), ncol=min(len(series), 4), fontsize=5, frameon=False)
    for spine in ax.spines.values():
        spine.set_color("#666666")
        spine.set_linewidth(0.6)

    fig.text(0.02, 0.02, "Host=MotifPRDDBN", fontsize=5, color="#333333")
    plt.tight_layout(pad=0.8)
    plt.savefig(output_path, facecolor=fig.get_facecolor())
    plt.close(fig)
    return output_path


def _max_disk_samples_by_time(samples):
    by_key = {}
    for row in samples:
        key = (row.get("day", ""), row.get("time", ""))
        current = by_key.get(key)
        if current is None or float(row.get("svctm", 0) or 0) > float(current.get("svctm", 0) or 0):
            by_key[key] = row
    return list(by_key.values())


def _rate_percent(samples, key, threshold):
    if not samples:
        return 0.0
    count = sum(1 for row in samples if float(row.get(key, 0) or 0) > threshold)
    return (count / len(samples)) * 100.0


def create_os_perf_charts(os_perf_data, output_dir):
    """Generate CPU, Memory, and Disk line charts using matplotlib"""
    if not os_perf_data or not os_perf_data.get("found"):
        return None, None, None
        
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    cpu_path = output_dir / "os_cpu_perf.png"
    mem_path = output_dir / "os_mem_perf.png"
    disk_path = output_dir / "os_disk_perf.png"
    runq_path = output_dir / "os_runq_perf.png"
    cpu_all_path = output_dir / "os_cpu_util_all.png"
    cpu_sys_path = output_dir / "os_cpu_util_system.png"
    cpu_usr_path = output_dir / "os_cpu_util_user.png"
    mem_free_path = output_dir / "os_mem_free.png"
    swap_used_path = output_dir / "os_swap_used.png"
    page_in_path = output_dir / "os_page_in.png"
    page_out_path = output_dir / "os_page_out.png"
    cpu_wait_io_path = output_dir / "os_cpu_wait_io.png"
    disk_service_path = output_dir / "os_disk_service_time.png"

    runq_samples = os_perf_data.get("runq_samples", [])
    cpu_samples = os_perf_data.get("cpu_samples", [])
    mem_free_samples = os_perf_data.get("mem_free_samples", [])
    swap_samples = os_perf_data.get("swap_samples", [])
    page_io_samples = os_perf_data.get("page_io_samples", [])
    disk_samples = _max_disk_samples_by_time(os_perf_data.get("disk_samples", []))
    os_perf_data["chart_paths"] = {
        "runq": str(_create_oswatcher_chart(runq_samples, "runq", "OSWatcher  CPU: Run Queue", runq_path) or ""),
        "cpu_all": str(_create_oswatcher_chart(cpu_samples, "util", "OSWatcher  CPU Utilization", cpu_all_path, y_limit=100) or ""),
        "cpu_sys": str(_create_oswatcher_chart(cpu_samples, "sys", "OSWatcher  CPU Utilization: System", cpu_sys_path, y_limit=100) or ""),
        "cpu_usr": str(_create_oswatcher_chart(cpu_samples, "usr", "OSWatcher  CPU Utilization: User", cpu_usr_path, y_limit=100) or ""),
        "mem_free": str(_create_oswatcher_multi_chart(mem_free_samples, [("free_gb", "Memory Free (GB)", "#005A7F"), ("inactive_gb", "Inactive Cache (GB)", "#F0782E")], "OSWatcher Memory Free (Gbytes)", mem_free_path) or ""),
        "swap_used": str(_create_oswatcher_chart(swap_samples, "used_gb", "OSWatcher  Memory: Swap Used (G Bytes)", swap_used_path) or ""),
        "page_in": str(_create_oswatcher_chart(page_io_samples, "pgpgin", "OSWatcher Memory: Page In Rate (Pages Per Second)", page_in_path) or ""),
        "page_out": str(_create_oswatcher_chart(page_io_samples, "pgpgout", "OSWatcher Memory: Page Out Rate (Pages Per Second)", page_out_path) or ""),
        "cpu_wait_io": str(_create_oswatcher_chart(cpu_samples, "iowait", "OSWatcher  CPU Utilization: Wait I/O", cpu_wait_io_path, y_limit=100) or ""),
        "disk_service": str(_create_oswatcher_multi_chart(disk_samples, [("svctm", "Service Time (ms)", "#1451FF"), ("await", "I/O Wait (ms)", "#E3342F")], "OSWatcher  IO: Devices With Highest Service Time In Milliseconds", disk_service_path) or ""),
    }
    
    # 1. CPU Chart
    cpu_trend = os_perf_data.get("cpu_trend", [])
    if cpu_trend:
        days = [r["day"] for r in cpu_trend]
        usr = [r["usr"] for r in cpu_trend]
        sys = [r["sys"] for r in cpu_trend]
        iowait = [r["iowait"] for r in cpu_trend]
        
        plt.figure(figsize=(10, 5), dpi=120)
        plt.plot(days, usr, marker='o', label='%usr', color='#00478F', linewidth=2)
        plt.plot(days, sys, marker='s', label='%sys', color='#FF8C00', linewidth=2)
        plt.plot(days, iowait, marker='^', label='%iowait', color='#E63946', linewidth=2)
        
        plt.title("CPU Utilization Trend (Daily Average)", fontsize=14, fontweight="bold", pad=15)
        plt.xlabel("Day of Month", fontsize=10)
        plt.ylabel("CPU %", fontsize=10)
        plt.legend()
        plt.grid(True, linestyle="--", alpha=0.6)
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig(cpu_path)
        plt.close()
    else:
        cpu_path = None
        
    # 2. Memory Chart (Page In/Out or Swap)
    mem_trend = os_perf_data.get("mem_trend", [])
    if mem_trend:
        days = [r["day"] for r in mem_trend]
        pgpgin = [r["pgpgin"] for r in mem_trend]
        pgpgout = [r["pgpgout"] for r in mem_trend]
        
        plt.figure(figsize=(10, 5), dpi=120)
        plt.plot(days, pgpgin, marker='o', label='Page In (pgpgin/s)', color='#2A9D8F', linewidth=2)
        plt.plot(days, pgpgout, marker='s', label='Page Out (pgpgout/s)', color='#E76F51', linewidth=2)
        
        plt.title("Memory Page In/Out Trend (Daily Average)", fontsize=14, fontweight="bold", pad=15)
        plt.xlabel("Day of Month", fontsize=10)
        plt.ylabel("Pages/s", fontsize=10)
        plt.legend()
        plt.grid(True, linestyle="--", alpha=0.6)
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig(mem_path)
        plt.close()
    else:
        mem_path = None
        
    # 3. Disk Chart (Service Time)
    disk_trend = os_perf_data.get("disk_trend", [])
    if disk_trend:
        days = [r["day"] for r in disk_trend]
        svctm = [r["svctm"] for r in disk_trend]
        await_val = [r["await"] for r in disk_trend]
        
        plt.figure(figsize=(10, 5), dpi=120)
        plt.plot(days, svctm, marker='o', label='Service Time (svctm)', color='#6D6875', linewidth=2)
        plt.plot(days, await_val, marker='s', label='I/O Wait (await)', color='#B5838D', linewidth=2)
        
        plt.title("Disk Performance Trend (Daily Average)", fontsize=14, fontweight="bold", pad=15)
        plt.xlabel("Day of Month", fontsize=10)
        plt.ylabel("Milliseconds (ms)", fontsize=10)
        plt.legend()
        plt.grid(True, linestyle="--", alpha=0.6)
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig(disk_path)
        plt.close()
    else:
        disk_path = None
        
    return cpu_path, mem_path, disk_path

def add_performance_analysis(doc, data, output_dir):
    sec53 = data.get("section53", {})
    
    add_heading2(doc, "5.3 Performance Analysis")
    
    # ----- 5.3.1 CPU Section -----
    add_heading3(doc, "5.3.1 CPU Section")
    add_normal(doc, "Instance Efficiency Percentages (Target 100%)", bold=True)
    inst_eff = sec53.get("instance_efficiency", {})
    if inst_eff:
        rows = [{"Metric": k, "Percentage (%)": v} for k, v in inst_eff.items()]
        add_small_table(doc, ["Metric", "Percentage (%)"], rows, header_fill=HEADER_BLUE)
        
        parse_cpu_elapsd = inst_eff.get("Parse CPU to Parse Elapsd %", "")
        try:
            parse_val = float(parse_cpu_elapsd)
            if parse_val < 50.0:
                add_normal(doc, f"Parse CPU to Parse Elapsed is low ratio ({parse_val}%), more CPU time is spending to parse SQL statements. It may mean investigation is warranted to determine why the CPU is spending this much time simply parsing SQL statements. However, if this is normal behavior of the database, it can be ignored.")
            else:
                add_normal(doc, "Parse CPU to Parse Elapsed ratio is normal.")
        except ValueError:
            pass
            
        add_normal(doc, "Other elements of instance efficiency percentages are fine.")
    else:
        add_normal(doc, "No instance efficiency data found in fast assessment logs.")
    
    # ----- 5.3.2 SQL Section -----
    add_heading3(doc, "5.3.2 SQL Section")
    
    add_normal(doc, "Top 5 Timed Foreground Events", bold=True)
    top5_events = sec53.get("top5_events", [])
    if top5_events:
        for event in top5_events:
            add_normal(doc, f"- {event}")
        if any("dblink" in ev.lower() for ev in top5_events):
            add_normal(doc, "- The SQL*Net message from dblink wait event means that your local system is waiting on the network to transfer the data across the network. It's a very normal wait event for this sort of query. We recommend checking and improving the Bandwidth network or creating a view on the remote site referencing the local tables.")
    else:
        add_normal(doc, "No Top 5 Timed Events found.")
    
    add_normal(doc, "")
    
    add_normal(doc, "Top SQL by Disk Reads:")
    sql_disk = sec53.get("top_sql_disk", [])
    if sql_disk:
        rows = [{"Reads": q.get("reads"), "SQL_ID": q.get("sql_id"), "Text": q.get("text")} for q in sql_disk]
        add_small_table(doc, ["Reads", "SQL_ID", "Text"], rows, header_fill=HEADER_BLUE)
    else:
        add_normal(doc, "No Top SQL by Disk Reads found.")
        
    add_normal(doc, "Top SQL by Buffer Gets (Memory):")
    sql_mem = sec53.get("top_sql_mem", [])
    if sql_mem:
        rows = [{"Buffer Gets": q.get("buffer_gets"), "SQL_ID": q.get("sql_id"), "Text": q.get("text")} for q in sql_mem]
        add_small_table(doc, ["Buffer Gets", "SQL_ID", "Text"], rows, header_fill=HEADER_BLUE)
    else:
        add_normal(doc, "No Top SQL by Buffer Gets found.")
        
    # ----- 5.3.3 Memory Section -----
    add_heading3(doc, "5.3.3 Memory Section")
    mem_pools = sec53.get("memory_pools", {})
    if mem_pools.get("shared_free_pct"):
        rows = [
            {"Pool Name": "Shared Pool Free %", "Value": mem_pools.get("shared_free_pct", "")},
            {"Pool Name": "Java Pool Free %", "Value": mem_pools.get("java_free_pct", "")},
            {"Pool Name": "Large Pool Free %", "Value": mem_pools.get("large_free_pct", "")},
        ]
        add_small_table(doc, ["Pool Name", "Value"], rows)
    else:
        add_normal(doc, "No detailed memory pool statistics found.")
        
    # ----- 5.3.4 Oracle Open Cursors -----
    add_heading3(doc, "5.3.4 Oracle Open Cursors")
    cursors = sec53.get("open_cursors", {})
    if cursors.get("limit"):
        rows = [
            {"Parameter": "Open Cursors Limit", "Value": cursors.get("limit", "")},
            {"Parameter": "Estimated Usage", "Value": cursors.get("usage", "")},
            {"Parameter": "Usage %", "Value": cursors.get("pct", "")},
        ]
        add_small_table(doc, ["Parameter", "Value"], rows)
        
        try:
            pct_val = float(cursors.get("pct", "0").replace("%", ""))
            if pct_val >= 95.0:
                add_normal(doc, f"As above table shows that percent usage of open_cursors and session_cached_cursors are {pct_val}%. It indicates that open_cursors and session_cached_cursors of the database are insufficient. Therefore, we recommend increasing the value of open_cursors and session_cached_cursors to 2000. It may improve these values. However, it will consume more memory in PGA. So, you should increase the PGA size too.")
            else:
                add_normal(doc, f"The percent usage of open_cursors is {pct_val}%, which is within acceptable limits.")
        except ValueError:
            pass
    else:
        add_normal(doc, "No open cursors data found.")
        
    # ----- 5.3.5 PGA Memory Advisory -----
    add_heading3(doc, "Program Global Area (PGA) Analysis")
    pga_rows = sec53.get("pga_advisory", [])
    if pga_rows:
        pga_display_rows = _advisory_window(pga_rows)
        current_pga = pga_rows[_current_index(pga_rows)]
        current_pga_mb = int(float(current_pga.get("pga_size_bytes", 0) or 0) / (1024 * 1024))
        current_pga_hit = current_pga.get("estd_pga_cache_hit_pct", 0)
        current_extra_mb = current_pga.get("estd_extra_mb_rw", 0)

        summary_table = add_small_table(
            doc,
            ["Allocate (MB)", "PGA cache Hit %"],
            [{"Allocate (MB)": _fmt_number(current_pga_mb), "PGA cache Hit %": _fmt_number(current_pga_hit, 0)}],
            header_fill=HEADER_BLUE,
        )
        summary_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        chart_path = create_pga_advisory_chart(pga_rows, output_dir)
        if chart_path:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(str(chart_path), width=Inches(6.7))

        rows = []
        for row in pga_display_rows:
            pga_size_mb = int(float(row.get("pga_size_bytes", 0) or 0) / (1024 * 1024))
            label = _fmt_number(pga_size_mb)
            if row.get("is_current"):
                label = f"{label} (Current)"
            rows.append(
                {
                    "PGA Target\nEst (MB)": label,
                    "Estd Extra   W/A MB\nRead/Written to Disk": _fmt_number(row.get("estd_extra_mb_rw", 0)),
                    "Estd PGA\nCache Hit %": _fmt_number(row.get("estd_pga_cache_hit_pct", 0), 1),
                    "_current": row.get("is_current", False),
                }
            )

        def highlight_current(row, h, value):
            return bool(row.get("_current"))

        pga_table = add_small_table(
            doc,
            ["PGA Target\nEst (MB)", "Estd Extra   W/A MB\nRead/Written to Disk", "Estd PGA\nCache Hit %"],
            rows,
            header_fill=HEADER_BLUE,
            highlight_rule=highlight_current,
        )
        pga_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        add_normal(
            doc,
            f"From the graph PGA Target size forecast shows that current PGA Target size is sufficient. "
            f"The current PGA Target size is {_fmt_number(current_pga_mb)} MB, it has estimated Read/Written to Disk "
            f"{_fmt_number(current_extra_mb)} and PGA cache hit {_fmt_number(current_pga_hit, 0)}%.",
        )
    else:
        add_normal(doc, "No PGA Advisory data found.")
        
    # ----- 5.3.6 SGA Memory Advisory -----
    add_heading3(doc, "System Global Area (SGA) Analysis")
    sga_rows = sec53.get("sga_advisory", [])
    if sga_rows:
        sga_display_rows = _advisory_window(sga_rows)
        current_sga = sga_rows[_current_index(sga_rows)]
        current_sga_mb = int(float(current_sga.get("sga_size_mb", 0) or 0))
        current_sga_reads = int(float(current_sga.get("estd_physical_reads", 0) or 0))

        chart_path = create_sga_advisory_chart(sga_rows, output_dir)
        if chart_path:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(str(chart_path), width=Inches(6.7))

        rows = []
        for row in sga_display_rows:
            sga_size_mb = int(float(row.get("sga_size_mb", 0) or 0))
            label = _fmt_number(sga_size_mb)
            if row.get("is_current"):
                label = f"{label} (Current)"
            rows.append(
                {
                    "SGA Target\nSize (MB)": label,
                    "Est Physical Reads": _fmt_number(row.get("estd_physical_reads", 0)),
                    "_current": row.get("is_current", False),
                }
            )

        def highlight_current_sga(row, h, value):
            return bool(row.get("_current"))

        sga_table = add_small_table(
            doc,
            ["SGA Target\nSize (MB)", "Est Physical Reads"],
            rows,
            header_fill=HEADER_BLUE,
            highlight_rule=highlight_current_sga,
        )
        sga_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        add_normal(
            doc,
            f"From the graph, the SGA Target size forecast shows that the current SGA Target size is sufficient. "
            f"The current SGA Target size is {_fmt_number(current_sga_mb)} MB, it has estimate physical reads "
            f"{_fmt_number(current_sga_reads)}.",
        )
    else:
        add_normal(doc, "No SGA Advisory data found.")
        
    doc.add_page_break()


# =========================================================
# Tablespace / Registry / Appendix
# =========================================================

def add_tablespace_free_space(doc, data):
    add_heading1(doc, "6. Tablespace Free Space")
    add_heading2(doc, "6.1 Tablespace Free Space")

    rows = []
    for name, info in data.get("tablespace_free", {}).items():
        try:
            pct_free = float(info.get("pct_free", 0))
        except Exception:
            pct_free = 0

        rows.append({
            "TABLESPACE_NAME": name,
            "Percentage of Free space (%)": pct_free,
            "Status": "Critical" if pct_free < 10 else "Warning" if pct_free < 20 else "OK",
        })

    rows.sort(key=lambda x: x["Percentage of Free space (%)"])

    def highlight_tbs(row, h, value):
        return row.get("Status") in ["Critical", "Warning"]

    add_small_table(doc, ["TABLESPACE_NAME", "Percentage of Free space (%)", "Status"], rows, header_fill=HEADER_BLUE, highlight_rule=highlight_tbs)

    add_normal(doc, "Additional Suggestion:", bold=True)
    add_normal(doc, "- Refer tablespace free space. Tablespaces with low free space should be monitored. If datafile auto extend is enabled, risk may be reduced.")

    doc.add_page_break()


def add_default_tablespace(doc, data):
    add_heading1(doc, "7. Default tablespace and temporary tablespace")
    add_heading2(doc, "7.1 Default tablespace and temporary tablespace")

    rows = []
    for r in data.get("user_tablespaces", []):
        rows.append({
            "User Name": r.get("user_name", ""),
            "Default Tablespace": r.get("default_tbs", ""),
            "Temporary Tablespace": r.get("temp_tbs", ""),
        })

    add_small_table(doc, ["User Name", "Default Tablespace", "Temporary Tablespace"], rows)

    doc.add_page_break()


def add_database_registry(doc, data):
    add_heading1(doc, "8. Database Registry")
    add_heading2(doc, "8.1 Check Database Registry")

    rows = []
    for r in data.get("registry_list", []):
        rows.append({
            "Comp ID": r.get("comp_id", ""),
            "Version": r.get("version", ""),
            "Status": r.get("status", ""),
            "Last Modified": r.get("last_modified", ""),
        })

    add_small_table(doc, ["Comp ID", "Version", "Status", "Last Modified"], rows)

    doc.add_page_break()


def add_appendix_invalid_objects(doc, data):
    add_heading1(doc, "APPENDIX A – Invalid Object")
    add_heading2(doc, "A.a Invalid Objects report.")

    rows = []
    for r in data.get("invalid_objects_list", []):
        rows.append({
            "Creator": r.get("creator", ""),
            "Object Name": r.get("object_name", ""),
            "Object Type": r.get("object_type", ""),
            "Status": r.get("status", ""),
        })

    add_small_table(doc, ["Creator", "Object Name", "Object Type", "Status"], rows[:350])

    if len(rows) > 350:
        add_normal(doc, f"Showing first 350 of {len(rows)} invalid objects.")

    add_heading2(doc, "A.b Disabled Constraints.")
    add_normal(doc, "No disabled constraints data found.")

    add_normal(doc, "Additional Suggestion:", bold=True)
    add_normal(doc, "- You should verify invalid objects and disabled constraints that are used by application or not.")

    doc.add_page_break()


def add_appendix_defaults(doc, data, output_dir):
    add_heading1(doc, "APPENDIX B – Information from alert log")
    add_normal(doc, "- There is no alert Log to concern in this quarter.")
    doc.add_page_break()

    add_heading1(doc, "APPENDIX C – SQL Statement should be to investigate")
    add_normal(doc, "- There is no SQL statement to concern in this quarter.")
    doc.add_page_break()

    add_heading1(doc, "APPENDIX D – Operating system log")
    add_normal(doc, "- There is no Operating System log to concern in this quarter.")
    doc.add_page_break()

    add_heading1(doc, "APPENDIX E – Backup Configuration")
    backup_info = data.get("backup_info", {})
    if backup_info.get("found"):
        title = backup_info.get("title", "")
        if title:
            add_normal(doc, title, bold=True)
        log_content = backup_info.get("log_content", "")
        add_monospaced_box(doc, log_content, font_size=7)
    else:
        add_normal(doc, "No backup configuration data found.")
    doc.add_page_break()

    add_heading1(doc, "APPENDIX F – OS Performance Summary")
    
    os_info = data.get("os_info", {})
    os_perf = data.get("os_perf", {})
    
    if os_perf.get("found"):
        hostname = os_info.get("hostname", "N/A")
        os_release = os_info.get("os_version") or os_info.get("os_release") or "N/A"
        cpu_count = os_info.get("cpu_count") or os_info.get("cpu_cores") or "N/A"
        cores = os_info.get("cores_per_socket", "")
        ram_gb = os_info.get("total_ram_gb") or os_info.get("physical_memory_gb") or "N/A"

        cpu_label = f"{cpu_count} CPUs"
        if cores:
            cpu_label += f" ({cores} cores)"

        summary_table = add_small_table(
            doc,
            ["Name", "Value"],
            [
                {"Name": "Hostname", "Value": hostname},
                {"Name": "OS Version", "Value": os_release},
                {"Name": "CPU COUNT", "Value": cpu_label},
            ],
        )
        summary_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        cpu_path, mem_path, disk_path = create_os_perf_charts(os_perf, output_dir)
        chart_paths = os_perf.get("chart_paths", {})
        
        if chart_paths.get("cpu_all") or cpu_path:
            add_heading2(doc, "F.1 CPU Performance")
            
            runq_samples = os_perf.get("runq_samples", [])
            runq_trend = os_perf.get("runq_trend", [])
            if runq_samples or runq_trend:
                add_heading3(doc, "Process Run Queue")
                if chart_paths.get("runq"):
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.add_run().add_picture(chart_paths["runq"], width=Inches(6.8))

                runq_values = [float(r.get("runq", 0) or 0) for r in runq_samples]
                max_runq = max(runq_values, default=max((r.get("max", 0) for r in runq_trend), default=0))
                avg_runq = sum(runq_values) / max(len(runq_values), 1) if runq_values else sum(r.get("avg", 0) for r in runq_trend) / max(len(runq_trend), 1)

                add_normal(doc, "The following snaps recorded very high run queue values:")
                runq_table = add_small_table(doc, ["Rate", "Percent(%)"], [
                    {"Rate": "High (>3)", "Percent(%)": f"{_rate_percent(runq_samples, 'runq', 3):.2f}"},
                    {"Rate": "Very High (>6)", "Percent(%)": f"{_rate_percent(runq_samples, 'runq', 6):.2f}"},
                ], header_fill=HEADER_BLUE)
                runq_table.alignment = WD_TABLE_ALIGNMENT.CENTER

                try:
                    cpu_c = int(float(cpu_count))
                except ValueError:
                    cpu_c = 2

                if max_runq > cpu_c * 2:
                    add_normal(doc, f"The run queue value is very high, indicating a severe CPU bottleneck. With only {cpu_c} CPUs, the run queue should ideally remain below {cpu_c}. However, the graph shows that the queue frequently spikes and reaches a maximum of {max_runq:.2f}. This means there are often more processes waiting for a CPU than CPUs available, which leads to significant performance delays.")
                    add_normal(doc, f"The average value of {avg_runq:.2f} does not reflect the severity of the frequent peak loads. CPU resources are insufficient for the current workload and upgrading or optimizing CPU usage is strongly recommended.")
                else:
                    add_normal(doc, f"The run queue value is normal. Maximum run queue is {max_runq:.2f} with an average of {avg_runq:.2f}. CPU resources are sufficient for the current workload.")
            
            add_heading3(doc, "CPU Utilization")
            for key, heading in [
                ("cpu_all", None),
                ("cpu_sys", "CPU Utilization by System"),
                ("cpu_usr", "CPU Utilization by User"),
            ]:
                if heading:
                    add_heading3(doc, heading)
                if chart_paths.get(key):
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.add_run().add_picture(chart_paths[key], width=Inches(6.8))

            cpu_samples = os_perf.get("cpu_samples", [])
            if cpu_samples:
                add_normal(doc, "The following recorded high CPU utilization values:")
                cpu_rate_table = add_small_table(doc, ["Rate", "Percent(%)"], [
                    {"Rate": "High (>95%)", "Percent(%)": f"{_rate_percent(cpu_samples, 'util', 95):.2f}"},
                    {"Rate": "Very High (100%)", "Percent(%)": f"{_rate_percent(cpu_samples, 'util', 99.99):.2f}"},
                ], header_fill=HEADER_BLUE)
                cpu_rate_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            cpu_trend = os_perf.get("cpu_trend", [])
            usr_vals = [r["usr"] for r in cpu_trend] or [r["usr"] for r in cpu_samples]
            sys_vals = [r["sys"] for r in cpu_trend] or [r["sys"] for r in cpu_samples]
            iow_vals = [r["iowait"] for r in cpu_trend] or [r["iowait"] for r in cpu_samples]
            
            sample_util = [r.get("util", 0) for r in cpu_samples]
            max_cpu = max(sample_util, default=max([u + s for u, s in zip(usr_vals, sys_vals)], default=0))
            avg_cpu = sum(sample_util) / max(len(sample_util), 1) if sample_util else sum([u + s for u, s in zip(usr_vals, sys_vals)]) / max(len(usr_vals), 1)
            
            if max_cpu >= 95:
                add_normal(doc, f"CPU utilization is high. The graph shows that CPU usage frequently reaches and sustains near 100% for extended periods. Highest CPU usage is {max_cpu:.2f}% while the average CPU usage is around {avg_cpu:.2f}%. The higher the CPU utilization, the longer it will take processes to run, indicating that the system is operating at full capacity and may experience performance bottlenecks.")
            else:
                add_normal(doc, f"CPU utilization is within acceptable limits. Highest CPU usage is {max_cpu:.2f}% while the average CPU usage is around {avg_cpu:.2f}%.")
            
            add_small_table(doc, ["Metric", "Min (%)", "Max (%)", "Avg (%)"], [
                {"Metric": "CPU Utilization by User (%usr)", "Min (%)": f"{min(usr_vals):.2f}", "Max (%)": f"{max(usr_vals):.2f}", "Avg (%)": f"{(sum(usr_vals)/len(usr_vals)):.2f}"},
                {"Metric": "CPU Utilization by System (%sys)", "Min (%)": f"{min(sys_vals):.2f}", "Max (%)": f"{max(sys_vals):.2f}", "Avg (%)": f"{(sum(sys_vals)/len(sys_vals)):.2f}"},
                {"Metric": "CPU Wait I/O (%iowait)", "Min (%)": f"{min(iow_vals):.2f}", "Max (%)": f"{max(iow_vals):.2f}", "Avg (%)": f"{(sum(iow_vals)/len(iow_vals)):.2f}"},
            ])
            
        if chart_paths.get("mem_free") or mem_path:
            add_heading2(doc, "F.2 Memory Performance")
            add_heading3(doc, "Physical Memory")
            if chart_paths.get("mem_free"):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture(chart_paths["mem_free"], width=Inches(6.8))

            add_heading3(doc, "Swap")
            if chart_paths.get("swap_used"):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture(chart_paths["swap_used"], width=Inches(6.8))

            add_heading3(doc, "Page in")
            if chart_paths.get("page_in"):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture(chart_paths["page_in"], width=Inches(6.8))

            add_heading3(doc, "Page out")
            if chart_paths.get("page_out"):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture(chart_paths["page_out"], width=Inches(6.8))

            total_ram = os_info.get("total_ram_mb", "")
            swap_total = os_info.get("swap_total_mb", "")
            if total_ram or swap_total:
                add_normal(doc, "System configuration:", bold=True)
                if total_ram:
                    try:
                        add_normal(doc, f"Total memory = {_fmt_number(float(total_ram), 2)} Megabytes ({_fmt_number(float(total_ram) / 1024, 2)} Gigabytes)")
                    except ValueError:
                        add_normal(doc, f"Total memory = {total_ram} Megabytes")
                if swap_total:
                    try:
                        add_normal(doc, f"Total swap = {_fmt_number(float(swap_total), 2)} Megabytes ({_fmt_number(float(swap_total) / 1024, 2)} Gigabytes)")
                    except ValueError:
                        add_normal(doc, f"Total swap = {swap_total} Megabytes")
            
            free_trend = os_perf.get("free_mem_trend", [])
            swap_trend = os_perf.get("swap_trend", [])
            mem_free_samples = os_perf.get("mem_free_samples", [])
            swap_samples = os_perf.get("swap_samples", [])
            
            if free_trend and swap_trend:
                min_free = min((r.get("free_gb", 0) for r in mem_free_samples), default=0)
                avg_free = sum(r.get("free_gb", 0) for r in mem_free_samples) / max(len(mem_free_samples), 1)
                max_swap_gb = max((r.get("used_gb", 0) for r in swap_samples), default=0)
                max_swap_pct = max((r.get("used_pct", 0) for r in swap_samples), default=max((r.get("max_pct", 0) for r in swap_trend), default=0))
                
                if max_swap_pct > 50.0 or min_free < 0.5:
                    add_normal(doc, f"At the highest workload, the free memory remains around {min_free:.2f} GB. This indicates that the current workload is close to or in excess of available physical memory, causing the system to use swap and resulting in degraded performance.")
                    add_normal(doc, f"The swap memory load reached around {max_swap_gb:.2f} GB. In some moments of time, high memory page in and page out rates are observed. That means OS is low on free memory. We strongly recommend consulting with the system team to monitor memory usage trends, investigate memory-intensive processes, and plan for a memory capacity upgrade.")
                    add_normal(doc, f"As a result of the above graph, memory should be monitored closely. Average free memory is around {avg_free:.2f} GB, and swap usage is significant under peak workload.")
                else:
                    add_normal(doc, f"Memory is sufficient for the current workload. At high workload, free memory remains around {min_free:.2f} GB and swap usage peak is {max_swap_pct:.2f}%. No immediate memory capacity upgrade is required.")
            
            mem_trend = os_perf.get("mem_trend", [])
            pgpgin_vals = [r["pgpgin"] for r in mem_trend]
            pgpgout_vals = [r["pgpgout"] for r in mem_trend]
            
            add_small_table(doc, ["Metric", "Min (Pages/s)", "Max (Pages/s)", "Avg (Pages/s)"], [
                {"Metric": "Page In (pgpgin/s)", "Min (Pages/s)": f"{min(pgpgin_vals):.2f}", "Max (Pages/s)": f"{max(pgpgin_vals):.2f}", "Avg (Pages/s)": f"{(sum(pgpgin_vals)/len(pgpgin_vals)):.2f}"},
                {"Metric": "Page Out (pgpgout/s)", "Min (Pages/s)": f"{min(pgpgout_vals):.2f}", "Max (Pages/s)": f"{max(pgpgout_vals):.2f}", "Avg (Pages/s)": f"{(sum(pgpgout_vals)/len(pgpgout_vals)):.2f}"},
            ])
            
        if chart_paths.get("cpu_wait_io") or chart_paths.get("disk_service") or disk_path:
            add_heading2(doc, "F.3 Disk Performance")
            add_heading3(doc, "CPU Wait I/O")
            if chart_paths.get("cpu_wait_io"):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture(chart_paths["cpu_wait_io"], width=Inches(6.8))

            add_heading3(doc, "Disk Service Time")
            if chart_paths.get("disk_service"):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture(chart_paths["disk_service"], width=Inches(6.8))
            
            disk_trend = os_perf.get("disk_trend", [])
            disk_samples = _max_disk_samples_by_time(os_perf.get("disk_samples", []))
            svctm_vals = [r["svctm"] for r in disk_samples] or [r["svctm"] for r in disk_trend]
            await_vals = [r["await"] for r in disk_samples] or [r["await"] for r in disk_trend]
            iowait_vals = [r["iowait"] for r in os_perf.get("cpu_samples", [])]
            
            max_await = max(await_vals, default=0)
            max_iowait = max(iowait_vals, default=0)
            if max_iowait > 30 or max_await > 100:
                add_normal(doc, f"As the result from above graph, CPU Wait I/O occasionally peaks at high values. This suggests that system performance may still be impacted by high I/O demand or insufficient memory during peak workload.")
            else:
                add_normal(doc, "As the result from the graph, the disk service time and I/O wait time are generally low, indicating that disk devices are responding quickly. There are no continuous periods of high disk service time.")
            add_normal(doc, "In summary, disk performance is sufficient, but overall system performance may be degraded during periods of heavy memory pressure. Monitoring memory expansion may help reduce high I/O wait times.")
            
            add_small_table(doc, ["Metric", "Min (ms)", "Max (ms)", "Avg (ms)"], [
                {"Metric": "Service Time (svctm)", "Min (ms)": f"{min(svctm_vals):.2f}", "Max (ms)": f"{max(svctm_vals):.2f}", "Avg (ms)": f"{(sum(svctm_vals)/len(svctm_vals)):.2f}"},
                {"Metric": "I/O Wait (await)", "Min (ms)": f"{min(await_vals):.2f}", "Max (ms)": f"{max(await_vals):.2f}", "Avg (ms)": f"{(sum(await_vals)/len(await_vals)):.2f}"},
            ])
            
    else:
        add_normal(doc, "No OS performance summary data found.")


# =========================================================
# Main entry
# =========================================================

def create_docx_report(data: dict, output_docx: Path, source_zip: Path, output_dir: Path = None):
    output_docx = Path(output_docx)
    source_zip = Path(source_zip)
    output_dir = Path(output_dir) if output_dir else output_docx.parent

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    doc = Document()
    set_doc_layout(doc)

    add_table_of_contents(doc)
    add_defect_classification(doc)
    add_suggestion_summary(doc, data)
    add_general_project_information(doc, data, source_zip)
    add_oracle_minimum_requirement(doc, data)
    add_system_checklist(doc, data)
    add_database_information(doc, data)
    add_database_growth_rate(doc, data, output_dir)
    add_performance_analysis(doc, data, output_dir)
    add_tablespace_free_space(doc, data)
    add_default_tablespace(doc, data)
    add_database_registry(doc, data)
    add_appendix_invalid_objects(doc, data)
    add_appendix_defaults(doc, data, output_dir)

    doc.save(output_docx)
