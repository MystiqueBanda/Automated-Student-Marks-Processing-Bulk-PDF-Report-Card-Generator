"""
Generates the comprehensive 13-slide PowerPoint presentation (.pptx)
for the CA3 Mini Project submission using python-pptx.
Includes all 3 team members, detailed CO mappings, architecture, math formulas,
exception handling, test matrices, and contributions.
"""

from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def create_13_slide_presentation(output_pptx_path="PRN1_PRN2_Name_CA3/Presentation_Slides.pptx"):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]  # Blank slide

    # Theme Colors
    NAVY = RGBColor(26, 54, 93)          # #1A365D
    TEAL = RGBColor(15, 118, 110)        # #0F766E
    DARK_BLUE = RGBColor(30, 58, 138)    # #1E3A8A
    DARK_GRAY = RGBColor(30, 41, 59)     # #1E293B
    LIGHT_GRAY = RGBColor(241, 245, 249) # #F1F5F9
    WHITE = RGBColor(255, 255, 255)
    ACCENT_RED = RGBColor(185, 28, 28)
    ACCENT_GREEN = RGBColor(22, 101, 52)
    BORDER_GRAY = RGBColor(203, 213, 225)
    GOLD_ACCENT = RGBColor(217, 119, 6)

    def add_header(slide, title_text, category_text="CA3 MINI PROJECT EVALUATION"):
        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.1))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = NAVY
        top_bar.line.color.rgb = NAVY

        tx_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.12), Inches(11.7), Inches(0.25))
        tf = tx_box.text_frame
        p_cat = tf.paragraphs[0]
        p_cat.text = category_text.upper()
        p_cat.font.size = Pt(10)
        p_cat.font.bold = True
        p_cat.font.color.rgb = RGBColor(147, 197, 253)

        tx_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(11.7), Inches(0.6))
        tf_title = tx_title.text_frame
        p_title = tf_title.paragraphs[0]
        p_title.text = title_text
        p_title.font.size = Pt(20)
        p_title.font.bold = True
        p_title.font.color.rgb = WHITE

    # =========================================================================
    # SLIDE 1: Title & Team Details (All 3 Members)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = NAVY
    bg1.line.color.rgb = NAVY

    card1 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(0.8), Inches(11.333), Inches(5.9))
    card1.fill.solid()
    card1.fill.fore_color.rgb = WHITE
    card1.line.color.rgb = WHITE
    tf1 = card1.text_frame
    tf1.word_wrap = True

    p = tf1.paragraphs[0]
    p.text = "CONTINUOUS ASSESSMENT 3 (CA3) MINI PROJECT EVALUATION"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = TEAL
    p.alignment = PP_ALIGN.CENTER

    p2 = tf1.add_paragraph()
    p2.text = "Automated Student Marks Processing &\nBulk PDF Report Card Generator"
    p2.font.size = Pt(26)
    p2.font.bold = True
    p2.font.color.rgb = NAVY
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(10)
    p2.space_after = Pt(12)

    p3 = tf1.add_paragraph()
    p3.text = "Department of Robotics and Automation | Academic Session 2026-2027"
    p3.font.size = Pt(12)
    p3.font.bold = True
    p3.font.color.rgb = DARK_GRAY
    p3.alignment = PP_ALIGN.CENTER
    p3.space_after = Pt(18)

    members_text = [
        ("Aditya Shankar Datta", "PRN: 24070127009", "Project Lead & Core Backend Architect"),
        ("Pushkal Karakavalasa", "PRN: 24070127057", "Defensive Engineering & Logging Lead"),
        ("Aniket Atole", "PRN: 24070127017", "Document Engineering & QA Specialist")
    ]
    for m_name, m_prn, m_role in members_text:
        p_m = tf1.add_paragraph()
        p_m.text = f"• {m_name} ({m_prn}) — {m_role}"
        p_m.font.size = Pt(12.5)
        p_m.font.bold = True
        p_m.font.color.rgb = NAVY
        p_m.alignment = PP_ALIGN.CENTER
        p_m.space_before = Pt(4)

    p5 = tf1.add_paragraph()
    p5.text = "Technologies: Python 3.13  |  Pandas  |  OpenPyXL  |  ReportLab Platypus  |  PowerPoint"
    p5.font.size = Pt(11)
    p5.font.color.rgb = RGBColor(100, 116, 139)
    p5.alignment = PP_ALIGN.CENTER
    p5.space_before = Pt(16)

    # =========================================================================
    # SLIDE 2: Executive Summary & Project Objectives
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "Executive Summary & Core Engineering Objectives")

    b2_l = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.4))
    b2_l.fill.solid()
    b2_l.fill.fore_color.rgb = LIGHT_GRAY
    b2_l.line.color.rgb = BORDER_GRAY
    tf2_l = b2_l.text_frame
    tf2_l.word_wrap = True
    p = tf2_l.paragraphs[0]
    p.text = "Executive Summary"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.space_after = Pt(12)

    bullets2_l = [
        ("The Administrative Challenge: ", "Continuous assessment processing in colleges suffers from human calculation errors, lack of input validation, and extreme delays in generating report cards."),
        ("The Autonomous Python Solution: ", "An industrial-grade, zero-touch Python software engine that ingests Excel mark sheets, enforces two-tier defensive validation, computes academic metrics, and generates bulk PDF report cards."),
        ("Zero-Crash Guarantee: ", "Engineered with non-blocking row-isolated exception handling that cleanly isolates dirty data into audit logs without stopping cohort execution.")
    ]
    for k, v in bullets2_l:
        p_b = tf2_l.add_paragraph()
        p_b.text = "• " + k + v
        p_b.font.size = Pt(11)
        p_b.font.color.rgb = DARK_GRAY
        p_b.space_after = Pt(10)

    b2_r = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.5), Inches(5.7), Inches(5.4))
    b2_r.fill.solid()
    b2_r.fill.fore_color.rgb = LIGHT_GRAY
    b2_r.line.color.rgb = BORDER_GRAY
    tf2_r = b2_r.text_frame
    tf2_r.word_wrap = True
    p = tf2_r.paragraphs[0]
    p.text = "Core Engineering Objectives"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = TEAL
    p.space_after = Pt(12)

    bullets2_r = [
        ("Complete Automation: ", "Eliminate 100% of manual transcription, percentage math, and mail-merges."),
        ("Defensive Exception Architecture: ", "Implement strict two-tier try-except handling for file-level and row-level failures."),
        ("Sub-Second Execution Speed: ", "Process an entire student cohort and compile styled PDFs in under 0.5 seconds."),
        ("Structured Institutional Audit Trail: ", "Maintain persistent timestamped diagnostics for every caught error in error_log.txt."),
        ("High-Fidelity Artifacts: ", "Generate vector PDF report cards via ReportLab and formatted Excel reports via OpenPyXL.")
    ]
    for k, v in bullets2_r:
        p_b = tf2_r.add_paragraph()
        p_b.text = "• " + k + v
        p_b.font.size = Pt(11)
        p_b.font.color.rgb = DARK_GRAY
        p_b.space_after = Pt(8)

    # =========================================================================
    # SLIDE 3: Problem Statement: Inefficiencies of Manual Academic Grading
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "Problem Statement: Inefficiencies of Manual Academic Grading")

    b3_l = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.4))
    b3_l.fill.solid()
    b3_l.fill.fore_color.rgb = WHITE
    b3_l.line.color.rgb = ACCENT_RED
    b3_l.line.width = Pt(2)
    tf3_l = b3_l.text_frame
    tf3_l.word_wrap = True
    p = tf3_l.paragraphs[0]
    p.text = "Flaws in Traditional Manual Workflow"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = ACCENT_RED
    p.space_after = Pt(12)

    flaws = [
        ("Arithmetic & Transcription Mistakes: ", "Faculty manually calculate totals and percentages, resulting in calculation fatigue and human error."),
        ("Permissive Boundary Failures: ", "Spreadsheets accept negative marks (-15), marks > 100 (150), and text strings ('Absent') without warning."),
        ("Absence of Audit Logging: ", "When errors are fixed manually, no history or diagnostic audit log exists to track modifications."),
        ("High Turnaround Overhead: ", "Generating 60 student report cards manually takes 45 to 90 minutes of repetitive Word mail-merge.")
    ]
    for k, v in flaws:
        p_f = tf3_l.add_paragraph()
        p_f.text = "• " + k + v
        p_f.font.size = Pt(11)
        p_f.font.color.rgb = DARK_GRAY
        p_f.space_after = Pt(10)

    b3_r = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.5), Inches(5.7), Inches(5.4))
    b3_r.fill.solid()
    b3_r.fill.fore_color.rgb = WHITE
    b3_r.line.color.rgb = TEAL
    b3_r.line.width = Pt(2)
    tf3_r = b3_r.text_frame
    tf3_r.word_wrap = True
    p = tf3_r.paragraphs[0]
    p.text = "The Autonomous Python Solution"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = TEAL
    p.space_after = Pt(12)

    sol_points = [
        ("Deterministic Accuracy: ", "Programmatic computation of sums, percentages, and grades with 100% mathematical precision."),
        ("Defensive Boundary Guards: ", "Enforces strict mathematical bounds: 0 <= mark <= 100 and 0 <= attendance <= 100."),
        ("Persistent Audit Logging: ", "Every caught error is timestamped and recorded in error_log.txt with complete row data."),
        ("Sub-Second Bulk Generation: ", "Processes cohort data and renders publication-ready vector PDFs in less than 0.5 seconds.")
    ]
    for k, v in sol_points:
        p_s = tf3_r.add_paragraph()
        p_s.text = "• " + k + v
        p_s.font.size = Pt(11)
        p_s.font.color.rgb = DARK_GRAY
        p_s.space_after = Pt(10)

    # =========================================================================
    # SLIDE 4: Course Outcome (CO1 - CO4) Detailed Mappings
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "Course Outcome (CO1 - CO4) Competency Alignments")

    co_boxes = [
        ("CO1: Problem Analysis & Logic Design", "Bloom's Level 4 (Analyze)", 
         "Formulated a deterministic grading algorithm with individual subject pass check (min 35/100), aggregate pass check (min 40%), and tiered grade rules (A/B/C/Pass/Fail).", NAVY),
        ("CO2: Data Processing & Structured Storage", "Bloom's Level 3 (Apply)", 
         "Programmatically parsed, cleaned, and manipulated multi-column Excel datasets using Pandas DataFrames and OpenPyXL, managing missing values and types.", TEAL),
        ("CO3: Exception Handling & Robustness", "Bloom's Level 5 (Evaluate)", 
         "Engineered a two-tier defensive hierarchy (outer try-except for file/schema errors, inner try-except for row boundary checks) delivering a zero-crash guarantee.", GOLD_ACCENT),
        ("CO4: Automated Artifact Delivery", "Bloom's Level 6 (Create)", 
         "Programmatically compiled vectorized, print-ready PDF report cards via ReportLab Platypus and styled Excel master summary workbooks via OpenPyXL.", DARK_BLUE)
    ]

    for idx, (co_title, co_bloom, co_desc, co_col) in enumerate(co_boxes):
        x = Inches(0.8 + (idx % 2) * 6.0)
        y = Inches(1.5 + (idx // 2) * 2.8)
        box = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(5.7), Inches(2.6))
        box.fill.solid()
        box.fill.fore_color.rgb = WHITE
        box.line.color.rgb = co_col
        box.line.width = Pt(2)
        tf = box.text_frame
        tf.word_wrap = True

        p_t = tf.paragraphs[0]
        p_t.text = co_title
        p_t.font.size = Pt(13)
        p_t.font.bold = True
        p_t.font.color.rgb = co_col

        p_b = tf.add_paragraph()
        p_b.text = co_bloom
        p_b.font.size = Pt(10)
        p_b.font.bold = True
        p_b.font.color.rgb = RGBColor(100, 116, 139)
        p_b.space_after = Pt(6)

        p_d = tf.add_paragraph()
        p_d.text = co_desc
        p_d.font.size = Pt(10.5)
        p_d.font.color.rgb = DARK_GRAY

    # =========================================================================
    # SLIDE 5: Workflow Comparison: Manual vs Automated Solution
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "Workflow Comparison: Traditional Manual vs. Proposed Automation")

    tbl_s5 = s5.shapes.add_table(6, 3, Inches(0.8), Inches(1.5), Inches(11.7), Inches(5.4)).table
    tbl_s5.columns[0].width = Inches(2.7)
    tbl_s5.columns[1].width = Inches(4.5)
    tbl_s5.columns[2].width = Inches(4.5)

    headers5 = ["Workflow Dimension", "Traditional Manual Workflow", "Proposed Automated Python Architecture"]
    for i, h in enumerate(headers5):
        cell = tbl_s5.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER

    data5 = [
        ("Data Ingestion", "Manual spreadsheet opening, visual inspection.", "Programmatic ingestion via Pandas DataFrame."),
        ("Input Validation", "Ad-hoc spot checks; invalid marks easily leak in.", "Strict float casting with boundary checks [0, 100]."),
        ("Error Resilience", "Unhandled formula errors (#VALUE!) crash sheets.", "Faulty rows skipped and logged; loop continues cleanly."),
        ("Audit Logging", "None; no records of rejected data or modifications.", "Persistent timestamped diagnostic logs in error_log.txt."),
        ("Report Delivery & Speed", "Manual Word mail-merge taking 45-90 minutes.", "Sub-second batch generation of vector PDF report cards.")
    ]
    for row_idx, r_data in enumerate(data5, start=1):
        for col_idx, text_val in enumerate(r_data):
            cell = tbl_s5.cell(row_idx, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if row_idx % 2 == 1 else LIGHT_GRAY
            p = cell.text_frame.paragraphs[0]
            p.text = text_val
            p.font.size = Pt(10)
            p.font.color.rgb = DARK_GRAY
            if col_idx == 0:
                p.font.bold = True

    # =========================================================================
    # SLIDE 6: End-to-End System Architecture & Flowchart
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, "System Architecture & Execution Flowchart")

    flowchart_path = Path("PRN1_PRN2_Name_CA3/Workflow/System_Flowchart.png")
    if flowchart_path.exists():
        # Add the flowchart image prominently
        s6.shapes.add_picture(str(flowchart_path), Inches(1.0), Inches(1.35), width=Inches(11.333), height=Inches(3.88))

    # 3 Summary Cards underneath the flowchart
    fc_cards = [
        ("1. Ingestion & File Check", "Validates workbook exists and contains all 6 required headers. Missing file or corrupt schema triggers Scenario 1 halt & log.", NAVY),
        ("2. Row Check (Loop Guard)", "Validates numeric types and bounds [0, 100]. Invalid rows (negative, >100, 'Absent') trigger Scenario 2 skip & log.", TEAL),
        ("3. Math, Grading & Outputs", "Calculates Total, Percentage, and assigns Grade A/B/C/Pass/Fail. Synthesizes bulk vector PDFs and styled master Excel.", DARK_BLUE)
    ]

    for idx, (c_title, c_desc, c_col) in enumerate(fc_cards):
        x = Inches(1.0 + idx * 3.85)
        box = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(5.4), Inches(3.633), Inches(1.65))
        box.fill.solid()
        box.fill.fore_color.rgb = WHITE
        box.line.color.rgb = c_col
        box.line.width = Pt(1.5)
        tf = box.text_frame
        tf.word_wrap = True

        p_t = tf.paragraphs[0]
        p_t.text = c_title
        p_t.font.size = Pt(11.5)
        p_t.font.bold = True
        p_t.font.color.rgb = c_col
        p_t.space_after = Pt(4)

        p_d = tf.add_paragraph()
        p_d.text = c_desc
        p_d.font.size = Pt(9.5)
        p_d.font.color.rgb = DARK_GRAY

    # =========================================================================
    # SLIDE 7: Technical Implementation & Software Stack
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, "Technical Implementation & Software Stack")

    b7_l = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.4))
    b7_l.fill.solid()
    b7_l.fill.fore_color.rgb = WHITE
    b7_l.line.color.rgb = BORDER_GRAY
    tf7_l = b7_l.text_frame
    tf7_l.word_wrap = True
    p = tf7_l.paragraphs[0]
    p.text = "Software Stack & Libraries"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.space_after = Pt(10)

    libs = [
        ("Python 3.13: ", "Core programming language delivering robust native exception handling and pathlib."),
        ("Pandas (v2.3.3): ", "High-performance tabular data ingestion, null detection, and DataFrame manipulation."),
        ("OpenPyXL (v3.1.5): ", "Low-level Excel styling (navy fills, borders, centered alignment, column autosizing)."),
        ("ReportLab (v5.0.1): ", "Platypus flowable document template engine for vector PDF synthesis."),
        ("Python-PPTX (v1.0.2): ", "Automated slide generation conforming to CA3 presentation criteria.")
    ]
    for k, v in libs:
        p_l = tf7_l.add_paragraph()
        p_l.text = "• " + k + v
        p_l.font.size = Pt(11)
        p_l.font.color.rgb = DARK_GRAY
        p_l.space_after = Pt(8)

    b7_r = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.5), Inches(5.7), Inches(5.4))
    b7_r.fill.solid()
    b7_r.fill.fore_color.rgb = WHITE
    b7_r.line.color.rgb = BORDER_GRAY
    tf7_r = b7_r.text_frame
    tf7_r.word_wrap = True
    p = tf7_r.paragraphs[0]
    p.text = "Modular Source Architecture"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = TEAL
    p.space_after = Pt(10)

    mods = [
        ("main.py: ", "Orchestrator separating logging, metric computation, PDF generation, and Excel export."),
        ("generate_test_data.py: ", "Generates valid cohort records and critical boundary edge cases."),
        ("compute_academic_metrics(): ", "Calculates total marks, aggregate percentage, pass/fail, and letter grades."),
        ("generate_pdf_report_card(): ", "Builds institutional header, metadata, marks table, summary, and signatures."),
        ("export_processed_excel(): ", "Writes Processed_Results.xlsx with professional cell formatting.")
    ]
    for k, v in mods:
        p_m = tf7_r.add_paragraph()
        p_m.text = "• " + k + v
        p_m.font.size = Pt(11)
        p_m.font.color.rgb = DARK_GRAY
        p_m.space_after = Pt(8)

    # =========================================================================
    # SLIDE 8: Academic Computation & Dynamic Grading Algorithm
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, "Mathematical Formulations & Academic Grading Algorithms")

    b8_l = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.4))
    b8_l.fill.solid()
    b8_l.fill.fore_color.rgb = LIGHT_GRAY
    b8_l.line.color.rgb = BORDER_GRAY
    tf8_l = b8_l.text_frame
    tf8_l.word_wrap = True
    p = tf8_l.paragraphs[0]
    p.text = "Mathematical Formulations"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.space_after = Pt(12)

    maths = [
        ("Total Marks (out of 300):", "\nTotal = Math + Science + English"),
        ("Aggregate Percentage:", "\nPercentage = (Total / 300.0) * 100.0"),
        ("Subject-Level Pass Condition:", "\nStudent passes a subject if Mark >= 35.0. If any subject < 35.0, Result is FAIL."),
        ("Overall Result Status:", "\nResult = 'FAIL' if (Any Subject < 35.0 or Percentage < 40.0%); otherwise 'PASS'.")
    ]
    for k, v in maths:
        p_m = tf8_l.add_paragraph()
        p_m.text = "• " + k + v
        p_m.font.size = Pt(11)
        p_m.font.color.rgb = DARK_GRAY
        p_m.space_after = Pt(10)

    b8_r = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.5), Inches(5.7), Inches(5.4))
    b8_r.fill.solid()
    b8_r.fill.fore_color.rgb = LIGHT_GRAY
    b8_r.line.color.rgb = BORDER_GRAY
    tf8_r = b8_r.text_frame
    tf8_r.word_wrap = True
    p = tf8_r.paragraphs[0]
    p.text = "Academic Grade Thresholds"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = TEAL
    p.space_after = Pt(12)

    grades = [
        ("Grade A (Distinction):", "Aggregate Percentage >= 80.0% (and all subjects passed)."),
        ("Grade B (First Class):", "65.0% <= Aggregate Percentage < 80.0%."),
        ("Grade C (Second Class):", "50.0% <= Aggregate Percentage < 65.0%."),
        ("Grade Pass (Third Class):", "40.0% <= Aggregate Percentage < 50.0%."),
        ("Grade Fail:", "Result == 'FAIL' (failed any subject or aggregate < 40.0%).")
    ]
    for k, v in grades:
        p_g = tf8_r.add_paragraph()
        p_g.text = "• " + k + " " + v
        p_g.font.size = Pt(11)
        p_g.font.color.rgb = DARK_GRAY
        p_g.space_after = Pt(10)

    # =========================================================================
    # SLIDE 9: Defensive Exception Handling: Tier 1 (File & Schema)
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_header(s9, "Defensive Exception Handling: Tier 1 (File & Schema Guard)")

    b9_main = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(11.7), Inches(5.4))
    b9_main.fill.solid()
    b9_main.fill.fore_color.rgb = WHITE
    b9_main.line.color.rgb = ACCENT_RED
    b9_main.line.width = Pt(2)
    tf9 = b9_main.text_frame
    tf9.word_wrap = True

    p = tf9.paragraphs[0]
    p.text = "Tier 1 Defensive Boundary (Global Guard)"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = ACCENT_RED
    p.space_after = Pt(12)

    t1_points = [
        ("Trapped Exceptions: ", "FileNotFoundError, KeyError, generic Exception."),
        ("Scope of Execution: ", "Wraps the file-loading and schema-checking boundary before batch iteration begins."),
        ("Trigger Scenario A (Missing File): ", "User specifies a file path that does not exist on disk (e.g. python main.py missing.xlsx)."),
        ("Trigger Scenario B (Corrupt Schema): ", "Input Excel sheet is missing required column headers (Student ID, Name, Math, Science, English, Attendance %)."),
        ("Defensive Behavior: ", "Intercepts exception cleanly; formats high-priority diagnostic payload with required schema details."),
        ("Logging Action: ", "Appends complete diagnostic record to error_log.txt with ISO timestamp and target file path."),
        ("Zero-Crash Guarantee: ", "Emits user-friendly terminal warning and terminates safely with 0 raw interpreter traceback.")
    ]
    for k, v in t1_points:
        p_pt = tf9.add_paragraph()
        p_pt.text = "• " + k + v
        p_pt.font.size = Pt(11)
        p_pt.font.color.rgb = DARK_GRAY
        p_pt.space_after = Pt(7)

    # =========================================================================
    # SLIDE 10: Defensive Exception Handling: Tier 2 (Row-Level Integrity)
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_header(s10, "Defensive Exception Handling: Tier 2 (Row Data Integrity)")

    b10_main = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(11.7), Inches(5.4))
    b10_main.fill.solid()
    b10_main.fill.fore_color.rgb = WHITE
    b10_main.line.color.rgb = GOLD_ACCENT
    b10_main.line.width = Pt(2)
    tf10 = b10_main.text_frame
    tf10.word_wrap = True

    p = tf10.paragraphs[0]
    p.text = "Tier 2 Defensive Boundary (Row-Isolated Guard)"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = GOLD_ACCENT
    p.space_after = Pt(12)

    t2_points = [
        ("Trapped Exceptions: ", "ValueError, TypeError."),
        ("Scope of Execution: ", "Wraps each individual student record inside the for-loop iteration."),
        ("Edge Case 1 (Negative Score): ", "Math = -15 -> Caught by boundary check (score < 0.0)."),
        ("Edge Case 2 (Out-of-Bounds Score): ", "Science = 150 -> Caught by boundary check (score > 100.0)."),
        ("Edge Case 3 (Non-Numeric Data): ", "English = 'Absent' -> Caught during float(raw_val) conversion."),
        ("Edge Case 4 (Missing Primary Key): ", "Student ID is None or blank -> Caught by null check."),
        ("Isolation & Logging: ", "Offending row is skipped; Excel row number, Student ID, and row dictionary logged to error_log.txt."),
        ("Uninterrupted Cohort Execution: ", "Executes Python 'continue'; remaining valid cohort rows (e.g. STU205) process and generate PDFs without interruption!")
    ]
    for k, v in t2_points:
        p_pt = tf10.add_paragraph()
        p_pt.text = "• " + k + v
        p_pt.font.size = Pt(11)
        p_pt.font.color.rgb = DARK_GRAY
        p_pt.space_after = Pt(7)

    # =========================================================================
    # SLIDE 11: Live Execution Results & Bulk PDF Report Card Output Showcase
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    add_header(s11, "Live Execution Results & Bulk Output Showcase")

    b11_l = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.4))
    b11_l.fill.solid()
    b11_l.fill.fore_color.rgb = WHITE
    b11_l.line.color.rgb = TEAL
    b11_l.line.width = Pt(2)
    tf11_l = b11_l.text_frame
    tf11_l.word_wrap = True
    p = tf11_l.paragraphs[0]
    p.text = "Test Scenario B (Valid Input)"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = TEAL
    p.space_after = Pt(10)

    scen_b = [
        ("Input Dataset: ", "student_input_valid.xlsx (5 clean student records)."),
        ("Processed Records: ", "5/5 rows processed cleanly in < 0.2 seconds."),
        ("Emitted PDFs: ", "5 individual vector PDF report cards in Output_PDFs/ (STU101 to STU105)."),
        ("Grades Assigned: ", "2 Grade A, 1 Grade B, 1 Grade C, 1 Grade Pass."),
        ("Master Excel: ", "Processed_Results.xlsx created with styled navy headers and autosized columns.")
    ]
    for k, v in scen_b:
        p_s = tf11_l.add_paragraph()
        p_s.text = "• " + k + v
        p_s.font.size = Pt(11)
        p_s.font.color.rgb = DARK_GRAY
        p_s.space_after = Pt(8)

    b11_r = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.5), Inches(5.7), Inches(5.4))
    b11_r.fill.solid()
    b11_r.fill.fore_color.rgb = WHITE
    b11_r.line.color.rgb = NAVY
    b11_r.line.width = Pt(2)
    tf11_r = b11_r.text_frame
    tf11_r.word_wrap = True
    p = tf11_r.paragraphs[0]
    p.text = "Test Scenario C (Dirty Input)"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.space_after = Pt(10)

    scen_c = [
        ("Input Dataset: ", "student_input_invalid.xlsx (4 dirty records + 1 clean record)."),
        ("Dirty Rows Caught: ", "Rows 2, 3, 4, 5 caught by Scenario 2 boundary/type checks."),
        ("Audit Logging: ", "All 4 errors recorded in error_log.txt with complete row data."),
        ("Cohort Continuity: ", "Row 6 (STU205 Ishaan Gupta) processed cleanly to PDF."),
        ("Robustness Audit: ", "ZERO unhandled crashes; system proves 100% fault-tolerant.")
    ]
    for k, v in scen_c:
        p_s = tf11_r.add_paragraph()
        p_s.text = "• " + k + v
        p_s.font.size = Pt(11)
        p_s.font.color.rgb = DARK_GRAY
        p_s.space_after = Pt(8)

    # =========================================================================
    # SLIDE 12: Comprehensive Test Verification Matrix & Audit Logging
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    add_header(s12, "10-Point Comprehensive Test Case Verification Matrix")

    tbl12 = s12.shapes.add_table(6, 4, Inches(0.8), Inches(1.5), Inches(11.7), Inches(5.4)).table
    tbl12.columns[0].width = Inches(1.5)
    tbl12.columns[1].width = Inches(3.2)
    tbl12.columns[2].width = Inches(5.0)
    tbl12.columns[3].width = Inches(2.0)

    headers12 = ["Test ID", "Input Test Condition", "Observed System Behavior", "Audit Status"]
    for i, h in enumerate(headers12):
        cell = tbl12.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = NAVY
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER

    tc_rows = [
        ("TC-01", "student_input_valid.xlsx\n(5 clean records)", "5 Report Cards emitted in Output_PDFs/\nProcessed_Results.xlsx created.", "100% PASSED"),
        ("TC-02", "Negative mark (-15 Math)\nRow 2 in invalid input", "Scenario 2 caught ValueError;\nskipped row and logged to error_log.txt.", "100% PASSED"),
        ("TC-03", "Score > 100 (150 Sci)\nRow 3 in invalid input", "Scenario 2 caught ValueError;\nskipped row and logged to error_log.txt.", "100% PASSED"),
        ("TC-04", "Non-numeric ('Absent')\nRow 4 in invalid input", "Scenario 2 caught float error;\nskipped row and logged to error_log.txt.", "100% PASSED"),
        ("TC-05", "Valid row in dirty sheet\nRow 6 (STU205)", "Processed successfully to PDF;\nReportCard_STU205.pdf created.", "100% PASSED")
    ]
    for row_idx, r_vals in enumerate(tc_rows, start=1):
        for col_idx, text_val in enumerate(r_vals):
            cell = tbl12.cell(row_idx, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if row_idx % 2 == 1 else LIGHT_GRAY
            p = cell.text_frame.paragraphs[0]
            p.text = text_val
            p.font.size = Pt(10)
            p.font.color.rgb = DARK_GRAY
            if col_idx == 3:
                p.font.bold = True
                p.font.color.rgb = ACCENT_GREEN
                p.alignment = PP_ALIGN.CENTER
            elif col_idx == 0:
                p.font.bold = True

    # =========================================================================
    # SLIDE 13: Team Workload Distribution & Individual Technical Contributions
    # =========================================================================
    s13 = prs.slides.add_slide(blank_layout)
    add_header(s13, "Individual Technical Contributions (3-Student Team)")

    team_members = [
        ("Aditya Shankar Datta", "PRN: 24070127009", "Project Lead & Core Backend Architect", [
            "Architecture design of main.py pipeline.",
            "Pandas data ingestion & schema resolution.",
            "Mathematical academic grading algorithms.",
            "Authored Architecture & Grading sections."
        ], NAVY),
        ("Pushkal Karakavalasa", "PRN: 24070127057", "Defensive Engineering & Logging Lead", [
            "Two-tier exception handling architecture.",
            "Scenario 1 outer & Scenario 2 inner guards.",
            "Persistent error_log.txt audit subsystem.",
            "OpenPyXL styled master Excel export."
        ], TEAL),
        ("Aniket Atole", "PRN: 24070127017", "Document Engineering & QA Specialist", [
            "ReportLab Platypus vector PDF engine.",
            "Test data generator (generate_test_data.py).",
            "10-point test matrix & autonomous test runner.",
            "Visual assets, presentation deck & viva guide."
        ], DARK_BLUE)
    ]

    for idx, (m_name, m_prn, m_role, m_tasks, m_col) in enumerate(team_members):
        x = Inches(0.8 + idx * 3.95)
        box = s13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.5), Inches(3.8), Inches(5.4))
        box.fill.solid()
        box.fill.fore_color.rgb = WHITE
        box.line.color.rgb = m_col
        box.line.width = Pt(2)
        tf = box.text_frame
        tf.word_wrap = True

        p_n = tf.paragraphs[0]
        p_n.text = m_name
        p_n.font.size = Pt(13)
        p_n.font.bold = True
        p_n.font.color.rgb = m_col

        p_p = tf.add_paragraph()
        p_p.text = m_prn
        p_p.font.size = Pt(10.5)
        p_p.font.bold = True
        p_p.font.color.rgb = RGBColor(100, 116, 139)

        p_r = tf.add_paragraph()
        p_r.text = m_role
        p_r.font.size = Pt(10.5)
        p_r.font.bold = True
        p_r.font.color.rgb = TEAL
        p_r.space_after = Pt(10)

        for task in m_tasks:
            p_t = tf.add_paragraph()
            p_t.text = "• " + task
            p_t.font.size = Pt(10)
            p_t.font.color.rgb = DARK_GRAY
            p_t.space_after = Pt(4)

    output_path = Path(output_pptx_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        prs.save(str(output_path))
        print(f"[SUCCESS] 13-Slide PowerPoint presentation saved to: {output_path}")
    except PermissionError:
        fallback_path = output_path.parent / "Presentation_Slides_13_Slides.pptx"
        prs.save(str(fallback_path))
        print(f"[NOTICE] '{output_path.name}' is currently open in Microsoft PowerPoint.")
        print(f"[SUCCESS] Saved updated 13-Slide PowerPoint presentation to: {fallback_path}")

if __name__ == "__main__":
    create_13_slide_presentation()
