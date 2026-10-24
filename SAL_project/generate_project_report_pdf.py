"""
Compiles Project_Report_Draft.txt into a publication-quality Project_Report.pdf
using ReportLab Platypus.
"""

from pathlib import Path
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, PageBreak, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY

def compile_project_report(output_pdf_path="PRN1_PRN2_Name_CA3/Project_Report.pdf"):
    pdf_path = Path(output_pdf_path)
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=A4,
        leftMargin=40,
        rightMargin=40,
        topMargin=40,
        bottomMargin=40
    )

    styles = getSampleStyleSheet()

    # Custom styles
    title_style = ParagraphStyle(
        "DocTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=20,
        leading=24,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#1E3A8A")
    )

    sub_title_style = ParagraphStyle(
        "DocSubTitle",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=12,
        leading=16,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#475569")
    )

    h1_style = ParagraphStyle(
        "Heading1_Custom",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=14,
        leading=18,
        textColor=colors.HexColor("#1E40AF"),
        spaceBefore=14,
        spaceAfter=6
    )

    h2_style = ParagraphStyle(
        "Heading2_Custom",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=15,
        textColor=colors.HexColor("#0F766E"),
        spaceBefore=8,
        spaceAfter=4
    )

    body_style = ParagraphStyle(
        "Body_Custom",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9.5,
        leading=13.5,
        alignment=TA_JUSTIFY,
        textColor=colors.HexColor("#1E293B"),
        spaceAfter=6
    )

    bullet_style = ParagraphStyle(
        "Bullet_Custom",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9.5,
        leading=13.5,
        alignment=TA_LEFT,
        textColor=colors.HexColor("#1E293B"),
        leftIndent=15,
        spaceAfter=4
    )

    code_style = ParagraphStyle(
        "Code_Custom",
        parent=styles["Normal"],
        fontName="Courier",
        fontSize=8.5,
        leading=11,
        alignment=TA_LEFT,
        textColor=colors.HexColor("#0F172A")
    )

    th_style = ParagraphStyle(
        "TH_Custom",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8.5,
        leading=11,
        alignment=TA_CENTER,
        textColor=colors.white
    )

    td_style = ParagraphStyle(
        "TD_Custom",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=10.5,
        alignment=TA_LEFT,
        textColor=colors.HexColor("#1E293B")
    )

    td_center = ParagraphStyle(
        "TDC_Custom",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8,
        leading=10.5,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#1E293B")
    )

    story = []

    # Title & Subtitle Banner
    story.append(Paragraph("CONTINUOUS ASSESSMENT 3 (CA3) MINI PROJECT REPORT", title_style))
    story.append(Spacer(1, 4))
    story.append(Paragraph("Automated Student Marks Processing & Bulk PDF Report Card Generator", sub_title_style))
    story.append(Spacer(1, 8))
    story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor("#1E3A8A"), spaceAfter=12))

    # Section 1: Team & Metadata
    story.append(Paragraph("1. Project Title & Team Details", h1_style))
    meta_data = [
        [Paragraph("<b>Course Module:</b>", td_style), Paragraph("Continuous Assessment 3 (CA3) - Engineering Software Automation", td_style)],
        [Paragraph("<b>Group Size:</b>", td_style), Paragraph("2 Students (Student 1 & Student 2)", td_style)],
        [Paragraph("<b>Student 1 Contribution:</b>", td_style), Paragraph("Core Automation Architecture, Pandas Ingestion Engine & Exception Handling (Scenarios 1 & 2)", td_style)],
        [Paragraph("<b>Student 2 Contribution:</b>", td_style), Paragraph("ReportLab PDF Layout Engine, Test Data Engineering (Valid/Invalid) & Documentation", td_style)],
        [Paragraph("<b>Tech Stack:</b>", td_style), Paragraph("Python 3.13, Pandas, OpenPyXL, ReportLab Platypus, Python-PPTX", td_style)],
    ]
    meta_table = Table(meta_data, colWidths=[140, 375])
    meta_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor("#CBD5E1")),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 10))

    # Section 2: CO Mappings
    story.append(Paragraph("2. Course Outcome (CO) Mappings", h1_style))
    story.append(Paragraph("<b>• CO1 (Problem Analysis & Logic Design):</b> Developed a deterministic batch automation pipeline that validates schema structure, calculates aggregates (Total, Percentage, Grade), and assigns academic classification.", bullet_style))
    story.append(Paragraph("<b>• CO2 (Data Ingestion & Processing):</b> Leveraged pandas and openpyxl to manipulate structured tabular Excel workbooks, dynamically extracting student marks matrices.", bullet_style))
    story.append(Paragraph("<b>• CO3 (Exception Handling & Robust Automation):</b> Implemented a production-grade 2-tier exception architecture: outer handling for missing files/columns (Scenario 1) and inner loop handling for data corruption/boundary violations (Scenario 2).", bullet_style))
    story.append(Paragraph("<b>• CO4 (Artifact Generation & Delivery):</b> Programmatically constructed publication-ready, styled PDF report cards with ReportLab and consolidated Excel performance sheets.", bullet_style))
    story.append(Spacer(1, 10))

    # Section 3: Problem Statement
    story.append(Paragraph("3. Problem Statement & Motivation", h1_style))
    story.append(Paragraph("Academic examination boards frequently encounter significant delays and transcription errors when manually processing student mark sheets. Typical manual challenges include human calculation mistakes in percentage computation, failure to catch negative or >100 out-of-range marks, absence of audit logging when dirty records are found, and labor-intensive manual generation of student report cards. This project automates the entire marks lifecycle, guaranteeing zero-error calculations, resilient error isolation, and bulk PDF synthesis in seconds.", body_style))
    story.append(Spacer(1, 10))

    # Section 4: Manual vs Automated Workflow
    story.append(Paragraph("4. Existing (Manual) vs. Proposed (Automated) Workflow", h1_style))
    cmp_data = [
        [Paragraph("Workflow Dimension", th_style), Paragraph("Traditional Manual Workflow", th_style), Paragraph("Proposed Automated System", th_style)],
        [Paragraph("File Ingestion", td_style), Paragraph("Manual opening of spreadsheets; visual verification", td_style), Paragraph("Automated parsing via Pandas & OpenPyXL", td_style)],
        [Paragraph("Input Validation", td_style), Paragraph("Ad-hoc spot-checks; dirty data often leaks in", td_style), Paragraph("Strict boundary [0, 100] & type checking via try-except", td_style)],
        [Paragraph("Error Handling", td_style), Paragraph("Unhandled sheet formula errors or crashes", td_style), Paragraph("Faulty rows skipped and logged to error_log.txt", td_style)],
        [Paragraph("Output Delivery", td_style), Paragraph("Manual typing/mail-merging report card PDFs", td_style), Paragraph("Sub-second bulk PDF generation with custom styling", td_style)],
        [Paragraph("Audit Trail", td_style), Paragraph("None; no records of skipped or invalid rows", td_style), Paragraph("Timestamped diagnostic records in error_log.txt", td_style)]
    ]
    cmp_table = Table(cmp_data, colWidths=[110, 200, 205])
    cmp_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E3A8A")),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ('BACKGROUND', (0, 1), (-1, 1), colors.white),
        ('BACKGROUND', (0, 2), (-1, 2), colors.HexColor("#F8FAFC")),
        ('BACKGROUND', (0, 3), (-1, 3), colors.white),
        ('BACKGROUND', (0, 4), (-1, 4), colors.HexColor("#F8FAFC")),
        ('BACKGROUND', (0, 5), (-1, 5), colors.white),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    story.append(cmp_table)
    story.append(Spacer(1, 12))

    # Section 5: Exception Handling Architecture
    story.append(Paragraph("5. Exception Handling & Boundary Validation Logic", h1_style))
    story.append(Paragraph("The software implements exactly two rigorously tested exception scenarios:", body_style))
    story.append(Paragraph("<b>Scenario 1 (File System & Schema Exception):</b> Guarded by an outer try-except block handling <code>FileNotFoundError</code> and <code>KeyError</code>. If the input Excel file is missing or required columns (Student ID, Name, Math, Science, English, Attendance %) are corrupted, the system logs a high-priority halt message to <code>error_log.txt</code> and exits safely.", bullet_style))
    story.append(Paragraph("<b>Scenario 2 (Row Data Integrity Exception):</b> Guarded by an inner try-except block handling <code>ValueError</code> and <code>TypeError</code>. Catches negative marks (< 0), marks exceeding 100, string values (e.g., 'Absent'), and missing Student IDs. The invalid row is isolated, recorded with full metadata in <code>error_log.txt</code>, and remaining rows proceed smoothly.", bullet_style))
    story.append(Spacer(1, 10))

    # Section 6: Verification & Test Results
    story.append(Paragraph("6. Execution Results & Verification Matrix", h1_style))
    test_data = [
        [Paragraph("Test Case", th_style), Paragraph("Input Condition", th_style), Paragraph("Expected Result", th_style), Paragraph("Observed Status", th_style)],
        [Paragraph("TC-01", td_center), Paragraph("5 valid rows (student_input_valid.xlsx)", td_style), Paragraph("5 PDFs emitted, Processed_Results.xlsx created", td_style), Paragraph("<font color='#059669'><b>PASSED</b></font>", td_center)],
        [Paragraph("TC-02", td_center), Paragraph("Row 2: Math = -15 (negative boundary)", td_style), Paragraph("Caught ValueError, skipped row, logged to file", td_style), Paragraph("<font color='#059669'><b>PASSED</b></font>", td_center)],
        [Paragraph("TC-03", td_center), Paragraph("Row 3: Science = 150 (upper boundary)", td_style), Paragraph("Caught ValueError, skipped row, logged to file", td_style), Paragraph("<font color='#059669'><b>PASSED</b></font>", td_center)],
        [Paragraph("TC-04", td_center), Paragraph("Row 4: English = 'Absent' (type error)", td_style), Paragraph("Caught ValueError, skipped row, logged to file", td_style), Paragraph("<font color='#059669'><b>PASSED</b></font>", td_center)],
        [Paragraph("TC-05", td_center), Paragraph("Row 5: Missing Student ID", td_style), Paragraph("Caught ValueError, skipped row, logged to file", td_style), Paragraph("<font color='#059669'><b>PASSED</b></font>", td_center)],
        [Paragraph("TC-06", td_center), Paragraph("Row 6: STU205 (valid row in dirty sheet)", td_style), Paragraph("Generated ReportCard_STU205.pdf successfully", td_style), Paragraph("<font color='#059669'><b>PASSED</b></font>", td_center)],
        [Paragraph("TC-07", td_center), Paragraph("Missing input Excel file", td_style), Paragraph("Scenario 1 caught FileNotFoundError & logged", td_style), Paragraph("<font color='#059669'><b>PASSED</b></font>", td_center)],
    ]
    test_table = Table(test_data, colWidths=[55, 175, 205, 80])
    test_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0F766E")),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor("#F8FAFC")),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(test_table)
    story.append(Spacer(1, 15))

    # Section 7: Conclusion
    story.append(Paragraph("7. Conclusion", h1_style))
    story.append(Paragraph("The Automated Student Marks Processing System delivers an end-to-end robust solution satisfying all CA3 project constraints. It eliminates manual computation overhead, ensures fault-tolerant operation with comprehensive exception handling, and creates official academic artifacts ready for institutional deployment.", body_style))

    doc.build(story)
    print(f"[SUCCESS] Academic Project Report PDF generated at: {pdf_path}")

if __name__ == "__main__":
    compile_project_report()
