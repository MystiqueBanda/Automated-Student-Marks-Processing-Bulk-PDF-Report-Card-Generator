#!/usr/bin/env python3
"""
========================================================================================
Project Title   : Automated Student Marks Processing & Bulk PDF Report Card Generator
Course/Module   : Continuous Assessment 3 (CA3) Mini Project
Authors         : 1. Aditya Shankar Datta (PRN: 24070127009) - Lead & Core Architecture
                  2. Pushkal Karakavalasa (PRN: 24070127057) - Exception Handling & Audit
                  3. Aniket Atole         (PRN: 24070127017) - PDF Engine & QA Specialist
Technology Stack: Python 3, Pandas, OpenPyXL, ReportLab
========================================================================================
"""

import sys
import os
from pathlib import Path
from datetime import datetime
import pandas as pd
import numpy as np

# ReportLab components for programmatic PDF generation
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT


# ======================================================================================
# CONSTANTS & CONFIGURATION
# ======================================================================================
REQUIRED_COLUMNS = ["Student ID", "Name", "Math", "Science", "English", "Attendance %"]
PASSING_MARK_PER_SUBJECT = 35.0
MAX_MARK_PER_SUBJECT = 100.0
TOTAL_MAX_MARKS = 300.0


# ======================================================================================
# LOGGING UTILITY
# ======================================================================================
def log_error(log_file_path: Path, scenario_name: str, message: str, context: dict = None):
    """
    Appends structured error records to error_log.txt.
    """
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] [{scenario_name}]\n  Message: {message}\n"
    if context:
        for k, v in context.items():
            log_entry += f"  {k}: {v}\n"
    log_entry += "-" * 70 + "\n"

    try:
        with open(log_file_path, "a", encoding="utf-8") as f:
            f.write(log_entry)
    except Exception as e:
        print(f"[FATAL LOGGING ERROR] Could not write to log file {log_file_path}: {e}")


# ======================================================================================
# ACADEMIC LOGIC FUNCTIONS
# ======================================================================================
def compute_academic_metrics(math: float, science: float, english: float):
    """
    Computes total marks, percentage, grade, and overall pass/fail status.
    Grade Rules:
      - Fail: Any subject < 35 or Total Percentage < 40%
      - Grade A: Percentage >= 80%
      - Grade B: Percentage >= 65% and < 80%
      - Grade C: Percentage >= 50% and < 65%
      - Grade Pass: Percentage >= 40% and < 50%
    """
    total = math + science + english
    percentage = round((total / TOTAL_MAX_MARKS) * 100.0, 2)

    has_failed_subject = (
        math < PASSING_MARK_PER_SUBJECT or
        science < PASSING_MARK_PER_SUBJECT or
        english < PASSING_MARK_PER_SUBJECT
    )

    if has_failed_subject or percentage < 40.0:
        grade = "Fail"
        result = "FAIL"
    elif percentage >= 80.0:
        grade = "A"
        result = "PASS"
    elif percentage >= 65.0:
        grade = "B"
        result = "PASS"
    elif percentage >= 50.0:
        grade = "C"
        result = "PASS"
    else:
        grade = "Pass"
        result = "PASS"

    return total, percentage, grade, result


# ======================================================================================
# PDF REPORT CARD GENERATOR
# ======================================================================================
def generate_pdf_report_card(student_record: dict, output_filepath: Path):
    """
    Builds a high-quality, professional academic report card using ReportLab Platypus.
    """
    doc = SimpleDocTemplate(
        str(output_filepath),
        pagesize=A4,
        rightMargin=36,
        leftMargin=36,
        topMargin=36,
        bottomMargin=36
    )

    styles = getSampleStyleSheet()
    
    # Custom styles
    title_style = ParagraphStyle(
        "InstTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=18,
        leading=22,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#1A365D")
    )
    
    sub_title_style = ParagraphStyle(
        "SubTitle",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10,
        leading=14,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#4A5568")
    )

    report_heading_style = ParagraphStyle(
        "ReportHeading",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=13,
        leading=16,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#2B6CB0")
    )

    meta_label_style = ParagraphStyle(
        "MetaLabel",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#2D3748")
    )

    meta_val_style = ParagraphStyle(
        "MetaValue",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#1A202C")
    )

    table_header_style = ParagraphStyle(
        "TableHeader",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9,
        alignment=TA_CENTER,
        textColor=colors.white
    )

    table_cell_style = ParagraphStyle(
        "TableCell",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        alignment=TA_CENTER,
        textColor=colors.HexColor("#2D3748")
    )

    story = []

    # 1. Header Banner
    story.append(Paragraph("APEX INSTITUTE OF TECHNOLOGY & HIGHER LEARNING", title_style))
    story.append(Paragraph("Department of Robotics and Automation | Academic Year 2026-2027", sub_title_style))
    story.append(Spacer(1, 10))
    story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor("#1A365D"), spaceAfter=10))
    story.append(Paragraph("OFFICIAL STUDENT PERFORMANCE REPORT CARD (CA3)", report_heading_style))
    story.append(Spacer(1, 12))

    # 2. Student Info Box
    student_info = [
        [
            Paragraph("<b>Student ID:</b>", meta_label_style),
            Paragraph(str(student_record["Student ID"]), meta_val_style),
            Paragraph("<b>Evaluation Cycle:</b>", meta_label_style),
            Paragraph("Continuous Assessment 3", meta_val_style)
        ],
        [
            Paragraph("<b>Student Name:</b>", meta_label_style),
            Paragraph(str(student_record["Name"]), meta_val_style),
            Paragraph("<b>Attendance:</b>", meta_label_style),
            Paragraph(f"{student_record['Attendance %']:.1f}%", meta_val_style)
        ],
        [
            Paragraph("<b>Generated On:</b>", meta_label_style),
            Paragraph(datetime.now().strftime("%d-%b-%Y %H:%M"), meta_val_style),
            Paragraph("<b>Status:</b>", meta_label_style),
            Paragraph(
                f"<font color='{'#276749' if student_record['Result'] == 'PASS' else '#9B2C2C'}'><b>{student_record['Result']}</b></font>",
                meta_val_style
            )
        ]
    ]

    info_table = Table(student_info, colWidths=[90, 170, 110, 150])
    info_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#F7FAFC")),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor("#CBD5E0")),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 15))

    # 3. Marks Subject Breakdown Table
    math_status = "PASS" if student_record["Math"] >= PASSING_MARK_PER_SUBJECT else "FAIL"
    sci_status = "PASS" if student_record["Science"] >= PASSING_MARK_PER_SUBJECT else "FAIL"
    eng_status = "PASS" if student_record["English"] >= PASSING_MARK_PER_SUBJECT else "FAIL"

    marks_data = [
        [
            Paragraph("Sr. No.", table_header_style),
            Paragraph("Subject Name", table_header_style),
            Paragraph("Max Marks", table_header_style),
            Paragraph("Passing Marks", table_header_style),
            Paragraph("Marks Scored", table_header_style),
            Paragraph("Result", table_header_style)
        ],
        [
            Paragraph("1", table_cell_style),
            Paragraph("Mathematics", table_cell_style),
            Paragraph("100", table_cell_style),
            Paragraph("35", table_cell_style),
            Paragraph(f"{student_record['Math']:.1f}", table_cell_style),
            Paragraph(math_status, table_cell_style)
        ],
        [
            Paragraph("2", table_cell_style),
            Paragraph("Science", table_cell_style),
            Paragraph("100", table_cell_style),
            Paragraph("35", table_cell_style),
            Paragraph(f"{student_record['Science']:.1f}", table_cell_style),
            Paragraph(sci_status, table_cell_style)
        ],
        [
            Paragraph("3", table_cell_style),
            Paragraph("English", table_cell_style),
            Paragraph("100", table_cell_style),
            Paragraph("35", table_cell_style),
            Paragraph(f"{student_record['English']:.1f}", table_cell_style),
            Paragraph(eng_status, table_cell_style)
        ],
        [
            Paragraph("<b>TOTAL</b>", table_cell_style),
            Paragraph("<b>All Subjects Combined</b>", table_cell_style),
            Paragraph("<b>300</b>", table_cell_style),
            Paragraph("<b>105</b>", table_cell_style),
            Paragraph(f"<b>{student_record['Total Marks']:.1f}</b>", table_cell_style),
            Paragraph(f"<b>{student_record['Result']}</b>", table_cell_style)
        ]
    ]

    marks_table = Table(marks_data, colWidths=[45, 175, 75, 75, 75, 75])
    marks_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1A365D")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 7),
        ('TOPPADDING', (0, 0), (-1, 0), 7),
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor("#FFFFFF")),
        ('BACKGROUND', (0, 2), (-1, 2), colors.HexColor("#F7FAFC")),
        ('BACKGROUND', (0, 3), (-1, 3), colors.HexColor("#FFFFFF")),
        ('BACKGROUND', (0, 4), (-1, 4), colors.HexColor("#EDF2F7")),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
        ('TOPPADDING', (0, 1), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
    ]))
    story.append(marks_table)
    story.append(Spacer(1, 15))

    # 4. Performance Summary Metric Cards
    result_bg = colors.HexColor("#C6F6D5") if student_record["Result"] == "PASS" else colors.HexColor("#FED7D7")
    result_fg = "#22543D" if student_record["Result"] == "PASS" else "#742A2A"

    summary_data = [
        [
            Paragraph("<b>Total Scored</b>", meta_label_style),
            Paragraph("<b>Aggregate Percentage</b>", meta_label_style),
            Paragraph("<b>Assigned Grade</b>", meta_label_style),
            Paragraph("<b>Final Outcome</b>", meta_label_style)
        ],
        [
            Paragraph(f"<font size=12><b>{student_record['Total Marks']:.1f} / 300</b></font>", table_cell_style),
            Paragraph(f"<font size=12><b>{student_record['Percentage']:.2f}%</b></font>", table_cell_style),
            Paragraph(f"<font size=12><b>Grade {student_record['Grade']}</b></font>", table_cell_style),
            Paragraph(f"<font size=12 color='{result_fg}'><b>{student_record['Result']}</b></font>", table_cell_style)
        ]
    ]

    summary_table = Table(summary_data, colWidths=[130, 130, 130, 130])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#E2E8F0")),
        ('BACKGROUND', (0, 1), (2, 1), colors.HexColor("#F7FAFC")),
        ('BACKGROUND', (3, 1), (3, 1), result_bg),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor("#CBD5E0")),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
        ('TOPPADDING', (0, 0), (-1, -1), 7),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 7),
    ]))
    story.append(summary_table)
    story.append(Spacer(1, 30))

    # 5. Sign-off Blocks
    sig_data = [
        [
            Paragraph("____________________________<br/><b>Class Teacher / Faculty In-Charge</b>", table_cell_style),
            Paragraph("____________________________<br/><b>Head of Department (Robotics & Automation)</b>", table_cell_style),
            Paragraph("____________________________<br/><b>Controller of Examinations</b>", table_cell_style)
        ]
    ]
    sig_table = Table(sig_data, colWidths=[173, 173, 174])
    sig_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'BOTTOM'),
        ('TOPPADDING', (0, 0), (-1, -1), 20),
    ]))
    story.append(sig_table)

    # 6. Build Document
    doc.build(story)


# ======================================================================================
# EXCEL EXPORT WITH FORMATTING
# ======================================================================================
def export_processed_excel(processed_records: list, output_path: Path):
    """
    Exports clean processed records into Processed_Results.xlsx with styled columns.
    """
    df_out = pd.DataFrame(processed_records)
    
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_out.to_excel(writer, sheet_name="CA3_Results", index=False)
        worksheet = writer.sheets["CA3_Results"]

        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
        cell_alignment = Alignment(horizontal="center", vertical="center")
        border_thin = Border(
            left=Side(style='thin', color="CBD5E0"),
            right=Side(style='thin', color="CBD5E0"),
            top=Side(style='thin', color="CBD5E0"),
            bottom=Side(style='thin', color="CBD5E0")
        )

        for col_idx, col in enumerate(worksheet.columns, start=1):
            max_len = 0
            for row_idx, cell in enumerate(col, start=1):
                cell.border = border_thin
                if row_idx == 1:
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = cell_alignment
                else:
                    cell.alignment = cell_alignment
                val_str = str(cell.value or "")
                max_len = max(max_len, len(val_str))
            col_letter = worksheet.cell(row=1, column=col_idx).column_letter
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 14)


# ======================================================================================
# MAIN EXECUTION CONTROLLER
# ======================================================================================
def process_student_marks(input_file_path: str = None):
    """
    Core orchestrator implementing strict two-scenario exception handling.
    """
    current_dir = Path(__file__).resolve().parent
    error_log_path = current_dir / "error_log.txt"
    output_pdf_dir = current_dir / "Output_PDFs"
    output_excel_path = current_dir / "Processed_Results.xlsx"
    output_pdf_dir.mkdir(parents=True, exist_ok=True)

    print("=" * 80)
    print("  AUTOMATED STUDENT MARKS PROCESSING & BULK PDF REPORT CARD GENERATOR")
    print("  Academic Continuous Assessment 3 (CA3) - Engineering Evaluation")
    print("=" * 80)

    # Resolve target file path
    if input_file_path:
        target_path = Path(input_file_path)
        if not target_path.is_absolute():
            target_path = current_dir / target_path
    else:
        target_path = current_dir / "student_input_valid.xlsx"

    print(f"[*] Target Input File : {target_path}")
    print(f"[*] Error Log Target  : {error_log_path}")
    print(f"[*] PDF Output Folder : {output_pdf_dir}")
    print(f"[*] Summary Excel Path: {output_excel_path}\n")

    # ----------------------------------------------------------------------------------
    # EXCEPTION SCENARIO 1: FileNotFoundError or Missing/Corrupt Required Columns
    # ----------------------------------------------------------------------------------
    try:
        if not target_path.exists():
            raise FileNotFoundError(f"Input file does not exist at specified path: '{target_path}'")

        df_raw = pd.read_excel(target_path, engine="openpyxl")

        # Validate schema completeness
        missing_columns = [col for col in REQUIRED_COLUMNS if col not in df_raw.columns]
        if missing_columns:
            raise KeyError(f"Input file schema is corrupted. Missing required columns: {missing_columns}")

        print(f"[SUCCESS: SCENARIO 1 PASSED] Input file verified with {len(df_raw)} records and complete schema.")

    except (FileNotFoundError, KeyError, Exception) as exc:
        err_msg = str(exc)
        print(f"[!] [EXCEPTION SCENARIO 1 TRIGGERED] File/Schema Error: {err_msg}")
        log_error(
            log_file_path=error_log_path,
            scenario_name="EXCEPTION SCENARIO 1: FILE / SCHEMA ERROR",
            message=err_msg,
            context={"Target File": str(target_path), "Required Columns": REQUIRED_COLUMNS}
        )
        print(f"[!] Warning logged to '{error_log_path.name}'. Halting execution for this file safely.\n")
        return False

    # ----------------------------------------------------------------------------------
    # EXCEPTION SCENARIO 2: Data Integrity / ValueError (Row-by-Row Processing)
    # ----------------------------------------------------------------------------------
    print("\n[*] Commencing Row-Level Validation & PDF Generation Loop...")
    valid_records = []
    skipped_count = 0
    generated_pdf_count = 0

    for idx, row in df_raw.iterrows():
        row_number = idx + 2  # Excel row numbering (1-indexed + header)

        try:
            # 1. Check for Student ID & Name presence
            raw_id = row.get("Student ID")
            raw_name = row.get("Name")

            if pd.isna(raw_id) or str(raw_id).strip() == "" or str(raw_id).lower() == "nan":
                raise ValueError(f"Student ID is missing or blank in row {row_number}.")

            if pd.isna(raw_name) or str(raw_name).strip() == "" or str(raw_name).lower() == "nan":
                raise ValueError(f"Student Name is missing or blank for Student ID '{raw_id}'.")

            student_id = str(raw_id).strip()
            student_name = str(raw_name).strip()

            # 2. Parse & Validate Subject Marks (Math, Science, English)
            subject_marks = {}
            for subject in ["Math", "Science", "English"]:
                raw_val = row.get(subject)
                
                # Check for null or non-numeric conversion
                try:
                    score = float(raw_val)
                except (ValueError, TypeError):
                    raise ValueError(
                        f"Non-numeric score '{raw_val}' detected in subject '{subject}' for student '{student_id}'."
                    )

                if np.isnan(score):
                    raise ValueError(f"Missing score (NaN) in subject '{subject}' for student '{student_id}'.")

                # Boundary condition: 0.0 <= score <= 100.0
                if score < 0.0 or score > MAX_MARK_PER_SUBJECT:
                    raise ValueError(
                        f"Score {score} in subject '{subject}' violates boundary [0, 100] for student '{student_id}'."
                    )

                subject_marks[subject] = score

            # 3. Parse & Validate Attendance %
            raw_att = row.get("Attendance %")
            try:
                attendance = float(raw_att)
            except (ValueError, TypeError):
                raise ValueError(
                    f"Non-numeric attendance '{raw_att}' for student '{student_id}'."
                )

            if attendance < 0.0 or attendance > 100.0:
                raise ValueError(
                    f"Attendance {attendance}% violates boundary [0, 100] for student '{student_id}'."
                )

            # 4. Compute Metrics
            total_marks, percentage, grade, result = compute_academic_metrics(
                subject_marks["Math"],
                subject_marks["Science"],
                subject_marks["English"]
            )

            record = {
                "Student ID": student_id,
                "Name": student_name,
                "Math": subject_marks["Math"],
                "Science": subject_marks["Science"],
                "English": subject_marks["English"],
                "Total Marks": total_marks,
                "Percentage": percentage,
                "Grade": grade,
                "Result": result,
                "Attendance %": attendance,
                "Remarks": "Promoted" if result == "PASS" else "Needs Improvement"
            }

            # 5. Generate Individual PDF Report Card
            safe_id = "".join([c if c.isalnum() else "_" for c in student_id])
            pdf_path = output_pdf_dir / f"ReportCard_{safe_id}.pdf"
            generate_pdf_report_card(record, pdf_path)
            generated_pdf_count += 1
            valid_records.append(record)

            print(f"  [OK] Row {row_number}: Student '{student_id}' ({student_name}) -> Total: {total_marks}/300, {percentage}% (Grade {grade}) -> Generated {pdf_path.name}")

        except ValueError as val_err:
            skipped_count += 1
            err_msg = str(val_err)
            print(f"  [!] [EXCEPTION SCENARIO 2 TRIGGERED] Row {row_number} Skipped: {err_msg}")
            
            # Extract row data for context
            row_dict = {k: (None if pd.isna(v) else v) for k, v in row.items()}
            log_error(
                log_file_path=error_log_path,
                scenario_name="EXCEPTION SCENARIO 2: ROW DATA INTEGRITY ERROR",
                message=err_msg,
                context={
                    "Excel Row Number": row_number,
                    "Student ID": str(row.get("Student ID")),
                    "Row Data": str(row_dict)
                }
            )
            continue  # Continue processing subsequent rows cleanly

    # ----------------------------------------------------------------------------------
    # SUMMARY EXCEL EXPORT & CONCLUSION
    # ----------------------------------------------------------------------------------
    print("\n" + "=" * 80)
    print("  EXECUTION BATCH SUMMARY")
    print("=" * 80)
    print(f"  Total Rows Examined     : {len(df_raw)}")
    print(f"  Successfully Processed  : {len(valid_records)}")
    print(f"  Skipped (Integrity Err) : {skipped_count}")
    print(f"  PDF Report Cards Emitted: {generated_pdf_count}")

    if valid_records:
        export_processed_excel(valid_records, output_excel_path)
        print(f"  Master Excel Saved      : {output_excel_path.name}")
    else:
        print("  [NOTICE] No valid rows found. Master Excel report skipped.")

    if skipped_count > 0:
        print(f"  [NOTICE] {skipped_count} row errors were captured and isolated in '{error_log_path.name}'.")

    print("=" * 80 + "\n")
    return True


if __name__ == "__main__":
    cli_arg = sys.argv[1] if len(sys.argv) > 1 else None
    process_student_marks(cli_arg)
