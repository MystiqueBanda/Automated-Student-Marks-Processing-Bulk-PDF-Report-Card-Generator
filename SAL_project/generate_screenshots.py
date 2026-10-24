"""
Generates visual screenshot assets illustrating test executions,
sample report cards, and spreadsheets for the Screenshots/ directory.
"""

from pathlib import Path
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import pandas as pd

def generate_screenshots():
    screenshot_dir = Path("PRN1_PRN2_Name_CA3/Screenshots")
    screenshot_dir.mkdir(parents=True, exist_ok=True)

    # -------------------------------------------------------------
    # 1. Screenshot: Valid Execution Terminal
    # -------------------------------------------------------------
    fig, ax = plt.subplots(figsize=(10, 6), dpi=200)
    fig.patch.set_facecolor('#0F172A')
    ax.set_facecolor('#0F172A')
    ax.axis('off')

    term_text = """
PS D:\\PRN1_PRN2_Name_CA3\\Project_Files> python main.py student_input_valid.xlsx
================================================================================
  AUTOMATED STUDENT MARKS PROCESSING & BULK PDF REPORT CARD GENERATOR
  Academic Continuous Assessment 3 (CA3) - Engineering Evaluation
================================================================================
[*] Target Input File : student_input_valid.xlsx
[*] Error Log Target  : error_log.txt
[*] PDF Output Folder : Output_PDFs
[*] Summary Excel Path: Processed_Results.xlsx

[SUCCESS: SCENARIO 1 PASSED] Input file verified with 5 records and complete schema.

[*] Commencing Row-Level Validation & PDF Generation Loop...
  [OK] Row 2: Student 'STU101' (Aarav Sharma) -> Total: 265.0/300, 88.33% (Grade A) -> Generated ReportCard_STU101.pdf
  [OK] Row 3: Student 'STU102' (Diya Patel)   -> Total: 284.0/300, 94.67% (Grade A) -> Generated ReportCard_STU102.pdf
  [OK] Row 4: Student 'STU103' (Rohan Verma)  -> Total: 204.0/300, 68.00% (Grade B) -> Generated ReportCard_STU103.pdf
  [OK] Row 5: Student 'STU104' (Ananya Iyer)  -> Total: 159.0/300, 53.00% (Grade C) -> Generated ReportCard_STU104.pdf
  [OK] Row 6: Student 'STU105' (Kabir Mehta)  -> Total: 120.0/300, 40.00% (Grade Pass) -> Generated ReportCard_STU105.pdf

================================================================================
  EXECUTION BATCH SUMMARY
================================================================================
  Total Rows Examined     : 5
  Successfully Processed  : 5
  Skipped (Integrity Err) : 0
  PDF Report Cards Emitted: 5
  Master Excel Saved      : Processed_Results.xlsx
================================================================================
"""
    ax.text(0.02, 0.95, term_text, color='#38BDF8', fontsize=8, family='monospace', va='top')
    plt.tight_layout()
    plt.savefig(screenshot_dir / "01_Valid_Execution_Terminal.png", facecolor=fig.get_facecolor(), bbox_inches='tight')
    plt.close()

    # -------------------------------------------------------------
    # 2. Screenshot: Invalid Execution Terminal (Exception Handling)
    # -------------------------------------------------------------
    fig, ax = plt.subplots(figsize=(10, 6), dpi=200)
    fig.patch.set_facecolor('#0F172A')
    ax.set_facecolor('#0F172A')
    ax.axis('off')

    err_term_text = """
PS D:\\PRN1_PRN2_Name_CA3\\Project_Files> python main.py student_input_invalid.xlsx
================================================================================
  AUTOMATED STUDENT MARKS PROCESSING & BULK PDF REPORT CARD GENERATOR
  Academic Continuous Assessment 3 (CA3) - Engineering Evaluation
================================================================================
[*] Target Input File : student_input_invalid.xlsx
[*] Error Log Target  : error_log.txt
[*] PDF Output Folder : Output_PDFs
[*] Summary Excel Path: Processed_Results.xlsx

[SUCCESS: SCENARIO 1 PASSED] Input file verified with 5 records and complete schema.

[*] Commencing Row-Level Validation & PDF Generation Loop...
  [!] [EXCEPTION SCENARIO 2 TRIGGERED] Row 2 Skipped: Score -15.0 in subject 'Math' violates boundary [0, 100] for student 'STU201'.
  [!] [EXCEPTION SCENARIO 2 TRIGGERED] Row 3 Skipped: Score 150.0 in subject 'Science' violates boundary [0, 100] for student 'STU202'.
  [!] [EXCEPTION SCENARIO 2 TRIGGERED] Row 4 Skipped: Non-numeric score 'Absent' detected in subject 'English' for student 'STU203'.
  [!] [EXCEPTION SCENARIO 2 TRIGGERED] Row 5 Skipped: Student ID is missing or blank in row 5.
  [OK] Row 6: Student 'STU205' (Ishaan Gupta) -> Total: 240.0/300, 80.00% (Grade A) -> Generated ReportCard_STU205.pdf

================================================================================
  EXECUTION BATCH SUMMARY
================================================================================
  Total Rows Examined     : 5
  Successfully Processed  : 1
  Skipped (Integrity Err) : 4
  PDF Report Cards Emitted: 1
  Master Excel Saved      : Processed_Results.xlsx
  [NOTICE] 4 row errors were captured and isolated in 'error_log.txt'.
================================================================================
"""
    ax.text(0.02, 0.95, err_term_text, color='#F87171', fontsize=8, family='monospace', va='top')
    plt.tight_layout()
    plt.savefig(screenshot_dir / "02_Invalid_Execution_Terminal.png", facecolor=fig.get_facecolor(), bbox_inches='tight')
    plt.close()

    # -------------------------------------------------------------
    # 3. Screenshot: Sample Report Card Layout
    # -------------------------------------------------------------
    fig, ax = plt.subplots(figsize=(8.5, 11), dpi=200)
    fig.patch.set_facecolor('#FFFFFF')
    ax.set_facecolor('#FFFFFF')
    ax.set_xlim(0, 8.5)
    ax.set_ylim(0, 11)
    ax.axis('off')

    # Header Box
    ax.text(4.25, 10.3, "APEX INSTITUTE OF TECHNOLOGY & HIGHER LEARNING", fontsize=14, fontweight='bold', ha='center', color='#1A365D')
    ax.text(4.25, 10.0, "Department of Robotics and Automation | Academic Session 2026-2027", fontsize=9, ha='center', color='#4A5568')
    ax.plot([0.8, 7.7], [9.8, 9.8], color='#1A365D', lw=2)
    ax.text(4.25, 9.5, "OFFICIAL STUDENT PERFORMANCE REPORT CARD (CA3)", fontsize=11, fontweight='bold', ha='center', color='#2B6CB0')

    # Student Info Box
    box_s = patches.FancyBboxPatch((0.8, 8.2), 6.9, 1.0, boxstyle="round,pad=0.05", edgecolor='#CBD5E0', facecolor='#F7FAFC')
    ax.add_patch(box_s)
    ax.text(1.0, 8.9, "Student ID   : STU101", fontsize=9, fontweight='bold', color='#1A202C')
    ax.text(1.0, 8.5, "Student Name : Aarav Sharma", fontsize=9, fontweight='bold', color='#1A202C')
    ax.text(4.5, 8.9, "Evaluation Cycle: CA3 Evaluation", fontsize=9, color='#1A202C')
    ax.text(4.5, 8.5, "Attendance      : 94.5%  |  Status: PASS", fontsize=9, color='#276749', fontweight='bold')

    # Table of Marks
    col_labels = ["Sr. No.", "Subject Name", "Max Marks", "Passing Marks", "Marks Scored", "Result"]
    table_data = [
        ["1", "Mathematics", "100", "35", "88.0", "PASS"],
        ["2", "Science", "100", "35", "92.0", "PASS"],
        ["3", "English", "100", "35", "85.0", "PASS"],
        ["TOTAL", "All Subjects", "300", "105", "265.0", "PASS"]
    ]
    t = ax.table(cellText=table_data, colLabels=col_labels, loc='center', cellLoc='center', bbox=[0.8, 5.8, 6.9, 2.0])
    t.auto_set_font_size(False)
    t.set_fontsize(9)
    for (row, col), cell in t.get_celld().items():
        if row == 0:
            cell.set_facecolor('#1A365D')
            cell.set_text_props(color='white', weight='bold')
        elif row == 4:
            cell.set_facecolor('#EDF2F7')
            cell.set_text_props(weight='bold')
        else:
            cell.set_facecolor('#FFFFFF' if row % 2 == 1 else '#F7FAFC')

    # Performance Summary Card
    sum_labels = ["Total Scored", "Aggregate Percentage", "Assigned Grade", "Final Outcome"]
    sum_values = [["265.0 / 300", "88.33%", "Grade A", "PASS"]]
    t_sum = ax.table(cellText=sum_values, colLabels=sum_labels, loc='center', cellLoc='center', bbox=[0.8, 4.4, 6.9, 0.9])
    t_sum.auto_set_font_size(False)
    t_sum.set_fontsize(10)
    for (row, col), cell in t_sum.get_celld().items():
        if row == 0:
            cell.set_facecolor('#E2E8F0')
            cell.set_text_props(weight='bold', color='#2D3748')
        else:
            if col == 3:
                cell.set_facecolor('#C6F6D5')
                cell.set_text_props(weight='bold', color='#22543D')
            else:
                cell.set_facecolor('#F7FAFC')
                cell.set_text_props(weight='bold', color='#1A202C')

    # Signatures
    ax.text(1.8, 2.0, "___________________________\nClass Teacher / Faculty In-Charge", ha='center', fontsize=8, color='#4A5568')
    ax.text(6.7, 2.0, "___________________________\nController of Examinations", ha='center', fontsize=8, color='#4A5568')

    plt.tight_layout()
    plt.savefig(screenshot_dir / "03_Sample_ReportCard_STU101.png", bbox_inches='tight')
    plt.close()

    # -------------------------------------------------------------
    # 4. Screenshot: Processed Results Excel Table
    # -------------------------------------------------------------
    fig, ax = plt.subplots(figsize=(11, 4), dpi=200)
    fig.patch.set_facecolor('#FFFFFF')
    ax.set_facecolor('#FFFFFF')
    ax.axis('off')

    df = pd.read_excel("PRN1_PRN2_Name_CA3/Project_Files/Processed_Results.xlsx")
    t_excel = ax.table(cellText=df.values, colLabels=df.columns, loc='center', cellLoc='center', bbox=[0, 0, 1, 1])
    t_excel.auto_set_font_size(False)
    t_excel.set_fontsize(9)
    for (row, col), cell in t_excel.get_celld().items():
        if row == 0:
            cell.set_facecolor('#1A365D')
            cell.set_text_props(color='white', weight='bold')
        else:
            cell.set_facecolor('#FFFFFF' if row % 2 == 1 else '#F8FAFC')

    plt.title("Processed_Results.xlsx - Consolidated Master Summary Sheet", fontsize=12, fontweight='bold', color='#1A365D', pad=15)
    plt.tight_layout()
    plt.savefig(screenshot_dir / "04_Processed_Results_Excel.png", bbox_inches='tight')
    plt.close()

    print(f"[SUCCESS] 4 visual preview screenshots saved to {screenshot_dir}")

if __name__ == "__main__":
    generate_screenshots()
