"""
Generates additional visual screenshot assets:
- 05_Error_Log_Preview.png
- 06_Folder_Structure_Tree.png
"""

from pathlib import Path
import matplotlib.pyplot as plt
import matplotlib.patches as patches

def generate_extra_screenshots():
    screenshot_dir = Path("PRN1_PRN2_Name_CA3/Screenshots")
    screenshot_dir.mkdir(parents=True, exist_ok=True)

    # 1. Error Log Preview
    fig, ax = plt.subplots(figsize=(10, 6), dpi=200)
    fig.patch.set_facecolor('#1E1E1E')
    ax.set_facecolor('#1E1E1E')
    ax.axis('off')

    log_sample = """error_log.txt - AUDIT LOGGING SUBSYSTEM
----------------------------------------------------------------------------------------
[2026-10-14 21:39:55] [EXCEPTION SCENARIO 2: ROW DATA INTEGRITY ERROR]
  Message: Score -15.0 in subject 'Math' violates boundary [0, 100] for student 'STU201'.
  Excel Row Number: 2
  Student ID: STU201
  Row Data: {'Student ID': 'STU201', 'Name': 'Vikram Singh', 'Math': -15, 'Science': 82, ...}
----------------------------------------------------------------------------------------
[2026-10-14 21:39:55] [EXCEPTION SCENARIO 2: ROW DATA INTEGRITY ERROR]
  Message: Score 150.0 in subject 'Science' violates boundary [0, 100] for student 'STU202'.
  Excel Row Number: 3
  Student ID: STU202
  Row Data: {'Student ID': 'STU202', 'Name': 'Pooja Nair', 'Math': 78, 'Science': 150, ...}
----------------------------------------------------------------------------------------
[2026-10-14 21:39:55] [EXCEPTION SCENARIO 2: ROW DATA INTEGRITY ERROR]
  Message: Non-numeric score 'Absent' detected in subject 'English' for student 'STU203'.
  Excel Row Number: 4
  Student ID: STU203
  Row Data: {'Student ID': 'STU203', 'Name': 'Sameer Joshi', 'English': 'Absent', ...}
----------------------------------------------------------------------------------------
[2026-10-14 21:40:00] [EXCEPTION SCENARIO 1: FILE / SCHEMA ERROR]
  Message: Input file does not exist at specified path: 'non_existent_file.xlsx'
  Target File: D:\\Downloads\\SAL_project\\PRN1_PRN2_Name_CA3\\Project_Files\\non_existent_file.xlsx
  Required Columns: ['Student ID', 'Name', 'Math', 'Science', 'English', 'Attendance %']
----------------------------------------------------------------------------------------"""

    ax.text(0.02, 0.95, log_sample, color='#E5C07B', fontsize=8, family='monospace', va='top')
    plt.tight_layout()
    plt.savefig(screenshot_dir / "05_Error_Log_Preview.png", facecolor=fig.get_facecolor(), bbox_inches='tight')
    plt.close()

    # 2. Folder Structure Graphic
    fig, ax = plt.subplots(figsize=(9, 6), dpi=200)
    fig.patch.set_facecolor('#0F172A')
    ax.set_facecolor('#0F172A')
    ax.axis('off')

    tree_text = """
PROJECT ARTIFACT DIRECTORY STRUCTURE
====================================================================
PRN1_PRN2_Name_CA3/
│
├── Project_Report.pdf            <-- Publication-Quality Report PDF
├── Project_Report_Draft.txt      <-- Comprehensive Academic Report Text
├── Presentation_Slides.pptx      <-- 13-Slide Modern Presentation Deck
├── Presentation_Slides.txt       <-- Presentation Slide Transcript
├── Demo_Video_Script.txt         <-- Timestamped 5:00 Demo Video Script
│
├── Workflow/
│   └── Workflow_Architecture.png <-- Full Visual Pipeline Architecture
│
├── Screenshots/
│   ├── 01_Valid_Execution_Terminal.png
│   ├── 02_Invalid_Execution_Terminal.png
│   ├── 03_Sample_ReportCard_STU101.png
│   ├── 04_Processed_Results_Excel.png
│   ├── 05_Error_Log_Preview.png
│   └── 06_Folder_Structure_Tree.png
│
└── Project_Files/
    ├── main.py                   <-- Core Engine with 2-Tier Exception Handling
    ├── generate_test_data.py     <-- Clean & Dirty Test Data Synthesizer
    ├── student_input_valid.xlsx  <-- Clean 5-Student Input Excel
    ├── student_input_invalid.xlsx<-- Edge Cases & Dirty Rows Input Excel
    ├── Processed_Results.xlsx    <-- Master Formatted Output Excel
    ├── error_log.txt             <-- Persistent Diagnostic Audit Log
    └── Output_PDFs/              <-- Bulk Vectorized Report Card PDFs
        ├── ReportCard_STU101.pdf
        ├── ReportCard_STU102.pdf
        ├── ... (Individual PDFs for all valid cohort records)
====================================================================
"""
    ax.text(0.02, 0.95, tree_text, color='#A7F3D0', fontsize=8.5, family='monospace', va='top')
    plt.tight_layout()
    plt.savefig(screenshot_dir / "06_Folder_Structure_Tree.png", facecolor=fig.get_facecolor(), bbox_inches='tight')
    plt.close()

    print("[SUCCESS] Extra screenshots generated.")

if __name__ == "__main__":
    generate_extra_screenshots()
