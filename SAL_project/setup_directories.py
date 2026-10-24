"""
Directory Setup Script for CA3 Mini Project
Creates the exact standardized folder structure:
PRN1_PRN2_Name_CA3/
│── Project_Report.pdf
│── Demo_Video.mp4
│── Workflow/
│── Screenshots/
│── Project_Files/
│   ├── main.py
│   ├── generate_test_data.py
│   ├── student_input_valid.xlsx
│   ├── student_input_invalid.xlsx
│   ├── Processed_Results.xlsx
│   ├── error_log.txt
│   └── Output_PDFs/
"""

import os
from pathlib import Path

def create_project_structure(base_dir="PRN1_PRN2_Name_CA3"):
    base_path = Path(base_dir)
    directories = [
        base_path,
        base_path / "Workflow",
        base_path / "Screenshots",
        base_path / "Project_Files",
        base_path / "Project_Files" / "Output_PDFs",
    ]
    
    for directory in directories:
        directory.mkdir(parents=True, exist_ok=True)
        print(f"[CREATED] Directory: {directory}")

    # Create placeholder for demo video note
    demo_video_note = base_path / "Demo_Video_README.txt"
    with open(demo_video_note, "w", encoding="utf-8") as f:
        f.write("Place your recorded demonstration video file here named 'Demo_Video.mp4'.\n"
                "Follow the script provided in 'Demo_Video_Script.txt'.\n")
    print(f"[CREATED] Placeholder: {demo_video_note}")

    print("\n[SUCCESS] Project directory structure initialized successfully.")

if __name__ == "__main__":
    create_project_structure()
