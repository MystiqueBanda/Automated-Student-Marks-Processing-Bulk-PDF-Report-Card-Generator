"""
Test Data Generator for Automated Student Marks Processing System
Generates:
1. student_input_valid.xlsx   (5 valid student records)
2. student_input_invalid.xlsx (5 records containing various boundary and edge case errors)
"""

import os
from pathlib import Path
import pandas as pd

def generate_datasets(output_dir=None):
    if output_dir is None:
        output_dir = Path(__file__).resolve().parent
    else:
        output_dir = Path(output_dir)

    output_dir.mkdir(parents=True, exist_ok=True)

    # -------------------------------------------------------------
    # 1. Valid Dataset (5 clean rows)
    # -------------------------------------------------------------
    valid_data = [
        {
            "Student ID": "STU101",
            "Name": "Aarav Sharma",
            "Math": 88,
            "Science": 92,
            "English": 85,
            "Attendance %": 94.5
        },
        {
            "Student ID": "STU102",
            "Name": "Diya Patel",
            "Math": 95,
            "Science": 98,
            "English": 91,
            "Attendance %": 98.0
        },
        {
            "Student ID": "STU103",
            "Name": "Rohan Verma",
            "Math": 68,
            "Science": 74,
            "English": 62,
            "Attendance %": 85.0
        },
        {
            "Student ID": "STU104",
            "Name": "Ananya Iyer",
            "Math": 52,
            "Science": 58,
            "English": 49,
            "Attendance %": 78.5
        },
        {
            "Student ID": "STU105",
            "Name": "Kabir Mehta",
            "Math": 38,
            "Science": 42,
            "English": 40,
            "Attendance %": 72.0
        }
    ]

    df_valid = pd.DataFrame(valid_data)
    valid_file_path = output_dir / "student_input_valid.xlsx"
    df_valid.to_excel(valid_file_path, index=False, engine="openpyxl")
    print(f"[SUCCESS] Valid test data written to: {valid_file_path} ({len(df_valid)} records)")

    # -------------------------------------------------------------
    # 2. Invalid Dataset (5 rows containing critical edge cases)
    # -------------------------------------------------------------
    # Row 1: Negative mark in Math (-15)
    # Row 2: Score > 100 in Science (150)
    # Row 3: Non-numeric mark in English ('Absent')
    # Row 4: Missing Student ID (None/NaN)
    # Row 5: Clean valid record (to verify continuing execution)
    invalid_data = [
        {
            "Student ID": "STU201",
            "Name": "Vikram Singh",
            "Math": -15,          # Boundary Error: Negative score
            "Science": 82,
            "English": 75,
            "Attendance %": 90.0
        },
        {
            "Student ID": "STU202",
            "Name": "Pooja Nair",
            "Math": 78,
            "Science": 150,        # Boundary Error: Score > 100
            "English": 80,
            "Attendance %": 88.0
        },
        {
            "Student ID": "STU203",
            "Name": "Sameer Joshi",
            "Math": 65,
            "Science": 70,
            "English": "Absent",   # Data Type Error: String instead of numeric
            "Attendance %": 65.0
        },
        {
            "Student ID": None,    # Structural Error: Missing Student ID
            "Name": "Meera Rao",
            "Math": 85,
            "Science": 89,
            "English": 92,
            "Attendance %": 95.0
        },
        {
            "Student ID": "STU205",
            "Name": "Ishaan Gupta", # Valid record: Should process successfully despite earlier errors
            "Math": 75,
            "Science": 80,
            "English": 85,
            "Attendance %": 91.0
        }
    ]

    df_invalid = pd.DataFrame(invalid_data)
    invalid_file_path = output_dir / "student_input_invalid.xlsx"
    df_invalid.to_excel(invalid_file_path, index=False, engine="openpyxl")
    print(f"[SUCCESS] Invalid test data written to: {invalid_file_path} ({len(df_invalid)} records)")

if __name__ == "__main__":
    generate_datasets()
