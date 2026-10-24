@echo off
title Student Marks & PDF Generator - UI Dashboard
cd /d "%~dp0\PRN1_PRN2_Name_CA3\Project_Files"
echo Starting Minimalistic UI Dashboard...
python -m streamlit run app.py
pause
