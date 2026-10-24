@echo off
title Student Marks & PDF Generator - UI Dashboard
cd /d "%~dp0"
echo Starting Minimalistic UI Dashboard...
python -m streamlit run app.py
pause
