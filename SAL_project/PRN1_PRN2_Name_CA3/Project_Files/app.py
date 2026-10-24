#!/usr/bin/env python3
"""
========================================================================================
Minimalistic Web UI Dashboard for Automated Student Marks Processing & PDF Generator
Department of Robotics and Automation | Continuous Assessment 3 (CA3)
Authors:
  1. Aditya Shankar Datta (PRN: 24070127009)
  2. Pushkal Karakavalasa (PRN: 24070127057)
  3. Aniket Atole         (PRN: 24070127017)

Usage:
  Run via Streamlit : streamlit run app.py
  Or directly       : python app.py
========================================================================================
"""

import sys
import os
import io
import zipfile
import base64
import subprocess
from pathlib import Path
from contextlib import redirect_stdout
import pandas as pd


def _is_running_in_streamlit() -> bool:
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        return get_script_run_ctx() is not None
    except Exception:
        return False


if __name__ == "__main__" and not _is_running_in_streamlit():
    # Allow running directly with `python app.py`
    print("[*] Launching Minimalistic UI Dashboard via Streamlit...")
    sys.exit(subprocess.call([sys.executable, "-m", "streamlit", "run", str(Path(__file__).resolve())]))


import streamlit as st

BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from main import process_student_marks, REQUIRED_COLUMNS
from generate_test_data import generate_datasets

# ======================================================================================
# PAGE CONFIGURATION & MINIMALISTIC STYLING
# ======================================================================================
OUTPUT_PDF_DIR = BASE_DIR / "Output_PDFs"
PROCESSED_EXCEL_PATH = BASE_DIR / "Processed_Results.xlsx"
ERROR_LOG_PATH = BASE_DIR / "error_log.txt"

st.set_page_config(
    page_title="Student Marks & PDF Generator",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1.8rem;
        padding-bottom: 2rem;
        max-width: 1180px;
    }
    .minimal-header {
        border-bottom: 1px solid #E2E8F0;
        padding-bottom: 0.8rem;
        margin-bottom: 1.2rem;
    }
    .minimal-title {
        font-size: 1.65rem;
        font-weight: 700;
        color: #0F172A;
        margin: 0;
    }
    .minimal-subtitle {
        font-size: 0.9rem;
        color: #64748B;
        margin-top: 0.25rem;
    }
    .status-pill-pass {
        display: inline-block;
        padding: 2px 10px;
        border-radius: 999px;
        font-size: 0.78rem;
        font-weight: 600;
        background: #DCFCE7;
        color: #166534;
    }
    .status-pill-warn {
        display: inline-block;
        padding: 2px 10px;
        border-radius: 999px;
        font-size: 0.78rem;
        font-weight: 600;
        background: #FEF3C7;
        color: #92400E;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ======================================================================================
# INITIALIZE SESSION STATE & ENSURE TEST FILES EXIST
# ======================================================================================
if not (BASE_DIR / "student_input_valid.xlsx").exists() or not (BASE_DIR / "student_input_invalid.xlsx").exists():
    generate_datasets(BASE_DIR)

if "last_console_log" not in st.session_state:
    st.session_state.last_console_log = ""
if "last_run_file" not in st.session_state:
    st.session_state.last_run_file = None
if "last_run_ok" not in st.session_state:
    st.session_state.last_run_ok = None
if "run_stats" not in st.session_state:
    st.session_state.run_stats = {"examined": 0, "valid": 0, "skipped": 0, "pdfs": 0}


def parse_stats_from_output(console_text: str) -> dict:
    stats = {"examined": 0, "valid": 0, "skipped": 0, "pdfs": 0}
    for line in console_text.splitlines():
        if "Total Rows Examined" in line:
            stats["examined"] = int(line.split(":")[-1].strip())
        elif "Successfully Processed" in line:
            stats["valid"] = int(line.split(":")[-1].strip())
        elif "Skipped (Integrity Err)" in line:
            stats["skipped"] = int(line.split(":")[-1].strip())
        elif "PDF Report Cards Emitted" in line:
            stats["pdfs"] = int(line.split(":")[-1].strip())
    return stats


def create_pdfs_zip() -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        if OUTPUT_PDF_DIR.exists():
            for pdf_file in sorted(OUTPUT_PDF_DIR.glob("*.pdf")):
                zf.write(pdf_file, arcname=pdf_file.name)
    buf.seek(0)
    return buf.read()


# ======================================================================================
# HEADER
# ======================================================================================
st.markdown(
    """
    <div class="minimal-header">
        <p class="minimal-title">Student Marks Processing & Bulk PDF Generator</p>
        <p class="minimal-subtitle">
            Department of Robotics and Automation • CA3 Automation Engine &nbsp;|&nbsp;
            Aditya Shankar Datta (24070127009) • Pushkal Karakavalasa (24070127057) • Aniket Atole (24070127017)
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ======================================================================================
# SIDEBAR CONTROLS
# ======================================================================================
with st.sidebar:
    st.subheader("⚙️ Control Panel")

    mode = st.radio(
        "Select Input Source",
        options=[
            "Valid Dataset (5 Clean Rows)",
            "Invalid Dataset (Edge Cases)",
            "Missing File Test (Scenario 1)",
            "Upload Custom Excel (.xlsx)",
        ],
        index=0,
    )

    target_file_path = None
    uploaded_df_preview = None

    if mode == "Valid Dataset (5 Clean Rows)":
        target_file_path = BASE_DIR / "student_input_valid.xlsx"
        st.caption("Tests normal processing of 5 valid student records.")
    elif mode == "Invalid Dataset (Edge Cases)":
        target_file_path = BASE_DIR / "student_input_invalid.xlsx"
        st.caption("Tests Scenario 2: negative marks (-15), >100 (150), 'Absent', missing ID.")
    elif mode == "Missing File Test (Scenario 1)":
        target_file_path = BASE_DIR / "non_existent_file.xlsx"
        st.caption("Tests Scenario 1: FileNotFoundError guard & safe logging.")
    else:
        uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])
        if uploaded_file is not None:
            custom_path = BASE_DIR / f"_uploaded_{uploaded_file.name}"
            with open(custom_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            target_file_path = custom_path
            st.success(f"Loaded: {uploaded_file.name}")

    st.divider()

    run_clicked = st.button(
        "▶️ Run Processing Engine",
        type="primary",
        width="stretch",
        disabled=(target_file_path is None),
    )

    col_sb1, col_sb2 = st.columns(2)
    with col_sb1:
        if st.button("🔄 Reset Test Data", width="stretch", help="Regenerate clean and invalid test Excel files"):
            buf = io.StringIO()
            with redirect_stdout(buf):
                generate_datasets(BASE_DIR)
            st.toast("Test Excel datasets regenerated!", icon="✅")

    with col_sb2:
        if st.button("🧹 Clear Error Log", width="stretch", help="Clear error_log.txt"):
            if ERROR_LOG_PATH.exists():
                ERROR_LOG_PATH.write_text("", encoding="utf-8")
            st.toast("Cleared error_log.txt", icon="🧹")

    st.divider()
    st.caption("**Required Excel Columns:**")
    st.code(", ".join(REQUIRED_COLUMNS), language="text")


# ======================================================================================
# EXECUTE ENGINE WHEN TRIGGERED
# ======================================================================================
if run_clicked and target_file_path is not None:
    # Clear old PDFs before a fresh valid/invalid run so the user sees exact run outputs
    if target_file_path.exists() and OUTPUT_PDF_DIR.exists():
        for old_pdf in OUTPUT_PDF_DIR.glob("*.pdf"):
            try:
                old_pdf.unlink()
            except Exception:
                pass

    console_buffer = io.StringIO()
    with redirect_stdout(console_buffer):
        run_ok = process_student_marks(str(target_file_path))

    captured_output = console_buffer.getvalue()
    st.session_state.last_console_log = captured_output
    st.session_state.last_run_file = target_file_path.name
    st.session_state.last_run_ok = run_ok
    st.session_state.run_stats = parse_stats_from_output(captured_output)


# ======================================================================================
# TOP KPI METRICS
# ======================================================================================
stats = st.session_state.run_stats
k1, k2, k3, k4 = st.columns(4)
k1.metric("Rows Examined", stats["examined"])
k2.metric("Valid Processed", stats["valid"])
k3.metric("Skipped (Errors)", stats["skipped"])
k4.metric("PDFs Generated", stats["pdfs"])

if st.session_state.last_run_file:
    if st.session_state.last_run_ok and stats["skipped"] == 0:
        st.success(f"Last Run (`{st.session_state.last_run_file}`): Completed cleanly with 0 errors.")
    elif st.session_state.last_run_ok and stats["skipped"] > 0:
        st.warning(
            f"Last Run (`{st.session_state.last_run_file}`): Processed {stats['valid']} valid row(s); "
            f"isolated & skipped {stats['skipped']} invalid row(s) into `error_log.txt` (Scenario 2)."
        )
    else:
        st.error(
            f"Last Run (`{st.session_state.last_run_file}`): Halted safely by Scenario 1 (File/Schema Error). "
            f"Details logged to `error_log.txt`."
        )

# ======================================================================================
# MAIN WORKSPACE TABS
# ======================================================================================
tab_data, tab_pdfs, tab_logs = st.tabs([
    "📊 Data & Processed Results",
    "📄 Generated PDF Report Cards",
    "🛡️ Exception Audit Log & Console",
])

# --------------------------------------------------------------------------------------
# TAB 1: DATA & PROCESSED RESULTS
# --------------------------------------------------------------------------------------
with tab_data:
    col_in, col_out = st.columns([1, 1.25], gap="large")

    with col_in:
        st.markdown("##### 📥 Input Excel Preview")
        if target_file_path is not None and target_file_path.exists():
            try:
                df_in = pd.read_excel(target_file_path, engine="openpyxl")
                st.caption(f"File: `{target_file_path.name}` ({len(df_in)} rows)")
                st.dataframe(df_in.astype(str), width="stretch", hide_index=True)
            except Exception as e:
                st.error(f"Could not preview input file: {e}")
        else:
            st.info(f"Selected file `{target_file_path.name if target_file_path else 'None'}` does not exist on disk (Scenario 1 test).")

    with col_out:
        st.markdown("##### ✅ Processed Output (`Processed_Results.xlsx`)")
        if PROCESSED_EXCEL_PATH.exists():
            try:
                df_out = pd.read_excel(PROCESSED_EXCEL_PATH, engine="openpyxl")
                st.caption(f"Master Summary: `{PROCESSED_EXCEL_PATH.name}` ({len(df_out)} valid records)")
                st.dataframe(df_out, width="stretch", hide_index=True)

                with open(PROCESSED_EXCEL_PATH, "rb") as f:
                    st.download_button(
                        label="⬇️ Download Processed_Results.xlsx",
                        data=f.read(),
                        file_name="Processed_Results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        width="stretch",
                    )
            except Exception as e:
                st.error(f"Error reading Processed_Results.xlsx: {e}")
        else:
            st.info("No `Processed_Results.xlsx` found yet. Click **Run Processing Engine** in the sidebar.")

# --------------------------------------------------------------------------------------
# TAB 2: GENERATED PDF REPORT CARDS
# --------------------------------------------------------------------------------------
with tab_pdfs:
    pdf_files = sorted(OUTPUT_PDF_DIR.glob("*.pdf")) if OUTPUT_PDF_DIR.exists() else []

    if not pdf_files:
        st.info("No PDF report cards in `Output_PDFs/` yet. Run the processing engine to generate them.")
    else:
        top_c1, top_c2 = st.columns([3, 1])
        with top_c1:
            st.markdown(f"##### 🗂️ Bulk PDF Report Cards (`{len(pdf_files)}` files in `Output_PDFs/`)")
        with top_c2:
            st.download_button(
                label="📦 Download All PDFs (.zip)",
                data=create_pdfs_zip(),
                file_name="Output_PDFs_ReportCards.zip",
                mime="application/zip",
                width="stretch",
            )

        selected_pdf_name = st.selectbox(
            "Select a Student Report Card to Preview or Download:",
            options=[p.name for p in pdf_files],
        )
        selected_pdf_path = OUTPUT_PDF_DIR / selected_pdf_name
        pdf_bytes = selected_pdf_path.read_bytes()

        dl_col, _ = st.columns([1, 2])
        with dl_col:
            st.download_button(
                label=f"⬇️ Download {selected_pdf_name}",
                data=pdf_bytes,
                file_name=selected_pdf_name,
                mime="application/pdf",
                width="stretch",
            )

        # Inline PDF Viewer
        b64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")
        pdf_iframe = f'<iframe src="data:application/pdf;base64,{b64_pdf}" width="100%" height="620" type="application/pdf" style="border: 1px solid #CBD5E1; border-radius: 8px;"></iframe>'
        st.markdown(pdf_iframe, unsafe_allow_html=True)

# --------------------------------------------------------------------------------------
# TAB 3: EXCEPTION AUDIT LOG & CONSOLE OUTPUT
# --------------------------------------------------------------------------------------
with tab_logs:
    log_col, term_col = st.columns(2, gap="large")

    with log_col:
        st.markdown("##### 🛡️ Persistent Audit Log (`error_log.txt`)")
        if ERROR_LOG_PATH.exists() and ERROR_LOG_PATH.read_text(encoding="utf-8").strip():
            log_content = ERROR_LOG_PATH.read_text(encoding="utf-8")
            st.download_button(
                label="⬇️ Download error_log.txt",
                data=log_content,
                file_name="error_log.txt",
                mime="text/plain",
            )
            st.code(log_content, language="text")
        else:
            st.success("`error_log.txt` is currently empty (no errors logged).")

    with term_col:
        st.markdown("##### 💻 Live Engine Console Output")
        if st.session_state.last_console_log:
            st.code(st.session_state.last_console_log, language="text")
        else:
            st.info("Click **▶️ Run Processing Engine** in the sidebar to view live execution logs.")
