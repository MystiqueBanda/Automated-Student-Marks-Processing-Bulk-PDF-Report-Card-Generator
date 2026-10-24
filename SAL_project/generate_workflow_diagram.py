"""
Generates a professional Workflow Architecture diagram for the CA3 project.
Saves to PRN1_PRN2_Name_CA3/Workflow/Workflow_Architecture.png
"""

from pathlib import Path
import matplotlib.pyplot as plt
import matplotlib.patches as patches

def draw_workflow():
    fig, ax = plt.subplots(figsize=(14, 8), dpi=300)
    ax.set_xlim(0, 14)
    ax.set_ylim(0, 8)
    ax.axis('off')

    # Background color
    fig.patch.set_facecolor('#F8FAFC')
    ax.set_facecolor('#F8FAFC')

    # Title
    ax.text(7, 7.5, "SYSTEM ARCHITECTURE & EXECUTION WORKFLOW", 
            fontsize=16, fontweight='bold', ha='center', color='#1E293B', fontfamily='sans-serif')
    ax.text(7, 7.15, "Automated Student Marks Processing & Bulk PDF Report Card Generator", 
            fontsize=11, ha='center', color='#64748B', fontfamily='sans-serif')

    # Helper function for drawing rounded boxes
    def draw_box(x, y, w, h, bg, border, title, desc, title_col='#FFFFFF', desc_col='#E2E8F0'):
        box = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.1,rounding_size=0.15",
                                     linewidth=1.5, edgecolor=border, facecolor=bg)
        ax.add_patch(box)
        ax.text(x + w/2, y + h*0.65, title, fontsize=10, fontweight='bold', ha='center', va='center',
                color=title_col, fontfamily='sans-serif')
        ax.text(x + w/2, y + h*0.35, desc, fontsize=8, ha='center', va='center',
                color=desc_col, fontfamily='sans-serif')

    # Stage 1: Data Ingestion
    draw_box(0.5, 4.8, 2.5, 1.4, '#1E3A8A', '#1E40AF', 
             "1. DATA INGESTION", "student_input_*.xlsx\nCLI Argument / Default")

    # Arrow 1 -> 2
    ax.annotate("", xy=(3.4, 5.5), xytext=(3.0, 5.5),
                arrowprops=dict(arrowstyle="->", color="#3B82F6", lw=2.5))

    # Stage 2: Exception Scenario 1 (File/Schema Check)
    draw_box(3.4, 4.6, 2.8, 1.8, '#D97706', '#B45309',
             "2. SCENARIO 1 CHECK", "File Existence & Schema\nCheck Required Columns\n[try ... except]")

    # Branch down from Stage 2: File Error -> Error Log
    ax.annotate("", xy=(4.8, 2.7), xytext=(4.8, 4.6),
                arrowprops=dict(arrowstyle="->", color="#DC2626", lw=2, linestyle='dashed'))
    ax.text(5.0, 3.6, "Missing File /\nBad Schema", fontsize=8, color="#DC2626", fontweight='bold')

    # Arrow 2 -> 3
    ax.annotate("", xy=(6.6, 5.5), xytext=(6.2, 5.5),
                arrowprops=dict(arrowstyle="->", color="#10B981", lw=2.5))
    ax.text(6.4, 5.7, "Valid", fontsize=8, color="#059669", fontweight='bold', ha='center')

    # Stage 3: Row Iteration & Scenario 2 Check
    draw_box(6.6, 4.4, 3.2, 2.2, '#0D9488', '#0F766E',
             "3. SCENARIO 2 VALIDATION", "Row-Level Integrity Loop\nType Check (float)\nBoundary Check [0, 100]\n[try ... except ValueError]")

    # Branch down from Stage 3: Row Error -> Error Log
    ax.annotate("", xy=(8.2, 2.7), xytext=(8.2, 4.4),
                arrowprops=dict(arrowstyle="->", color="#DC2626", lw=2, linestyle='dashed'))
    ax.text(8.4, 3.6, "Dirty Row\n(Skip & Log)", fontsize=8, color="#DC2626", fontweight='bold')

    # Error Log Box at bottom
    draw_box(4.5, 1.3, 4.6, 1.4, '#FEF2F2', '#EF4444',
             "ERROR LOGGING SUBSYSTEM", "error_log.txt\nTimestamp, Scenario, Row #, Value, Trace",
             title_col='#991B1B', desc_col='#7F1D1D')

    # Arrow 3 -> 4
    ax.annotate("", xy=(10.2, 5.5), xytext=(9.8, 5.5),
                arrowprops=dict(arrowstyle="->", color="#10B981", lw=2.5))
    ax.text(10.0, 5.7, "Valid Row", fontsize=8, color="#059669", fontweight='bold', ha='center')

    # Stage 4: Academic Computation & Output Generation
    draw_box(10.2, 4.0, 3.3, 3.0, '#312E81', '#3730A3',
             "4. OUTPUT GENERATION", 
             "Academic Metrics Engine:\nTotal, %, Grade (A/B/C/Fail)\n\n"
             "• ReportLab PDF Engine:\n  Output_PDFs/ReportCard_*.pdf\n\n"
             "• Excel OpenPyXL Engine:\n  Processed_Results.xlsx")

    # Legend / Key Info at bottom left
    key_box = patches.FancyBboxPatch((0.5, 0.5), 3.5, 2.2, boxstyle="round,pad=0.08",
                                     linewidth=1, edgecolor='#CBD5E1', facecolor='#FFFFFF')
    ax.add_patch(key_box)
    ax.text(0.7, 2.4, "PROJECT HIGHLIGHTS", fontsize=9, fontweight='bold', color='#1E293B')
    ax.text(0.7, 2.0, "• CO1: Problem Logic & Flow", fontsize=8, color='#475569')
    ax.text(0.7, 1.6, "• CO2: Pandas Data Ingestion", fontsize=8, color='#475569')
    ax.text(0.7, 1.2, "• CO3: 2-Tier Try-Except Blocks", fontsize=8, color='#475569')
    ax.text(0.7, 0.8, "• CO4: Bulk PDF & Styled Excel", fontsize=8, color='#475569')

    # Legend / Verification note at bottom right
    ver_box = patches.FancyBboxPatch((9.7, 0.5), 3.8, 2.2, boxstyle="round,pad=0.08",
                                     linewidth=1, edgecolor='#CBD5E1', facecolor='#FFFFFF')
    ax.add_patch(ver_box)
    ax.text(9.9, 2.4, "BOUNDARY & EXCEPTION RULES", fontsize=9, fontweight='bold', color='#1E293B')
    ax.text(9.9, 2.0, "• Marks Range: [0, 100] per Subject", fontsize=8, color='#475569')
    ax.text(9.9, 1.6, "• Attendance: [0, 100] % Required", fontsize=8, color='#475569')
    ax.text(9.9, 1.2, "• Passing Mark: 35 / 100 per Subject", fontsize=8, color='#475569')
    ax.text(9.9, 0.8, "• Zero-Crash Guarantee: Safe Continuation", fontsize=8, color='#475569')

    workflow_path = Path("PRN1_PRN2_Name_CA3/Workflow/Workflow_Architecture.png")
    workflow_path.parent.mkdir(parents=True, exist_ok=True)
    plt.tight_layout()
    plt.savefig(workflow_path, dpi=300, facecolor=fig.get_facecolor(), edgecolor='none')
    plt.close()
    print(f"[SUCCESS] Workflow Architecture diagram saved to {workflow_path}")

if __name__ == "__main__":
    draw_workflow()
