import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

def create_rtm():
    wb = openpyxl.Workbook()

    # -------------------------------------------------------------
    # STYLES DEFINITION
    # -------------------------------------------------------------
    font_title = Font(name="Calibri", size=16, bold=True, color="1B365D")
    font_subtitle = Font(name="Calibri", size=11, italic=True, color="4A5568")
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_sub_header = Font(name="Calibri", size=11, bold=True, color="1B365D")
    font_body = Font(name="Calibri", size=9.5, color="2D3748")
    font_body_bold = Font(name="Calibri", size=9.5, bold=True, color="1B365D")
    font_future = Font(name="Calibri", size=9, italic=True, color="A0AEC0")

    fill_header = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    fill_future_header = PatternFill(start_color="4A5568", end_color="4A5568", fill_type="solid")
    fill_zebra_even = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_zebra_odd = PatternFill(start_color="F7FAFC", end_color="F7FAFC", fill_type="solid")

    fill_must = PatternFill(start_color="FED7D7", end_color="FED7D7", fill_type="solid") # soft red
    font_must = Font(name="Calibri", size=9.5, bold=True, color="9B2C2C")

    fill_should = PatternFill(start_color="BEE3F8", end_color="BEE3F8", fill_type="solid") # soft blue
    font_should = Font(name="Calibri", size=9.5, bold=True, color="2B6CB0")

    fill_could = PatternFill(start_color="EDF2F7", end_color="EDF2F7", fill_type="solid") # soft gray
    font_could = Font(name="Calibri", size=9.5, bold=True, color="4A5568")

    fill_status = PatternFill(start_color="E6FFFA", end_color="E6FFFA", fill_type="solid") # soft teal
    font_status = Font(name="Calibri", size=9.5, bold=True, color="234E52")

    thin_border_side = Side(border_style="thin", color="CBD5E0")
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_bottom_double = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=Side(border_style="double", color="1B365D"))

    # -------------------------------------------------------------
    # SHEET 1: RTM (Requirements Traceability Matrix)
    # -------------------------------------------------------------
    ws_rtm = wb.active
    ws_rtm.title = "RTM"
    ws_rtm.views.sheetView[0].showGridLines = True

    # Title Block
    ws_rtm["A1"] = "REQUIREMENTS TRACEABILITY MATRIX (RTM)"
    ws_rtm["A1"].font = font_title
    ws_rtm["A2"] = "Project: Insurance Claim Processing Analysis | Learner: Noman Zaid (2500320140103) | Baseline: Step 4 (BRD to Testing Pipeline)"
    ws_rtm["A2"].font = font_subtitle

    headers = [
        "OQ ID", 
        "HMW ID", 
        "Req ID (BR)", 
        "Requirement Description", 
        "Stakeholder (SH)", 
        "Priority", 
        "Status", 
        "Feature ID (Step 6)", 
        "DB table.column (Step 8)", 
        "User Story (Step 11)", 
        "UI Screen (Step 12)", 
        "Test Case (Step 13)"
    ]

    header_row = 4
    for col_idx, h in enumerate(headers, 1):
        cell = ws_rtm.cell(row=header_row, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_future_header if col_idx >= 8 else fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_cell
        ws_rtm.row_dimensions[header_row].height = 28

    # Rows (One row per HMW-BR link)
    rtm_data = [
        ("OQ-01", "HMW-01", "BR-01", "The business needs to track and dissect end-to-end and stage-wise turnaround time (submission to verification, verification to assessment, assessment to final decision) across insurance product lines so that processing bottlenecks are isolated and claims settlement latency is systematically reduced.", "SH-01", "Must", "Identified"),
        ("OQ-02", "HMW-02", "BR-02", "The business needs to classify and decompose claim repudiation reasons using frequency and Pareto analysis across policy lines so that the primary root causes of claim rejections can be identified, reducing avoidable customer disputes and policy misunderstandings.", "SH-04, SH-01", "Must", "Identified"),
        ("OQ-03", "HMW-03", "BR-03", "The business needs to evaluate claim settlement payout ratios (amount disbursed vs. amount requested) across policy categories and severity tiers so that unjustified payout deductions are prevented and accurate financial claim reserves are maintained.", "SH-05, SH-03", "Should", "Identified"),
        ("OQ-04", "HMW-04", "BR-04", "The business needs to benchmark empaneled healthcare hospitals, automobile garages, and third-party administrators (TPAs) on average claim costs, repudiation frequency, and processing duration so that abusive overbilling, tariff violations, and underperforming partners can be identified for contract renegotiation or de-empanelment.", "SH-02, SH-01", "Must", "Identified"),
        ("OQ-05", "HMW-05", "BR-05", "The business needs to automatically identify and score claims exhibiting high-risk fraud attributes (including claims filed within 30 days of policy inception, duplicate claims, and suspicious loss descriptions) so that potential fraud leakage is intercepted prior to financial disbursement.", "SH-02", "Must", "Identified"),
        ("OQ-06", "HMW-06", "BR-06", "The business needs to analyze monthly and seasonal claim volume fluctuations across product lines so that operational management can accurately forecast seasonal surges and dynamically reallocate claims adjudication staffing.", "SH-01", "Should", "Identified"),
        ("OQ-07", "HMW-07", "BR-07", "The business needs to isolate and model the statistical distribution of claim settlement amounts, identifying extreme positive outliers so that actuarial loss reserving models and reinsurance catastrophe triggers are accurately calibrated.", "SH-05, SH-03", "Should", "Identified"),
        ("OQ-08", "HMW-08", "BR-08", "The business needs to statistically test and quantify the turnaround time differential between cashless network claims and reimbursement claims so that executive leadership can evaluate the return on investment of expanding digital hospital and garage networks.", "SH-01, SH-05", "Must", "Identified"),
        ("OQ-09", "HMW-09", "BR-09", "The business needs to monitor customer satisfaction ratings and claim dispute escalation rates across geographic territories so that regional customer relationship teams and operational auditors can intervene in high-friction zones.", "SH-04", "Should", "Identified"),
        ("OQ-10", "HMW-10", "BR-10", "The business needs to correlate policyholder tenure, premium tier, and age demographic cohorts against claim frequency and loss severity so that underwriters can optimize risk pricing and policy renewal terms.", "SH-03", "Must", "Identified"),
        ("OQ-01", "HMW-01", "BR-11", "The business needs to maintain a data quality and validation log tracking missing values, duplicate entries, and data anomalies across claims records so that analytical decisions are founded on verified, trustworthy operational data.", "SH-01", "Must", "Identified"),
        ("OQ-05", "HMW-05", "BR-11", "The business needs to maintain a data quality and validation log tracking missing values, duplicate entries, and data anomalies across claims records so that analytical decisions are founded on verified, trustworthy operational data.", "SH-02, SH-05", "Must", "Identified"),
        ("OQ-01", "HMW-01", "BR-12", "The business needs to export filtered claims operational datasets, summary KPI tables, and compliance reports into standardized tabular formats (CSV/Excel) so that external regulatory auditors and executives can perform offline reviews.", "SH-01", "Could", "Identified"),
        ("OQ-03", "HMW-03", "BR-12", "The business needs to export filtered claims operational datasets, summary KPI tables, and compliance reports into standardized tabular formats (CSV/Excel) so that external regulatory auditors and executives can perform offline reviews.", "SH-05", "Could", "Identified")
    ]

    start_row = 5
    for idx, row_item in enumerate(rtm_data):
        curr_row = start_row + idx
        bg_fill = fill_zebra_even if idx % 2 == 0 else fill_zebra_odd
        ws_rtm.row_dimensions[curr_row].height = 42

        oq, hmw, br, req_desc, sh, prio, status = row_item
        
        # Populate Identified columns
        c_oq = ws_rtm.cell(row=curr_row, column=1, value=oq)
        c_hmw = ws_rtm.cell(row=curr_row, column=2, value=hmw)
        c_br = ws_rtm.cell(row=curr_row, column=3, value=br)
        c_desc = ws_rtm.cell(row=curr_row, column=4, value=req_desc)
        c_sh = ws_rtm.cell(row=curr_row, column=5, value=sh)
        c_prio = ws_rtm.cell(row=curr_row, column=6, value=prio)
        c_status = ws_rtm.cell(row=curr_row, column=7, value=status)

        # Populate Empty Future columns
        c_feat = ws_rtm.cell(row=curr_row, column=8, value="[Pending Step 6]")
        c_db = ws_rtm.cell(row=curr_row, column=9, value="[Pending Step 8]")
        c_us = ws_rtm.cell(row=curr_row, column=10, value="[Pending Step 11]")
        c_ui = ws_rtm.cell(row=curr_row, column=11, value="[Pending Step 12]")
        c_tc = ws_rtm.cell(row=curr_row, column=12, value="[Pending Step 13]")

        for c in [c_oq, c_hmw, c_br, c_desc, c_sh, c_prio, c_status]:
            c.border = border_cell
            c.fill = bg_fill
            c.font = font_body
            c.alignment = Alignment(vertical="center")

        c_oq.alignment = Alignment(horizontal="center", vertical="center")
        c_oq.font = font_body_bold
        c_hmw.alignment = Alignment(horizontal="center", vertical="center")
        c_hmw.font = font_body_bold
        c_br.alignment = Alignment(horizontal="center", vertical="center")
        c_br.font = font_body_bold
        c_desc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        c_sh.alignment = Alignment(horizontal="center", vertical="center")

        # Priority styling
        if prio == "Must":
            c_prio.fill = fill_must
            c_prio.font = font_must
        elif prio == "Should":
            c_prio.fill = fill_should
            c_prio.font = font_should
        else:
            c_prio.fill = fill_could
            c_prio.font = font_could
        c_prio.alignment = Alignment(horizontal="center", vertical="center")

        # Status styling
        c_status.fill = fill_status
        c_status.font = font_status
        c_status.alignment = Alignment(horizontal="center", vertical="center")

        # Future columns styling
        for c in [c_feat, c_db, c_us, c_ui, c_tc]:
            c.border = border_cell
            c.fill = bg_fill
            c.font = font_future
            c.alignment = Alignment(horizontal="center", vertical="center")

    # Column Widths for RTM sheet
    col_widths_rtm = {
        1: 12,  # OQ
        2: 13,  # HMW
        3: 13,  # Req ID
        4: 48,  # Requirement
        5: 16,  # Stakeholder
        6: 12,  # Priority
        7: 14,  # Status
        8: 18,  # Feature
        9: 24,  # DB table.column
        10: 18, # User Story
        11: 18, # UI Screen
        12: 22  # Test Case
    }
    for col_idx, width in col_widths_rtm.items():
        col_letter = get_column_letter(col_idx)
        ws_rtm.column_dimensions[col_letter].width = width

    # Enable AutoFilter on header row
    ws_rtm.auto_filter.ref = f"A4:L{start_row + len(rtm_data) - 1}"
    ws_rtm.freeze_panes = "A5"

    # -------------------------------------------------------------
    # SHEET 2: Coverage_Check
    # -------------------------------------------------------------
    ws_cov = wb.create_sheet(title="Coverage_Check")
    ws_cov.views.sheetView[0].showGridLines = True

    ws_cov["A1"] = "REQUIREMENTS TRACEABILITY & COVERAGE AUDIT"
    ws_cov["A1"].font = font_title
    ws_cov["A2"] = "Verification of Bidirectional Alignment: Operational Questions, HMWs, and Business Requirements"
    ws_cov["A2"].font = font_subtitle

    # Section 1: KPI Summary Cards Table
    ws_cov["A4"] = "1. EXECUTIVE COVERAGE SCORECARD"
    ws_cov["A4"].font = font_sub_header

    cov_kpis = [
        ("Metric Description", "Baseline Count", "Mapped / Covered", "Coverage Ratio", "Audit Status"),
        ("Operational Questions (OQ)", 10, 10, "100.0%", "PASSED - Full Traceability"),
        ("Design Thinking Statements (HMW)", 10, 10, "100.0%", "PASSED - Full Traceability"),
        ("Business Requirements (BR)", 12, 12, "100.0%", "PASSED - Full Traceability"),
        ("Must-Have Core MVP Requirements", 7, 7, "58.3% (~60%)", "PASSED - MoSCoW Standard"),
        ("HMW Statements with No Linked BR (Orphan HMW)", 0, 0, "0.0%", "PASSED - Zero Defects"),
        ("Business Requirements with No Source HMW (Orphan BR)", 0, 0, "0.0%", "PASSED - Zero Defects"),
        ("Uncovered Operational Questions (Orphan OQ)", 0, 0, "0.0%", "PASSED - Zero Defects")
    ]

    for r_idx, row_vals in enumerate(cov_kpis, 5):
        ws_cov.row_dimensions[r_idx].height = 24
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws_cov.cell(row=r_idx, column=c_idx, value=val)
            cell.border = border_cell
            if r_idx == 5:
                cell.fill = fill_header
                cell.font = font_header
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.font = font_body
                if c_idx == 1:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                    cell.font = font_body_bold
                elif c_idx in [2, 3, 4]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.fill = fill_status
                    cell.font = font_status

    # Section 2: Detailed OQ to BR Mapping Table
    ws_cov["A15"] = "2. OPERATIONAL QUESTION (OQ) COVERAGE MATRIX"
    ws_cov["A15"].font = font_sub_header

    oq_cov_headers = ["OQ ID", "Operational Question Title", "Linked HMW", "Linked BR(s)", "Key Stakeholder", "Coverage Status"]
    ws_cov.row_dimensions[16].height = 24
    for c_idx, h in enumerate(oq_cov_headers, 1):
        cell = ws_cov.cell(row=16, column=c_idx, value=h)
        cell.fill = fill_header
        cell.font = font_header
        cell.border = border_cell
        cell.alignment = Alignment(horizontal="center", vertical="center")

    oq_matrix = [
        ("OQ-01", "Turnaround Time (TAT) & Processing Bottlenecks", "HMW-01", "BR-01, BR-11, BR-12", "SH-01", "Covered"),
        ("OQ-02", "Claim Approval vs. Repudiation & Denial Reasons", "HMW-02", "BR-02", "SH-04, SH-01", "Covered"),
        ("OQ-03", "Settlement Payout Ratio & Haircut Deductions", "HMW-03", "BR-03, BR-12", "SH-05, SH-03", "Covered"),
        ("OQ-04", "Network Provider (Hospital/Garage) Cost & TAT", "HMW-04", "BR-04", "SH-02, SH-01", "Covered"),
        ("OQ-05", "Fraud Risk Scoring & Financial Exposure", "HMW-05", "BR-05, BR-11", "SH-02", "Covered"),
        ("OQ-06", "Monthly & Seasonal Claim Surge Forecasting", "HMW-06", "BR-06", "SH-01", "Covered"),
        ("OQ-07", "Extreme Outlier Claims Distribution & Reserves", "HMW-07", "BR-07", "SH-05, SH-03", "Covered"),
        ("OQ-08", "Cashless vs. Reimbursement Statistical Hypothesis", "HMW-08", "BR-08", "SH-01, SH-05", "Covered"),
        ("OQ-09", "Geographic Territory Grievance & Escalation Rates", "HMW-09", "BR-09", "SH-04", "Covered"),
        ("OQ-10", "Policyholder Demographic & Tenure Risk Correlation", "HMW-10", "BR-10", "SH-03", "Covered")
    ]

    for idx, oq_row in enumerate(oq_matrix):
        curr_r = 17 + idx
        ws_cov.row_dimensions[curr_r].height = 22
        bg = fill_zebra_even if idx % 2 == 0 else fill_zebra_odd
        for c_idx, val in enumerate(oq_row, 1):
            cell = ws_cov.cell(row=curr_r, column=c_idx, value=val)
            cell.border = border_cell
            cell.fill = bg
            cell.font = font_body
            if c_idx in [1, 3, 4, 5]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if c_idx == 1:
                    cell.font = font_body_bold
            elif c_idx == 2:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif c_idx == 6:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.fill = fill_status
                cell.font = font_status

    # Column widths for Coverage sheet
    col_widths_cov = {
        1: 14,
        2: 44,
        3: 16,
        4: 24,
        5: 18,
        6: 18
    }
    for col_idx, width in col_widths_cov.items():
        col_letter = get_column_letter(col_idx)
        ws_cov.column_dimensions[col_letter].width = width

    output_path = r"c:\Users\noman\OneDrive\Desktop\python\hcl-pro\docs\04_Traceability_Matrix.xlsx"
    wb.save(output_path)
    print(f"Successfully generated {output_path}")

if __name__ == "__main__":
    create_rtm()
