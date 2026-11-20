import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

def create_database_design_workbook():
    wb = openpyxl.Workbook()

    # Base Styles
    font_title = Font(name="Calibri", size=16, bold=True, color="1B365D")
    font_subtitle = Font(name="Calibri", size=11, italic=True, color="4A5568")
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_sub_header = Font(name="Calibri", size=11, bold=True, color="1B365D")
    font_body = Font(name="Calibri", size=9.5, color="2D3748")
    font_body_bold = Font(name="Calibri", size=9.5, bold=True, color="1B365D")
    font_code = Font(name="Consolas", size=9, color="1A202C")

    fill_header = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    fill_sub_header = PatternFill(start_color="2B6CB0", end_color="2B6CB0", fill_type="solid")
    fill_zebra_even = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_zebra_odd = PatternFill(start_color="F7FAFC", end_color="F7FAFC", fill_type="solid")
    fill_pk = PatternFill(start_color="FED7D7", end_color="FED7D7", fill_type="solid") # light red
    fill_fk = PatternFill(start_color="FEFCBF", end_color="FEFCBF", fill_type="solid") # light yellow
    fill_active = PatternFill(start_color="E6FFFA", end_color="E6FFFA", fill_type="solid") # light teal

    thin_border_side = Side(border_style="thin", color="CBD5E0")
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    # -------------------------------------------------------------
    # SHEET 1: Overview
    # -------------------------------------------------------------
    ws1 = wb.active
    ws1.title = "Overview"
    ws1.views.sheetView[0].showGridLines = True

    ws1["A1"] = "DATABASE ARCHITECTURE & DESIGN SPECIFICATION"
    ws1["A1"].font = font_title
    ws1["A2"] = "Project: Insurance Claim Processing Analysis | Stack: SQLite 3 (data/app.db) | Model: Relational Star-Snowflake (3NF Normalized)"
    ws1["A2"].font = font_subtitle

    ws1["A4"] = "1. DESIGN PHILOSOPHY & NORMALIZATION"
    ws1["A4"].font = font_sub_header

    overview_text = [
        ("Architecture Pattern", "Hybrid Star-Relational schema modeled in Third Normal Form (3NF) to eliminate update anomalies while preserving high-speed analytical joins."),
        ("Storage Engine", "Embedded SQLite 3 datastore stored directly in the repository at data/app.db, zero-configuration and platform independent."),
        ("Concurrency & Logging", "Configured with PRAGMA journal_mode=WAL (Write-Ahead Logging) and PRAGMA foreign_keys=ON to ensure strict ACID integrity and rapid concurrent reads."),
        ("Analytical View (Denormalized)", "v_claims_analysis_master joining all 4 tables into a comprehensive 36-column flat analytical view consumed directly by Streamlit pages."),
        ("OQ & BR Traceability", "Every column explicitly maps to an Operational Question (OQ-01 to OQ-10) and Business Requirement (BR-01 to BR-12), ensuring zero orphaned schema attributes.")
    ]

    for idx, (k, v) in enumerate(overview_text, 5):
        ws1.row_dimensions[idx].height = 24
        c1 = ws1.cell(row=idx, column=1, value=k)
        c2 = ws1.cell(row=idx, column=2, value=v)
        c1.font = font_body_bold
        c1.fill = fill_zebra_odd
        c1.border = border_cell
        c2.font = font_body
        c2.fill = fill_zebra_even
        c2.border = border_cell

    ws1["A12"] = "2. ENTITY RELATIONSHIP SUMMARY"
    ws1["A12"].font = font_sub_header

    ws1.column_dimensions["A"].width = 28
    ws1.column_dimensions["B"].width = 90

    # -------------------------------------------------------------
    # SHEET 2: Tables_Summary
    # -------------------------------------------------------------
    ws2 = wb.create_sheet(title="Tables_Summary")
    ws2.views.sheetView[0].showGridLines = True

    ws2["A1"] = "DATABASE TABLES SUMMARY"
    ws2["A1"].font = font_title
    ws2["A2"] = "Structural summary of entities, grains, keys, and row volumes"
    ws2["A2"].font = font_subtitle

    t_headers = ["Table Name", "Entity Type", "Grain / Level of Detail", "Primary Key (PK)", "Foreign Keys (FK)", "Column Count", "Planned Rows", "Core Purpose"]
    ws2.row_dimensions[4].height = 26
    for c_idx, h in enumerate(t_headers, 1):
        c = ws2.cell(row=4, column=c_idx, value=h)
        c.fill = fill_header
        c.font = font_header
        c.border = border_cell
        c.alignment = Alignment(horizontal="center", vertical="center")

    summary_rows = [
        ("policyholders_master", "Dimension / Master", "One row per unique insurance policyholder", "policyholder_id", "None", 8, "100", "Stores demographic attributes: age, gender, geographic location, tenure."),
        ("policies_master", "Dimension / Master", "One row per insurance contract issued", "policy_id", "policyholder_id", 8, "120", "Stores policy contractual terms: type, sum insured, annual premium, start date."),
        ("providers_master", "Dimension / Master", "One row per empaneled healthcare hospital or garage", "provider_id", "None", 7, "25", "Stores service partner details: facility name, category, region, empanelment status."),
        ("claims_fact", "Fact Table (Main)", "One row per insurance claim event submitted", "claim_id", "policy_id, provider_id", 20, "250", "Captures transactional lifecycle: amounts, dates, stage TATs, fraud scores, status."),
        ("v_claims_analysis_master", "Analytical View (SQL View)", "One row per enriched claim transaction", "claim_id", "N/A (Joined)", 36, "250", "Denormalized view joining fact + dimensions for Streamlit visual analytics.")
    ]

    for r_idx, r_vals in enumerate(summary_rows, 5):
        ws2.row_dimensions[r_idx].height = 28
        bg = fill_zebra_even if r_idx % 2 == 1 else fill_zebra_odd
        for c_idx, val in enumerate(r_vals, 1):
            c = ws2.cell(row=r_idx, column=c_idx, value=val)
            c.border = border_cell
            c.fill = bg
            c.font = font_body
            if c_idx in [1, 2, 4, 5, 6, 7]:
                c.alignment = Alignment(horizontal="center", vertical="center")
                if c_idx == 1:
                    c.font = font_body_bold
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")

    widths_ws2 = {1: 26, 2: 20, 3: 36, 4: 18, 5: 22, 6: 14, 7: 14, 8: 48}
    for col_idx, width in widths_ws2.items():
        ws2.column_dimensions[get_column_letter(col_idx)].width = width

    # Helper function for data dictionaries
    def add_data_dict_sheet(sheet_title, entity_title, columns_data):
        ws = wb.create_sheet(title=sheet_title)
        ws.views.sheetView[0].showGridLines = True
        ws["A1"] = f"DATA DICTIONARY: {entity_title}"
        ws["A1"].font = font_title
        ws["A2"] = f"Detailed technical column attributes, constraints, and business traceability"
        ws["A2"].font = font_subtitle

        dd_headers = ["Column Name", "SQLite Data Type", "Key Type", "Nullable", "Allowed Values / Range", "Column Description & Business Meaning", "Linked OQ / BR"]
        ws.row_dimensions[4].height = 26
        for c_idx, h in enumerate(dd_headers, 1):
            c = ws.cell(row=4, column=c_idx, value=h)
            c.fill = fill_header
            c.font = font_header
            c.border = border_cell
            c.alignment = Alignment(horizontal="center", vertical="center")

        for r_idx, r_vals in enumerate(columns_data, 5):
            ws.row_dimensions[r_idx].height = 24
            bg = fill_zebra_even if r_idx % 2 == 1 else fill_zebra_odd
            for c_idx, val in enumerate(r_vals, 1):
                c = ws.cell(row=r_idx, column=c_idx, value=val)
                c.border = border_cell
                c.fill = bg
                c.font = font_body
                if c_idx == 1:
                    c.font = font_code
                elif c_idx in [2, 3, 4]:
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    if val == "PK":
                        c.fill = fill_pk
                        c.font = font_body_bold
                    elif "FK" in str(val):
                        c.fill = fill_fk
                        c.font = font_body_bold
                elif c_idx == 7:
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    c.fill = fill_active
                    c.font = font_body_bold
                else:
                    c.alignment = Alignment(horizontal="left", vertical="center")

        widths = {1: 26, 2: 16, 3: 12, 4: 12, 5: 30, 6: 45, 7: 20}
        for col_idx, width in widths.items():
            ws.column_dimensions[get_column_letter(col_idx)].width = width

    # -------------------------------------------------------------
    # SHEET 3: Data_Dict_policyholders
    # -------------------------------------------------------------
    policyholders_cols = [
        ("policyholder_id", "VARCHAR(20)", "PK", "NOT NULL", "POL-HLD-1001 to POL-HLD-9999", "Unique alphanumeric identifier for insured customer", "BR-10 / OQ-10"),
        ("full_name", "VARCHAR(100)", "None", "NOT NULL", "Realistic Indian names (Faker)", "Full legal name of the insured citizen", "BR-11 / Demographics"),
        ("age", "INTEGER", "None", "NOT NULL", "18 to 75", "Chronological age of policyholder in completed years", "BR-10 / OQ-10"),
        ("gender", "VARCHAR(10)", "None", "NOT NULL", "'Male', 'Female', 'Other'", "Gender identification of policyholder", "BR-10 / OQ-10"),
        ("city", "VARCHAR(50)", "None", "NOT NULL", "Tier 1 & 2 Indian Cities (Delhi, Mumbai, etc.)", "Residential city of the policyholder", "BR-09 / OQ-09"),
        ("state", "VARCHAR(50)", "None", "NOT NULL", "Indian States (Maharashtra, UP, Karnataka, etc.)", "State jurisdiction of the policyholder", "BR-09 / OQ-09"),
        ("region", "VARCHAR(20)", "None", "NOT NULL", "'North', 'South', 'East', 'West', 'Central'", "Macro geographic zone for regional clustering", "BR-09 / OQ-09"),
        ("tenure_months", "INTEGER", "None", "NOT NULL", "1 to 120", "Duration of customer relationship with insurer in months", "BR-10 / OQ-10")
    ]
    add_data_dict_sheet("Dict_Policyholders", "policyholders_master", policyholders_cols)

    # -------------------------------------------------------------
    # SHEET 4: Data_Dict_policies
    # -------------------------------------------------------------
    policies_cols = [
        ("policy_id", "VARCHAR(20)", "PK", "NOT NULL", "POL-10001 to POL-99999", "Unique alphanumeric identifier for policy contract", "BR-01 to 10"),
        ("policyholder_id", "VARCHAR(20)", "FK", "NOT NULL", "Valid policyholders_master.policyholder_id", "Foreign key reference to policy owner", "BR-10 / OQ-10"),
        ("policy_type", "VARCHAR(20)", "None", "NOT NULL", "'Health', 'Motor', 'Life', 'Property'", "Primary insurance business product line", "BR-01, 02, 03, 06"),
        ("policy_tier", "VARCHAR(20)", "None", "NOT NULL", "'Silver', 'Gold', 'Platinum'", "Coverage tier determining deductible and benefits", "BR-10 / OQ-10"),
        ("sum_insured", "DECIMAL(12,2)", "None", "NOT NULL", "100,000.00 to 5,000,000.00", "Maximum contractual indemnity liability limit (INR)", "BR-07 / OQ-07"),
        ("annual_premium", "DECIMAL(10,2)", "None", "NOT NULL", "3,500.00 to 120,000.00", "Annual premium paid by policyholder (INR)", "BR-10 / OQ-10"),
        ("policy_start_date", "DATE", "None", "NOT NULL", "YYYY-MM-DD (2023-01-01 to 2025-12-31)", "Effective inception date of policy coverage", "BR-05 / OQ-05"),
        ("policy_status", "VARCHAR(20)", "None", "NOT NULL", "'Active', 'Lapsed', 'Renewed'", "Operational status of policy contract", "BR-11 / Master")
    ]
    add_data_dict_sheet("Dict_Policies", "policies_master", policies_cols)

    # -------------------------------------------------------------
    # SHEET 5: Data_Dict_providers
    # -------------------------------------------------------------
    providers_cols = [
        ("provider_id", "VARCHAR(20)", "PK", "NOT NULL", "PRV-101 to PRV-999", "Unique identifier for empaneled network facility", "BR-04 / OQ-04"),
        ("provider_name", "VARCHAR(100)", "None", "NOT NULL", "Hospital & Auto Garage business names", "Legal business entity name of the provider", "BR-04 / OQ-04"),
        ("provider_type", "VARCHAR(30)", "None", "NOT NULL", "'Hospital', 'Automobile Garage', 'TPA'", "Category of healthcare or motor service provider", "BR-04 / OQ-04"),
        ("city", "VARCHAR(50)", "None", "NOT NULL", "Indian Cities", "Operating city of provider facility", "BR-04, 09"),
        ("state", "VARCHAR(50)", "None", "NOT NULL", "Indian States", "Operating state of provider facility", "BR-04, 09"),
        ("empanelment_status", "VARCHAR(20)", "None", "NOT NULL", "'Empaneled Network', 'Non-Network', 'Under Review'", "Contractual network partnership standing", "BR-04 / OQ-04"),
        ("tariff_tier", "VARCHAR(10)", "None", "NOT NULL", "'Tier 1', 'Tier 2', 'Tier 3'", "Agreed package rate billing schedule tier", "BR-04 / OQ-04")
    ]
    add_data_dict_sheet("Dict_Providers", "providers_master", providers_cols)

    # -------------------------------------------------------------
    # SHEET 6: Data_Dict_claims
    # -------------------------------------------------------------
    claims_cols = [
        ("claim_id", "VARCHAR(20)", "PK", "NOT NULL", "CLM-10001 to CLM-99999", "Primary key for individual claim transaction", "BR-01 to 12 / All OQs"),
        ("policy_id", "VARCHAR(20)", "FK", "NOT NULL", "Valid policies_master.policy_id", "Foreign key linking claim to insurance contract", "BR-01 to 10"),
        ("provider_id", "VARCHAR(20)", "FK", "NOT NULL", "Valid providers_master.provider_id", "Foreign key linking claim to servicing facility", "BR-04 / OQ-04"),
        ("incident_date", "DATE", "None", "NOT NULL", "YYYY-MM-DD", "Date of occurrence of covered loss event", "BR-05 / OQ-05"),
        ("submission_date", "DATE", "None", "NOT NULL", "YYYY-MM-DD (>= incident_date)", "Date claim documents received by insurer", "BR-01, 06 / OQ-01, 06"),
        ("verification_date", "DATE", "None", "NULLABLE", "YYYY-MM-DD (>= submission_date)", "Date document intake verification completed", "BR-01 / OQ-01"),
        ("assessment_date", "DATE", "None", "NULLABLE", "YYYY-MM-DD (>= verification_date)", "Date medical/technical audit completed", "BR-01 / OQ-01"),
        ("settlement_date", "DATE", "None", "NULLABLE", "YYYY-MM-DD (>= assessment_date)", "Date final approval/repudiation executed", "BR-01 / OQ-01"),
        ("claim_amount_requested", "DECIMAL(12,2)", "None", "NOT NULL", "5,000.00 to 2,500,000.00", "Gross amount claimed by policyholder (INR)", "BR-03, 05 / OQ-03, 05"),
        ("claim_amount_settled", "DECIMAL(12,2)", "None", "NOT NULL", "0.00 to 2,500,000.00 (<= requested)", "Final net amount disbursed by insurer (INR)", "BR-03, 07 / OQ-03, 07"),
        ("settlement_ratio", "DECIMAL(5,4)", "None", "NOT NULL", "0.0000 to 1.0000 (settled / requested)", "Payout haircut fraction disbursed to claimant", "BR-03 / OQ-03"),
        ("settlement_mode", "VARCHAR(20)", "None", "NOT NULL", "'Cashless', 'Reimbursement'", "Operational payment channel used", "BR-08 / OQ-08"),
        ("adjudication_stage", "VARCHAR(30)", "None", "NOT NULL", "'Submission', 'Verification', 'Assessment', 'Settled'", "Current operational workflow state", "BR-01 / OQ-01"),
        ("claim_status", "VARCHAR(20)", "None", "NOT NULL", "'Approved', 'Repudiated', 'Pending Review'", "Final claim adjudication decision", "BR-02 / OQ-02"),
        ("rejection_reason_code", "VARCHAR(20)", "None", "NULLABLE", "'EXCL_PRE_EXIST', 'DOC_FRAUD', 'LACK_PREAUTH', etc.", "Standardized IRDAI denial code for repudiations", "BR-02 / OQ-02"),
        ("rejection_reason_detail", "VARCHAR(150)", "None", "NULLABLE", "Plain-English explanation of rejection ground", "Full contextual description of repudiation ground", "BR-02 / OQ-02"),
        ("tat_days", "INTEGER", "None", "NOT NULL", "1 to 65", "Total elapsed days from submission to decision", "BR-01, 08 / OQ-01, 08"),
        ("fraud_risk_score", "INTEGER", "None", "NOT NULL", "0 to 100", "Composite anomaly score from fraud engine", "BR-05 / OQ-05"),
        ("fraud_flag", "BOOLEAN", "None", "NOT NULL", "0 (Normal) or 1 (High Risk / Fraud Anomaly)", "Binary flag for mandatory SIU investigation", "BR-05 / OQ-05"),
        ("csat_score", "INTEGER", "None", "NULLABLE", "1 to 5", "Post-claim customer satisfaction survey rating", "BR-09 / OQ-09")
    ]
    add_data_dict_sheet("Dict_Claims_Fact", "claims_fact", claims_cols)

    # -------------------------------------------------------------
    # SHEET 7: Data_Dict_Joined_View (36 Columns)
    # -------------------------------------------------------------
    ws7 = wb.create_sheet(title="Joined_Analysis_View")
    ws7.views.sheetView[0].showGridLines = True
    ws7["A1"] = "DENORMALIZED JOINED ANALYSIS VIEW (v_claims_analysis_master)"
    ws7["A1"].font = font_title
    ws7["A2"] = "Complete 36-column analytical schema consumed by Streamlit Visualizations"
    ws7["A2"].font = font_subtitle

    view_headers = ["View Column #", "View Column Name", "Source Table.Column", "Data Type", "Analytical Role", "Linked OQ / BR"]
    ws7.row_dimensions[4].height = 26
    for c_idx, h in enumerate(view_headers, 1):
        c = ws7.cell(row=4, column=c_idx, value=h)
        c.fill = fill_header
        c.font = font_header
        c.border = border_cell
        c.alignment = Alignment(horizontal="center", vertical="center")

    view_columns = [
        (1, "claim_id", "claims_fact.claim_id", "VARCHAR(20)", "Primary Analytical Transaction Key", "All OQs"),
        (2, "policy_id", "claims_fact.policy_id", "VARCHAR(20)", "Contract Foreign Key", "BR-01 to 10"),
        (3, "provider_id", "claims_fact.provider_id", "VARCHAR(20)", "Facility Foreign Key", "BR-04"),
        (4, "policyholder_id", "policies_master.policyholder_id", "VARCHAR(20)", "Customer Foreign Key", "BR-10"),
        (5, "full_name", "policyholders_master.full_name", "VARCHAR(100)", "Customer Identity", "BR-11"),
        (6, "age", "policyholders_master.age", "INTEGER", "Demographic Factor (Underwriting)", "OQ-10 / BR-10"),
        (7, "gender", "policyholders_master.gender", "VARCHAR(10)", "Demographic Slicing Factor", "OQ-10 / BR-10"),
        (8, "customer_city", "policyholders_master.city", "VARCHAR(50)", "Customer Location Dimension", "OQ-09 / BR-09"),
        (9, "customer_state", "policyholders_master.state", "VARCHAR(50)", "Geographic Heatmap Slicing", "OQ-09 / BR-09"),
        (10, "customer_region", "policyholders_master.region", "VARCHAR(20)", "Macro Zone Filter (North/South/etc.)", "OQ-09 / BR-09"),
        (11, "tenure_months", "policyholders_master.tenure_months", "INTEGER", "Customer Loyalty Metric", "OQ-10 / BR-10"),
        (12, "policy_type", "policies_master.policy_type", "VARCHAR(20)", "Primary Sidebar Filter (Health/Motor)", "OQ-01, 02, 03"),
        (13, "policy_tier", "policies_master.policy_tier", "VARCHAR(20)", "Plan Coverage Category", "OQ-10 / BR-10"),
        (14, "sum_insured", "policies_master.sum_insured", "DECIMAL(12,2)", "Coverage Exposure Ceiling", "OQ-07 / BR-07"),
        (15, "annual_premium", "policies_master.annual_premium", "DECIMAL(10,2)", "Underwriting Premium Revenue", "OQ-10 / BR-10"),
        (16, "policy_start_date", "policies_master.policy_start_date", "DATE", "Inception Date for Fraud Delta", "OQ-05 / BR-05"),
        (17, "provider_name", "providers_master.provider_name", "VARCHAR(100)", "Facility Benchmarking Name", "OQ-04 / BR-04"),
        (18, "provider_type", "providers_master.provider_type", "VARCHAR(30)", "Facility Category (Hospital/Garage)", "OQ-04 / BR-04"),
        (19, "empanelment_status", "providers_master.empanelment_status", "VARCHAR(20)", "Network Quality Standing", "OQ-04 / BR-04"),
        (20, "provider_city", "providers_master.city", "VARCHAR(50)", "Facility Location", "OQ-04 / BR-04"),
        (21, "incident_date", "claims_fact.incident_date", "DATE", "Loss Event Date", "OQ-05 / BR-05"),
        (22, "submission_date", "claims_fact.submission_date", "DATE", "Inflow Time-series Date", "OQ-01, 06"),
        (23, "verification_date", "claims_fact.verification_date", "DATE", "Milestone 1 Completion Date", "OQ-01 / BR-01"),
        (24, "assessment_date", "claims_fact.assessment_date", "DATE", "Milestone 2 Completion Date", "OQ-01 / BR-01"),
        (25, "settlement_date", "claims_fact.settlement_date", "DATE", "Adjudication Decision Date", "OQ-01 / BR-01"),
        (26, "claim_amount_requested", "claims_fact.claim_amount_requested", "DECIMAL(12,2)", "Gross Claim Billed", "OQ-03 / BR-03"),
        (27, "claim_amount_settled", "claims_fact.claim_amount_settled", "DECIMAL(12,2)", "Net Payout Disbursed", "OQ-03, 07"),
        (28, "settlement_ratio", "claims_fact.settlement_ratio", "DECIMAL(5,4)", "Payout Haircut Ratio", "OQ-03 / BR-03"),
        (29, "settlement_mode", "claims_fact.settlement_mode", "VARCHAR(20)", "Hypothesis Test Group (Cashless/Reimb)", "OQ-08 / BR-08"),
        (30, "adjudication_stage", "claims_fact.adjudication_stage", "VARCHAR(30)", "Stage Bottleneck Classification", "OQ-01 / BR-01"),
        (31, "claim_status", "claims_fact.claim_status", "VARCHAR(20)", "Approved vs. Repudiated Split", "OQ-02 / BR-02"),
        (32, "rejection_reason_code", "claims_fact.rejection_reason_code", "VARCHAR(20)", "Pareto Denial Cause Code", "OQ-02 / BR-02"),
        (33, "rejection_reason_detail", "claims_fact.rejection_reason_detail", "VARCHAR(150)", "Pareto Denial Text", "OQ-02 / BR-02"),
        (34, "tat_days", "claims_fact.tat_days", "INTEGER", "End-to-End Latency Target", "OQ-01, 08"),
        (35, "fraud_risk_score", "claims_fact.fraud_risk_score", "INTEGER", "Composite Anomaly Score", "OQ-05 / BR-05"),
        (36, "csat_score", "claims_fact.csat_score", "INTEGER", "Customer Grievance Metric", "OQ-09 / BR-09")
    ]

    for r_idx, r_vals in enumerate(view_columns, 5):
        ws7.row_dimensions[r_idx].height = 22
        bg = fill_zebra_even if r_idx % 2 == 1 else fill_zebra_odd
        for c_idx, val in enumerate(r_vals, 1):
            c = ws7.cell(row=r_idx, column=c_idx, value=val)
            c.border = border_cell
            c.fill = bg
            c.font = font_body
            if c_idx == 1:
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.font = font_body_bold
            elif c_idx in [2, 3]:
                c.font = font_code
                c.alignment = Alignment(horizontal="left", vertical="center")
            elif c_idx == 4:
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx == 6:
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.fill = fill_active
                c.font = font_body_bold
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")

    widths_ws7 = {1: 14, 2: 26, 3: 35, 4: 16, 5: 35, 6: 18}
    for col_idx, width in widths_ws7.items():
        ws7.column_dimensions[get_column_letter(col_idx)].width = width

    # -------------------------------------------------------------
    # SHEET 8: Relationships
    # -------------------------------------------------------------
    ws8 = wb.create_sheet(title="Relationships")
    ws8.views.sheetView[0].showGridLines = True
    ws8["A1"] = "DATABASE RELATIONSHIPS & CARDINALITY"
    ws8["A1"].font = font_title
    ws8["A2"] = "Referential integrity constraints, parent-child links, and business enforcement rules"
    ws8["A2"].font = font_subtitle

    rel_headers = ["Relationship ID", "Parent Table", "Child Table", "Join Key (Parent -> Child)", "Cardinality", "Integrity Rule", "Plain-English Relationship Statement"]
    ws8.row_dimensions[4].height = 26
    for c_idx, h in enumerate(rel_headers, 1):
        c = ws8.cell(row=4, column=c_idx, value=h)
        c.fill = fill_header
        c.font = font_header
        c.border = border_cell
        c.alignment = Alignment(horizontal="center", vertical="center")

    rel_data = [
        ("REL-01", "policyholders_master", "policies_master", "policyholder_id -> policyholder_id", "1 : N (One-to-Many)", "ON DELETE RESTRICT, ON UPDATE CASCADE", "One policyholder can purchase multiple insurance policy contracts over time."),
        ("REL-02", "policies_master", "claims_fact", "policy_id -> policy_id", "1 : N (One-to-Many)", "ON DELETE RESTRICT, ON UPDATE CASCADE", "One policy contract can experience multiple insurance claim events during its coverage term."),
        ("REL-03", "providers_master", "claims_fact", "provider_id -> provider_id", "1 : N (One-to-Many)", "ON DELETE RESTRICT, ON UPDATE CASCADE", "One empaneled hospital or garage can service and bill multiple policyholder claims.")
    ]

    for r_idx, r_vals in enumerate(rel_data, 5):
        ws8.row_dimensions[r_idx].height = 32
        bg = fill_zebra_even if r_idx % 2 == 1 else fill_zebra_odd
        for c_idx, val in enumerate(r_vals, 1):
            c = ws8.cell(row=r_idx, column=c_idx, value=val)
            c.border = border_cell
            c.fill = bg
            c.font = font_body
            if c_idx in [1, 2, 3, 5, 6]:
                c.alignment = Alignment(horizontal="center", vertical="center")
                if c_idx == 1:
                    c.font = font_body_bold
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")

    widths_ws8 = {1: 16, 2: 24, 3: 20, 4: 32, 5: 20, 6: 35, 7: 50}
    for col_idx, width in widths_ws8.items():
        ws8.column_dimensions[get_column_letter(col_idx)].width = width

    # -------------------------------------------------------------
    # SHEET 9: KPIs_Measures
    # -------------------------------------------------------------
    ws9 = wb.create_sheet(title="KPIs_Measures")
    ws9.views.sheetView[0].showGridLines = True
    ws9["A1"] = "ANALYTICAL KPIS & COMPUTED MEASURES"
    ws9["A1"].font = font_title
    ws9["A2"] = "Mathematical definitions, SQL aggregations, and business metrics"
    ws9["A2"].font = font_subtitle

    kpi_headers = ["Measure ID", "KPI / Measure Name", "Business Formula", "SQL / Python Aggregation Expression", "Unit", "Linked OQ / BR"]
    ws9.row_dimensions[4].height = 26
    for c_idx, h in enumerate(kpi_headers, 1):
        c = ws9.cell(row=4, column=c_idx, value=h)
        c.fill = fill_header
        c.font = font_header
        c.border = border_cell
        c.alignment = Alignment(horizontal="center", vertical="center")

    kpi_data = [
        ("MSR-01", "Average Turnaround Time (TAT)", "Total processing duration / Total settled claims", "AVG(tat_days)", "Days", "OQ-01 / BR-01"),
        ("MSR-02", "Median Processing Duration", "50th percentile of elapsed adjudication days", "np.median(df['tat_days'])", "Days", "OQ-01 / BR-01"),
        ("MSR-03", "Claim Approval Rate", "(Approved Claims Count / Total Decided Claims) * 100", "SUM(CASE WHEN claim_status='Approved' THEN 1 ELSE 0 END)*100.0/COUNT(*)", "Percentage (%)", "OQ-02 / BR-02"),
        ("MSR-04", "Claim Repudiation Rate", "(Repudiated Claims Count / Total Decided Claims) * 100", "SUM(CASE WHEN claim_status='Repudiated' THEN 1 ELSE 0 END)*100.0/COUNT(*)", "Percentage (%)", "OQ-02 / BR-02"),
        ("MSR-05", "Average Settlement Payout Ratio", "Average of (Claim Amount Settled / Claim Amount Requested)", "AVG(claim_amount_settled * 1.0 / claim_amount_requested)", "Ratio (0.0 - 1.0)", "OQ-03 / BR-03"),
        ("MSR-06", "Average Deduction Haircut %", "Average percentage deducted from requested bills", "AVG((claim_amount_requested - claim_amount_settled)*100.0 / claim_amount_requested)", "Percentage (%)", "OQ-03 / BR-03"),
        ("MSR-07", "Total Incurred Claims Payout", "Sum of all disbursements across settled claims", "SUM(claim_amount_settled)", "INR (₹)", "OQ-03, 07 / BR-03"),
        ("MSR-08", "High-Risk Fraud Exposure Sum", "Sum of requested amounts on claims flagged with fraud_flag=1", "SUM(CASE WHEN fraud_flag=1 THEN claim_amount_requested ELSE 0 END)", "INR (₹)", "OQ-05 / BR-05"),
        ("MSR-09", "Cashless vs. Reimb TAT Delta", "Mean TAT (Reimbursement) - Mean TAT (Cashless)", "df[df['settlement_mode']=='Reimbursement']['tat_days'].mean() - df[df['settlement_mode']=='Cashless']['tat_days'].mean()", "Days", "OQ-08 / BR-08"),
        ("MSR-10", "Regional Customer Grievance Rate", "(Claims with Ombudsman Dispute / Total Regional Claims) * 100", "SUM(CASE WHEN csat_score <= 2 THEN 1 ELSE 0 END)*100.0 / COUNT(*)", "Percentage (%)", "OQ-09 / BR-09")
    ]

    for r_idx, r_vals in enumerate(kpi_data, 5):
        ws9.row_dimensions[r_idx].height = 28
        bg = fill_zebra_even if r_idx % 2 == 1 else fill_zebra_odd
        for c_idx, val in enumerate(r_vals, 1):
            c = ws9.cell(row=r_idx, column=c_idx, value=val)
            c.border = border_cell
            c.fill = bg
            c.font = font_body
            if c_idx in [1, 5, 6]:
                c.alignment = Alignment(horizontal="center", vertical="center")
                if c_idx == 1:
                    c.font = font_body_bold
                elif c_idx == 6:
                    c.fill = fill_active
                    c.font = font_body_bold
            elif c_idx == 4:
                c.font = font_code
                c.alignment = Alignment(horizontal="left", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")

    widths_ws9 = {1: 14, 2: 28, 3: 40, 4: 48, 5: 16, 6: 18}
    for col_idx, width in widths_ws9.items():
        ws9.column_dimensions[get_column_letter(col_idx)].width = width

    # -------------------------------------------------------------
    # SHEET 10: SQLite_DDL
    # -------------------------------------------------------------
    ws10 = wb.create_sheet(title="SQLite_DDL")
    ws10.views.sheetView[0].showGridLines = True
    ws10["A1"] = "EXECUTABLE SQLITE DDL SCRIPT"
    ws10["A1"].font = font_title
    ws10["A2"] = "Production schema with table definitions, constraints, indices, and joined analytical view"
    ws10["A2"].font = font_subtitle

    ddl_statements = [
        "-- 1. Policyholders Master Dimension Table",
        "CREATE TABLE IF NOT EXISTS policyholders_master (",
        "    policyholder_id VARCHAR(20) PRIMARY KEY,",
        "    full_name VARCHAR(100) NOT NULL,",
        "    age INTEGER NOT NULL CHECK (age >= 18 AND age <= 100),",
        "    gender VARCHAR(10) NOT NULL CHECK (gender IN ('Male', 'Female', 'Other')),",
        "    city VARCHAR(50) NOT NULL,",
        "    state VARCHAR(50) NOT NULL,",
        "    region VARCHAR(20) NOT NULL CHECK (region IN ('North', 'South', 'East', 'West', 'Central')),",
        "    tenure_months INTEGER NOT NULL CHECK (tenure_months >= 0)",
        ");",
        "",
        "-- 2. Policies Master Dimension Table",
        "CREATE TABLE IF NOT EXISTS policies_master (",
        "    policy_id VARCHAR(20) PRIMARY KEY,",
        "    policyholder_id VARCHAR(20) NOT NULL,",
        "    policy_type VARCHAR(20) NOT NULL CHECK (policy_type IN ('Health', 'Motor', 'Life', 'Property')),",
        "    policy_tier VARCHAR(20) NOT NULL CHECK (policy_tier IN ('Silver', 'Gold', 'Platinum')),",
        "    sum_insured DECIMAL(12,2) NOT NULL CHECK (sum_insured > 0),",
        "    annual_premium DECIMAL(10,2) NOT NULL CHECK (annual_premium > 0),",
        "    policy_start_date DATE NOT NULL,",
        "    policy_status VARCHAR(20) NOT NULL CHECK (policy_status IN ('Active', 'Lapsed', 'Renewed')),",
        "    FOREIGN KEY (policyholder_id) REFERENCES policyholders_master(policyholder_id)",
        "        ON DELETE RESTRICT ON UPDATE CASCADE",
        ");",
        "",
        "-- 3. Providers Master Dimension Table",
        "CREATE TABLE IF NOT EXISTS providers_master (",
        "    provider_id VARCHAR(20) PRIMARY KEY,",
        "    provider_name VARCHAR(100) NOT NULL,",
        "    provider_type VARCHAR(30) NOT NULL CHECK (provider_type IN ('Hospital', 'Automobile Garage', 'TPA')),",
        "    city VARCHAR(50) NOT NULL,",
        "    state VARCHAR(50) NOT NULL,",
        "    empanelment_status VARCHAR(20) NOT NULL CHECK (empanelment_status IN ('Empaneled Network', 'Non-Network', 'Under Review')),",
        "    tariff_tier VARCHAR(10) NOT NULL CHECK (tariff_tier IN ('Tier 1', 'Tier 2', 'Tier 3'))",
        ");",
        "",
        "-- 4. Claims Fact Table (Main Transactional Table)",
        "CREATE TABLE IF NOT EXISTS claims_fact (",
        "    claim_id VARCHAR(20) PRIMARY KEY,",
        "    policy_id VARCHAR(20) NOT NULL,",
        "    provider_id VARCHAR(20) NOT NULL,",
        "    incident_date DATE NOT NULL,",
        "    submission_date DATE NOT NULL,",
        "    verification_date DATE,",
        "    assessment_date DATE,",
        "    settlement_date DATE,",
        "    claim_amount_requested DECIMAL(12,2) NOT NULL CHECK (claim_amount_requested > 0),",
        "    claim_amount_settled DECIMAL(12,2) NOT NULL CHECK (claim_amount_settled >= 0),",
        "    settlement_ratio DECIMAL(5,4) NOT NULL CHECK (settlement_ratio >= 0.0 AND settlement_ratio <= 1.0),",
        "    settlement_mode VARCHAR(20) NOT NULL CHECK (settlement_mode IN ('Cashless', 'Reimbursement')),",
        "    adjudication_stage VARCHAR(30) NOT NULL CHECK (adjudication_stage IN ('Submission', 'Verification', 'Assessment', 'Settled')),",
        "    claim_status VARCHAR(20) NOT NULL CHECK (claim_status IN ('Approved', 'Repudiated', 'Pending Review')),",
        "    rejection_reason_code VARCHAR(20),",
        "    rejection_reason_detail VARCHAR(150),",
        "    tat_days INTEGER NOT NULL CHECK (tat_days >= 0),",
        "    fraud_risk_score INTEGER NOT NULL CHECK (fraud_risk_score >= 0 AND fraud_risk_score <= 100),",
        "    fraud_flag INTEGER NOT NULL CHECK (fraud_flag IN (0, 1)),",
        "    csat_score INTEGER CHECK (csat_score >= 1 AND csat_score <= 5),",
        "    FOREIGN KEY (policy_id) REFERENCES policies_master(policy_id)",
        "        ON DELETE RESTRICT ON UPDATE CASCADE,",
        "    FOREIGN KEY (provider_id) REFERENCES providers_master(provider_id)",
        "        ON DELETE RESTRICT ON UPDATE CASCADE",
        ");",
        "",
        "-- 5. Performance Indices for Analytical Filters",
        "CREATE INDEX IF NOT EXISTS idx_claims_policy_id ON claims_fact(policy_id);",
        "CREATE INDEX IF NOT EXISTS idx_claims_provider_id ON claims_fact(provider_id);",
        "CREATE INDEX IF NOT EXISTS idx_claims_submission_date ON claims_fact(submission_date);",
        "CREATE INDEX IF NOT EXISTS idx_claims_status ON claims_fact(claim_status);",
        "",
        "-- 6. Denormalized Master Analytical View (36 Columns)",
        "CREATE VIEW IF NOT EXISTS v_claims_analysis_master AS",
        "SELECT",
        "    c.claim_id,",
        "    c.policy_id,",
        "    c.provider_id,",
        "    p.policyholder_id,",
        "    ph.full_name,",
        "    ph.age,",
        "    ph.gender,",
        "    ph.city AS customer_city,",
        "    ph.state AS customer_state,",
        "    ph.region AS customer_region,",
        "    ph.tenure_months,",
        "    p.policy_type,",
        "    p.policy_tier,",
        "    p.sum_insured,",
        "    p.annual_premium,",
        "    p.policy_start_date,",
        "    pr.provider_name,",
        "    pr.provider_type,",
        "    pr.empanelment_status,",
        "    pr.city AS provider_city,",
        "    c.incident_date,",
        "    c.submission_date,",
        "    c.verification_date,",
        "    c.assessment_date,",
        "    c.settlement_date,",
        "    c.claim_amount_requested,",
        "    c.claim_amount_settled,",
        "    c.settlement_ratio,",
        "    c.settlement_mode,",
        "    c.adjudication_stage,",
        "    c.claim_status,",
        "    c.rejection_reason_code,",
        "    c.rejection_reason_detail,",
        "    c.tat_days,",
        "    c.fraud_risk_score,",
        "    c.csat_score",
        "FROM claims_fact c",
        "JOIN policies_master p ON c.policy_id = p.policy_id",
        "JOIN policyholders_master ph ON p.policyholder_id = ph.policyholder_id",
        "JOIN providers_master pr ON c.provider_id = pr.provider_id;"
    ]

    for idx, stmt in enumerate(ddl_statements, 4):
        ws10.row_dimensions[idx].height = 18
        cell = ws10.cell(row=idx, column=1, value=stmt)
        cell.font = font_code
        if stmt.startswith("--"):
            cell.font = Font(name="Consolas", size=9.5, bold=True, color="2B6CB0")
        elif "CREATE TABLE" in stmt or "CREATE VIEW" in stmt:
            cell.font = Font(name="Consolas", size=9.5, bold=True, color="1B365D")

    ws10.column_dimensions["A"].width = 110

    output_path = r"c:\Users\noman\OneDrive\Desktop\python\hcl-pro\docs\08_Database_Design.xlsx"
    wb.save(output_path)
    print(f"Successfully generated {output_path}")

def update_rtm_with_db_columns():
    rtm_path = r"c:\Users\noman\OneDrive\Desktop\python\hcl-pro\docs\04_Traceability_Matrix.xlsx"
    wb = openpyxl.load_workbook(rtm_path)
    ws = wb["RTM"]

    # Mapping of BR to exact DB table.column
    db_mappings = {
        "BR-01": "claims_fact.tat_days, claims_fact.adjudication_stage",
        "BR-02": "claims_fact.claim_status, claims_fact.rejection_reason_code",
        "BR-03": "claims_fact.claim_amount_settled, claims_fact.settlement_ratio",
        "BR-04": "providers_master.provider_name, claims_fact.claim_amount_settled",
        "BR-05": "claims_fact.fraud_risk_score, claims_fact.fraud_flag",
        "BR-06": "claims_fact.submission_date, policies_master.policy_type",
        "BR-07": "claims_fact.claim_amount_settled, policies_master.sum_insured",
        "BR-08": "claims_fact.settlement_mode, claims_fact.tat_days",
        "BR-09": "policyholders_master.state, claims_fact.csat_score",
        "BR-10": "policyholders_master.age, policies_master.annual_premium",
        "BR-11": "claims_fact.*, policies_master.* (All tables / PK-FK integrity)",
        "BR-12": "v_claims_analysis_master.* (36 Columns export)"
    }

    font_db = Font(name="Calibri", size=9, bold=True, color="2B6CB0")
    fill_db = PatternFill(start_color="FEFCBF", end_color="FEFCBF", fill_type="solid") # soft yellow

    for row in range(5, ws.max_row + 1):
        br_val = ws.cell(row=row, column=3).value
        if br_val in db_mappings:
            cell = ws.cell(row=row, column=9, value=db_mappings[br_val])
            cell.font = font_db
            cell.fill = fill_db
            cell.alignment = Alignment(horizontal="center", vertical="center")

    wb.save(rtm_path)
    print(f"Successfully updated RTM with DB table.column mappings: {rtm_path}")

if __name__ == "__main__":
    create_database_design_workbook()
    update_rtm_with_db_columns()
