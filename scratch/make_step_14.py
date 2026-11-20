import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import os

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color="D3D3D3"):
    tblPr = table._tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), color)
        tblBorders.append(border)
    tblPr.append(tblBorders)

def generate_build_plan_doc(output_docx_path):
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(6)

    # Document Header
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_after = Pt(2)
    run_meta = p_meta.add_run("ACADEMIC PROJECT ARTIFACT · MCA DATA ANALYTICS · ABES ENGINEERING COLLEGE")
    run_meta.font.size = Pt(9)
    run_meta.font.bold = True
    run_meta.font.color.rgb = RGBColor(0x71, 0x80, 0x96)

    # Title
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(6)
    run_title = p_title.add_run("Step 14: System Build Plan & Engineering Implementation Roadmap")
    run_title.font.name = 'Calibri'
    run_title.font.size = Pt(24)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(16)
    run_sub = p_sub.add_run("Comprehensive Multipage Architecture, Modular Component Specifications, Phased Construction Guide, and Traceability Gates")
    run_sub.font.size = Pt(13)
    run_sub.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)

    # Metadata Box
    meta_table = doc.add_table(rows=6, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(meta_table, "CBD5E0")
    meta_data = [
        ("Project Name", "Insurance Claim Processing Analysis"),
        ("Lead Systems Architect / Author", "Noman Zaid (Roll No: 2500320140103, MCA 2nd Year DA)"),
        ("Faculty Mentor", "Ms. Shilpa, Department of Computer Applications, ABES EC"),
        ("Implementation Stack", "Python 3.12+, Streamlit Multipage, Tailwind CSS CDN, Plotly Express, SQLite3, SciPy"),
        ("Build Phases", "6 Phased Construction Milestones (Foundation -> UI-01..07 -> QA Gates)"),
        ("Traceability Baseline", "100% Mapped to US-01..14, UI-01..07, BR-01..12, TC-F-01..14, TC-T-01..14")
    ]
    for i, (k, v) in enumerate(meta_data):
        row = meta_table.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = Inches(2.2), Inches(4.3)
        set_cell_background(c0, "F7FAFC")
        set_cell_background(c1, "FFFFFF")
        set_cell_margins(c0, 70, 70, 100, 100)
        set_cell_margins(c1, 70, 70, 100, 100)
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(k)
        r0.font.bold = True
        r0.font.size = Pt(9.5)
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(v)
        r1.font.size = Pt(9.5)
        if i == 5:
            r1.font.bold = True
            r1.font.color.rgb = RGBColor(0x27, 0x67, 0x49)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 1: Executive Summary & Architecture Paradigm
    h1 = doc.add_heading(level=1)
    r = h1.add_run("1. Executive Summary & Architectural Paradigm")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "The purpose of this engineering build plan is to formalize the phased development lifecycle for the "
        "Insurance Claim Processing Analysis system. Building directly upon the validated technical specifications (Step 7), "
        "relational SQLite schema (Step 8), UI design system (Step 12), and QA test suite (Step 13), this document governs "
        "the actual codebase construction for Step 15.\n\n"
        "The software architecture strictly adheres to a Decoupled Model-View-Controller (MVC) and Pipeline pattern:\n"
        "• Data Access Layer (Model): Encapsulated within 'src/data_loader.py', handling SQLite connection pooling, "
        "complex multi-table relational joins, data type coercion, and Streamlit reactive caching (@st.cache_data).\n"
        "• Analytical & Statistics Engine (Controller/Service): Pure Python mathematical routines in 'src/analytics.py' "
        "governing cycle time milestones, Pareto 80/20 frequencies, haircut percentages, Welch's t-test hypothesis evaluations, "
        "and fraud anomaly scoring without any presentation dependencies.\n"
        "• Component Presentation Layer (View): Modularized in 'src/ui.py' utilizing Tailwind CSS CDN via isolated iframes "
        "to deliver modern responsive typography, metric cards, and executive banners.\n"
        "• Multipage Application Shell: Managed by 'app.py' and 'pages/*.py' serving as high-level visual orchestrators "
        "implementing global state synchronization via st.session_state."
    )

    # Section 2: Target File System Architecture
    h1 = doc.add_heading(level=1)
    r = h1.add_run("2. Target Directory & Module Specification")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "The production codebase is organized into cleanly partitioned modules to maintain separation of concerns:"
    )

    mod_table = doc.add_table(rows=11, cols=4)
    mod_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(mod_table, "CBD5E0")

    m_headers = ["File / Module Path", "Layer / Role", "Primary Responsibility", "Linked Requirements"]
    for j, h in enumerate(m_headers):
        c = mod_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 60, 60)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    modules_data = [
        ("app.py", "View / Controller", "UI-01 Executive Overview: Global sidebar filters, portfolio KPI cards, haircut scatter plot, and claim status donut.", "BR-01, BR-03, US-01, US-04, US-14"),
        ("pages/1_TAT_Analytics.py", "View / Controller", "UI-02 Turnaround Time: Stage milestone stacked bar charts, Cashless vs Reimbursement violin distributions, and 30-day moving average.", "BR-01, BR-06, US-02, US-07"),
        ("pages/2_Denial_Analysis.py", "View / Controller", "UI-03 Denial Analysis: Dual-axis Pareto 80/20 repudiation chart, state-level grievance heatmap, and root cause audit table.", "BR-02, BR-09, US-03, US-10"),
        ("pages/3_Providers.py", "View / Controller", "UI-04 Network Providers: Provider cost vs TAT quadrant benchmarking, top 10 expensive hospital leaderboard, and tariff audit badges.", "BR-04, US-05"),
        ("pages/4_Fraud_Screening.py", "View / Controller", "UI-05 Fraud Screening: Risk threshold slider, high-risk triage queue, early inception exposure KPI, and anomaly composition pie.", "BR-05, US-06"),
        ("pages/5_Underwriting.py", "View / Controller", "UI-06 Actuarial Underwriting: Welch's t-test hypothesis card, Tukey catastrophic outlier fences, and correlation matrix heatmap.", "BR-07, BR-08, BR-10, US-08, US-09, US-11"),
        ("pages/6_Data_Quality.py", "View / Controller", "UI-07 Data Quality: Data health scorecard, pre/post cleaning transformation logs, table browsers, and CSV/Excel export buttons.", "BR-11, BR-12, US-12, US-13"),
        ("src/data_loader.py", "Data Access (Model)", "Encapsulates SQLite connection pooling, SQL joining of fact/dimension tables, derived metrics computation, and caching (@st.cache_data).", "BR-01..12, TC-T-01..06"),
        ("src/analytics.py", "Analytical Engine", "Pure statistical/mathematical routines: TAT calculations, Pareto percentages, haircut formulas, Welch's t-test, and anomaly scoring.", "BR-01..10, TC-T-07..11"),
        ("src/ui.py", "UI Presentation", "Reusable Tailwind CSS component library rendering responsive headers, 4-column metric KPI cards, and insight alert callouts.", "UI-01..07, TC-T-12, TC-T-13")
    ]

    col_w_mod = [Inches(1.8), Inches(1.3), Inches(2.4), Inches(1.0)]
    for i, row in enumerate(modules_data):
        r = mod_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w_mod[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 60, 60, 60, 60)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(8)
            if j == 0:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 3: Phased Construction Roadmap
    h1 = doc.add_heading(level=1)
    r = h1.add_run("3. Phased Implementation Roadmap (Phases 1 to 6)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    phases = [
        ("Phase 1: Core Foundation & Analytical Engine (src/data_loader.py & src/analytics.py)",
         "Deliverables: Backend data connectors, caching routines, and pure mathematical calculation functions.\n"
         "• Task 1.1: Implement 'get_database_connection()' and '@st.cache_data(ttl=3600) def load_master_claims_data()' joining claims_fact with policies_master, policyholders_master, and providers_master.\n"
         "• Task 1.2: Compute calculated fields: turnaround time ('tat_days'), haircut deduction percentage ('haircut_pct'), and days from inception ('days_from_inception').\n"
         "• Task 1.3: Build analytical functions: 'compute_kpi_summary()', 'compute_pareto_repudiation()', 'perform_welchs_ttest()', 'calculate_fraud_score()', and 'calculate_tukey_fences()'.\n"
         "• Quality Gate: Pytest test suite (TC-T-01 to TC-T-11) must execute with 100% pass rate."),

        ("Phase 2: Global Shell, Navigation Router & Executive Overview (app.py & src/ui.py)",
         "Deliverables: Main multi-page entry point, persistent sidebar filters, and UI-01 Executive Dashboard.\n"
         "• Task 2.1: Configure Streamlit page layout ('wide', 'initial_sidebar_state=expanded', corporate page title & icon).\n"
         "• Task 2.2: Establish global sidebar filters synced to st.session_state (Policy Type, Settlement Mode, Submission Date Range).\n"
         "• Task 2.3: Assemble UI-01 Executive Overview: render executive header, 4 Tailwind KPI cards (Total Claims, Net Payout, Median TAT, Approval Rate), Claim Haircut Scatter Plot, Claim Status Donut, and Monthly Volume Area Chart.\n"
         "• Quality Gate: Functional tests TC-F-01, TC-F-02, and TC-F-05 verified under dynamic filter toggles."),

        ("Phase 3: Operational Cycle Times & Provider Benchmarking (pages/1_ & pages/3_)",
         "Deliverables: Deep-dive operational cycle time analysis (UI-02) and empaneled network provider scorecard (UI-04).\n"
         "• Task 3.1: Build 'pages/1_TAT_Analytics.py': Milestone cycle time stacked horizontal bar chart (Document Review vs Medical Assessment vs Settlement), Violin distribution comparing Cashless vs Reimbursement, and 30-day moving average time-series.\n"
         "• Task 3.2: Build 'pages/3_Providers.py': Cost vs TAT performance quadrant scatter plot with benchmark crosshairs, top 10 expensive hospital leaderboard, and tariff audit badges (z > 2.0).\n"
         "• Quality Gate: Functional tests TC-F-03, TC-F-06, and TC-F-08 verified; Apex Multispeciality correctly flagged."),

        ("Phase 4: Repudiation, Fraud & Underwriting Analytics (pages/2_, pages/4_, pages/5_)",
         "Deliverables: Repudiation root-cause drill-down (UI-03), fraud risk triage queue (UI-05), and actuarial hypothesis testing (UI-06).\n"
         "• Task 4.1: Build 'pages/2_Denial_Analysis.py': Dual-axis Pareto 80/20 chart isolating top rejection causes, regional repudiation heatmap, and IRDAI dispute drill-down table.\n"
         "• Task 4.2: Build 'pages/4_Fraud_Screening.py': Interactive risk slider (>= 70), suspicious claim triage queue table, total exposure KPI card, and anomaly type breakdown pie chart.\n"
         "• Task 4.3: Build 'pages/5_Underwriting.py': Welch's two-sample t-test hypothesis card (Cashless vs Reimbursement TAT, p < 0.001), catastrophic claim outlier isolation (> ₹20L via Tukey IQR fences), and Pearson correlation matrix.\n"
         "• Quality Gate: Functional tests TC-F-04, TC-F-07, TC-F-09, TC-F-10, and TC-F-12 verified."),

        ("Phase 5: Data Governance, Provenance & Export Utility (pages/6_Data_Quality.py)",
         "Deliverables: Data provenance scorecard, ETL transformation audit log, and tabular data export triggers (UI-07).\n"
         "• Task 5.1: Build 'pages/6_Data_Quality.py': Data health metric ribbon (Completeness 99.8%, Uniqueness 100%, Referential Integrity 100%), pre/post cleaning audit comparison table, and interactive database table browser.\n"
         "• Task 5.2: Implement secure, sanitized download triggers: 'st.download_button' for CSV and Excel formats based on active sidebar filter selections.\n"
         "• Quality Gate: Functional tests TC-F-13 and TC-F-14 verified; clean exports verified without PII leaks."),

        ("Phase 6: Full System Integration, Performance Tuning & Final QA Verification",
         "Deliverables: End-to-end regression validation, performance SLA benchmarking, and production packaging.\n"
         "• Task 6.1: Run complete automated test suite ('pytest -v') validating all 14 technical test cases (TC-T-01..14).\n"
         "• Task 6.2: Conduct full end-to-end browser walkthrough across all 7 screens validating all 14 functional test cases (TC-F-01..14).\n"
         "• Task 6.3: Benchmark performance SLAs (Initial load < 2.0s, filter update reflow < 1.0s, memory usage < 250MB).\n"
         "• Quality Gate: Zero unhandled exceptions, zero broken charts, and 100% requirement traceability verified.")
    ]

    for title, desc in phases:
        h2 = doc.add_heading(level=2)
        r = h2.add_run(title)
        r.font.size = Pt(12)
        r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
        p = doc.add_paragraph(desc)
        p.paragraph_format.space_after = Pt(8)

    # Section 4: Requirements & Test Traceability Matrix
    h1 = doc.add_heading(level=1)
    r = h1.add_run("4. Comprehensive Build Traceability Matrix")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "Every single business requirement, user story, UI screen, target code file, and verification test case "
        "is deterministically mapped across the implementation matrix:"
    )

    rtm_table = doc.add_table(rows=13, cols=6)
    rtm_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(rtm_table, "CBD5E0")

    rtm_headers = ["Req ID", "User Story", "UI Screen", "Target File / Module", "Automated Pytest", "Functional Test"]
    for j, h in enumerate(rtm_headers):
        c = rtm_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 50, 50)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    rtm_data = [
        ("BR-01", "US-01, US-14", "UI-01, UI-02", "app.py, pages/1_TAT_Analytics.py", "TC-T-07, TC-T-12", "TC-F-01, TC-F-02, TC-F-03"),
        ("BR-02", "US-03", "UI-03", "pages/2_Denial_Analysis.py", "TC-T-08", "TC-F-04"),
        ("BR-03", "US-04", "UI-01", "app.py, src/analytics.py", "TC-T-09", "TC-F-05"),
        ("BR-04", "US-05", "UI-04", "pages/3_Providers.py", "TC-T-02, TC-T-04", "TC-F-06"),
        ("BR-05", "US-06", "UI-05", "pages/4_Fraud_Screening.py", "TC-T-11", "TC-F-07"),
        ("BR-06", "US-07", "UI-02", "pages/1_TAT_Analytics.py", "TC-T-07", "TC-F-08"),
        ("BR-07", "US-08", "UI-06", "pages/5_Underwriting.py", "TC-T-02", "TC-F-09"),
        ("BR-08", "US-09", "UI-06", "pages/5_Underwriting.py, src/analytics.py", "TC-T-10", "TC-F-10"),
        ("BR-09", "US-10", "UI-03", "pages/2_Denial_Analysis.py", "TC-T-02", "TC-F-11"),
        ("BR-10", "US-11", "UI-06", "pages/5_Underwriting.py", "TC-T-02", "TC-F-12"),
        ("BR-11", "US-12", "UI-07", "pages/6_Data_Quality.py, src/data_loader.py", "TC-T-01..06", "TC-F-13"),
        ("BR-12", "US-13", "UI-07", "pages/6_Data_Quality.py", "TC-T-14", "TC-F-14")
    ]

    col_w_rtm = [Inches(0.8), Inches(1.0), Inches(0.9), Inches(1.7), Inches(1.1), Inches(1.0)]
    for i, row in enumerate(rtm_data):
        r = rtm_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w_rtm[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 50, 50, 50, 50)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(8)
            if j == 0:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
            elif j == 4:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0x27, 0x67, 0x49)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 5: Risk Management & Fallback Contingency
    h1 = doc.add_heading(level=1)
    r = h1.add_run("5. Technical Risk Management & Fallback Protocols")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "To guarantee uninterrupted operation during evaluation, viva defense, and cloud deployment, "
        "the following engineering safeguards are implemented:\n"
        "1. Filter Starvation Handling: In scenarios where user filter combinations yield zero records, "
        "views intercept empty DataFrames immediately and display a polite warning ('No records match active filter criteria') "
        "rather than allowing Plotly charts to crash with IndexError or KeyError.\n"
        "2. Cross-Page Session State Synchronization: To avoid filter reset bugs during multipage transitions, "
        "filter widgets are keyed to st.session_state, preserving user choices across all page views.\n"
        "3. SQLite Concurrency Safeguards: SQLite connection queries are strictly read-only and wrapped in @st.cache_data(ttl=3600), "
        "preventing database table locks or disk I/O bottlenecks during rapid user interactions.\n"
        "4. Plotly Responsiveness & Theme Parity: All Plotly visualizations employ 'responsive=True' within 'st.plotly_chart', "
        "with consistent corporate navy color sequences matching the design framework tokens."
    )

    # Section 6: Execution Checklist for Step 15
    h1 = doc.add_heading(level=1)
    r = h1.add_run("6. Step 15 Implementation Checklist & Verification Gates")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "The following execution checklist governs the active coding phase (Step 15):\n"
        " [X] Step 14 Deliverables Approved: 'BUILD_PLAN.md' and 'docs/14_BUILD_PLAN.docx' established.\n"
        " [ ] Build 'src/data_loader.py' with SQLite relational queries and @st.cache_data.\n"
        " [ ] Build 'src/analytics.py' housing pure calculation functions for TAT, Pareto, Haircut, Welch's t-test, and Fraud.\n"
        " [ ] Construct 'app.py' (UI-01 Executive Overview) with global sidebar filters and Tailwind KPI cards.\n"
        " [ ] Construct 'pages/1_TAT_Analytics.py' (UI-02) featuring milestone cycle time stacked bars and violin distributions.\n"
        " [ ] Construct 'pages/2_Denial_Analysis.py' (UI-03) featuring the dual-axis Pareto 80/20 repudiation chart.\n"
        " [ ] Construct 'pages/3_Providers.py' (UI-04) featuring the provider performance quadrant and tariff audit badges.\n"
        " [ ] Construct 'pages/4_Fraud_Screening.py' (UI-05) featuring the interactive risk threshold slider and triage queue.\n"
        " [ ] Construct 'pages/5_Underwriting.py' (UI-06) featuring Welch's t-test hypothesis card and Tukey outlier fences.\n"
        " [ ] Construct 'pages/6_Data_Quality.py' (UI-07) featuring the data health scorecard and CSV/Excel exports.\n"
        " [ ] Execute full pytest suite ('pytest -v') ensuring all 14 tests pass.\n"
        " [ ] Perform complete multi-page browser smoke test validating all 14 functional acceptance scenarios."
    )

    doc.save(output_docx_path)
    print(f"Successfully generated {output_docx_path}")

if __name__ == "__main__":
    output_path = r"c:\Users\noman\OneDrive\Desktop\python\hcl-pro\docs\14_BUILD_PLAN.docx"
    generate_build_plan_doc(output_path)
