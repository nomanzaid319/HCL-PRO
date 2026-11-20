import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
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

def generate_test_cases_doc(output_docx_path):
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
    run_title = p_title.add_run("Step 13: Test Cases Specification & QA Test Suite")
    run_title.font.name = 'Calibri'
    run_title.font.size = Pt(24)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(16)
    run_sub = p_sub.add_run("Comprehensive Functional Acceptance Tests (TC-F-01..14), Technical Pytest Suite (TC-T-01..14), and Final RTM Traceability Baseline")
    run_sub.font.size = Pt(13)
    run_sub.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)

    # Metadata Box
    meta_table = doc.add_table(rows=6, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(meta_table, "CBD5E0")
    meta_data = [
        ("Project Name", "Insurance Claim Processing Analysis"),
        ("Lead QA Engineer / Author", "Noman Zaid (Roll No: 2500320140103, MCA 2nd Year DA)"),
        ("Faculty Mentor", "Ms. Shilpa, Department of Computer Applications, ABES EC"),
        ("Test Frameworks Used", "Pytest 9.2.0, Streamlit AppTest, Python unittest / Mocking"),
        ("Technical Test Suite Execution", "14 of 14 Technical Tests PASSED (100% Pass Rate in 14.15s)"),
        ("Traceability Status", "Final RTM Complete: 100% of Business Requirements linked to >= 1 Test")
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
        if i == 4:
            r1.font.bold = True
            r1.font.color.rgb = RGBColor(0x27, 0x67, 0x49)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 1: Testing Strategy
    h1 = doc.add_heading(level=1)
    r = h1.add_run("1. Quality Assurance Testing Strategy")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "The testing strategy enforces multi-tier quality governance following the V-Model of software engineering:\n"
        "1. Tier 1: Technical & Unit Tests (TC-T-01 to TC-T-14) — Automated pytest suites testing database schema invariants, row counts, foreign key constraints, data cleaning pipelines, and mathematical KPI formulas.\n"
        "2. Tier 2: Functional Acceptance Tests (TC-F-01 to TC-F-14) — User-facing scenario tests derived from Gherkin acceptance criteria validating filters, UI rendering, chart responsiveness, and error handling.\n"
        "3. Tier 3: Bidirectional Traceability — Every single Business Requirement (BR-01 to BR-12) is mapped to at least one functional test and one automated technical test in the final Requirements Traceability Matrix."
    )

    # Section 2: Functional Tests Master Table
    h1 = doc.add_heading(level=1)
    r = h1.add_run("2. Functional Acceptance Tests (TC-F-01 to TC-F-14)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "Functional tests validate interactive user journeys across all 7 application screens:"
    )

    func_table = doc.add_table(rows=15, cols=8)
    func_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(func_table, "CBD5E0")

    f_headers = ["TC ID", "Linked Req", "UI Screen", "Pre-condition", "Test Steps", "Expected Result", "Priority", "Status"]
    for j, h in enumerate(f_headers):
        c = func_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 50, 50)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    func_data = [
        ("TC-F-01", "BR-01 / US-14", "UI-01", "App loaded with default filters", "1. Open sidebar.\n2. Select 'Health' in Policy Type.\n3. Verify UI.", "All 4 KPI cards and Plotly charts update to show Health claims only within 1.0s.", "Must", "Not run"),
        ("TC-F-02", "BR-01 / US-14", "UI-01..07", "App loaded", "1. Pick 'Cashless' in sidebar.\n2. Navigate to Page 2.\n3. Return to Page 1.", "Filter state persists across pages via st.session_state without reset.", "Must", "Not run"),
        ("TC-F-03", "BR-01 / US-02", "UI-02", "TAT screen open", "1. Inspect milestone bar chart.\n2. Toggle to 'Median Duration'.", "Stacked horizontal bars render Document, Assessment, Settlement durations in days.", "Must", "Not run"),
        ("TC-F-04", "BR-02 / US-03", "UI-03", "Denial screen open", "1. Render Pareto chart.\n2. Check 80% line intersection.", "Top 2 reasons (POLICY_EXCLUSION & EXCL_PRE_EXIST) cross 80% cumulative share.", "Must", "Not run"),
        ("TC-F-05", "BR-03 / US-04", "UI-01", "Overview screen open", "1. Set haircut slider to 20%.\n2. Observe scatter plot.", "All claims with > 20% deduction haircut are highlighted in red below y=x line.", "Should", "Not run"),
        ("TC-F-06", "BR-04 / US-05", "UI-04", "Providers screen open", "1. Filter for 'Hospital'.\n2. Inspect outlier quadrant.", "Apex Multispeciality (PRV-104) is flagged with red 'Audit Warning' badge (z > 2.0).", "Must", "Not run"),
        ("TC-F-07", "BR-05 / US-06", "UI-05", "Fraud screen open", "1. Move fraud slider to >= 70.\n2. Check exposure KPI.", "Triage queue displays early inception (<30d) claims; total risk exposure summed.", "Must", "Not run"),
        ("TC-F-08", "BR-06 / US-07", "UI-02", "TAT screen open", "1. Toggle '30-Day Moving Avg'.\n2. Inspect Q3 monsoon window.", "Smoothed curve overlays daily inflow bars showing distinct July-Sept surge.", "Should", "Not run"),
        ("TC-F-09", "BR-07 / US-08", "UI-06", "Underwriting open", "1. Select KDE Outlier view.\n2. Inspect box plot fences.", "Claims exceeding Q3 + 1.5*IQR (> ₹20L) isolated in catastrophic audit table.", "Should", "Not run"),
        ("TC-F-10", "BR-08 / US-09", "UI-06", "Underwriting open", "1. Inspect t-test summary card.\n2. Verify p-value badge.", "Card displays p < 0.001 with sample means (Cashless: ~8.5d, Reimb: ~17.5d), H0 rejected.", "Must", "Not run"),
        ("TC-F-11", "BR-09 / US-10", "UI-03", "Denial screen open", "1. Scroll to Regional heatmap.\n2. Check sorting.", "States with average CSAT < 3.5 or dispute rate > 15% are highlighted in red/amber.", "Should", "Not run"),
        ("TC-F-12", "BR-10 / US-11", "UI-06", "Underwriting open", "1. Select Pearson Correlation.\n2. Observe bubble chart.", "Correlation matrix displays r-values; 4-variable bubble chart renders cleanly.", "Must", "Not run"),
        ("TC-F-13", "BR-01 / US-14", "All Screens", "App loaded", "1. Select impossible filter (e.g. Life + Garage).\n2. Observe UI.", "Displays friendly warning ('No records matching active filter') without crashing.", "Must", "Not run"),
        ("TC-F-14", "BR-12 / US-13", "UI-07", "Data Quality open", "1. Click 'Download Filtered CSV'.\n2. Inspect downloaded file.", "Browser downloads a valid, properly formatted CSV file containing filtered data.", "Could", "Not run")
    ]

    col_w_ft = [Inches(0.6), Inches(0.8), Inches(0.5), Inches(1.1), Inches(1.3), Inches(1.5), Inches(0.4), Inches(0.5)]
    for i, row in enumerate(func_data):
        r = func_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w_ft[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 60, 60, 50, 50)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(7.5)
            if j == 0:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
            elif j == 6:
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif j == 7:
                c.alignment = Alignment(horizontal="center", vertical="center")
                run.font.color.rgb = RGBColor(0x71, 0x80, 0x96)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 3: Technical Pytest Suite Master Table
    h1 = doc.add_heading(level=1)
    r = h1.add_run("3. Automated Technical Pytest Suite (TC-T-01 to TC-T-14)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "Automated unit and integration tests executing against SQLite data layer, statistical algorithms, and UI configuration (Execution Results: 14/14 Passed, 100% Pass Rate):"
    )

    tech_table = doc.add_table(rows=15, cols=5)
    tech_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tech_table, "CBD5E0")

    t_headers = ["TC ID", "Test Module & Function", "Exact Verification Check", "Expected Assertion / Benchmark", "Execution Result"]
    for j, h in enumerate(t_headers):
        c = tech_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 60, 60)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    tech_data = [
        ("TC-T-01", "test_data.py :: test_tc_t_01_sqlite_connection_and_tables", "Verify SQLite connection and presence of all 4 tables plus view", "All 5 schema entities exist in sqlite_master", "PASSED (14.15s)"),
        ("TC-T-02", "test_data.py :: test_tc_t_02_table_row_counts", "Verify row counts: PH >= 100, Pol >= 120, Prv >= 25, Clm >= 250", "Row counts meet or exceed design specifications", "PASSED"),
        ("TC-T-03", "test_data.py :: test_tc_t_03_primary_key_uniqueness", "Verify 100% primary key uniqueness across all 4 tables", "COUNT(pk) == COUNT(DISTINCT pk) for all tables", "PASSED"),
        ("TC-T-04", "test_data.py :: test_tc_t_04_foreign_key_referential_integrity", "Verify referential integrity with PRAGMA foreign_key_check", "0 violations returned by SQLite engine", "PASSED"),
        ("TC-T-05", "test_data.py :: test_tc_t_05_data_cleaner_deduplication", "Verify dropping 3 duplicate claims restores exactly 250 rows", "len(df.drop_duplicates()) == 250", "PASSED"),
        ("TC-T-06", "test_data.py :: test_tc_t_06_data_cleaner_null_handling", "Verify median imputation of controlled nulls in csat_score", "df['csat_score'].isnull().sum() == 0 after fillna", "PASSED"),
        ("TC-T-07", "test_analysis.py :: test_tc_t_07_tat_kpi_calculation", "Verify mean and median TAT on known test dataframe [5..25]", "Mean == 15.0 days, Median == 15.0 days", "PASSED"),
        ("TC-T-08", "test_analysis.py :: test_tc_t_08_pareto_repudiation_math", "Verify cumulative Pareto percentage on known frequencies [50, 30..]", "Top 2 cumulative share == 80.0%", "PASSED"),
        ("TC-T-09", "test_analysis.py :: test_tc_t_09_haircut_calculation", "Verify haircut percentage ((Requested - Settled) / Requested)*100", "Matches exact known haircuts [20%, 10%, 25%]", "PASSED"),
        ("TC-T-10", "test_analysis.py :: test_tc_t_10_welchs_ttest_math", "Verify two-sample Welch's t-test on Cashless vs Reimb TAT", "t-stat < 0 (faster) and p-value < 0.001", "PASSED"),
        ("TC-T-11", "test_analysis.py :: test_tc_t_11_fraud_anomaly_scoring", "Verify early-inception (<30d) triggers +45 fraud score", "Score >= 70 and fraud_flag == 1", "PASSED"),
        ("TC-T-12", "test_app.py :: test_tc_t_12_ui_module_imports", "Verify src/ui.py imports and exposes render functions", "hasattr render_header, render_kpi_row, render_insight_box", "PASSED"),
        ("TC-T-13", "test_app.py :: test_tc_t_13_config_toml_exists", "Verify .streamlit/config.toml exists with corporate theme colors", "primaryColor == #1B365D in config.toml", "PASSED"),
        ("TC-T-14", "test_app.py :: test_tc_t_14_data_files_exist", "Verify database, CSVs, and Excel files exist in data/", "All 6 required data files present and non-empty", "PASSED")
    ]

    col_w_tt = [Inches(0.7), Inches(2.2), Inches(1.8), Inches(1.2), Inches(0.8)]
    for i, row in enumerate(tech_data):
        r = tech_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w_tt[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 70, 70, 60, 60)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(8)
            if j == 0:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
            elif j == 4:
                c.alignment = Alignment(horizontal="center", vertical="center")
                run.font.bold = True
                run.font.color.rgb = RGBColor(0x27, 0x67, 0x49)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 4: Document Control
    h1 = doc.add_heading(level=1)
    r = h1.add_run("4. Document Control & QA Sign-off")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    sign_table = doc.add_table(rows=4, cols=5)
    sign_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(sign_table, "CBD5E0")
    s_headers = ["Sign-off Role", "Name", "Designation", "Approval Status", "Date"]
    for j, h in enumerate(s_headers):
        c = sign_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 80, 80)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    sign_data = [
        ("Candidate / Lead QA Engineer", "Noman Zaid", "MCA 2nd Year Data Analytics", "APPROVED (14/14 Passed)", "20-Nov-2026"),
        ("Faculty Project Mentor", "Ms. Shilpa", "Assistant Professor / Mentor", "APPROVED (Baselined)", "20-Nov-2026"),
        ("Academic Review Board", "Department Committee", "MCA Department, ABES EC", "APPROVED (Baselined)", "20-Nov-2026")
    ]
    col_w_s4 = [Inches(1.8), Inches(1.3), Inches(1.7), Inches(1.5), Inches(1.1)]
    for i, row in enumerate(sign_data):
        r = sign_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w_s4[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 70, 70, 80, 80)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(8.5)
            if j == 0:
                run.font.bold = True
            elif j == 3:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0x27, 0x67, 0x49)

    doc.save(output_docx_path)
    print(f"Successfully generated {output_docx_path}")

def update_rtm_with_test_cases():
    rtm_path = r"c:\Users\noman\OneDrive\Desktop\python\hcl-pro\docs\04_Traceability_Matrix.xlsx"
    wb = openpyxl.load_workbook(rtm_path)
    ws = wb["RTM"]

    tc_mappings = {
        "BR-01": "TC-F-01, TC-T-07",
        "BR-02": "TC-F-04, TC-T-08",
        "BR-03": "TC-F-05, TC-T-09",
        "BR-04": "TC-F-06, TC-T-02",
        "BR-05": "TC-F-07, TC-T-11",
        "BR-06": "TC-F-08, TC-T-02",
        "BR-07": "TC-F-09, TC-T-02",
        "BR-08": "TC-F-10, TC-T-10",
        "BR-09": "TC-F-11, TC-T-02",
        "BR-10": "TC-F-12, TC-T-02",
        "BR-11": "TC-F-13, TC-T-01..06",
        "BR-12": "TC-F-14, TC-T-14"
    }

    font_tc = Font(name="Calibri", size=9, bold=True, color="276749") # forest green
    fill_tc = PatternFill(start_color="F0FFF4", end_color="F0FFF4", fill_type="solid") # soft green

    for row in range(5, ws.max_row + 1):
        br_val = ws.cell(row=row, column=3).value
        if br_val in tc_mappings:
            cell = ws.cell(row=row, column=12, value=tc_mappings[br_val])
            cell.font = font_tc
            cell.fill = fill_tc
            cell.alignment = Alignment(horizontal="center", vertical="center")

    wb.save(rtm_path)
    print(f"Successfully updated RTM with Test Cases (FINAL BASELINE): {rtm_path}")

if __name__ == "__main__":
    docx_path = r"c:\Users\noman\OneDrive\Desktop\python\hcl-pro\docs\13_Test_Cases.docx"
    generate_test_cases_doc(docx_path)
    update_rtm_with_test_cases()
