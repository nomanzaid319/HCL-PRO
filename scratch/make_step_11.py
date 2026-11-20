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

def generate_user_stories_doc(output_docx_path):
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

    # Header
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
    run_title = p_title.add_run("Step 11: Agile User Stories & Acceptance Criteria")
    run_title.font.name = 'Calibri'
    run_title.font.size = Pt(24)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(16)
    run_sub = p_sub.add_run("Persona-Driven Agile Epics, User Stories (US-01 to US-14), Gherkin Acceptance Criteria, and 2-Week Sprint Plan")
    run_sub.font.size = Pt(13)
    run_sub.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)

    # Metadata Box
    meta_table = doc.add_table(rows=5, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(meta_table, "CBD5E0")
    meta_data = [
        ("Project Name", "Insurance Claim Processing Analysis"),
        ("Learner / Scrum Master", "Noman Zaid (Roll No: 2500320140103, MCA 2nd Year DA)"),
        ("Faculty Mentor", "Ms. Shilpa, Department of Computer Applications, ABES EC"),
        ("Upstream Baselines", "docs/06_PRD.docx (Features F-01 to F-12), docs/08_Database_Design.xlsx"),
        ("Agile Framework", "Scrum / Kanban hybrid, 2-Week Sprint, Fibonacci Story Points (1-8)")
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

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 1: Agile Epics Breakdown
    h1 = doc.add_heading(level=1)
    r = h1.add_run("1. Agile Epics Structure")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "To manage software delivery systematically, features and user stories are organized into five strategic Epics:"
    )

    epic_table = doc.add_table(rows=6, cols=4)
    epic_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(epic_table, "CBD5E0")

    e_headers = ["Epic ID", "Epic Name", "Core Capabilities", "Story Points Total"]
    for j, h in enumerate(e_headers):
        c = epic_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 80, 80)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    epics = [
        ("EPIC-01", "Core Dashboard & Global Controls", "Executive KPI scorecards, multi-line policy overview, persistent sidebar filters.", "6 Points"),
        ("EPIC-02", "Operational TAT & Bottleneck Analytics", "Stage-wise cycle time tracking (Submission -> Decision), seasonal inflow surge forecasting.", "8 Points"),
        ("EPIC-03", "Risk, Fraud & Network Governance", "Pareto claim repudiation engine, provider benchmarking matrix, early-inception fraud screening.", "13 Points"),
        ("EPIC-04", "Actuarial Insights & Hypothesis Testing", "Cashless vs. Reimbursement Welch's t-test, catastrophic outlier modeler, demographic correlations.", "11 Points"),
        ("EPIC-05", "Data Architecture, Quality & Export", "SQLite connection & @st.cache_data, data cleaning audit trail, CSV/Excel reporting export.", "5 Points")
    ]

    col_w_e = [Inches(1.0), Inches(2.2), Inches(2.3), Inches(1.0)]
    for i, row in enumerate(epics):
        r = epic_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w_e[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 70, 70, 80, 80)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(8.5)
            if j == 0:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 2: User Stories Master Table
    h1 = doc.add_heading(level=1)
    r = h1.add_run("2. User Stories Master Table (US-01 to US-14)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "The 14 user stories cover 100% of the MVP features and operational questions. Each story is sized using the Fibonacci sequence (1-8 story points) and can be implemented in under half a day:"
    )

    us_table = doc.add_table(rows=15, cols=7)
    us_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(us_table, "CBD5E0")

    us_headers = ["US ID", "Target Persona", "User Story (\"As a <persona>, I want <goal> so that <benefit>\")", "Feature ID", "BR ID", "Priority", "Points"]
    for j, h in enumerate(us_headers):
        c = us_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 60, 60)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    stories_data = [
        ("US-01", "Suresh Menon (CFO)", "As a CFO, I want to view an executive KPI summary card row displaying total claims, gross payout, mean TAT, and approval percentage so that I can evaluate portfolio financial health at a glance.", "F-03", "BR-03", "Must", "3"),
        ("US-02", "Rajesh Sharma (Ops)", "As a Claims Operations Manager, I want to analyze stage-wise cycle times across submission, verification, assessment, and settlement so that I can isolate the exact operational bottlenecks causing processing latency.", "F-01", "BR-01", "Must", "5"),
        ("US-03", "Ananya Verma (CX)", "As a Customer Experience Lead, I want to view a Pareto 80/20 chart of claim repudiation causes so that I can identify the vital few denial reasons driving the majority of policyholder rejections.", "F-02", "BR-02", "Must", "3"),
        ("US-04", "Suresh Menon (CFO)", "As a CFO, I want to monitor claim deduction haircuts and filter claims with deductions exceeding 20% so that I can prevent unfair claim haircutting and maintain financial reserve accuracy.", "F-03", "BR-03", "Should", "3"),
        ("US-05", "Priya Nair (SIU)", "As an SIU Fraud Investigator, I want to benchmark empaneled hospitals and garages on average claim cost and turnaround time so that I can identify providers exhibiting abnormal tariff inflation and long delays.", "F-04", "BR-04", "Must", "5"),
        ("US-06", "Priya Nair (SIU)", "As an SIU Fraud Investigator, I want to screen claims with early-inception (<30 days) and high anomaly scores so that I can intercept fraudulent disbursements and quantify total monetary risk exposure.", "F-05", "BR-05", "Must", "5"),
        ("US-07", "Rajesh Sharma (Ops)", "As a Claims Operations Manager, I want to view monthly claim inflow trends with a 30-day moving average so that I can forecast seasonal surge periods (such as monsoon motor spikes) and dynamically schedule adjuster capacity.", "F-06", "BR-06", "Should", "3"),
        ("US-08", "Vikram Malhotra (UW)", "As a Chief Underwriting Officer, I want to isolate catastrophic outlier claims using Kernel Density Estimation and Tukey IQR fences so that actuarial loss reserves and reinsurance triggers are accurately maintained.", "F-07", "BR-07", "Should", "3"),
        ("US-09", "Suresh Menon (CFO)", "As a CFO, I want to execute a two-sample independent t-test comparing Cashless and Reimbursement processing turnaround times so that I have empirical statistical proof to justify capital investments in expanding digital hospital networks.", "F-08", "BR-08", "Must", "5"),
        ("US-10", "Ananya Verma (CX)", "As a Customer Experience Lead, I want to analyze customer satisfaction ratings and escalation rates across geographic territories so that I can deploy targeted operational audits to high-friction regions.", "F-09", "BR-09", "Should", "3"),
        ("US-11", "Vikram Malhotra (UW)", "As a Chief Underwriting Officer, I want to explore correlations between policyholder age, customer tenure, premium tier, and claim loss severity so that underwriting risk pricing and premium loadings can be optimized.", "F-10", "BR-10", "Must", "3"),
        ("US-12", "Rajesh Sharma (Ops)", "As a Claims Operations Manager, I want to view an automated data quality audit scorecard and cleaning transformation log so that I can verify data completeness, uniqueness, and integrity before reviewing analytics.", "F-11", "BR-11", "Must", "3"),
        ("US-13", "Suresh Menon (CFO)", "As a CFO, I want to export filtered claims datasets and summary KPI reports into CSV and Excel formats so that compliance teams can perform statutory reviews and external audits.", "F-12", "BR-12", "Could", "2"),
        ("US-14", "All Personas", "As an analytics user, I want a persistent global sidebar filter for policy type, date range, and settlement mode so that all charts and KPI cards across every page dynamically reflect the active cohort without page reloads.", "F-01..12", "BR-01", "Must", "3")
    ]

    col_w_us = [Inches(0.6), Inches(1.3), Inches(2.5), Inches(0.6), Inches(0.6), Inches(0.5), Inches(0.4)]
    for i, row in enumerate(stories_data):
        r = us_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w_us[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 70, 70, 60, 60)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(8)
            if j == 0:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
            elif j in [3, 4, 6]:
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif j == 5:
                c.alignment = Alignment(horizontal="center", vertical="center")
                if val == "Must":
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(0xC5, 0x30, 0x30)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 3: Acceptance Criteria per Story (Gherkin Syntax)
    h1 = doc.add_heading(level=1)
    r = h1.add_run("3. Acceptance Criteria (Gherkin Syntax)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "Each user story is governed by testable acceptance criteria formatted strictly in Given / When / Then syntax:"
    )

    gherkin_specs = [
        ("US-01 (Executive KPI Summary)",
         "Given the Executive Overview dashboard is open, "
         "When the user selects 'Health' in the policy filter, "
         "Then four Tailwind KPI cards display Total Health Claims, Total Settled Amount (₹), Mean Turnaround Time (Days), and Approval Rate (%), recalculating in under 1.0 second."),

        ("US-02 (Stage TAT Bottleneck Analyzer)",
         "Given the Turnaround Time Analytics page is open, "
         "When the user toggles between 'Mean Duration' and 'Median Duration', "
         "Then the stacked bar chart updates to show cycle times for Document Verification, Assessment, and Settlement, highlighting stages exceeding 5 days in red."),

        ("US-03 (Pareto Repudiation Engine)",
         "Given the Denial Analysis screen is rendered, "
         "When the Pareto chart displays repudiated claims, "
         "Then the cumulative percentage curve crosses 80% across the top 2 denial reasons ('POLICY_EXCLUSION' and 'EXCL_PRE_EXIST'), displaying frequency counts and percentage shares."),

        ("US-04 (Settlement Haircut Monitor)",
         "Given approved claims on the Overview page, "
         "When the user adjusts the haircut threshold slider to 20%, "
         "Then all claims where Settled Amount is < 80% of Requested Amount are highlighted in red on the Claimed vs. Settled scatter plot with reference line y = x."),

        ("US-05 (Provider Benchmarking Matrix)",
         "Given the Network Providers page is open, "
         "When the user selects 'Hospital' in the provider type selector, "
         "Then the Cost vs. TAT scatter plot renders with quadrant benchmarking, and Apex Multispeciality Hospital (PRV-104) is tagged with an 'Audit Warning' badge for z > 2.0."),

        ("US-06 (Pre-Settlement Fraud Detector)",
         "Given incoming claims on the Fraud Screening page, "
         "When the user filters for High-Risk claims (Fraud Score >= 70), "
         "Then the triage queue displays claims with early-inception (<30 days) and duplicate flags, summing total monetary exposure in a prominent red risk card."),

        ("US-07 (Seasonal Surge Forecaster)",
         "Given historical claims data, "
         "When the user selects the 30-Day Moving Average view, "
         "Then a smoothed time-series line chart overlays daily claims inflow, showing a pronounced volume surge during Q3 (Monsoon months: July, August, September)."),

        ("US-08 (Catastrophic Outlier Modeler)",
         "Given settled claim amounts on the Underwriting page, "
         "When the outlier distribution is plotted, "
         "Then claims exceeding Q3 + 1.5*IQR (> ₹20 Lakhs) are isolated on the box plot and listed in a dedicated catastrophic reinsurance review table."),

        ("US-09 (Cashless vs. Reimbursement t-test)",
         "Given claims split into Cashless and Reimbursement modes, "
         "When the two-sample independent Welch's t-test executes, "
         "Then the system renders a summary badge confirming p < 0.001 with sample means (Cashless: ~8.5 days, Reimbursement: ~17.5 days) and rejects H0 at alpha = 0.05."),

        ("US-10 (Geographic Grievance Heatmap)",
         "Given claims with customer feedback, "
         "When mapped across Indian territories on the CX screen, "
         "Then states with average CSAT < 3.5 or escalation rate > 15% are shaded in amber/red with actionable ombudsman case counts."),

        ("US-11 (Demographic Correlation Explorer)",
         "Given policyholder demographic and claims data, "
         "When the correlation matrix renders, "
         "Then pairwise Pearson correlation coefficients (r) relating age, customer tenure, annual premium, and claim amounts are displayed with color-graded significance."),

        ("US-12 (Data Quality & Cleaning Log)",
         "Given raw synthetic claims data containing known quality defects, "
         "When the cleaning pipeline executes, "
         "Then the Data Quality page renders a before/after audit table detailing 6 nulls handled, 3 duplicate rows dropped, and 3 outliers flagged."),

        ("US-13 (Tabular Data & Report Exporter)",
         "Given any filtered view on the Data Quality screen, "
         "When the user clicks 'Download Filtered Data (CSV)', "
         "Then a valid, properly formatted CSV file containing the active records downloads to the local browser within 2 seconds."),

        ("US-14 (Global Persistent Sidebar Filter)",
         "Given the Streamlit multipage application is loaded, "
         "When the user selects 'Health' and 'Cashless' in the sidebar, "
         "Then the filter state persists across page navigation using st.session_state, updating all charts without page refresh.")
    ]

    for us_id, gherkin in gherkin_specs:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(f"{us_id}: ")
        r.font.bold = True
        r.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
        rd = p.add_run(gherkin)
        rd.font.size = Pt(9.5)
        rd.font.italic = True

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 4: 2-Week Sprint Implementation Plan
    h1 = doc.add_heading(level=1)
    r = h1.add_run("4. Two-Week Sprint Implementation Roadmap")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "Stories are scheduled across two 5-day sprints to deliver the full production application within a 10-day timeline:"
    )

    sprint_table = doc.add_table(rows=11, cols=4)
    sprint_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(sprint_table, "CBD5E0")

    s_headers = ["Sprint & Day", "Scheduled User Stories", "Primary Technical Focus", "Daily Milestone / Deliverable"]
    for j, h in enumerate(s_headers):
        c = sprint_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 80, 80)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    sprint_days = [
        ("Sprint 1: Day 1", "US-12, US-14", "Data Loader & Sidebar Filters", "SQLite @st.cache_data pipeline & global st.session_state sidebar."),
        ("Sprint 1: Day 2", "US-01, US-04", "Executive Overview Dashboard", "Tailwind KPI metric cards & Claimed vs. Settled scatter plot."),
        ("Sprint 1: Day 3", "US-02, US-07", "Turnaround Time Analytics", "Stacked milestone cycle time chart & 30-day moving average forecaster."),
        ("Sprint 1: Day 4", "US-03, US-10", "Denial & Customer Grievances", "Pareto 80/20 repudiation chart & regional escalation heatmap."),
        ("Sprint 1: Day 5", "Review & QA", "Sprint 1 Regression Testing", "Pytest unit tests on data loader and stage-wise TAT math."),
        ("Sprint 2: Day 6", "US-05", "Provider Benchmarking", "Hospital and garage Cost vs. TAT quadrant benchmarking scatter plot."),
        ("Sprint 2: Day 7", "US-06", "Fraud & Anomaly Screening", "Rule-based fraud scoring engine & high-risk triage exposure queue."),
        ("Sprint 2: Day 8", "US-08, US-09", "Hypothesis & Outlier Testing", "Two-sample Welch's t-test engine & Gaussian KDE outlier distribution."),
        ("Sprint 2: Day 9", "US-11, US-13", "Demographics & Exporting", "Multi-variable bubble chart & in-memory CSV/Excel report exporter."),
        ("Sprint 2: Day 10", "Final Polish", "End-to-End Cloud Deployment", "Pytest AppTest suite pass, README docs, Streamlit Cloud live link.")
    ]

    col_w_s = [Inches(1.3), Inches(1.3), Inches(1.8), Inches(2.1)]
    for i, row in enumerate(sprint_days):
        r = sprint_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w_s[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 70, 70, 70, 70)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(8)
            if j == 0:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 5: Document Control
    h1 = doc.add_heading(level=1)
    r = h1.add_run("5. Document Control & Sprint Baseline Sign-off")
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
        ("Candidate / Scrum Master", "Noman Zaid", "MCA 2nd Year Data Analytics", "APPROVED (Baselined)", "20-Nov-2026"),
        ("Faculty Project Mentor", "Ms. Shilpa", "Assistant Professor / Mentor", "APPROVED (Baselined)", "20-Nov-2026"),
        ("Academic Review Board", "Department Committee", "MCA Department, ABES EC", "APPROVED (Baselined)", "20-Nov-2026")
    ]
    col_w_s2 = [Inches(1.8), Inches(1.3), Inches(1.7), Inches(1.5), Inches(1.1)]
    for i, row in enumerate(sign_data):
        r = sign_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w_s2[j]
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

def update_rtm_with_user_stories():
    rtm_path = r"c:\Users\noman\OneDrive\Desktop\python\hcl-pro\docs\04_Traceability_Matrix.xlsx"
    wb = openpyxl.load_workbook(rtm_path)
    ws = wb["RTM"]

    us_mappings = {
        "BR-01": "US-02, US-14",
        "BR-02": "US-03",
        "BR-03": "US-01, US-04",
        "BR-04": "US-05",
        "BR-05": "US-06",
        "BR-06": "US-07",
        "BR-07": "US-08",
        "BR-08": "US-09",
        "BR-09": "US-10",
        "BR-10": "US-11",
        "BR-11": "US-12",
        "BR-12": "US-13"
    }

    font_us = Font(name="Calibri", size=9, bold=True, color="1B365D")
    fill_us = PatternFill(start_color="EDF2F7", end_color="EDF2F7", fill_type="solid")

    for row in range(5, ws.max_row + 1):
        br_val = ws.cell(row=row, column=3).value
        if br_val in us_mappings:
            cell = ws.cell(row=row, column=10, value=us_mappings[br_val])
            cell.font = font_us
            cell.fill = fill_us
            cell.alignment = Alignment(horizontal="center", vertical="center")

    wb.save(rtm_path)
    print(f"Successfully updated RTM with User Stories: {rtm_path}")

if __name__ == "__main__":
    docx_path = r"c:\Users\noman\OneDrive\Desktop\python\hcl-pro\docs\11_User_Stories.docx"
    generate_user_stories_doc(docx_path)
    update_rtm_with_user_stories()
