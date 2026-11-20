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

def generate_ui_framework_doc(output_docx_path):
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
    run_title = p_title.add_run("Step 12: UI Framework & Design System")
    run_title.font.name = 'Calibri'
    run_title.font.size = Pt(24)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(16)
    run_sub = p_sub.add_run("Streamlit Multipage Architecture, Tailwind CSS UI Components, Screen Wireframes (UI-01 to UI-07), and Chart Selection Rationale")
    run_sub.font.size = Pt(13)
    run_sub.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)

    # Metadata Box
    meta_table = doc.add_table(rows=5, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(meta_table, "CBD5E0")
    meta_data = [
        ("Project Name", "Insurance Claim Processing Analysis"),
        ("Learner / UI Architect", "Noman Zaid (Roll No: 2500320140103, MCA 2nd Year DA)"),
        ("Faculty Mentor", "Ms. Shilpa, Department of Computer Applications, ABES EC"),
        ("UI Technology Stack", "Streamlit 1.40.2 Multipage, Tailwind CSS CDN via iframe, Plotly 5.24.1"),
        ("Target Display Standard", "Desktop 1080p primary, responsive laptop 720p & mobile viewport support")
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

    # Section 1: Screens Master Table
    h1 = doc.add_heading(level=1)
    r = h1.add_run("1. UI Screens Master Architecture (UI-01 to UI-07)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "The application organizes user functionality into 7 cohesive analytical screens unified by persistent global state management:"
    )

    screen_table = doc.add_table(rows=8, cols=6)
    screen_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(screen_table, "CBD5E0")

    s_headers = ["UI ID", "Screen Name & File", "Screen Purpose", "Key Visual Components", "Interactive Filters", "User Stories"]
    for j, h in enumerate(s_headers):
        c = screen_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 60, 60)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    screens = [
        ("UI-01", "Executive Overview (app.py)", "Portfolio-wide operational scorecard and claim haircut monitoring", "Tailwind KPI cards (Claims, Payout, TAT, Approval %), Haircut scatter plot, Claim status donut", "Policy Type, Date Range, Settlement Mode", "US-01, US-04, US-14"),
        ("UI-02", "Turnaround Time (pages/1_TAT_Analytics.py)", "Cycle time tracking, stage delay diagnosis, and surge forecasting", "Stacked horizontal bar of milestones, TAT distribution violin plot, 30-day moving average time-series", "Stage toggle (Mean vs Median), Time Horizon", "US-02, US-07"),
        ("UI-03", "Denial Analysis (pages/2_Denial_Analysis.py)", "Root cause repudiation analysis and regional grievance tracking", "Dual-axis Pareto 80/20 chart, Denial causes ranking table, Regional grievance heatmap", "Policy Type, Denial Category, Minimum Threshold", "US-03, US-10"),
        ("UI-04", "Network Providers (pages/3_Providers.py)", "Empaneled hospital and auto garage performance benchmarking", "Cost vs. TAT quadrant scatter plot, Top 10 expensive providers leaderboard, Tariff audit badges", "Provider Type (Hospital/Garage), Region, Status", "US-05"),
        ("UI-05", "Fraud Screening (pages/4_Fraud_Screening.py)", "Anomaly detection, early-inception auditing, and financial risk exposure", "High-risk triage queue table, Total exposure metric card, Anomaly breakdown pie, Fraud score histogram", "Fraud Score Slider (>=70), Anomaly Type", "US-06"),
        ("UI-06", "Actuarial & Underwriting (pages/5_Underwriting.py)", "Statistical hypothesis testing, demographic correlations, and outlier fences", "Welch's t-test summary card, Correlation matrix heatmap, Multi-variable bubble chart, Box/KDE plot", "Significance Alpha (0.05), Bubble size variable", "US-08, US-09, US-11"),
        ("UI-07", "Data Quality & Export (pages/6_Data_Quality.py)", "Data provenance audit, cleaning transformation log, and tabular report export", "Data health scorecard (Completeness/Uniqueness), Before/After audit table, CSV & Excel download buttons", "Table selector, Export format (CSV/XLSX)", "US-12, US-13")
    ]

    col_w_sc = [Inches(0.6), Inches(1.5), Inches(1.4), Inches(1.5), Inches(1.0), Inches(0.8)]
    for i, row in enumerate(screens):
        r = screen_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w_sc[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 70, 70, 60, 60)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(8)
            if j == 0:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 2: Detailed ASCII Wireframes
    h1 = doc.add_heading(level=1)
    r = h1.add_run("2. Detailed Screen ASCII Wireframes")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "Complete structural ASCII layouts specifying headers, sidebar filters, Tailwind KPI rows, Plotly charts, and tables across all primary screens:"
    )

    wireframes = [
        ("2.1 Wireframe: UI-01 Executive Overview (app.py)",
         "+---------------------------------------------------------------------------------------------------------+\n"
         "| [SIDEBAR FILTERS]              | [TAILWIND HEADER: Insurance Claim Processing Analysis Workbench]        |\n"
         "| - Policy Type: [All / Health]  | Noman Zaid (2500320140103) · MCA 2nd Yr DA · ABES EC · Mentor: Ms. Shilpa |\n"
         "| - Settlement: [All / Cashless] +-------------------------------------------------------------------------+\n"
         "| - Date Range: 2023-01..2024-12 | [TAILWIND KPI ROW (via st.iframe)]                                      |\n"
         "| - Status: [All / Approved]     | [Total Claims: 250] [Total Settled: ₹1.84 Cr] [Mean TAT: 12.2d] [Appr: 72%] |\n"
         "|                                +-------------------------------------------------------------------------+\n"
         "| Navigation:                    | [PLOTLY ROW 1]                                                          |\n"
         "| > 1_TAT_Analytics              | +-----------------------------------+ +---------------------------------+\n"
         "| > 2_Denial_Analysis            | | Claims by Status & Policy Line    | | Claimed vs. Settled Amount      |\n"
         "| > 3_Providers                  | | (Grouped Bar: Policy x Status)    | | (Scatter Plot with y=x line)    |\n"
         "| > 4_Fraud_Screening            | +-----------------------------------+ +---------------------------------+\n"
         "| > 5_Underwriting               | [ANALYTICAL CALLOUT: Portfolio Payout & Haircut Summary]                |\n"
         "| > 6_Data_Quality               | Average settlement haircut is 11.4%. 18 claims flagged with haircut >20%.|\n"
         "+--------------------------------+-------------------------------------------------------------------------+\n"),

        ("2.2 Wireframe: UI-02 Turnaround Time Analytics (pages/1_TAT_Analytics.py)",
         "+---------------------------------------------------------------------------------------------------------+\n"
         "| [SIDEBAR FILTERS]              | [HEADER: Stage-wise Turnaround Time & Processing Bottlenecks]           |\n"
         "| - View: [Mean / Median]        +-------------------------------------------------------------------------+\n"
         "| - Policy Line: [Health]        | [KPI ROW] [Verification: 3.4d] [Assessment: 4.8d] [Settlement: 4.0d]     |\n"
         "| - SLA Breach: [Show > 15d]     +-------------------------------------------------------------------------+\n"
         "|                                | [CHART: Stage-wise Cycle Time Breakdown across Policy Types]            |\n"
         "|                                | [Stacked Horizontal Bar Chart: Document Intake | Medical Audit | Pay]   |\n"
         "|                                +-------------------------------------------------------------------------+\n"
         "|                                | [CHART: Seasonal Inflow Volume & 30-Day Moving Average]                 |\n"
         "|                                | [Time-Series Dual Plot: Daily Inflow Bars + Rolling 30d Trend Line]     |\n"
         "|                                +-------------------------------------------------------------------------+\n"
         "|                                | [TABLE: SLA Breach Exception Queue - 14 Claims Exceeding 15-Day Target] |\n"
         "+--------------------------------+-------------------------------------------------------------------------+\n"),

        ("2.3 Wireframe: UI-03 Denial & Repudiation Analysis (pages/2_Denial_Analysis.py)",
         "+---------------------------------------------------------------------------------------------------------+\n"
         "| [SIDEBAR FILTERS]              | [HEADER: Claim Repudiation & Root Cause Decomposition]                  |\n"
         "| - Policy Line: [Motor]         +-------------------------------------------------------------------------+\n"
         "| - Min Rejection Threshold: 5   | [KPI ROW] [Repudiated Claims: 70] [Denial Rate: 28.0%] [Top 2 Share: 81.5%]|\n"
         "|                                +-------------------------------------------------------------------------+\n"
         "|                                | [CHART: Dual-Axis Pareto 80/20 Repudiation Analysis]                     |\n"
         "|                                | [Bars: Denial Frequency] [Line: Cumulative % crossing 80% threshold]   |\n"
         "|                                +-------------------------------------------------------------------------+\n"
         "|                                | [CHART: Regional Customer Grievance Heatmap & CSAT Ratings]             |\n"
         "|                                | [State-wise Horizontal Bar Heatmap: Escalation % + Mean CSAT Score]     |\n"
         "+--------------------------------+-------------------------------------------------------------------------+\n"),

        ("2.4 Wireframe: UI-04 Network Provider Benchmarking (pages/3_Providers.py)",
         "+---------------------------------------------------------------------------------------------------------+\n"
         "| [SIDEBAR FILTERS]              | [HEADER: Network Healthcare Hospital & Auto Garage Benchmarking]        |\n"
         "| - Facility Type: [Hospital]    +-------------------------------------------------------------------------+\n"
         "| - Region: [North / West]       | [CHART: Provider Cost vs. Turnaround Time Quadrant Benchmark]           |\n"
         "| - Outlier Threshold: [z > 2.0] | [Quadrant Scatter: X=Mean TAT, Y=Mean Cost, Size=Volume, Color=Status]  |\n"
         "|                                | * Apex Multispeciality Hospital (PRV-104) flagged in Red (z = 2.45)     |\n"
         "|                                +-------------------------------------------------------------------------+\n"
         "|                                | [TABLE: Provider Performance Leaderboard & Tariff Audit Warnings]       |\n"
         "+--------------------------------+-------------------------------------------------------------------------+\n"),

        ("2.5 Wireframe: UI-05 Fraud & Anomaly Screening (pages/4_Fraud_Screening.py)",
         "+---------------------------------------------------------------------------------------------------------+\n"
         "| [SIDEBAR FILTERS]              | [HEADER: Pre-Settlement Fraud Detection & Anomaly Triage Queue]         |\n"
         "| - Fraud Score Slider: [>= 70]  +-------------------------------------------------------------------------+\n"
         "| - Inception Window: [< 30d]    | [RED ALERT CARD: Total Monetary Risk Exposure = ₹ 38,42,000 across 22 Claims] |\n"
         "|                                +-------------------------------------------------------------------------+\n"
         "|                                | [CHART: Fraud Risk Score Distribution & Anomaly Flag Breakdown]         |\n"
         "|                                +-------------------------------------------------------------------------+\n"
         "|                                | [TABLE: High-Risk Claims Triage Queue (Early Inception, Duplicates)]     |\n"
         "+--------------------------------+-------------------------------------------------------------------------+\n"),

        ("2.6 Wireframe: UI-06 Actuarial Insights & Hypothesis (pages/5_Underwriting.py)",
         "+---------------------------------------------------------------------------------------------------------+\n"
         "| [SIDEBAR FILTERS]              | [HEADER: Statistical Hypothesis Testing & Actuarial Loss Reserving]     |\n"
         "| - Significance Alpha: [0.05]   +-------------------------------------------------------------------------+\n"
         "| - Bubble Size: [Premium]       | [HYPOTHESIS TEST BADGE: Welch's t-test p < 0.001 (Reject H0 - Stat. Sig)]|\n"
         "|                                | Cashless Mean = 8.49d vs. Reimbursement Mean = 17.49d (Delta = 9.0 days)|\n"
         "|                                +-------------------------------------------------------------------------+\n"
         "|                                | [CHART: Side-by-Side Violin & Box Plots: Cashless vs. Reimbursement]    |\n"
         "|                                +-------------------------------------------------------------------------+\n"
         "|                                | [CHART: Kernel Density Estimation (KDE) & Tukey Outlier Fences]         |\n"
         "|                                +-------------------------------------------------------------------------+\n"
         "|                                | [CHART: Pearson Demographic Correlation Matrix & Bubble Scatter]        |\n"
         "+--------------------------------+-------------------------------------------------------------------------+\n"),

        ("2.7 Wireframe: UI-07 Data Quality & Export (pages/6_Data_Quality.py)",
         "+---------------------------------------------------------------------------------------------------------+\n"
         "| [SIDEBAR CONTROLS]             | [HEADER: Data Quality Audit Scorecard & Compliance Report Exporter]     |\n"
         "| - Select Table: [Claims Fact]  +-------------------------------------------------------------------------+\n"
         "| - Export Format: [CSV / XLSX]  | [QUALITY SCORECARD: Completeness: 97.6% | Uniqueness: 100% | FK: Valid] |\n"
         "|                                +-------------------------------------------------------------------------+\n"
         "|                                | [AUDIT TABLE: Before vs. After Cleaning Transformation Log]             |\n"
         "|                                | - 6 Null CSAT scores handled | 3 Duplicates dropped | 3 Outliers flagged|\n"
         "|                                +-------------------------------------------------------------------------+\n"
         "|                                | [BUTTON: Download Filtered Dataset (CSV)] [BUTTON: Download Excel (.xlsx)]|\n"
         "+--------------------------------+-------------------------------------------------------------------------+\n")
    ]

    for title, wf_ascii in wireframes:
        h2 = doc.add_heading(level=2)
        r2 = h2.add_run(title)
        r2.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)

        box = doc.add_table(rows=1, cols=1)
        box.alignment = WD_TABLE_ALIGNMENT.CENTER
        c = box.rows[0].cells[0]
        c.width = Inches(6.5)
        set_cell_background(c, "2D3748") # dark slate
        set_cell_margins(c, 70, 70, 90, 90)
        set_table_borders(box, "1A202C")
        cp = c.paragraphs[0]
        cp.paragraph_format.space_after = Pt(0)
        cr = cp.add_run(wf_ascii)
        cr.font.name = "Consolas"
        cr.font.size = Pt(7.5)
        cr.font.color.rgb = RGBColor(0xED, 0xF2, 0xF7)

        doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Section 3: Global Design System & Theme
    h1 = doc.add_heading(level=1)
    r = h1.add_run("3. Global Design System & State Management")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "• Shared Sidebar Filter State: Filters for Policy Type, Date Range, and Settlement Mode reside in `st.session_state`. Selecting a filter on the Overview page immediately propagates across all child pages without resetting.\n"
        "• 4-Color Corporate Palette:\n"
        "  1. Primary Deep Navy (`#1B365D`): Page headers, primary chart traces, active navigation.\n"
        "  2. Secondary Slate Blue (`#2B6CB0`): Sub-headers, secondary chart bars, table accents.\n"
        "  3. Alert Coral / Crimson (`#C53030`): High-risk fraud alerts, SLA breaches, repudiations.\n"
        "  4. Success Emerald Green (`#276749`): Approved claims, cashless status, data quality passes.\n"
        "• Background Palette: Soft Off-White (`#F8FAFC`) with Pure White (`#FFFFFF`) card containers.\n"
        "• Typography: System clean sans-serif (Inter, Calibri, Segoe UI) with tabular numbers for financial figures."
    )

    # Section 4: Best Chart for Each OQ
    h1 = doc.add_heading(level=1)
    r = h1.add_run("4. Analytical Chart Selection Rationale (OQ-01 to OQ-10)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "Every visualization was selected based on strict perceptual and statistical design principles:"
    )

    chart_table = doc.add_table(rows=11, cols=4)
    chart_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(chart_table, "CBD5E0")

    c_headers = ["OQ ID", "Operational Question", "Recommended Visualization", "Design Rationale / Why This Wins"]
    for j, h in enumerate(c_headers):
        c = chart_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 70, 70)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    chart_rationale = [
        ("OQ-01", "Stage Turnaround Times", "Stacked Horizontal Bar Chart", "Stacked bars cleanly represent sequential milestone progression (Verification -> Assessment -> Decision) while maintaining a common zero baseline for total TAT comparison."),
        ("OQ-02", "Repudiation Root Causes", "Dual-Axis Pareto 80/20 Chart", "Combines discrete frequency bars with a cumulative percentage ogive line, immediately highlighting the 'vital few' denial reasons responsible for 80% of repudiations."),
        ("OQ-03", "Payout Haircuts & Deductions", "Scatter Plot with 45° Reference Line (y=x)", "Plotting Requested (X) vs. Settled (Y) against a y=x reference line visually exposes the magnitude of deduction haircuts as the vertical distance from the line."),
        ("OQ-04", "Network Provider Benchmarking", "Quadrant Scatter Plot with Z-Score Bands", "Maps Average Cost (Y) vs. Processing TAT (X) with quadrant threshold lines (z=2.0), instantly isolating abusive, high-cost and slow providers in the top-right quadrant."),
        ("OQ-05", "Fraud Risk & Exposure", "Donut Chart + High-Risk Triage Queue", "A donut chart displays the proportion of normal vs. high-risk claims, paired with an actionable interactive table summing total monetary risk exposure in real time."),
        ("OQ-06", "Seasonal Surge Forecasting", "Time-Series Dual Plot (Bars + 30d Moving Avg)", "Daily inflow volume bars capture short-term operational volatility, while the 30-day moving average curve filters out noise to reveal predictable seasonal monsoon surges."),
        ("OQ-07", "Catastrophic Outlier Claims", "Kernel Density Estimation (KDE) + Box Plot", "KDE plots visualize continuous parametric shape and right-skewness, while box plots with Tukey fences (Q3 + 1.5*IQR) formally mark catastrophic reinsurance outliers."),
        ("OQ-08", "Cashless vs. Reimbursement", "Side-by-Side Violin & Box Plots", "Violin plots illustrate the complete probability distribution and modality of both settlement modes, proving the lower mean and tighter variance of Cashless claims."),
        ("OQ-09", "Territorial Grievances & CSAT", "Horizontal Bar Heatmap / Choropleth", "Color-coded bars sorted from worst-to-best immediately highlight regional operating branches with dispute escalation rates > 15% and average CSAT < 3.5."),
        ("OQ-10", "Demographic Risk Correlation", "Correlation Heatmap + Multi-Variable Bubble", "A correlation matrix quantifies pairwise linear associations (r), while a 4-variable bubble chart (Age, Claim, Premium, Type) captures multi-dimensional underwriting risk.")
    ]

    col_w_cr = [Inches(0.7), Inches(1.8), Inches(2.0), Inches(2.0)]
    for i, row in enumerate(chart_rationale):
        r = chart_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w_cr[j]
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
    r = h1.add_run("5. Document Control & Sign-off")
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
        ("Candidate / UI Architect", "Noman Zaid", "MCA 2nd Year Data Analytics", "APPROVED (Baselined)", "20-Nov-2026"),
        ("Faculty Project Mentor", "Ms. Shilpa", "Assistant Professor / Mentor", "APPROVED (Baselined)", "20-Nov-2026"),
        ("Academic Review Board", "Department Committee", "MCA Department, ABES EC", "APPROVED (Baselined)", "20-Nov-2026")
    ]
    col_w_s3 = [Inches(1.8), Inches(1.3), Inches(1.7), Inches(1.5), Inches(1.1)]
    for i, row in enumerate(sign_data):
        r = sign_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w_s3[j]
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

def update_rtm_with_ui_screens():
    rtm_path = r"c:\Users\noman\OneDrive\Desktop\python\hcl-pro\docs\04_Traceability_Matrix.xlsx"
    wb = openpyxl.load_workbook(rtm_path)
    ws = wb["RTM"]

    ui_mappings = {
        "BR-01": "UI-02 (TAT Analytics)",
        "BR-02": "UI-03 (Denial Analysis)",
        "BR-03": "UI-01 (Executive Overview)",
        "BR-04": "UI-04 (Provider Matrix)",
        "BR-05": "UI-05 (Fraud Screening)",
        "BR-06": "UI-02 (TAT Analytics)",
        "BR-07": "UI-06 (Underwriting & Outliers)",
        "BR-08": "UI-06 (Underwriting & t-test)",
        "BR-09": "UI-03 (Denial Analysis)",
        "BR-10": "UI-06 (Underwriting Demographics)",
        "BR-11": "UI-07 (Data Quality Audit)",
        "BR-12": "UI-07 (Data Quality & Export)"
    }

    font_ui = Font(name="Calibri", size=9, bold=True, color="2B6CB0")
    fill_ui = PatternFill(start_color="EBF8FF", end_color="EBF8FF", fill_type="solid")

    for row in range(5, ws.max_row + 1):
        br_val = ws.cell(row=row, column=3).value
        if br_val in ui_mappings:
            cell = ws.cell(row=row, column=11, value=ui_mappings[br_val])
            cell.font = font_ui
            cell.fill = fill_ui
            cell.alignment = Alignment(horizontal="center", vertical="center")

    wb.save(rtm_path)
    print(f"Successfully updated RTM with UI Screens: {rtm_path}")

if __name__ == "__main__":
    docx_path = r"c:\Users\noman\OneDrive\Desktop\python\hcl-pro\docs\12_UI_Framework.docx"
    generate_ui_framework_doc(docx_path)
    update_rtm_with_ui_screens()
