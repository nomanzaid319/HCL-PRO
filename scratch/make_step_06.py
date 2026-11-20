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

def generate_prd():
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
    run_title = p_title.add_run("Step 6: Product Requirements Document (PRD)")
    run_title.font.name = 'Calibri'
    run_title.font.size = Pt(24)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(16)
    run_sub = p_sub.add_run("Translating Business Requirements into User Personas, Feature Specifications (F-01 to F-12), and UX Architecture")
    run_sub.font.size = Pt(13)
    run_sub.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)

    # Section 1: Document Control
    h1 = doc.add_heading(level=1)
    r = h1.add_run("1. Document Control & Product Metadata")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    ctrl_table = doc.add_table(rows=6, cols=2)
    ctrl_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(ctrl_table, "CBD5E0")

    ctrl_data = [
        ("Product Name", "Insurance Claim Processing Analysis Workbench"),
        ("Version & Status", "Version 1.0 (Officially Baselined)"),
        ("Product Owner / Author", "Noman Zaid (Roll No: 2500320140103, MCA 2nd Year DA)"),
        ("Faculty Mentor", "Ms. Shilpa, Department of Computer Applications, ABES EC"),
        ("Upstream Baseline", "docs/05_BRD.docx (BR-01 to BR-12, Rules BR-RULE-01 to 08)"),
        ("Target Delivery Stack", "Streamlit Multipage, Tailwind CSS UI via iframe, SQLite (app.db), Plotly, pytest")
    ]
    for i, (k, v) in enumerate(ctrl_data):
        row = ctrl_table.rows[i]
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

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 2: Product Vision & Goals
    h1 = doc.add_heading(level=1)
    r = h1.add_run("2. Product Vision & Value Proposition")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "Product Vision: To deliver an intuitive, responsive, and analytically rigorous claims intelligence platform that transforms fragmented insurance transaction records into actionable operational decisions, enabling insurers to compress claim settlement turnaround time by 45%, intercept fraudulent leakages prior to disbursement, and eliminate unfair claim repudiations."
    )
    p = doc.add_paragraph(
        "Key Value Drivers:\n"
        "1. Real-time Adjudication Transparency: Granular tracking of cycle times across four operational milestones.\n"
        "2. Automated Fraud Screening: Rule-based anomaly engine identifying early-inception and duplicate submissions.\n"
        "3. Network Partner Optimization: Benchmarking healthcare hospitals and repair garages on average costs and rejection rates.\n"
        "4. Statistical Proof of ROI: Rigorous hypothesis testing demonstrating the operational latency reduction of cashless networks."
    )

    # Section 3: Detailed User Personas
    h1 = doc.add_heading(level=1)
    r = h1.add_run("3. User Personas & Target Segments")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "The application is engineered around five distinct user personas across claims operations, risk management, underwriting, customer support, and executive finance:"
    )

    persona_table = doc.add_table(rows=6, cols=4)
    persona_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(persona_table, "CBD5E0")

    p_headers = ["Persona Name", "Job Role", "Primary Motivation / Goal", "Key UI Feature Used"]
    for j, h in enumerate(p_headers):
        c = persona_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 80, 80)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    personas = [
        ("Rajesh Sharma", "Claims Operations Head (SH-01)", "Eliminate processing backlogs, reallocate adjusters, meet IRDAI SLA guidelines.", "Stage TAT Bottleneck Analyzer (F-01) & Surge Forecaster (F-06)"),
        ("Priya Nair", "SIU Fraud Investigation Lead (SH-02)", "Intercept high-risk claims pre-settlement, audit abusive hospital billing tariffs.", "Fraud Anomaly Engine (F-05) & Provider Benchmark Matrix (F-04)"),
        ("Vikram Malhotra", "Chief Underwriting Officer (SH-03)", "Evaluate loss distributions, align premium loadings with demographic cohorts.", "Demographic Explorer (F-10) & Solvency Outlier Modeler (F-07)"),
        ("Ananya Verma", "Customer Experience & SLA Lead (SH-04)", "Reduce policyholder disputes, analyze denial causes, monitor regional CSAT.", "Pareto Repudiation Engine (F-02) & Regional Grievance Heatmap (F-09)"),
        ("Suresh Menon", "Chief Financial Officer (SH-05)", "Maintain capital reserve solvency, validate cashless investment returns.", "Executive KPI Summary (F-03) & Hypothesis Testing Engine (F-08)")
    ]

    col_w_p = [Inches(1.4), Inches(1.5), Inches(2.2), Inches(1.4)]
    for i, row in enumerate(personas):
        r = persona_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w_p[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 70, 70, 80, 80)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(8.5)
            if j == 0:
                run.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 4: Feature Specifications (F-01 to F-12)
    h1 = doc.add_heading(level=1)
    r = h1.add_run("4. Product Feature Specifications (F-01 to F-12)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "Each product feature translates a specific business requirement into concrete UI interactions, data processing logic, and acceptance criteria:"
    )

    feature_specs = [
        ("F-01", "Stage-wise TAT Bottleneck Analyzer", "BR-01", "Must", "Rajesh Sharma (Ops Head)",
         "Calculates and visualizes the average and median cycle time across four distinct claims adjudication stages (Submission -> Document Verification -> Assessment -> Decision). Highlights stages breaching SLA thresholds (> 5 days in a single stage).",
         "Sidebar filters (Policy Type, Date Range), Stage Breakdown View toggle (Mean vs. Median).",
         "Interactive stacked horizontal bar chart, stage cycle time KPI cards, adjudicator backlog count table.",
         "Given the operational claims dataset, When the user selects 'Health Insurance' in the sidebar, Then the system updates the stage TAT chart within 1.0 second, displaying exact average days for each milestone."),

        ("F-02", "Pareto Claim Repudiation Engine", "BR-02", "Must", "Ananya Verma (CX Lead)",
         "Categorizes rejected claims by standardized IRDAI denial codes and renders a dual-axis Pareto 80/20 chart to identify the minority of denial reasons driving the majority of rejections.",
         "Policy Type filter, Repudiation category drill-down selector.",
         "Pareto chart (bar chart of denial counts + cumulative percentage line at 80% threshold), top 5 denial reasons breakdown table.",
         "Given repudiated claims, When the Pareto chart renders, Then the cumulative percentage line crosses 80% across the top 2-3 reasons, displaying the exact frequency and share of total repudiations."),

        ("F-03", "Settlement Payout Haircut Monitor", "BR-03", "Should", "Suresh Menon (CFO)",
         "Evaluates the ratio of claim amount settled to claim amount requested. Computes the haircut percentage ((Claimed - Settled) / Claimed) and highlights claims with deductions > 20%.",
         "Deduction haircut threshold slider (default: 20%), Policy line filter.",
         "Distribution histogram of payout ratios, scatter plot of Claimed vs. Settled Amount with reference diagonal (y = x), haircut summary KPI card.",
         "Given approved claims, When the haircut threshold is adjusted to 15%, Then all claims with > 15% deductions are dynamically highlighted in red on the scatter plot."),

        ("F-04", "Provider Cost & Performance Benchmarking Matrix", "BR-04", "Must", "Priya Nair (SIU Lead)",
         "Aggregates claim volume, average settlement cost, repudiation frequency, and average processing TAT across empaneled hospitals and auto repair garages, flagging statistical outliers.",
         "Provider Category selector (Hospitals vs. Garages vs. TPAs), Region filter, Sort by Cost/TAT/Rejection rate.",
         "Scatter plot of Provider Average Cost vs. Processing TAT with quadrant benchmarking, tabular leaderboard of top 10 highest-cost providers.",
         "Given network provider data, When filtered for 'Hospitals' in North zone, Then providers exceeding 2 standard deviations above the regional peer mean are tagged with an 'Audit Warning' badge."),

        ("F-05", "Pre-Settlement Fraud Anomaly Detector", "BR-05", "Must", "Priya Nair (SIU Lead)",
         "Scans claims against business rules: early inception (< 30 days), duplicate policyholder/incident date, and inflated bills. Assigns a composite fraud score (0-100) and aggregates financial exposure.",
         "Fraud Score threshold slider (default: >= 70), Anomaly type multi-select.",
         "High-Risk Claims Queue table with one-click SIU review button, Total Financial Risk Exposure metric card, Anomaly breakdown donut chart.",
         "Given incoming claims, When scored by the anomaly engine, Then all synthetic records with early-inception or duplicate flags are displayed in the high-risk triage table with score >= 70."),

        ("F-06", "Seasonal Surge & Volume Forecaster", "BR-06", "Should", "Rajesh Sharma (Ops Head)",
         "Aggregates claims inflow over time (monthly and quarterly), computes a 30-day moving average, and highlights historical surge windows (monsoon flood auto spikes, post-monsoon dengue health spikes).",
         "Date Granularity selector (Monthly, Quarterly), Policy Type filter.",
         "Time-series line chart with 30-day rolling moving average trend line, Month-over-Month (MoM) volume growth bar chart, Surge warning banner.",
         "Given historical claims dates, When the user toggles '30-day Moving Average', Then a smoothed trend line renders over daily bars showing peak inflow periods."),

        ("F-07", "Catastrophic Outlier & Solvency Modeler", "BR-07", "Should", "Vikram Malhotra (Underwriting)",
         "Models the statistical distribution of claim amounts using Kernel Density Estimation (KDE) and box plots. Isolates extreme positive outliers exceeding Q3 + 1.5×IQR and 99th percentile for reinsurance review.",
         "Log scale toggle for skewed financial data, Policy line filter.",
         "KDE distribution plot with normal curve overlay, box plot showing Tukey fences, catastrophic claims detail table.",
         "Given settled claim amounts, When the outlier modeler runs, Then claims exceeding the upper IQR fence are listed with their coverage limit and solvency reserve impact."),

        ("F-08", "Cashless vs. Reimbursement Hypothesis Testing Engine", "BR-08", "Must", "Suresh Menon (CFO)",
         "Executes a formal two-sample independent t-test comparing the processing turnaround time of Cashless claims vs. Reimbursement claims. Outputs sample means, t-statistic, p-value, and effect size.",
         "Significance level alpha selector (default: 0.05), Policy Type filter.",
         "Hypothesis test summary card (Reject/Fail to Reject H0), side-by-side violin/box plots, statistical parameters table.",
         "Given claims split by settlement mode, When the t-test executes, Then the system outputs p < 0.05, confirming that Cashless processing is statistically significantly faster than Reimbursement."),

        ("F-09", "Geographic Grievance & Escalation Heatmap", "BR-09", "Should", "Ananya Verma (CX Lead)",
         "Maps claim disputes, customer satisfaction (CSAT) scores, and regulatory ombudsman escalations across geographic states and operational zones.",
         "Metric selector (Escalation Rate % vs. Average CSAT), Policy Type filter.",
         "Regional choropleth / bar heatmap, territory grievance scorecard ranking worst-to-best performing zones.",
         "Given customer service records, When mapped by state, Then territories with average CSAT < 3.5 are shaded in amber/red with actionable escalation counts."),

        ("F-10", "Underwriting Demographic Correlation Explorer", "BR-10", "Must", "Vikram Malhotra (Underwriting)",
         "Analyzes relationships between policyholder age, customer tenure, annual premium, and claim loss severity using Pearson correlation matrices and multi-variable bubble charts.",
         "Correlation method selector (Pearson vs. Spearman), Bubble size metric selector.",
         "Interactive correlation heatmap with r-values, multi-variable bubble chart (X: Age, Y: Claim Amount, Size: Premium, Color: Policy Type).",
         "Given policyholder demographic data, When the correlation matrix renders, Then pairwise correlation coefficients are displayed with color-graded significance indicators."),

        ("F-11", "Data Quality Audit & Validation Pipeline", "BR-11", "Must", "Rajesh Sharma (Ops Head)",
         "Monitors dataset integrity, detects missing values, duplicate identifier keys, and out-of-range anomalies, providing a transparent before/after cleaning audit trail.",
         "View selector (Raw Data vs. Cleaned Data vs. Audit Log).",
         "Data quality scorecard (Completeness %, Uniqueness %, Validity %), tabular before/after transformation log detailing rows affected per cleaning rule.",
         "Given raw synthetic claims data containing known quality defects, When the validation pipeline executes, Then 100% of anomalies are flagged and logged in the transformation audit table."),

        ("F-12", "Tabular Dataset & Compliance Report Exporter", "BR-12", "Could", "Suresh Menon (CFO)",
         "Enables users to download filtered claims operational datasets, summary KPI tables, and compliance reports in CSV and Excel formats.",
         "Export Format selector (CSV, XLSX), Export Scope toggle (Active Filtered View vs. Full Dataset).",
         "Download button triggering immediate browser file download with timestamped filename.",
         "Given any active filtered dashboard view, When the user clicks 'Download Filtered Data (CSV)', Then a valid CSV matching the filter criteria downloads within 2 seconds.")
    ]

    for f_id, f_name, linked_br, prio, persona, desc, inputs, outputs, ac in feature_specs:
        h2 = doc.add_heading(level=2)
        r2 = h2.add_run(f"{f_id}: {f_name}")
        r2.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)

        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("Linked Business Requirement: ")
        r.font.bold = True
        p.add_run(f"{linked_br}  |  ")
        r = p.add_run("MoSCoW Priority: ")
        r.font.bold = True
        r_prio = p.add_run(f"{prio} Have")
        if prio == "Must":
            r_prio.font.bold = True
            r_prio.font.color.rgb = RGBColor(0xC5, 0x30, 0x30)
        p.add_run("  |  ")
        r = p.add_run("Primary Persona: ")
        r.font.bold = True
        p.add_run(persona)

        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("Functional Description: ")
        r.font.bold = True
        p.add_run(desc)

        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("User Controls & Inputs: ")
        r.font.bold = True
        p.add_run(inputs)

        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("Visual Display & Outputs: ")
        r.font.bold = True
        p.add_run(outputs)

        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        r = p.add_run("Acceptance Criteria: ")
        r.font.bold = True
        r_ac = p.add_run(ac)
        r_ac.font.italic = True

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 5: Non-Functional Requirements (NFR-01 to NFR-06)
    h1 = doc.add_heading(level=1)
    r = h1.add_run("5. Non-Functional Requirements (NFR-01 to NFR-06)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    nfr_table = doc.add_table(rows=7, cols=4)
    nfr_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(nfr_table, "CBD5E0")

    nfr_headers = ["NFR ID", "Quality Dimension", "Non-Functional Specification", "Measurement / Target Metric"]
    for j, h in enumerate(nfr_headers):
        c = nfr_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 80, 80)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    nfr_data = [
        ("NFR-01", "Performance Latency", "Initial app load under 3.0 seconds; filter updates and chart re-rendering under 1.0 second.", "Streamlit render timer < 1.0s on cached queries."),
        ("NFR-02", "Memory Efficiency", "Total runtime in-memory consumption must remain strictly under 600 MB to prevent Streamlit Cloud eviction.", "Peak RAM < 600 MB (safety margin on 1 GB limit)."),
        ("NFR-03", "Data Privacy & Ethics", "Zero actual customer PII; all names, contacts, and claim identifiers must be synthetically generated.", "100% synthetic validation audit."),
        ("NFR-04", "Responsive Usability", "UI layout must dynamically adapt across standard desktop 1080p, laptop 720p, and tablet viewports.", "Tailwind CSS responsive grid system."),
        ("NFR-05", "Automated Testability", "Every core calculation and page must be validated via pytest with >= 90% pass rate.", "Pytest test suite execution pass rate >= 90%."),
        ("NFR-06", "Code Maintainability", "Modular code structure separating database queries (src/data_loader.py), analysis (src/analytics.py), and UI components.", "Flake8 / clean module separation.")
    ]

    col_w_nfr = [Inches(0.9), Inches(1.5), Inches(2.6), Inches(1.5)]
    for i, row in enumerate(nfr_data):
        r = nfr_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w_nfr[j]
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

    # Section 6: Information Architecture & Screen Navigation Map
    h1 = doc.add_heading(level=1)
    r = h1.add_run("6. Information Architecture & Navigation Structure")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "The application organizes features into a multipage Streamlit architecture across 7 specialized screens, unified by a global persistent sidebar filter:"
    )

    nav_table = doc.add_table(rows=8, cols=4)
    nav_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(nav_table, "CBD5E0")

    nav_headers = ["Screen ID", "Screen Name / Page File", "Target Persona", "Hosted Feature Modules"]
    for j, h in enumerate(nav_headers):
        c = nav_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 80, 80)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    nav_data = [
        ("UI-01", "Executive Overview (app.py)", "All Personas / CFO", "High-level KPI metric cards, Total Payout, Claim Status split, Settlement Haircut Monitor (F-03)."),
        ("UI-02", "Turnaround Time Analytics (pages/1_TAT_Analytics.py)", "Rajesh Sharma (Ops)", "Stage-wise TAT Bottleneck Analyzer (F-01), Milestone duration breakdown, SLA breach table."),
        ("UI-03", "Denial & Repudiation Analysis (pages/2_Denial_Analysis.py)", "Ananya Verma (CX)", "Pareto Claim Repudiation Engine (F-02), Denial cause breakdown by policy line."),
        ("UI-04", "Network Provider Benchmarking (pages/3_Providers.py)", "Priya Nair (SIU)", "Provider Cost & Performance Matrix (F-04), Hospital/Garage outlier tariff scatter plot."),
        ("UI-05", "Fraud & Anomaly Screening (pages/4_Fraud_Screening.py)", "Priya Nair (SIU)", "Pre-Settlement Fraud Detector (F-05), Early inception queue, Total monetary exposure sum."),
        ("UI-06", "Underwriting & Demographics (pages/5_Underwriting.py)", "Vikram Malhotra (Underwriting)", "Demographic Explorer (F-10), Catastrophic Outlier Modeler (F-07), Cashless t-test (F-08)."),
        ("UI-07", "Data Quality & Export (pages/6_Data_Quality.py)", "All Personas", "Data Quality Audit Pipeline (F-11), Before/After cleaning log, CSV/Excel Data Exporter (F-12).")
    ]

    col_w_nav = [Inches(1.0), Inches(2.2), Inches(1.5), Inches(1.8)]
    for i, row in enumerate(nav_data):
        r = nav_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w_nav[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 70, 70, 80, 80)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(8.5)
            if j == 0:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 7: Sign-off
    h1 = doc.add_heading(level=1)
    r = h1.add_run("7. Product Baseline Sign-off")
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
        ("Candidate / Product Lead", "Noman Zaid", "MCA 2nd Year Data Analytics", "APPROVED (Baselined)", "20-Nov-2026"),
        ("Faculty Project Mentor", "Ms. Shilpa", "Assistant Professor / Mentor", "APPROVED (Baselined)", "20-Nov-2026"),
        ("Academic Review Board", "Department Committee", "MCA Department, ABES EC", "APPROVED (Baselined)", "20-Nov-2026")
    ]
    col_w_s = [Inches(1.8), Inches(1.3), Inches(1.7), Inches(1.5), Inches(1.1)]
    for i, row in enumerate(sign_data):
        r = sign_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w_s[j]
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

    output_path = r"c:\Users\noman\OneDrive\Desktop\python\hcl-pro\docs\06_PRD.docx"
    doc.save(output_path)
    print(f"Successfully generated {output_path}")

def update_rtm_with_features():
    rtm_path = r"c:\Users\noman\OneDrive\Desktop\python\hcl-pro\docs\04_Traceability_Matrix.xlsx"
    wb = openpyxl.load_workbook(rtm_path)
    ws = wb["RTM"]

    feature_mappings = {
        "BR-01": "F-01 (Stage TAT Analyzer)",
        "BR-02": "F-02 (Pareto Denial Engine)",
        "BR-03": "F-03 (Payout Haircut Monitor)",
        "BR-04": "F-04 (Provider Matrix)",
        "BR-05": "F-05 (Fraud Anomaly Detector)",
        "BR-06": "F-06 (Surge Forecaster)",
        "BR-07": "F-07 (Catastrophic Outlier Modeler)",
        "BR-08": "F-08 (Cashless t-test Engine)",
        "BR-09": "F-09 (Regional Grievance Heatmap)",
        "BR-10": "F-10 (Demographic Explorer)",
        "BR-11": "F-11 (Data Quality Pipeline)",
        "BR-12": "F-12 (Dataset Exporter)"
    }

    font_feat = Font(name="Calibri", size=9.5, bold=True, color="1B365D")
    fill_feat = PatternFill(start_color="EBF8FF", end_color="EBF8FF", fill_type="solid")

    for row in range(5, ws.max_row + 1):
        br_val = ws.cell(row=row, column=3).value
        if br_val in feature_mappings:
            cell = ws.cell(row=row, column=8, value=feature_mappings[br_val])
            cell.font = font_feat
            cell.fill = fill_feat
            cell.alignment = Alignment(horizontal="center", vertical="center")

    wb.save(rtm_path)
    print(f"Successfully updated RTM with Feature IDs: {rtm_path}")

if __name__ == "__main__":
    generate_prd()
    update_rtm_with_features()
