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

def create_document():
    doc = Document()

    # Page Margins
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Base Normal Style
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(6)

    # Document Pre-Header
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
    run_title = p_title.add_run("Step 5: Business Requirements Document (BRD)")
    run_title.font.name = 'Calibri'
    run_title.font.size = Pt(24)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(16)
    run_sub = p_sub.add_run("Enterprise Business Requirements Specification for Insurance Claim Processing & Operational Analytics")
    run_sub.font.size = Pt(13)
    run_sub.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)

    # Section 1: Document Control
    h1 = doc.add_heading(level=1)
    r = h1.add_run("1. Document Control & Metadata")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    ctrl_table = doc.add_table(rows=6, cols=2)
    ctrl_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(ctrl_table, "CBD5E0")
    
    ctrl_data = [
        ("Document Title", "Business Requirements Document (BRD) - Insurance Claim Processing Analysis"),
        ("Version & Status", "Version 1.0 (Officially Baselined)"),
        ("Primary Author", "Noman Zaid (Roll No: 2500320140103, MCA 2nd Year Data Analytics)"),
        ("Faculty Project Mentor", "Ms. Shilpa, Department of Computer Applications, ABES Engineering College"),
        ("Release Date", "20-November-2026"),
        ("Target System Baseline", "Python 3.12, Streamlit Multipage, SQLite (data/app.db), Plotly, pytest")
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

    # Section 2: Executive Summary
    h1 = doc.add_heading(level=1)
    r = h1.add_run("2. Executive Summary")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "In modern insurance carriers, claims adjudication represents both the primary cost center and the defining customer experience touchpoint. Operational inefficiency, prolonged Turnaround Times (TAT), opaque repudiation grounds, and undetected fraudulent billings severely impair loss ratios, erode policyholder trust, and trigger regulatory sanctions from the Insurance Regulatory and Development Authority of India (IRDAI)."
    )
    p = doc.add_paragraph(
        "This Business Requirements Document (BRD) formalizes the operational and analytical specifications for the Insurance Claim Processing Analysis application. The application delivers an interactive, data-driven analytical workbench designed to empower Claims Operations Managers, Fraud Investigators (SIU), Actuaries, and Customer Grievance Officers. By uniting multi-line claims transactions (Health, Motor, Life, Property) with stage-wise bottleneck tracking, Pareto denial decomposition, automated anomaly scoring, and statistical hypothesis testing, the platform transitions claims management from reactive fire-fighting to proactive, data-driven decision excellence."
    )

    # Section 3: Business Problem & Opportunity
    h1 = doc.add_heading(level=1)
    r = h1.add_run("3. Business Problem & Opportunity")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "3.1 Business Problem Context:\n"
        "Claims processing workflows currently suffer from fragmented transactional logs and disconnected legacy tools. Key operational pain points include:\n"
        "• Unmonitored Stage Latency: Claims average 18+ days in adjudication because managers lack visibility into whether delays stem from initial document verification, medical review, or TPA communication.\n"
        "• Fraudulent Payout Leakage: Up to 10-15% of disbursed claims contain unflagged fraud indicators, including early-inception claims (loss filed < 30 days from policy purchase), duplicate submissions, and inflated provider billing.\n"
        "• Customer Grievances & Litigation: Unexplained claim rejections and unpredictable settlement payout haircuts create widespread policyholder disputes, ombudsman complaints, and customer attrition.\n"
        "• Inefficient Provider Networks: Healthcare hospitals and auto repair garages exhibit wide cost variance without empirical benchmarking, draining financial reserves."
    )
    p = doc.add_paragraph(
        "3.2 Strategic Business Opportunity:\n"
        "Deploying an integrated analytics workbench enables the insurer to:\n"
        "1. Compress Average Turnaround Time from 18 days to under 10 days by isolating stage-wise bottlenecks.\n"
        "2. Intercept >= 95% of high-risk fraudulent claims prior to fund disbursement, safeguarding capital reserves.\n"
        "3. Reduce non-technical claim rejections by 15% through transparent Pareto denial analysis.\n"
        "4. Validate the operational efficiency of Cashless networks with rigorous statistical significance (p < 0.05), guiding strategic network investments."
    )

    # Section 4: Objectives & Success Metrics
    h1 = doc.add_heading(level=1)
    r = h1.add_run("4. Objectives & Success Metrics (KPIs with Targets)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "The project establishes explicit, quantifiable Key Performance Indicators (KPIs) mapped to the core project objectives:"
    )

    kpi_table = doc.add_table(rows=6, cols=5)
    kpi_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(kpi_table, "CBD5E0")

    kpi_headers = ["Objective ID", "Strategic Objective", "Key Performance Indicator (KPI)", "Current Baseline", "Target Benchmark"]
    for j, h in enumerate(kpi_headers):
        c = kpi_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 80, 80)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    kpi_data = [
        ("OBJ-01", "Turnaround Time Optimization", "Mean End-to-End Claim Processing Days", "18.2 Days", "<= 10.0 Days"),
        ("OBJ-02", "Fraud & Leakage Interception", "High-Risk Anomaly Detection Rate (%)", "45% (Manual Audit)", ">= 95.0% Flagged"),
        ("OBJ-03", "Repudiation Root Cause Reduction", "Non-Technical Claim Repudiation Share", "16.5% of Claims", "<= 10.0% (15% Rel. Drop)"),
        ("OBJ-04", "Statistical Channel Validation", "Cashless vs. Reimbursement TAT p-value", "Unverified (Subjective)", "p < 0.05 (Stat. Significant)"),
        ("OBJ-05", "Software Reliability & QA", "Automated Pytest Pass Rate & Latency", "Manual Testing", ">= 90% Pass, < 3.0s Load")
    ]

    col_w_kpi = [Inches(1.0), Inches(1.5), Inches(1.7), Inches(1.1), Inches(1.2)]
    for i, row in enumerate(kpi_data):
        r = kpi_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w_kpi[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 70, 70, 80, 80)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(8.5)
            if j == 0:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
            elif j == 4:
                run.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 5: Stakeholders RACI Table
    h1 = doc.add_heading(level=1)
    r = h1.add_run("5. Stakeholder Governance (RACI Matrix)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "A formal RACI matrix (Responsible, Accountable, Consulted, Informed) defines decision rights and operational accountability across all project phases:"
    )

    raci_table = doc.add_table(rows=8, cols=6)
    raci_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(raci_table, "CBD5E0")

    raci_headers = ["Project Decision / Workstream", "SH-01 (Ops Head)", "SH-02 (SIU Lead)", "SH-03 (Underwriter)", "SH-04 (Cust Exp)", "SH-05 (CFO)"]
    for j, h in enumerate(raci_headers):
        c = raci_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 70, 70)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    raci_data = [
        ("Business Requirements (BRD) & Scope Sign-off", "Accountable (A)", "Consulted (C)", "Consulted (C)", "Consulted (C)", "Informed (I)"),
        ("Stage-wise Turnaround Time SLAs & Bottlenecks", "Responsible (R)", "Informed (I)", "Informed (I)", "Consulted (C)", "Informed (I)"),
        ("Fraud Detection Rules & Provider Blacklisting", "Consulted (C)", "Accountable (A)", "Consulted (C)", "Informed (I)", "Consulted (C)"),
        ("Underwriting Risk Tiers & Demographic Pricing", "Informed (I)", "Informed (I)", "Accountable (A)", "Informed (I)", "Consulted (C)"),
        ("Claim Repudiation Guidelines & Grievance Audits", "Responsible (R)", "Consulted (C)", "Informed (I)", "Accountable (A)", "Informed (I)"),
        ("Financial Solvency Reserves & Outlier Thresholds", "Informed (I)", "Informed (I)", "Consulted (C)", "Informed (I)", "Accountable (A)"),
        ("User Acceptance Testing (UAT) & Production Sign-off", "Accountable (A)", "Responsible (R)", "Responsible (R)", "Responsible (R)", "Informed (I)")
    ]

    raci_w = [Inches(2.5), Inches(0.8), Inches(0.8), Inches(0.8), Inches(0.8), Inches(0.8)]
    for i, row in enumerate(raci_data):
        r = raci_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = raci_w[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 70, 70, 70, 70)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(8)
            if j == 0:
                run.font.bold = True
            elif "Accountable (A)" in val:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0xC5, 0x30, 0x30)
            elif "Responsible (R)" in val:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 6: Scope
    h1 = doc.add_heading(level=1)
    r = h1.add_run("6. Project Scope (In-Scope vs. Out-of-Scope)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    scope_table = doc.add_table(rows=6, cols=2)
    scope_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(scope_table, "CBD5E0")
    scope_table.rows[0].cells[0].width = Inches(3.25)
    scope_table.rows[0].cells[1].width = Inches(3.25)
    set_cell_background(scope_table.rows[0].cells[0], "2B6CB0")
    set_cell_background(scope_table.rows[0].cells[1], "718096")
    set_cell_margins(scope_table.rows[0].cells[0], 80, 80, 100, 100)
    set_cell_margins(scope_table.rows[0].cells[1], 80, 80, 100, 100)

    p0 = scope_table.rows[0].cells[0].paragraphs[0]
    p0.paragraph_format.space_after = Pt(0)
    r0 = p0.add_run("IN-SCOPE (Committed Deliverables)")
    r0.font.bold = True
    r0.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    p1 = scope_table.rows[0].cells[1].paragraphs[0]
    p1.paragraph_format.space_after = Pt(0)
    r1 = p1.add_run("OUT-OF-SCOPE (Explicit Exclusions)")
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    scope_data = [
        ("Multi-line relational claims database in SQLite (data/app.db) across 4 core tables: Claims Fact, Policies, Policyholders, Providers.",
         "Direct integration with live banking payment gateways or real-time core insurer underwriting mainframes."),
        ("Synthetic data generator producing 200+ realistic claims with built-in seasonal spikes, provider anomalies, and fraud flags.",
         "Processing of real living individual personal identifiable information (PII) or confidential policy contracts."),
        ("Interactive Streamlit dashboard featuring global sidebar filters, Tailwind CSS metric cards, and responsive Plotly visual analytics.",
         "Autonomous machine-executable financial payout disbursement without human claims adjudicator approval."),
        ("Inferential statistical hypothesis testing (two-sample t-test) validating Cashless vs. Reimbursement processing TAT.",
         "Deep learning computer vision models for automated mobile image-based motor accident damage assessment."),
        ("Comprehensive automated test suite (pytest) covering database integrity, analytical formulas, and Streamlit AppTest page renders.",
         "Enterprise multi-tenant single sign-on (SSO/SAML) identity access management infrastructure.")
    ]

    for i, (in_s, out_s) in enumerate(scope_data):
        row = scope_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = Inches(3.25), Inches(3.25)
        set_cell_background(c0, bg)
        set_cell_background(c1, bg)
        set_cell_margins(c0, 70, 70, 80, 80)
        set_cell_margins(c1, 70, 70, 80, 80)
        p = c0.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(f"• {in_s}")
        r.font.size = Pt(8.5)
        p = c1.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(f"✕ {out_s}")
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(0x71, 0x80, 0x96)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 7: Business Requirements
    h1 = doc.add_heading(level=1)
    r = h1.add_run("7. Business Requirements Specification (BR-01 to BR-12)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "Each requirement is formally baselined with its business rationale, MoSCoW priority, originating HMW statement, and testable acceptance criteria (Given / When / Then):"
    )

    br_specs = [
        ("BR-01", "Turnaround Time (TAT) Tracking & Bottleneck Analysis",
         "The business needs to track and dissect end-to-end and stage-wise turnaround time across insurance product lines so that processing bottlenecks are isolated and claims settlement latency is systematically reduced.",
         "Claims processing latency is the number one cause of customer attrition and regulatory penalties. Isolating stage-specific duration allows management to deploy targeted operational remedies.",
         "Must", "HMW-01",
         "Given the operational claims dataset, When the claims manager filters by policy type, Then the system calculates mean and median TAT across each stage (Submission -> Verification -> Assessment -> Decision) and displays stage-wise cycle times in days."),

        ("BR-02", "Claim Repudiation Root Cause Decomposition",
         "The business needs to classify and decompose claim repudiation reasons using frequency and Pareto analysis across policy lines so that the primary root causes of claim rejections can be identified, reducing avoidable customer disputes.",
         "Unexplained or non-technical claim repudiations drive costly Ombudsman disputes. Pareto decomposition isolates the vital 20% of denial causes contributing to 80% of customer disputes.",
         "Must", "HMW-02",
         "Given a cohort of repudiated claims, When analyzed on the denial dashboard, Then the system displays a Pareto 80/20 chart ranking denial categories and shows the exact percentage contribution of each rejection ground."),

        ("BR-03", "Settlement Payout Haircut & Ratio Evaluation",
         "The business needs to evaluate claim settlement payout ratios across policy categories and severity tiers so that unjustified payout deductions are prevented and accurate financial claim reserves are maintained.",
         "Arbitrary or excessive deduction haircuts risk litigation and reputational harm, while under-deducting inflates claims expense variance.",
         "Should", "HMW-03",
         "Given approved claims, When the settlement ratio is evaluated, Then the system computes Settled Amount / Requested Amount and identifies records with deduction haircuts exceeding 20%."),

        ("BR-04", "Network Provider Benchmarking (Hospitals & Garages)",
         "The business needs to benchmark empaneled healthcare hospitals, automobile garages, and TPAs on average claim costs, repudiation frequency, and processing duration so that abusive overbilling and underperforming partners can be identified.",
         "Third-party network billing variance represents a massive leak of claims capital. Benchmarking enforces tariff discipline and guides policyholder triage.",
         "Must", "HMW-04",
         "Given network provider transactions, When evaluated by provider category and region, Then the system highlights providers whose average settled claim cost exceeds the regional peer mean by > 2 standard deviations."),

        ("BR-05", "Automated Fraud & Anomaly Scoring Engine",
         "The business needs to automatically identify and score claims exhibiting high-risk fraud attributes (claims within 30 days of inception, duplicate claims, suspicious losses) so that potential fraud leakage is intercepted prior to disbursement.",
         "Detecting fraud after funds are disbursed yields less than 5% recovery. Intercepting high-risk claims pre-settlement preserves corporate reserves.",
         "Must", "HMW-05",
         "Given incoming claims, When scored against fraud business rules, Then the system flags claims with early-inception (< 30 days), duplicate incident indicators, or fraud scores >= 70, displaying the total monetary risk exposure."),

        ("BR-06", "Monthly & Seasonal Surge Forecasting",
         "The business needs to analyze monthly and seasonal claim volume fluctuations across product lines so that operational management can accurately forecast seasonal surges and dynamically reallocate claims adjudication staffing.",
         "Insurance claims follow pronounced seasonal cycles (e.g., monsoon auto accidents, post-festival health claims). Failure to anticipate surges causes severe backlog spikes.",
         "Should", "HMW-06",
         "Given multi-year claims records, When aggregated by submission month and quarter, Then the system displays a 30-day moving average and month-over-month growth rate highlighting predictable surge periods."),

        ("BR-07", "Extreme Outlier Claims Distribution & Modeling",
         "The business needs to isolate and model the statistical distribution of claim settlement amounts, identifying extreme positive outliers so that actuarial loss reserving models and reinsurance catastrophe triggers are accurately calibrated.",
         "A handful of catastrophic claims can distort standard descriptive statistics and compromise solvency margins if not isolated.",
         "Should", "HMW-07",
         "Given settled claims data, When plotted via histogram, KDE, and box plots, Then the system flags claims exceeding Q3 + 1.5×IQR and above the 99th percentile for actuarial catastrophe review."),

        ("BR-08", "Cashless vs. Reimbursement Hypothesis Testing",
         "The business needs to statistically test and quantify the turnaround time differential between cashless network claims and reimbursement claims so that executive leadership can evaluate the return on investment of expanding digital hospital and garage networks.",
         "Executive decisions on digital provider network expansion require empirical mathematical proof beyond superficial average comparisons.",
         "Must", "HMW-08",
         "Given claims partitioned into Cashless and Reimbursement modes, When a two-sample independent t-test is executed, Then the system outputs the sample means, t-statistic, p-value, and effect size, verifying significance at alpha = 0.05."),

        ("BR-09", "Geographic Territory Grievance & Escalation Monitoring",
         "The business needs to monitor customer satisfaction ratings and claim dispute escalation rates across geographic territories so that regional customer relationship teams and operational auditors can intervene in high-friction zones.",
         "Customer dissatisfaction often concentrates in specific regional operating branches due to understaffed adjusters or uncooperative local provider networks.",
         "Should", "HMW-09",
         "Given claims mapped to geographic states/regions, When filtered for grievances, Then the system computes regional dispute escalation rates and highlights territories with average CSAT ratings below 3.5."),

        ("BR-10", "Policyholder Demographic & Risk Tier Correlation",
         "The business needs to correlate policyholder tenure, premium tier, and age demographic cohorts against claim frequency and loss severity so that underwriters can optimize risk pricing and policy renewal terms.",
         "Actuarial pricing requires empirical claims experience validation across demographic cohorts to prevent adverse selection and mispricing.",
         "Must", "HMW-10",
         "Given policyholder demographic and claims data, When multi-variable correlation analysis is performed, Then the system generates Pearson correlation coefficients and bubble charts relating age, tenure, premium, and claim amounts."),

        ("BR-11", "Data Quality, Validation & Cleaning Audit Trail",
         "The business needs to maintain a data quality and validation log tracking missing values, duplicate entries, and data anomalies across claims records so that analytical decisions are founded on verified, trustworthy operational data.",
         "Analytical dashboards built on unverified data generate flawed executive decisions. A transparent cleaning log establishes data provenance and regulatory credibility.",
         "Must", "HMW-01, HMW-05",
         "Given raw synthetic claims data, When the cleaning pipeline executes, Then the system displays a before/after audit log recording nulls handled, duplicates removed, and out-of-range values corrected."),

        ("BR-12", "Operational Data & Compliance Report Exporting",
         "The business needs to export filtered claims operational datasets, summary KPI tables, and compliance reports into standardized tabular formats (CSV/Excel) so that external regulatory auditors and executives can perform offline reviews.",
         "Enterprise compliance and external actuarial audits require downloadable tabular extracts matching active dashboard filter states.",
         "Could", "HMW-01, HMW-03",
         "Given any filtered view on the dashboard, When the user clicks the Export Data button, Then the system generates a downloadable CSV/Excel file containing the active records and summary metrics.")
    ]

    for br_id, title, desc, rationale, prio, hmw, ac in br_specs:
        h2 = doc.add_heading(level=2)
        r2 = h2.add_run(f"{br_id}: {title}")
        r2.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)

        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("Requirement Description: ")
        r.font.bold = True
        p.add_run(desc)

        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("Business Rationale: ")
        r.font.bold = True
        p.add_run(rationale)

        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(f"MoSCoW Priority: ")
        r.font.bold = True
        r_val = p.add_run(f"{prio} Have")
        if prio == "Must":
            r_val.font.bold = True
            r_val.font.color.rgb = RGBColor(0xC5, 0x30, 0x30)
        p.add_run("  |  ")
        r = p.add_run("Source HMW: ")
        r.font.bold = True
        p.add_run(hmw)

        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        r = p.add_run("Acceptance Criteria (Given/When/Then): ")
        r.font.bold = True
        r_ac = p.add_run(ac)
        r_ac.font.italic = True

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 8: Business Rules
    h1 = doc.add_heading(level=1)
    r = h1.add_run("8. Business Rules Governing Claims Processing")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "The application models and enforces 8 fundamental insurance business rules:"
    )

    rules = [
        ("BR-RULE-01 (Sequential Lifecycle Progression):", 
         "A claim must progress sequentially through defined operational milestones: Claim Inflow -> Document Verification -> Assessment -> Settlement/Repudiation. Timestamps must strictly satisfy: Submission Date <= Verification Date <= Assessment Date <= Decision Date."),
        ("BR-RULE-02 (Turnaround Time SLA Ceilings):", 
         "Standard claims should achieve adjudication within 10 business days. Cashless pre-authorizations must achieve turnaround within 24 hours. Any claim exceeding 15 days is automatically classified as Breach of Operational SLA."),
        ("BR-RULE-03 (Early-Inception Fraud Flag):", 
         "Any claim where the incident date occurs within 30 calendar days of policy inception date must be automatically flagged as High Risk (Rule Score +40 points) for mandatory SIU investigation."),
        ("BR-RULE-04 (Duplicate Submission Detection):", 
         "Any claim sharing an identical Policyholder ID, Provider ID, and Incident Date with an existing claim record must be flagged as a Potential Duplicate Anomaly."),
        ("BR-RULE-05 (Settlement Payout Boundary):", 
         "Settled claim amount must never be negative and must never exceed Claim Amount Requested or Policy Sum Insured (Coverage Limit). Mathematical invariant: 0 <= Amount Settled <= Amount Requested <= Sum Insured."),
        ("BR-RULE-06 (Mandatory Repudiation Ground):", 
         "Every claim marked as 'Repudiated' must be assigned a standardized IRDAI reason code (e.g., Pre-Existing Condition Non-Disclosure, Policy Exclusion, Fraudulent Documentation, Lack of Pre-Authorization)."),
        ("BR-RULE-07 (Catastrophic Outlier Threshold):", 
         "Any claim settlement amount exceeding the upper Tukey fence (Q3 + 1.5 * IQR) or the 99th percentile of its policy line is flagged as a Catastrophic Outlier for reinsurance solvency auditing."),
        ("BR-RULE-08 (Provider Cost Variance Warning):", 
         "Any empaneled provider whose average claim cost exceeds the regional peer average by more than 2.0 standard deviations is flagged for Tariff Inflation Review.")
    ]

    for r_title, r_text in rules:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(3)
        rt = p.add_run(f"{r_title} ")
        rt.font.bold = True
        rt.font.color.rgb = RGBColor(0x1A, 0x20, 0x2C)
        rd = p.add_run(r_text)
        rd.font.size = Pt(9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 9: Data Requirements
    h1 = doc.add_heading(level=1)
    r = h1.add_run("9. Data Requirements (Entities, Source, Cadence)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "To support all 12 business requirements, the application specifies four relational data entities:"
    )

    data_summary_table = doc.add_table(rows=5, cols=5)
    data_summary_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(data_summary_table, "CBD5E0")

    d_headers = ["Entity Name", "Entity Type", "Core Attributes", "Data Source", "Refresh Cadence"]
    for j, h in enumerate(d_headers):
        c = data_summary_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 80, 80)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    d_data = [
        ("claims_fact", "Fact Table (Main)", "claim_id, policy_id, provider_id, dates, amounts, status, tat_days, fraud_score", "Synthetic Generator / SQLite", "Batch / Session Load"),
        ("policies_master", "Dimension / Master", "policy_id, policyholder_id, policy_type, sum_insured, premium, start_date", "Synthetic Generator / SQLite", "Static Reference"),
        ("policyholders_master", "Dimension / Master", "policyholder_id, full_name, age, gender, city, state, tenure_months", "Synthetic Generator / SQLite", "Static Reference"),
        ("providers_master", "Dimension / Master", "provider_id, provider_name, provider_type, city, state, empanelment_status", "Synthetic Generator / SQLite", "Static Reference")
    ]

    col_w_d = [Inches(1.4), Inches(1.2), Inches(2.2), Inches(1.3), Inches(1.1)]
    for i, row in enumerate(d_data):
        r = data_summary_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w_d[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 70, 70, 80, 80)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(8.5)
            if j == 0:
                run.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 10: Assumptions, Constraints, Dependencies
    h1 = doc.add_heading(level=1)
    r = h1.add_run("10. Assumptions, Constraints & Dependencies")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "10.1 Business Assumptions:\n"
        "• Claims records encompass four major lines of business: Health, Motor, Life, and Property.\n"
        "• Historical data covers a representative operational span of 12-24 months to enable seasonal trend analysis.\n"
        "• Currency values are denominated strictly in Indian Rupees (INR ₹).\n\n"
        "10.2 Technical Constraints:\n"
        "• Hosting Environment: Streamlit Community Cloud enforces a 1.0 GB RAM memory ceiling, demanding strict in-memory efficiency and `@st.cache_data` caching.\n"
        "• Database Engine: Embedded SQLite (`data/app.db`) is read-optimized without multi-tenant distributed write concurrency.\n"
        "• Zero Live PII: In full adherence to academic ethics, zero real customer records are ingested.\n\n"
        "10.3 Project Dependencies:\n"
        "• Dependency on Step 4 Traceability Matrix baseline to guarantee 100% downstream feature coverage.\n"
        "• Dependency on Step 8/10 database schema creation and synthetic dataset generation prior to UI code authoring."
    )

    # Section 11: Risks and Mitigation
    h1 = doc.add_heading(level=1)
    r = h1.add_run("11. Risk Management Matrix")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "Operational and technical risks are evaluated with targeted mitigation strategies:"
    )

    risk_table = doc.add_table(rows=6, cols=5)
    risk_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(risk_table, "CBD5E0")

    r_headers = ["Risk ID", "Risk Description", "Likelihood", "Impact", "Mitigation Strategy"]
    for j, h in enumerate(r_headers):
        c = risk_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 80, 80)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    r_data = [
        ("R-01", "Synthetic data lacks realistic variance, leading to trivial statistical results.", "Low", "High", "Engineer realistic distributions using NumPy/Faker with explicit seasonal waves, deliberate outlier injections, and correlated fraud flags (Step 10)."),
        ("R-02", "Streamlit Cloud memory exhaustion on complex data aggregations.", "Medium", "High", "Implement SQLite SQL aggregations (`GROUP BY`) and use `@st.cache_data` to minimize in-memory Pandas dataframe duplication."),
        ("R-03", "Statistical sample size insufficient for hypothesis test power (Type II error).", "Low", "Medium", "Generate a main fact dataset of >= 200 records, ensuring >= 50 records in both Cashless and Reimbursement cohorts for robust statistical power."),
        ("R-04", "Scope creep causing delivery delays beyond college academic schedule.", "Medium", "Medium", "Enforce strict MoSCoW prioritization; protect the ~60% Must-Have baseline and defer Could-Have features if timeline tightens."),
        ("R-05", "Database file locking or schema mismatch during automated pytest execution.", "Low", "High", "Use SQLite WAL (Write-Ahead Logging) mode and isolate test fixtures with automated teardown in `tests/test_data.py`.")
    ]

    col_w_r = [Inches(0.8), Inches(2.2), Inches(0.9), Inches(0.8), Inches(2.5)]
    for i, row in enumerate(r_data):
        r = risk_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w_r[j]
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

    # Section 12: Glossary of Terms
    h1 = doc.add_heading(level=1)
    r = h1.add_run("12. Domain & Technical Glossary")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "Definitions of core business and technical terminology used throughout the requirements:"
    )

    glossary = [
        ("Turnaround Time (TAT)", "The elapsed duration in business or calendar days between claim submission and final operational adjudication."),
        ("Claim Repudiation", "The formal operational rejection or denial of an insurance claim based on non-compliance, policy exclusions, or fraud."),
        ("Cashless Settlement", "A streamlined claims mode where the insurer settles approved medical/repair costs directly with the empaneled provider."),
        ("Reimbursement Claim", "A traditional settlement mode where the insured policyholder pays out-of-pocket and subsequently claims compensation."),
        ("Third-Party Administrator (TPA)", "An outsourced organization that processes insurance claims, network hospitalization, and provider billing on behalf of the insurer."),
        ("Settlement Payout Haircut", "The percentage reduction between the total claimed amount and the final amount disbursed due to copays, deductibles, or non-payable items."),
        ("Early Inception Claim", "A claim filed within 30 days of policy purchase, historically associated with higher probability of pre-existing condition non-disclosure."),
        ("Special Investigation Unit (SIU)", "The specialized enterprise investigative division dedicated to detecting and countering insurance fraud and collusion."),
        ("IQR (Interquartile Range) Fence", "A statistical outlier threshold defined as Q3 + 1.5 * (Q3 - Q1), used to isolate extreme catastrophic loss amounts."),
        ("Incurred Claim Ratio (ICR)", "The ratio of total net claims paid by an insurer to the total net earned premium, reflecting underwriting profitability."),
        ("Requirements Traceability Matrix (RTM)", "A grid mapping each business requirement through architecture, database columns, user stories, UI screens, and test cases.")
    ]

    for term, definition in glossary:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(3)
        rt = p.add_run(f"{term}: ")
        rt.font.bold = True
        rt.font.color.rgb = RGBColor(0x1A, 0x20, 0x2C)
        rd = p.add_run(definition)
        rd.font.size = Pt(9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 13: Sign-off
    h1 = doc.add_heading(level=1)
    r = h1.add_run("13. Formal Review & Sign-off")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "By signing below, the undersigned stakeholders formally approve this Business Requirements Document as the authoritative baseline for Step 6 (Product Requirements Document) and Step 7 (Technical Requirements Document):"
    )

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
        ("Project Lead / Candidate", "Noman Zaid", "MCA 2nd Year Data Analytics", "APPROVED (Baselined)", "20-Nov-2026"),
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

    output_path = r"c:\Users\noman\OneDrive\Desktop\python\hcl-pro\docs\05_BRD.docx"
    doc.save(output_path)
    print(f"Successfully generated {output_path}")

if __name__ == "__main__":
    create_document()
