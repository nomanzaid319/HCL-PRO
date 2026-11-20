import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
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

    # Set Margins
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Base Styles
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(0x2D, 0x37, 0x48) # Charcoal
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(6)

    # Document Header / Pre-Title
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_after = Pt(2)
    run_meta = p_meta.add_run("ACADEMIC PROJECT ARTIFACT · MCA DATA ANALYTICS · ABES ENGINEERING COLLEGE")
    run_meta.font.size = Pt(9)
    run_meta.font.bold = True
    run_meta.font.color.rgb = RGBColor(0x71, 0x80, 0x96)

    # Main Title
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(6)
    run_title = p_title.add_run("Step 1: Project Understanding Document")
    run_title.font.name = 'Calibri'
    run_title.font.size = Pt(24)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D) # Deep Navy

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(16)
    run_sub = p_sub.add_run("Comprehensive Problem Definition, Stakeholder Analysis, Operational Questions, and Project Boundary Specifications")
    run_sub.font.size = Pt(13)
    run_sub.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)

    # Metadata Box Table
    meta_table = doc.add_table(rows=5, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.autofit = False
    set_table_borders(meta_table, "CBD5E0")

    meta_data = [
        ("Project Name", "Insurance Claim Processing Analysis"),
        ("Learner / Candidate", "Noman Zaid (Roll No: 2500320140103, MCA 2nd Year Data Analytics)"),
        ("Institution & Department", "Department of Computer Applications, ABES Engineering College, Ghaziabad"),
        ("Faculty Mentor", "Ms. Shilpa"),
        ("Target Solution Architecture", "Python 3.12, Streamlit Multipage, Tailwind CSS UI, SQLite (app.db), Plotly, pytest")
    ]

    col_widths = [Inches(2.2), Inches(4.3)]
    for i, (k, v) in enumerate(meta_data):
        row = meta_table.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = col_widths[0], col_widths[1]
        set_cell_background(c0, "F7FAFC")
        set_cell_background(c1, "FFFFFF")
        set_cell_margins(c0, 80, 80, 120, 120)
        set_cell_margins(c1, 80, 80, 120, 120)
        
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(k)
        r0.font.bold = True
        r0.font.size = Pt(10)
        r0.font.color.rgb = RGBColor(0x1A, 0x20, 0x2C)

        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(v)
        r1.font.size = Pt(10)
        r1.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 1: Executive Summary
    h1 = doc.add_heading(level=1)
    r = h1.add_run("1. Executive Summary")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    
    p = doc.add_paragraph(
        "The insurance sector represents one of the most data-intensive domains within modern financial services, where operational profitability, regulatory compliance, and customer retention hinges upon the efficiency and integrity of claims adjudication. In an increasingly competitive landscape, insurers face a double-edged challenge: policyholders demand frictionless, rapid, and transparent settlements, while the enterprise must vigilantly mitigate fraudulent activities, leakage, and operational overheads."
    )
    p = doc.add_paragraph(
        "The Insurance Claim Processing Analysis application is an enterprise-grade analytical solution engineered to empower claims executives, fraud investigation units (SIU), underwriters, and customer grievance managers with granular, data-driven intelligence. By aggregating multi-line insurance records (Health, Motor, Life, and Property) into an interactive analytical workbench, this solution provides multi-dimensional diagnostic insights into Turnaround Times (TAT), denial rationales, network provider billing irregularities, payout ratios, and anomaly distributions. Built upon Python, Streamlit, Plotly, and SQLite, the platform bridges operational silos to convert raw transactional claims logs into high-impact strategic decisions."
    )

    # Section 2: Problem Statement
    h1 = doc.add_heading(level=1)
    r = h1.add_run("2. Problem Statement")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    # Callout Box for Problem Statement
    callout = doc.add_table(rows=1, cols=1)
    callout.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = callout.rows[0].cells[0]
    c.width = Inches(6.5)
    set_cell_background(c, "EBF8FF") # Light blue tint
    set_cell_margins(c, 120, 120, 160, 160)
    set_table_borders(callout, "3182CE")
    cp = c.paragraphs[0]
    cp.paragraph_format.space_after = Pt(0)
    cr_bold = cp.add_run("Core Problem Statement:\n")
    cr_bold.font.bold = True
    cr_bold.font.size = Pt(11)
    cr_bold.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
    cr = cp.add_run(
        "\"Insurance claims operations suffer from severe settlement latency, undetected fraudulent payout leakages, and elevated claim repudiation friction because operational managers and risk analysts lack an integrated, multi-dimensional analytical platform to pinpoint stage-wise processing bottlenecks, evaluate provider billing anomalies, and optimize turnaround times across policy portfolios.\""
    )
    cr.font.size = Pt(11)
    cr.font.italic = True
    cr.font.color.rgb = RGBColor(0x1A, 0x20, 0x2C)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 3: Intended Audience & User Personas
    h1 = doc.add_heading(level=1)
    r = h1.add_run("3. Intended Audience & User Personas")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "The application is engineered for distinct operational and executive personas within an insurance firm. The primary user groups, their operational rhythm, and their core interaction patterns are detailed below:"
    )

    aud_table = doc.add_table(rows=6, cols=4)
    aud_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(aud_table, "CBD5E0")
    
    aud_headers = ["User Persona", "Organizational Role", "Usage Frequency", "Primary Interaction with App"]
    for j, h in enumerate(aud_headers):
        cell = aud_table.rows[0].cells[j]
        set_cell_background(cell, "1B365D")
        set_cell_margins(cell, 100, 100, 100, 100)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    aud_data = [
        ("Claims Operations Manager", "Oversees claims adjudication teams across policy lines", "Daily (Morning review & real-time monitoring)", "Filters claims by status, monitors stage-wise TAT, identifies processing bottlenecks, and reallocates caseloads to clear backlog."),
        ("Fraud Investigation Officer (SIU)", "Investigates suspicious claims, billing fraud, and syndicate collusion", "Daily / Ad-hoc case triage", "Screens high-risk claims flagged by anomaly detection, inspects provider billing distributions, and audits early-inception claim spikes."),
        ("Underwriting & Actuarial Analyst", "Calculates loss ratios, refines policy terms, and prices portfolio risk", "Weekly / Monthly actuarial review", "Analyzes settlement ratios across policy types and age tiers, evaluates loss distributions, and recalibrates premium rating factors."),
        ("Customer Grievance & Quality Lead", "Monitors service level agreements (SLAs) and Ombudsman complaints", "Weekly / Bi-weekly SLA review", "Tracks claim rejection rates, reviews repudiation grounds, audits customer escalations by region, and ensures fair claims treatment."),
        ("Chief Operating Officer (COO)", "Drives enterprise efficiency, financial reserve stability, and loss management", "Monthly executive review", "Reviews executive KPI scorecards, total claims payout vs. reserve burn, overall settlement ratios, and macro portfolio health.")
    ]

    col_w = [Inches(1.5), Inches(1.5), Inches(1.2), Inches(2.3)]
    for i, row in enumerate(aud_data):
        r = aud_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 80, 80, 100, 100)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(9)
            if j == 0:
                run.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 4: Stakeholder Analysis
    h1 = doc.add_heading(level=1)
    r = h1.add_run("4. Stakeholder Analysis")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "A rigorous stakeholder mapping ensures the analytical outputs align directly with organizational decision rights and operational mandates. Each stakeholder is tracked using a persistent identifier (SH-xx):"
    )

    sh_table = doc.add_table(rows=6, cols=6)
    sh_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(sh_table, "CBD5E0")

    sh_headers = ["SH ID", "Stakeholder", "Role Description", "Core Interest", "Influence", "Decision Made via App"]
    for j, h in enumerate(sh_headers):
        cell = sh_table.rows[0].cells[j]
        set_cell_background(cell, "1B365D")
        set_cell_margins(cell, 100, 100, 80, 80)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    sh_data = [
        ("SH-01", "Claims Operations Head", "Oversees operational efficiency, adjudicators, and TPA SLAs", "Turnaround time reduction, backlog elimination, processing cost control", "High", "Reallocate claims processing staff, renegotiate TPA response SLAs, streamline verification stages."),
        ("SH-02", "Special Investigation Unit (SIU) Lead", "Directs fraud prevention, anti-fraud algorithms, and claims auditing", "Fraud detection, loss leakage reduction, provider collusion tracking", "High", "Flag claims for forensic field audits, initiate legal repudiation, delist abusive network hospitals/garages."),
        ("SH-03", "Chief Underwriting Officer", "Manages underwriting guidelines, risk assessment, and policy pricing", "Accurate risk pricing, loss ratio control, exposure management", "Medium", "Adjust underwriting deductibles, exclude high-risk coverage riders, revise demographic risk loadings."),
        ("SH-04", "Customer Experience Officer", "Ensures regulatory compliance, policyholder satisfaction, and grievance reduction", "Fair repudiation, transparent communication, reduced disputes", "Medium", "Refine claim denial documentation, mandate secondary reviews for disputed claims, train support reps."),
        ("SH-05", "Chief Financial Officer (CFO)", "Governs financial reserves, solvency margins, and payout disbursements", "Cash reserve solvency, loss ratio predictability, budget adherence", "High", "Set quarterly claims reserve provisions, approve reinsurance treaty settlements, audit operational expenditure.")
    ]

    sh_w = [Inches(0.6), Inches(1.2), Inches(1.3), Inches(1.2), Inches(0.7), Inches(1.5)]
    for i, row in enumerate(sh_data):
        r = sh_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = sh_w[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 80, 80, 80, 80)
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

    # Section 5: Operational Questions (OQs)
    h1 = doc.add_heading(level=1)
    r = h1.add_run("5. Operational Questions (OQ-01 to OQ-10)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "To ensure that the analytics suite directly drives business value, the system is designed around 10 mission-critical Operational Questions (OQs). Each question is mapped to underlying data attributes, analytical methodology, and business impact:"
    )

    oq_table = doc.add_table(rows=11, cols=5)
    oq_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(oq_table, "CBD5E0")

    oq_headers = ["OQ ID", "Operational Question", "Required Data Fields", "Analytical Method", "Business Decision / Impact"]
    for j, h in enumerate(oq_headers):
        cell = oq_table.rows[0].cells[j]
        set_cell_background(cell, "1B365D")
        set_cell_margins(cell, 100, 100, 80, 80)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    oq_data = [
        ("OQ-01", "What is the average Turnaround Time (TAT) across policy types (Health, Motor, Life, Property), and which adjudication stages cause the longest delays?", "submission_date, approval_date, settlement_date, stage, policy_type", "Time delta calculation, stage-wise cycle time analysis, grouped bar charts", "Eliminates processing bottlenecks, shortens customer waiting periods, and optimizes adjudicator workflows."),
        ("OQ-02", "What is the overall claim approval rate versus rejection rate, and what are the top 5 primary grounds for claim denial?", "claim_status, rejection_reason, policy_type, claim_id", "Frequency distribution, Pareto analysis (80/20 rule), percentage breakdown", "Identifies root causes of denial (e.g., policy exclusions, lack of pre-auth), refining policy wording and reducing disputes."),
        ("OQ-03", "How does the settlement payout ratio (Amount Settled vs. Amount Claimed) vary across policy types, claim severity, and demographic tiers?", "claim_amount_requested, claim_amount_settled, settlement_ratio, policy_type, severity_level", "Ratio computation, scatter plot with trend line, grouped box plots", "Monitors payout haircuts, detects excessive under-settlement risking litigation, and supports reserve forecasting."),
        ("OQ-04", "Which healthcare network hospitals, auto garages, or TPAs exhibit abnormal billing inflation, high rejection rates, or prolonged TAT?", "provider_id, provider_name, provider_type, claim_amount_settled, claim_status, tat_days", "Provider benchmarking, cost per claim outlier analysis, ranking matrix", "Delists non-compliant service providers, renegotiates tariff contracts, and directs claimants to efficient network partners."),
        ("OQ-05", "What proportion of claims exhibit high fraud risk indicators, and what is the total monetary risk exposure to the enterprise?", "fraud_risk_score, fraud_flag, inception_to_claim_days, duplicate_flag, claim_amount_requested", "Rule-based fraud scoring, financial exposure aggregation, risk tier distribution", "Triages claims for field investigations prior to disbursement, saving millions in fraudulent leakage."),
        ("OQ-06", "How do monthly and seasonal claim submission volumes fluctuate, and are there predictable surge periods stressing operational capacity?", "submission_date, claim_volume, policy_type, month, year", "Time-series trend analysis, 30-day moving average, month-over-month growth %", "Forecasts seasonal surges (e.g., monsoon motor claims, dengue/viral health spikes) and schedules adjusters."),
        ("OQ-07", "What is the parametric and non-parametric distribution of claim settlement amounts, and are extreme outliers distorting reserves?", "claim_amount_settled, policy_type, log_claim_amount", "Kernel Density Estimation (KDE), histogram, IQR & 99th percentile outlier fences", "Validates reinsurance catastrophe triggers and sets accurate prudential solvency reserve thresholds."),
        ("OQ-08", "Is there a statistically significant difference in processing turnaround time between Cashless and Reimbursement claims?", "settlement_mode, tat_days, policy_type", "Hypothesis testing (Independent two-sample t-test / Mann-Whitney U test, p-value)", "Quantifies efficiency advantage of digital cashless integration to drive customer adoption."),
        ("OQ-09", "Which geographic zones or regional operating hubs record the poorest customer satisfaction (CSAT) or highest escalation rates?", "region_state, grievance_flag, escalation_level, csat_score", "Cross-tabulation, geographic choropleth / heatmaps, escalation rate %", "Deploys targeted operational audits and regional customer relationship managers to high-friction territories."),
        ("OQ-10", "How does policyholder tenure, premium volume, and age cohort correlate with claim frequency and claim severity?", "tenure_months, annual_premium, policyholder_age, claim_count, claim_severity", "Pearson/Spearman correlation matrix, multi-variable bubble charts, age-tier aggregation", "Provides empirical loss experience data to actuarial teams for risk-adjusted pricing and retention strategies.")
    ]

    oq_w = [Inches(0.6), Inches(1.6), Inches(1.4), Inches(1.3), Inches(1.6)]
    for i, row in enumerate(oq_data):
        r = oq_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = oq_w[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 80, 80, 80, 80)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(8.5)
            if j == 0:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Feasibility Check
    p_chk = doc.add_paragraph()
    r_chk_title = p_chk.add_run("Data Feasibility Verification Check:\n")
    r_chk_title.font.bold = True
    r_chk_title.font.color.rgb = RGBColor(0x27, 0x67, 0x49) # Forest Green
    r_chk = p_chk.add_run(
        "✓ Every operational question (OQ-01 through OQ-10) is fully supported by transactional columns designed into the database schema (claims, policies, policyholders, providers). No synthetic external data assumptions or unobtainable variables are required."
    )
    r_chk.font.size = Pt(9.5)
    r_chk.font.italic = True

    # Section 6: SMART Objectives
    h1 = doc.add_heading(level=1)
    r = h1.add_run("6. Project Objectives (SMART)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "The project is governed by five Specific, Measurable, Achievable, Relevant, and Time-bound (SMART) objectives:"
    )

    smart_items = [
        ("OBJ-01 (Turnaround Time Optimization):", "Develop an interactive cycle-time diagnostic module that identifies stage-specific bottlenecks across 100% of claims records, targeting insights to reduce mean processing latency from 18 days to under 10 days."),
        ("OBJ-02 (Leakage & Fraud Mitigation):", "Implement a rule-based anomaly and fraud detection framework that flags at least 95% of synthetic suspicious claims (early inception, inflated charges, duplicate submissions), quantifying exposure risk."),
        ("OBJ-03 (Rejection & Dispute Reduction):", "Construct a comprehensive denial analysis engine utilizing Pareto decomposition to categorize top repudiation drivers across all 4 policy types, providing actionable recommendations to lower non-technical rejections by 15%."),
        ("OBJ-04 (Statistical Rigor & Validation):", "Execute formal hypothesis testing (t-test / Mann-Whitney U test) to evaluate cashless vs. reimbursement efficiency with p < 0.05 statistical significance, validated with 100% reproducible SciPy/statsmodels pipelines."),
        ("OBJ-05 (Production-Ready Architecture):", "Deliver a multi-page, responsive Streamlit dashboard with SQLite persistence, sub-3-second query response time, and >= 90% pytest automated test coverage deployed to Streamlit Community Cloud within the project schedule.")
    ]

    for title, desc in smart_items:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(4)
        rt = p.add_run(f"{title} ")
        rt.font.bold = True
        rt.font.color.rgb = RGBColor(0x1A, 0x20, 0x2C)
        rd = p.add_run(desc)
        rd.font.size = Pt(10)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 7: Scope Specification
    h1 = doc.add_heading(level=1)
    r = h1.add_run("7. Scope Specification (In-Scope vs. Out-of-Scope)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "Establishing definitive project boundaries ensures rigorous execution, prevents scope creep, and maintains alignment with college academic timeline constraints."
    )

    scope_table = doc.add_table(rows=6, cols=2)
    scope_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(scope_table, "CBD5E0")

    scope_table.rows[0].cells[0].width = Inches(3.25)
    scope_table.rows[0].cells[1].width = Inches(3.25)
    set_cell_background(scope_table.rows[0].cells[0], "2B6CB0")
    set_cell_background(scope_table.rows[0].cells[1], "718096")
    set_cell_margins(scope_table.rows[0].cells[0], 100, 100, 100, 100)
    set_cell_margins(scope_table.rows[0].cells[1], 100, 100, 100, 100)

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
        ("Multi-line synthetic dataset generation (Health, Motor, Life, Property) with 200+ realistic claims, relational keys, and built-in patterns.",
         "Integration with live third-party insurer APIs, proprietary core banking systems, or production TPA databases."),
        ("Relational database design (3NF schema) with masters, facts, constraints, and persistence in local SQLite (data/app.db).",
         "Real personal identifiable information (PII) of actual living citizens or confidential insurance contracts."),
        ("Multi-page Streamlit application with interactive Plotly charts, global filtering (dates, policy, status), and KPI metric cards.",
         "Deployment of autonomous automated financial payment disbursement or online banking clearing integrations."),
        ("Statistical hypothesis testing (t-test / Mann-Whitney U) evaluating processing time between settlement modes.",
         "Production-grade deep learning computer vision models for automated damaged vehicle image inspection."),
        ("Automated data validation, data cleaning before/after audit log, and automated pytest suite covering data integrity and UI components.",
         "Commercial enterprise user authentication with multi-factor biometric security (SAML/OAuth2 SSO).")
    ]

    for i, (in_s, out_s) in enumerate(scope_data):
        row = scope_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = Inches(3.25), Inches(3.25)
        set_cell_background(c0, bg)
        set_cell_background(c1, bg)
        set_cell_margins(c0, 80, 80, 100, 100)
        set_cell_margins(c1, 80, 80, 100, 100)
        
        p = c0.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(f"• {in_s}")
        r.font.size = Pt(9)
        
        p = c1.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(f"✕ {out_s}")
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0x71, 0x80, 0x96)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 8: Assumptions & Constraints
    h1 = doc.add_heading(level=1)
    r = h1.add_run("8. Assumptions & Constraints")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    h2 = doc.add_heading(level=2)
    r2 = h2.add_run("8.1 Project Assumptions")
    r2.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)

    assumptions = [
        ("Data Authenticity Assumption:", "The synthetic dataset generated via Faker and NumPy will emulate authentic Indian insurance operational characteristics (e.g., policy distributions, realistic metro vs. non-metro costs, believable denial ratios)."),
        ("Workflow Standardization:", "Claims workflow is assumed to transition sequentially through defined operational stages: Submission -> Document Verification -> Medical/Technical Assessment -> Settlement / Repudiation."),
        ("Currency & Units:", "All financial metrics are modeled in Indian Rupees (INR ₹) with standard rounding to 2 decimal places."),
        ("User Environment:", "End users operate modern Chromium-based web browsers (Chrome, Edge, Firefox) with JavaScript enabled at standard desktop 1080p display resolution.")
    ]
    for k, v in assumptions:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(3)
        rk = p.add_run(f"{k} ")
        rk.font.bold = True
        rv = p.add_run(v)
        rv.font.size = Pt(9.5)

    h2 = doc.add_heading(level=2)
    r2 = h2.add_run("8.2 Technical & Operational Constraints")
    r2.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)

    constraints = [
        ("Resource & Hosting Limits:", "Streamlit Community Cloud imposes resource ceilings (1 GB RAM, shared vCPU), requiring strict memory optimization, categorical downcasting, and cached queries (@st.cache_data)."),
        ("Single-File Database:", "SQLite is utilized for data storage without multi-user distributed concurrent write concurrency; data transactions during runtime are primarily read-optimized analytical queries."),
        ("No Live PII Mandate:", "Strict adherence to academic ethics dictates that zero actual consumer personal records or proprietary company data can be stored, processed, or uploaded."),
        ("No External Code Execution prior to Step 10:", "Per strict SDLC governance, design and analysis specifications (Steps 1-9) precede data generation and code implementation.")
    ]
    for k, v in constraints:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(3)
        rk = p.add_run(f"{k} ")
        rk.font.bold = True
        rv = p.add_run(v)
        rv.font.size = Pt(9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 9: Success Criteria (Measurable)
    h1 = doc.add_heading(level=1)
    r = h1.add_run("9. Measurable Success Criteria")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "The project's success is measured objectively against explicit technical, analytical, and functional benchmarks:"
    )

    sc_table = doc.add_table(rows=6, cols=5)
    sc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(sc_table, "CBD5E0")

    sc_headers = ["Metric ID", "Evaluation Dimension", "Target Benchmark / Threshold", "Measurement Method", "Verification Mechanism"]
    for j, h in enumerate(sc_headers):
        cell = sc_table.rows[0].cells[j]
        set_cell_background(cell, "1B365D")
        set_cell_margins(cell, 100, 100, 80, 80)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    sc_data = [
        ("SC-01", "Operational Question Coverage", "100% (10 of 10 OQs answered)", "Direct mapping of interactive charts/KPIs to OQ-01 through OQ-10", "Step 4 & Step 13 RTM Traceability Audit"),
        ("SC-02", "Automated Test Coverage", ">= 90% Pass Rate across Unit & Functional Tests", "Automated pytest suite execution (test_data.py, test_analysis.py, test_app.py)", "Terminal test run log artifact"),
        ("SC-03", "Application Latency", "< 3.0 seconds initial load; < 1.0 s on filter update", "Streamlit execution profiling and page render timer", "Benchmarking script and AppTest timer"),
        ("SC-04", "Data Quality & Integrity", "Zero unresolved foreign key anomalies; 100% PK uniqueness", "SQL consistency checks and data cleaning log comparison", "Step 10 Data Generation Validation Report"),
        ("SC-05", "Deployment & Uptime", "100% accessible public URL on Streamlit Community Cloud", "Public HTTP status 200 response and end-to-end cloud load", "Live URL verification in Step 16")
    ]

    sc_w = [Inches(0.7), Inches(1.5), Inches(1.5), Inches(1.4), Inches(1.4)]
    for i, row in enumerate(sc_data):
        r = sc_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = sc_w[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 80, 80, 80, 80)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(8.5)
            if j == 0:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 10: Sign-off & Revision History
    h1 = doc.add_heading(level=1)
    r = h1.add_run("10. Document Control & Sign-off")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    rev_table = doc.add_table(rows=2, cols=5)
    rev_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(rev_table, "CBD5E0")

    rev_headers = ["Version", "Release Date", "Author", "Reviewed By", "Status / Notes"]
    for j, h in enumerate(rev_headers):
        cell = rev_table.rows[0].cells[j]
        set_cell_background(cell, "1B365D")
        set_cell_margins(cell, 80, 80, 80, 80)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    rev_data = ["v1.0", "20-Nov-2026", "Noman Zaid (Roll: 2500320140103)", "Ms. Shilpa (Mentor)", "Approved Baseline for Step 2"]
    for j, val in enumerate(rev_data):
        c = rev_table.rows[1].cells[j]
        set_cell_background(c, "FFFFFF")
        set_cell_margins(c, 80, 80, 80, 80)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(val)
        run.font.size = Pt(8.5)

    output_path = r"c:\Users\noman\OneDrive\Desktop\python\hcl-pro\docs\01_Project_Understanding.docx"
    doc.save(output_path)
    print(f"Successfully generated {output_path}")

if __name__ == "__main__":
    create_document()
