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

    # Header / Meta
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
    run_title = p_title.add_run("Step 3: Needs Analysis Document")
    run_title.font.name = 'Calibri'
    run_title.font.size = Pt(24)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(16)
    run_sub = p_sub.add_run("Deriving Formal, Testable Business Requirements (BR-01 to BR-12) from Design-Thinking HMW Statements")
    run_sub.font.size = Pt(13)
    run_sub.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)

    # Meta Box
    meta_table = doc.add_table(rows=4, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.autofit = False
    set_table_borders(meta_table, "CBD5E0")

    meta_data = [
        ("Project Name", "Insurance Claim Processing Analysis"),
        ("Learner / Candidate", "Noman Zaid (Roll No: 2500320140103, MCA 2nd Year DA)"),
        ("Faculty Mentor", "Ms. Shilpa"),
        ("Source Artifacts", "docs/01_Project_Understanding.docx (OQ-01 to OQ-10), docs/02_HMW_Statements.docx (HMW-01 to HMW-10)")
    ]
    for i, (k, v) in enumerate(meta_data):
        row = meta_table.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = Inches(2.2), Inches(4.3)
        set_cell_background(c0, "F7FAFC")
        set_cell_background(c1, "FFFFFF")
        set_cell_margins(c0, 70, 70, 110, 110)
        set_cell_margins(c1, 70, 70, 110, 110)
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

    # Section 1: Methodology & Principles
    h1 = doc.add_heading(level=1)
    r = h1.add_run("1. Needs Analysis Methodology")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "Needs Analysis translates user-centered 'How Might We' opportunities into concrete Business Requirements (BRs). In strict accordance with professional Business Analysis Body of Knowledge (BABOK) standards, this document states WHAT capabilities the business requires to achieve strategic outcomes, strictly avoiding technology or implementation prescriptions (e.g., no references to Streamlit widgets or SQL queries at this stage)."
    )
    p = doc.add_paragraph(
        "Every requirement is articulated in the standardized canonical syntax:"
    )
    
    callout = doc.add_table(rows=1, cols=1)
    callout.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = callout.rows[0].cells[0]
    c.width = Inches(6.5)
    set_cell_background(c, "EBF8FF")
    set_cell_margins(c, 80, 80, 120, 120)
    set_table_borders(callout, "3182CE")
    cp = c.paragraphs[0]
    cp.paragraph_format.space_after = Pt(0)
    cr = cp.add_run("\"The business needs to <capability> so that <benefit>.\"")
    cr.font.bold = True
    cr.font.size = Pt(11)
    cr.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 2: Business Requirements Table
    h1 = doc.add_heading(level=1)
    r = h1.add_run("2. Business Requirements Specification (BR-01 to BR-12)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "The following 12 business requirements define the core operational scope. Requirements are prioritized using the MoSCoW framework (Must have, Should have, Could have, Won't have for now):"
    )

    br_table = doc.add_table(rows=13, cols=6)
    br_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(br_table, "CBD5E0")

    br_headers = ["BR ID", "Business Requirement (\"The business needs to...\")", "Stakeholder", "Source HMW", "Priority", "Measurement Method"]
    for j, h in enumerate(br_headers):
        c = br_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 90, 90, 70, 70)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    br_data = [
        ("BR-01", "The business needs to track and dissect end-to-end and stage-wise turnaround time (submission to verification, verification to assessment, assessment to final decision) across insurance product lines so that processing bottlenecks are isolated and claims settlement latency is systematically reduced.", "SH-01", "HMW-01", "Must", "Mean stage-wise duration (days) tracked; reduction in average overall TAT from 18 to <= 10 days."),
        ("BR-02", "The business needs to classify and decompose claim repudiation reasons using frequency and Pareto analysis across policy lines so that the primary root causes of claim rejections can be identified, reducing avoidable customer disputes and policy misunderstandings.", "SH-04, SH-01", "HMW-02", "Must", "Percentage of repudiations categorized by standard grounds; reduction in non-technical denial rate by 15%."),
        ("BR-03", "The business needs to evaluate claim settlement payout ratios (amount disbursed vs. amount requested) across policy categories and severity tiers so that unjustified payout deductions are prevented and accurate financial claim reserves are maintained.", "SH-05, SH-03", "HMW-03", "Should", "Median settlement ratio variance tracked within +/- 5% of historical benchmark; identification of severity-tier deductions exceeding 20%."),
        ("BR-04", "The business needs to benchmark empaneled healthcare hospitals, automobile garages, and third-party administrators (TPAs) on average claim costs, repudiation frequency, and processing duration so that abusive overbilling, tariff violations, and underperforming partners can be identified for contract renegotiation or de-empanelment.", "SH-02, SH-01", "HMW-04", "Must", "Provider risk score ranking; identification of providers with average claim amounts exceeding regional peers by > 2 standard deviations."),
        ("BR-05", "The business needs to automatically identify and score claims exhibiting high-risk fraud attributes (including claims filed within 30 days of policy inception, duplicate claims, and suspicious loss descriptions) so that potential fraud leakage is intercepted prior to financial disbursement.", "SH-02", "HMW-05", "Must", ">= 95% detection rate of synthetic high-risk fraud flags; quantified financial risk exposure sum tracked in real time."),
        ("BR-06", "The business needs to analyze monthly and seasonal claim volume fluctuations across product lines so that operational management can accurately forecast seasonal surges and dynamically reallocate claims adjudication staffing.", "SH-01", "HMW-06", "Should", "30-day moving average and month-over-month (MoM) volume growth % calculated; staffing surge preparedness within 14 days of historical peak."),
        ("BR-07", "The business needs to isolate and model the statistical distribution of claim settlement amounts, identifying extreme positive outliers so that actuarial loss reserving models and reinsurance catastrophe triggers are accurately calibrated.", "SH-05, SH-03", "HMW-07", "Should", "Identification and profiling of claims falling outside the 1.5x IQR upper fence and above the 99th percentile; variance from median reserve allocation."),
        ("BR-08", "The business needs to statistically test and quantify the turnaround time differential between cashless network claims and reimbursement claims so that executive leadership can evaluate the return on investment of expanding digital hospital and garage networks.", "SH-01, SH-05", "HMW-08", "Must", "Formal hypothesis test (p-value < 0.05 and effect size Cohen's d); delta in median settlement days between cashless and reimbursement."),
        ("BR-09", "The business needs to monitor customer satisfaction ratings and claim dispute escalation rates across geographic territories so that regional customer relationship teams and operational auditors can intervene in high-friction zones.", "SH-04", "HMW-09", "Should", "Escalation rate (%) per territory; count of regions with average CSAT score falling below 3.5 out of 5.0."),
        ("BR-10", "The business needs to correlate policyholder tenure, premium tier, and age demographic cohorts against claim frequency and loss severity so that underwriters can optimize risk pricing and policy renewal terms.", "SH-03", "HMW-10", "Must", "Multi-variate correlation coefficients (Pearson/Spearman r) between tenure, premium, age, and loss severity; loss ratio variance across demographic segments."),
        ("BR-11", "The business needs to maintain a data quality and validation log tracking missing values, duplicate entries, and data anomalies across claims records so that analytical decisions are founded on verified, trustworthy operational data.", "SH-01, SH-02, SH-05", "HMW-01, HMW-05", "Must", "100% of missing or duplicate synthetic records flagged and audited; before/after data cleaning audit trail logged."),
        ("BR-12", "The business needs to export filtered claims operational datasets, summary KPI tables, and compliance reports into standardized tabular formats (CSV/Excel) so that external regulatory auditors and executives can perform offline reviews.", "SH-01, SH-05", "HMW-01, HMW-03", "Could", "Single-click export capability producing standard compliant data formats matching active filter criteria.")
    ]

    br_w = [Inches(0.6), Inches(2.2), Inches(0.8), Inches(0.7), Inches(0.6), Inches(1.6)]
    for i, row in enumerate(br_data):
        r = br_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = br_w[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 70, 70, 70, 70)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(8)
            if j == 0:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
            elif j == 4:
                run.font.bold = True
                if val == "Must":
                    run.font.color.rgb = RGBColor(0xC5, 0x30, 0x30)
                elif val == "Should":
                    run.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
                else:
                    run.font.color.rgb = RGBColor(0x71, 0x80, 0x96)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 3: Priority Distribution Analysis
    h1 = doc.add_heading(level=1)
    r = h1.add_run("3. Requirement Priority Distribution (MoSCoW)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "To guarantee high academic delivery feasibility and enterprise-level focus, requirements are balanced per industry MoSCoW guidelines (~60% Must):"
    )

    prio_table = doc.add_table(rows=5, cols=4)
    prio_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(prio_table, "CBD5E0")

    prio_headers = ["MoSCoW Priority", "Count of Requirements", "Percentage Share", "Associated Requirement IDs"]
    for j, h in enumerate(prio_headers):
        c = prio_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 80, 80)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    prio_data = [
        ("Must Have (MVP Core)", "7", "58.3% (~60%)", "BR-01, BR-02, BR-04, BR-05, BR-08, BR-10, BR-11"),
        ("Should Have (High Impact)", "4", "33.3%", "BR-03, BR-06, BR-07, BR-09"),
        ("Could Have (Desirable)", "1", "8.3%", "BR-12"),
        ("Won't Have (Deferred/OOS)", "0", "0.0%", "Deferred to Phase 2 (Live Core APIs, Image AI)")
    ]

    prio_w = [Inches(1.8), Inches(1.2), Inches(1.2), Inches(2.3)]
    for i, row in enumerate(prio_data):
        r = prio_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = prio_w[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 70, 70, 80, 80)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(8.5)
            if j == 0:
                run.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 4: Data Needed per Requirement
    h1 = doc.add_heading(level=1)
    r = h1.add_run("4. Data Needs Specification (Field-Level Mapping)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "Each requirement demands specific data entities and transactional fields to become operational. The mapping below identifies these data needs from a purely conceptual perspective:"
    )

    data_table = doc.add_table(rows=13, cols=3)
    data_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(data_table, "CBD5E0")

    data_headers = ["BR ID", "Core Capability", "Required Data Attributes & Entities"]
    for j, h in enumerate(data_headers):
        c = data_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 80, 80)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    data_reqs = [
        ("BR-01", "Turnaround Time Tracking", "Claim ID, Policy Type, Submission Timestamp, Verification Timestamp, Assessment Timestamp, Decision Timestamp, Stage Durations (Days)."),
        ("BR-02", "Repudiation Root Cause Analysis", "Claim ID, Final Decision Status (Approved, Repudiated), Repudiation Reason Code, Repudiation Detailed Category, Policy Type."),
        ("BR-03", "Settlement Payout Haircut Analysis", "Claim ID, Claim Amount Claimed (Requested), Claim Amount Settled (Disbursed), Settlement Payout Ratio, Policy Type, Loss Severity Tier."),
        ("BR-04", "Provider Cost & TAT Benchmarking", "Provider ID, Provider Legal Name, Facility Type (Hospital, Garage, TPA), Empanelment Status, Claim Amounts Settled, Rejection Counts, Provider TAT (Days), City/Region."),
        ("BR-05", "Automated Fraud & Anomaly Scoring", "Claim ID, Policy Inception Date, Incident Date, Claim Submission Date, Inception-to-Loss Days Delta, Duplicate Claim Flag, Fraud Risk Score (0-100), Claim Amount Requested."),
        ("BR-06", "Seasonal Surge & Volume Forecasting", "Claim ID, Submission Date (Day, Month, Year, Quarter), Policy Product Line, Cumulative Inflow Count, Operational Adjuster Headcount."),
        ("BR-07", "Extreme Outlier Claims Modeling", "Claim ID, Settlement Amount, Policy Coverage Limit (Sum Insured), Policy Type, Statistical IQR Boundaries, Outlier Indicator Flag."),
        ("BR-08", "Cashless vs. Reimbursement Testing", "Claim ID, Settlement Processing Channel (Cashless, Reimbursement), Total Processing Days, Policy Product Type."),
        ("BR-09", "Geographic Grievance Monitoring", "Claim ID, Geographic Region / State / Territory, Customer Satisfaction Rating (1-5), Formal Grievance Indicator, Regulatory Ombudsman Escalation Level."),
        ("BR-10", "Policyholder Demographic Correlation", "Policyholder ID, Customer Age, Account Tenure (Months), Policy Annual Premium Amount, Lifetime Incurred Claims Count, Cumulative Payout Sum, Policy Type."),
        ("BR-11", "Data Quality & Integrity Auditing", "Dataset Record ID, Incomplete Record Flags (Null Counts), Duplicate Identifier Keys, Value Range Violation Codes, Cleaned Data Audit Log."),
        ("BR-12", "Data & Compliance Report Exporting", "Joined Transactional View Attributes (30-40 columns), Current Filter Criteria State, Formatted Summary Table Aggregations.")
    ]

    data_w = [Inches(0.8), Inches(1.8), Inches(3.9)]
    for i, row in enumerate(data_reqs):
        r = data_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = data_w[j]
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

    # Section 5: Gap Analysis
    h1 = doc.add_heading(level=1)
    r = h1.add_run("5. Bidirectional Traceability & Gap Analysis")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "A formal gap analysis verifies that no strategic opportunity is neglected and no requirement is fabricated without stakeholder justification:"
    )

    gap_table = doc.add_table(rows=3, cols=3)
    gap_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(gap_table, "CBD5E0")

    gap_headers = ["Audit Dimension", "Identified Gaps / Discrepancies", "Status & Resolution"]
    for j, h in enumerate(gap_headers):
        c = gap_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 80, 80)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    gap_data = [
        ("HMW Statements with No Linked BR", "0 (Zero)", "PASSED: All 10 HMW statements (HMW-01 to HMW-10) directly spawn one or more formal business requirements."),
        ("Business Requirements with No Source HMW", "0 (Zero)", "PASSED: All 12 business requirements (BR-01 to BR-12) explicitly trace back to an authorized design-thinking source.")
    ]

    gap_w = [Inches(2.5), Inches(1.5), Inches(2.5)]
    for i, row in enumerate(gap_data):
        r = gap_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = gap_w[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 70, 70, 80, 80)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(8.5)
            if j == 0:
                run.font.bold = True
            elif j == 2:
                run.font.color.rgb = RGBColor(0x27, 0x67, 0x49)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 6: Document Control & Sign-off
    h1 = doc.add_heading(level=1)
    r = h1.add_run("6. Document Control & Sign-off")
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

    rev_data = ["v1.0", "20-Nov-2026", "Noman Zaid (Roll: 2500320140103)", "Ms. Shilpa (Mentor)", "Approved Baseline for Step 4 (RTM)"]
    for j, val in enumerate(rev_data):
        c = rev_table.rows[1].cells[j]
        set_cell_background(c, "FFFFFF")
        set_cell_margins(c, 80, 80, 80, 80)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(val)
        run.font.size = Pt(8.5)

    output_path = r"c:\Users\noman\OneDrive\Desktop\python\hcl-pro\docs\03_Needs_Analysis.docx"
    doc.save(output_path)
    print(f"Successfully generated {output_path}")

if __name__ == "__main__":
    create_document()
