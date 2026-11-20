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
    run_title = p_title.add_run("Step 2: How Might We (HMW) Statements")
    run_title.font.name = 'Calibri'
    run_title.font.size = Pt(24)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(16)
    run_sub = p_sub.add_run("Design Thinking Facilitation: Translating Stakeholder Pain Points into Strategic Analytics Questions")
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
        ("Source Artifact", "docs/01_Project_Understanding.docx (Stakeholders SH-01 to SH-05, Questions OQ-01 to OQ-10)")
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

    # Section 1: Introduction to Design Thinking & HMW
    h1 = doc.add_heading(level=1)
    r = h1.add_run("1. Design Thinking Approach & Context")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "Design Thinking reframes real-world enterprise frustrations into actionable, human-centered opportunities. Rather than immediately proposing technical solutions (e.g., 'building a database query' or 'adding a chart'), the 'How Might We' (HMW) methodology dissects stakeholder friction points to identify core capabilities needed to empower users."
    )
    p = doc.add_paragraph(
        "By grounding each HMW statement in a validated Stakeholder Pain Point and aligning it to a primary Operational Question (OQ), this artifact bridges the Project Understanding (Step 1) to Business Needs Analysis (Step 3) without introducing premature solution bias."
    )

    # Section 2: Stakeholder Pain Points
    h1 = doc.add_heading(level=1)
    r = h1.add_run("2. Stakeholder Pain Point Analysis")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "Based on extensive workflow analysis of claims adjudication operations, each stakeholder group experiences distinct operational friction points (2-3 each):"
    )

    pp_table = doc.add_table(rows=12, cols=4)
    pp_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(pp_table, "CBD5E0")

    pp_headers = ["Pain Point ID", "Stakeholder", "Role", "Operational Pain Point Description"]
    for j, h in enumerate(pp_headers):
        c = pp_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 90, 90, 80, 80)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    pp_data = [
        ("PP-01", "SH-01", "Claims Operations Head", "Chronic claims settlement delays with zero visibility into which internal stage (verification vs assessment) causes the backlog."),
        ("PP-02", "SH-01", "Claims Operations Head", "Unpredictable monthly and seasonal claims spikes leaving processing teams overwhelmed and understaffed during peak periods."),
        ("PP-03", "SH-01", "Claims Operations Head", "Severe processing latency discrepancies between cashless network claims and physical reimbursement submissions."),
        ("PP-04", "SH-02", "SIU Fraud Detection Lead", "Fraudulent claims (early inception, duplicate claims, staged losses) slip through manual checks, causing massive payout leakage."),
        ("PP-05", "SH-02", "SIU Fraud Detection Lead", "Certain network hospitals and auto garages systematically inflate charges or bill unnecessary procedures without being flagged."),
        ("PP-06", "SH-03", "Chief Underwriting Officer", "Lack of empirical data linking policyholder tenure, premium volume, and age cohorts to claim frequency and loss severity."),
        ("PP-07", "SH-03", "Chief Underwriting Officer", "Inability to isolate extreme outlier claims that disproportionately skew product loss ratios and actuarial pricing."),
        ("PP-08", "SH-04", "Customer Experience Head", "High policyholder dissatisfaction and regulatory complaints due to opaque or arbitrary claim repudiations without clear reasons."),
        ("PP-09", "SH-04", "Customer Experience Head", "Regional customer grievance clusters and escalation hotspots go unnoticed until severe brand damage or ombudsman fines occur."),
        ("PP-10", "SH-05", "Chief Financial Officer (CFO)", "Sudden swings in settlement payout ratios and reserve burn rates threaten enterprise solvency and financial forecasting."),
        ("PP-11", "SH-05", "Chief Financial Officer (CFO)", "Excessive administrative overhead and capital leakage from inefficient reimbursement workflows compared to digital networks.")
    ]

    pp_w = [Inches(1.0), Inches(0.8), Inches(1.5), Inches(3.2)]
    for i, row in enumerate(pp_data):
        r = pp_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = pp_w[j]
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

    # Section 3: HMW Statements Table
    h1 = doc.add_heading(level=1)
    r = h1.add_run("3. Strategic How Might We (HMW) Statements")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "The 11 stakeholder pain points are synthesized into 10 precisely scoped 'How Might We' statements. Each statement adheres to design-thinking principles: problem-focused, solution-agnostic, and mapped to specific operational questions (OQ-01 to OQ-10):"
    )

    hmw_table = doc.add_table(rows=11, cols=6)
    hmw_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(hmw_table, "CBD5E0")

    hmw_headers = ["HMW ID", "How Might We Statement", "Stakeholder", "Pain Point", "Linked OQ", "Priority"]
    for j, h in enumerate(hmw_headers):
        c = hmw_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 90, 90, 70, 70)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    hmw_data = [
        ("HMW-01", "How might we help claims operations managers identify and isolate the exact adjudication milestones causing processing delays across different insurance lines?", "SH-01", "PP-01", "OQ-01", "PRIORITY (Top 5)"),
        ("HMW-02", "How might we enable customer experience and claims review teams to understand the root causes of claim repudiations so that avoidable rejections and customer disputes can be minimized?", "SH-04", "PP-08", "OQ-02", "PRIORITY (Top 5)"),
        ("HMW-03", "How might we provide underwriters and finance leads with visibility into claim settlement payout haircuts across policy types and severity bands to maintain equitable compensation?", "SH-05, SH-03", "PP-10", "OQ-03", "Standard"),
        ("HMW-04", "How might we assist fraud investigators in detecting suspicious billing anomalies, excessive costs, and prolonged cycle times among empaneled hospitals and repair garages?", "SH-02", "PP-05", "OQ-04", "PRIORITY (Top 5)"),
        ("HMW-05", "How might we empower the Special Investigation Unit (SIU) to proactively flag high-risk claims (such as early-inception policies and duplicate submissions) before funds are disbursed?", "SH-02", "PP-04", "OQ-05", "PRIORITY (Top 5)"),
        ("HMW-06", "How might we help claims operations leaders forecast seasonal and monthly claim surges so that staffing and processing capacity can be dynamically allocated?", "SH-01", "PP-02", "OQ-06", "Standard"),
        ("HMW-07", "How might we enable finance managers and actuaries to detect extreme outlier claims that disproportionately consume solvency reserves and necessitate reinsurance intervention?", "SH-05, SH-03", "PP-07, PP-10", "OQ-07", "Standard"),
        ("HMW-08", "How might we help operations executives statistically evaluate the turnaround time differential between cashless network claims and reimbursement claims to justify expanding digital integrations?", "SH-01, SH-05", "PP-03, PP-11", "OQ-08", "PRIORITY (Top 5)"),
        ("HMW-09", "How might we help quality and grievance officers identify regional operating territories experiencing high escalation rates and low customer satisfaction for targeted operational audits?", "SH-04", "PP-09", "OQ-09", "Standard"),
        ("HMW-10", "How might we assist underwriting and actuarial teams in assessing how policyholder tenure, premium tier, and age cohorts impact claim frequency and severity?", "SH-03", "PP-06", "OQ-10", "Standard")
    ]

    hmw_w = [Inches(0.7), Inches(2.3), Inches(0.8), Inches(0.8), Inches(0.8), Inches(1.1)]
    for i, row in enumerate(hmw_data):
        r = hmw_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = hmw_w[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 70, 70, 70, 70)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(8)
            if j == 0:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
            elif j == 5 and "PRIORITY" in val:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0xC5, 0x30, 0x30) # Crimson Red

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 4: Top 5 Priority HMW Statements
    h1 = doc.add_heading(level=1)
    r = h1.add_run("4. Priority Assessment (Top 5 Strategic Statements)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "To establish high-impact focus for product requirements and architectural prioritization, five statements are designated as top priority:"
    )

    priorities = [
        ("1. HMW-01 (Turnaround Time Optimization):", 
         "Rationale: Directly resolves core operational backlogs and customer wait friction, serving as the primary driver for regulatory compliance and operational cost reduction."),
        ("2. HMW-05 (Fraud & Risk Mitigation):", 
         "Rationale: Prevents irreversible capital leakage and unauthorized financial disbursements by intercepting suspicious claims before funds leave corporate reserves."),
        ("3. HMW-02 (Rejection & Dispute Analysis):", 
         "Rationale: Curbs unnecessary customer repudiations and protects enterprise reputation by diagnosing avoidable documentation gaps and policy misunderstandings."),
        ("4. HMW-04 (Network Provider Benchmarking):", 
         "Rationale: Controls institutional inflation and third-party fee gouging across empaneled hospitals and garages, directly safeguarding underwriting margins."),
        ("5. HMW-08 (Cashless vs. Reimbursement Statistical Validation):", 
         "Rationale: Delivers empirical statistical evidence (p-value validation) to justify executive investments in digital provider networks and cashless customer channels.")
    ]

    for title, reason in priorities:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(4)
        rt = p.add_run(f"{title} ")
        rt.font.bold = True
        rt.font.color.rgb = RGBColor(0x1A, 0x20, 0x2C)
        rr = p.add_run(reason)
        rr.font.size = Pt(10)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 5: Coverage Matrix Check
    h1 = doc.add_heading(level=1)
    r = h1.add_run("5. Operational Coverage Verification")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    callout = doc.add_table(rows=1, cols=1)
    callout.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = callout.rows[0].cells[0]
    c.width = Inches(6.5)
    set_cell_background(c, "F0FFF4") # Soft green
    set_cell_margins(c, 100, 100, 140, 140)
    set_table_borders(callout, "38A169")
    cp = c.paragraphs[0]
    cp.paragraph_format.space_after = Pt(0)
    cr_bold = cp.add_run("Coverage Verification Result: 100% Complete\n")
    cr_bold.font.bold = True
    cr_bold.font.size = Pt(10)
    cr_bold.font.color.rgb = RGBColor(0x22, 0x54, 0x3D)
    cr = cp.add_run(
        "✓ Every Operational Question from Step 1 (OQ-01 through OQ-10) is addressed by at least one dedicated How Might We statement.\n"
        "✓ Every key Stakeholder (SH-01 through SH-05) has their core operational pain points represented.\n"
        "✓ Zero orphan questions, zero orphan stakeholders. Ready for Needs Analysis (Step 3)."
    )
    cr.font.size = Pt(9.5)
    cr.font.color.rgb = RGBColor(0x27, 0x67, 0x49)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 6: Document Control
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

    rev_data = ["v1.0", "20-Nov-2026", "Noman Zaid (Roll: 2500320140103)", "Ms. Shilpa (Mentor)", "Approved Baseline for Step 3"]
    for j, val in enumerate(rev_data):
        c = rev_table.rows[1].cells[j]
        set_cell_background(c, "FFFFFF")
        set_cell_margins(c, 80, 80, 80, 80)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(val)
        run.font.size = Pt(8.5)

    output_path = r"c:\Users\noman\OneDrive\Desktop\python\hcl-pro\docs\02_HMW_Statements.docx"
    doc.save(output_path)
    print(f"Successfully generated {output_path}")

if __name__ == "__main__":
    create_document()
