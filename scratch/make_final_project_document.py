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

def set_table_borders(table, color="CBD5E0"):
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

def add_header_banner(doc, title_text, subtitle_text):
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_before = Pt(0)
    p_meta.paragraph_format.space_after = Pt(2)
    run_meta = p_meta.add_run("ABES ENGINEERING COLLEGE, GHAZIABAD · DEPARTMENT OF COMPUTER APPLICATIONS")
    run_meta.font.size = Pt(8.5)
    run_meta.font.bold = True
    run_meta.font.color.rgb = RGBColor(0x71, 0x80, 0x96)

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(2)
    p_title.paragraph_format.space_after = Pt(4)
    r_title = p_title.add_run(title_text)
    r_title.font.size = Pt(20)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run(subtitle_text)
    r_sub.font.size = Pt(11)
    r_sub.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)

def add_section_heading(doc, text, level=1):
    h = doc.add_heading(level=level)
    h.paragraph_format.space_before = Pt(14)
    h.paragraph_format.space_after = Pt(6)
    r = h.add_run(text)
    if level == 1:
        r.font.size = Pt(15)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    elif level == 2:
        r.font.size = Pt(12.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
    else:
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)
    return h

def build_project_document(output_docx_path):
    doc = Document()

    # Margin Setup (Standard 1.0 in)
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Base Typography
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(6)

    # ==========================================
    # 1. FORMAL COVER / TITLE PAGE
    # ==========================================
    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst.paragraph_format.space_before = Pt(12)
    p_inst.paragraph_format.space_after = Pt(4)
    r_inst = p_inst.add_run("ABES ENGINEERING COLLEGE, GHAZIABAD\nDEPARTMENT OF COMPUTER APPLICATIONS")
    r_inst.font.size = Pt(13)
    r_inst.font.bold = True
    r_inst.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p_aff = doc.add_paragraph()
    p_aff.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_aff.paragraph_format.space_after = Pt(24)
    r_aff = p_aff.add_run("(Affiliated to Dr. A.P.J. Abdul Kalam Technical University, Lucknow, Uttar Pradesh)")
    r_aff.font.size = Pt(9.5)
    r_aff.font.color.rgb = RGBColor(0x71, 0x80, 0x96)

    p_proj_title = doc.add_paragraph()
    p_proj_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_proj_title.paragraph_format.space_before = Pt(18)
    p_proj_title.paragraph_format.space_after = Pt(8)
    r_proj = p_proj_title.add_run("INSURANCE CLAIM PROCESSING ANALYSIS")
    r_proj.font.size = Pt(24)
    r_proj.font.bold = True
    r_proj.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(28)
    r_sub = p_sub.add_run("An Enterprise Multi-Page Operational Dashboard, Statistical Hypothesis Testing & Fraud Risk Scoring System")
    r_sub.font.size = Pt(13)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)

    p_req = doc.add_paragraph()
    p_req.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_req.paragraph_format.space_after = Pt(36)
    r_req = p_req.add_run("A Major Academic Project Report Submitted in Partial Fulfillment of the Requirements\nfor the Award of the Degree of\nMASTER OF COMPUTER APPLICATIONS (DATA ANALYTICS)")
    r_req.font.size = Pt(10.5)
    r_req.font.bold = True
    r_req.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)

    # Candidate & Mentor Table
    cand_table = doc.add_table(rows=1, cols=2)
    cand_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for cell in cand_table.rows[0].cells:
        cell.width = Inches(3.2)
        set_cell_background(cell, "FFFFFF")
    
    c_left, c_right = cand_table.rows[0].cells[0], cand_table.rows[0].cells[1]
    
    p_l = c_left.paragraphs[0]
    p_l.add_run("Submitted By:\n").font.bold = True
    r_nom = p_l.add_run("NOMAN ZAID\n")
    r_nom.font.bold = True
    r_nom.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    p_l.add_run("University Roll No: 2500320140103\nMCA 2nd Year (Data Analytics)\nDepartment of Computer Applications")
    p_l.runs[0].font.size = Pt(10)
    for run in p_l.runs[1:]:
        run.font.size = Pt(10)

    p_r = c_right.paragraphs[0]
    p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_r.add_run("Under the Guidance of:\n").font.bold = True
    r_ment = p_r.add_run("MS. SHILPA\n")
    r_ment.font.bold = True
    r_ment.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
    p_r.add_run("Faculty Mentor & Assistant Professor\nDepartment of Computer Applications\nABES Engineering College, Ghaziabad")
    p_r.runs[0].font.size = Pt(10)
    for run in p_r.runs[1:]:
        run.font.size = Pt(10)

    p_session = doc.add_paragraph()
    p_session.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_session.paragraph_format.space_before = Pt(40)
    r_sess = p_session.add_run("ACADEMIC SESSION: 2025 – 2026")
    r_sess.font.size = Pt(11)
    r_sess.font.bold = True
    r_sess.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    doc.add_page_break()

    # ==========================================
    # 2. CERTIFICATE OF APPROVAL
    # ==========================================
    add_header_banner(doc, "CERTIFICATE OF APPROVAL", "Department of Computer Applications · Faculty Mentorship Certification")

    p_cert = doc.add_paragraph(
        "This is to certify that the project entitled \"INSURANCE CLAIM PROCESSING ANALYSIS\" submitted by "
        "NOMAN ZAID (University Roll No: 2500320140103), a student of Master of Computer Applications (MCA 2nd Year, "
        "Data Analytics Specialization) at ABES Engineering College, Ghaziabad, represents bonafide work carried out "
        "under my guidance and supervision in partial fulfillment of the requirements for the award of the degree of "
        "Master of Computer Applications affiliated to Dr. A.P.J. Abdul Kalam Technical University, Lucknow.\n\n"
        "To the best of my knowledge, the matter embodied in this major project report has not been submitted elsewhere "
        "for the award of any other degree or diploma."
    )
    p_cert.paragraph_format.line_spacing = 1.3
    p_cert.paragraph_format.space_after = Pt(36)

    sign_table = doc.add_table(rows=1, cols=2)
    sign_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for cell in sign_table.rows[0].cells:
        cell.width = Inches(3.2)
        set_cell_background(cell, "FFFFFF")

    sl, sr = sign_table.rows[0].cells[0], sign_table.rows[0].cells[1]
    p_sl = sl.paragraphs[0]
    p_sl.add_run("_________________________\n").font.color.rgb = RGBColor(0x71, 0x80, 0x96)
    p_sl.add_run("Ms. Shilpa\n").font.bold = True
    p_sl.add_run("Faculty Mentor & Assistant Professor\nDepartment of Computer Applications\nABES Engineering College, Ghaziabad")
    p_sl.runs[1].font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p_sr = sr.paragraphs[0]
    p_sr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sr.add_run("_________________________\n").font.color.rgb = RGBColor(0x71, 0x80, 0x96)
    p_sr.add_run("Head of Department\n").font.bold = True
    p_sr.add_run("Department of Computer Applications\nABES Engineering College, Ghaziabad")
    p_sr.runs[1].font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    doc.add_page_break()

    # ==========================================
    # 3. CANDIDATE DECLARATION & ACKNOWLEDGEMENT
    # ==========================================
    add_header_banner(doc, "CANDIDATE DECLARATION", "Statement of Academic Integrity & Original Authorship")

    p_decl = doc.add_paragraph(
        "I, NOMAN ZAID (Roll No: 2500320140103), student of MCA 2nd Year (Data Analytics), Department of Computer "
        "Applications, ABES Engineering College, Ghaziabad, hereby declare that the major project work entitled "
        "\"INSURANCE CLAIM PROCESSING ANALYSIS\" is an authentic record of my own individual work carried out under "
        "the mentorship of Ms. Shilpa.\n\n"
        "I further declare that all synthetic data, analytical scripts, mathematical modeling, and documentation "
        "comply strictly with academic ethics, without plagiarized code or unauthorized third-party intellectual property."
    )
    p_decl.paragraph_format.line_spacing = 1.3
    p_decl.paragraph_format.space_after = Pt(24)

    p_sig = doc.add_paragraph()
    p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sig.add_run("NOMAN ZAID\nRoll No: 2500320140103\nMCA 2nd Year Data Analytics\nABES Engineering College, Ghaziabad").font.bold = True

    add_section_heading(doc, "ACKNOWLEDGEMENTS", level=2)
    p_ack = doc.add_paragraph(
        "I express my deepest gratitude to my faculty mentor, Ms. Shilpa, for her invaluable guidance, rigorous technical "
        "scrutiny, and persistent encouragement throughout the 18 steps of this project lifecycle.\n\n"
        "I am profoundly grateful to the Head of Department and the faculty members of the Department of Computer "
        "Applications, ABES Engineering College, for equipping me with advanced data analytics foundations, software "
        "engineering methodologies, and state-of-the-art computational infrastructure."
    )
    p_ack.paragraph_format.line_spacing = 1.25

    doc.add_page_break()

    # ==========================================
    # 4. EXECUTIVE SUMMARY & ABSTRACT
    # ==========================================
    add_header_banner(doc, "EXECUTIVE SUMMARY", "Abstract, Core Findings, and Strategic Business Value")

    p_sum = doc.add_paragraph(
        "The modern insurance enterprise operates in an increasingly complex environment characterized by stringent "
        "regulatory compliance (IRDAI), escalating consumer expectations for instant claim adjudication, and the ever-present "
        "menace of fraudulent leakage. This project delivers an enterprise-grade Data Analytics & Decision Support System "
        "designed to optimize the end-to-end insurance claim processing lifecycle across Health, Motor, Life, and Property portfolios.\n\n"
        "Built upon a Decoupled Model-View-Controller (MVC) and Pipeline architecture, the solution leverages Python 3.12, "
        "Streamlit Multipage Router, Tailwind CSS UI components, Plotly client-side graphics, and an embedded SQLite relational "
        "database. The analytical engine evaluates a realistic synthetic cohort of 250 claims across 4 relational tables, "
        "delivering key empirical findings:\n"
        "• Turnaround Time Optimization: Portfolio mean cycle time is 12.2 days with a median of 9.0 days. Welch's two-sample "
        "independent t-test formally rejects H0 with p = 9.02e-27 (p < 0.001), establishing that Cashless processing (Mean: 8.5d) "
        "is statistically significantly faster than Reimbursement (Mean: 17.5d).\n"
        "• Pareto Repudiation Root Causes: Across 70 denied claims, the top two rejection categories (POLICY_EXCLUSION: 30 claims, "
        "EXCL_PRE_EXIST: 27 claims) constitute 81.4% of all repudiations, strictly confirming the Pareto 80/20 rule and pinpointing "
        "pre-underwriting disclosure gaps.\n"
        "• Tariff Anomaly Detection: Apex Multispeciality Hospital (PRV-104) is isolated with a billing deviation z-score > 2.0, "
        "triggering an automated audit warning badge.\n"
        "• Quality Assurance: All 14 technical unit tests in the Pytest suite (TC-T-01..14) pass with a 100% pass rate in 4.43 seconds, "
        "maintaining 100% bidirectional traceability to business requirements."
    )
    p_sum.paragraph_format.line_spacing = 1.25

    # ==========================================
    # CHAPTER 1: INTRODUCTION & PROBLEM UNDERSTANDING
    # ==========================================
    add_section_heading(doc, "CHAPTER 1: INTRODUCTION & PROBLEM UNDERSTANDING", level=1)
    
    p_ch1 = doc.add_paragraph(
        "1.1 Industry Background & Domain Context\n"
        "In the Indian insurance landscape governed by the Insurance Regulatory and Development Authority of India (IRDAI), "
        "claim settlement efficiency serves as the ultimate litmus test of insurer solvency, operational integrity, and brand loyalty. "
        "Insurers process millions of claims annually across multiple lines of business—ranging from cashless hospitalizations to "
        "motor accidental repairs and life death benefits. Despite massive digitisation, operational friction persists in the form of "
        "unpredictable cycle times, high repudiation grievance rates, and fraudulent billing leakage.\n\n"
        "1.2 Problem Statement\n"
        "The underwriting and operations teams face critical operational bottlenecks: (a) absence of granular milestone cycle time "
        "visibility, (b) inability to identify root causes behind high denial rates, (c) tariff billing variances among empaneled "
        "hospitals, and (d) delayed detection of moral hazard in early-inception policies filed within 30 days of coverage start.\n\n"
        "1.3 Key Stakeholders & Operational Questions\n"
        "The system addresses five primary enterprise personas:\n"
        "• SH-01 Chief Claims Officer (CCO): Requires portfolio-wide operational visibility, settlement haircuts, and SLA tracking.\n"
        "• SH-02 Head of Underwriting & Actuarial: Needs statistical validation of channel latency and catastrophic loss fences.\n"
        "• SH-03 Special Investigation Unit (SIU) Lead: Demands rule-based anomaly scoring to prioritize forensic field audits.\n"
        "• SH-04 Network Provider Manager: Monitors tariff compliance and empanelment performance across hospitals and garages.\n"
        "• SH-05 Grievance & Compliance Officer: Tracks IRDAI regulatory compliance and customer satisfaction (CSAT)."
    )

    # ==========================================
    # CHAPTER 2: REQUIREMENTS SPECIFICATION & TRACEABILITY
    # ==========================================
    add_section_heading(doc, "CHAPTER 2: REQUIREMENTS ANALYSIS & TRACEABILITY MATRIX", level=1)
    
    p_ch2 = doc.add_paragraph(
        "2.1 Business & Functional Requirements Hierarchy\n"
        "Requirements were synthesized through extensive domain research and formalized across 12 Business Requirements (BR-01..12), "
        "14 Functional Requirements (FR-01..14), and 14 User Stories (US-01..14) with Gherkin acceptance criteria.\n\n"
        "2.2 Non-Functional Requirements (NFR)\n"
        "• NFR-01 Performance: Interactive filter updates and Plotly chart reflows must complete within 1.0 second.\n"
        "• NFR-02 Scalability & Headroom: System must execute comfortably within Streamlit Cloud's 1.0 GB RAM container (< 250 MB).\n"
        "• NFR-03 Data Integrity: 100% primary key uniqueness, foreign key referential integrity, and ISO 8000 compliance.\n"
        "• NFR-04 Security & Privacy: Zero real customer PII; 100% believable synthetic Indian data context."
    )

    # Master RTM Table
    rtm_table = doc.add_table(rows=13, cols=6)
    rtm_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(rtm_table, "CBD5E0")
    
    r_hdrs = ["BR ID", "Functional Req", "User Story", "Target Screen", "Pytest Suite", "Functional Test"]
    for j, h in enumerate(r_hdrs):
        c = rtm_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 70, 70, 50, 50)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h)
        r.font.bold = True
        r.font.size = Pt(8)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    rtm_rows = [
        ("BR-01", "FR-01, FR-02", "US-01, US-14", "UI-01, UI-02", "TC-T-07, TC-T-12", "TC-F-01, TC-F-02, TC-F-03"),
        ("BR-02", "FR-03, FR-04", "US-03", "UI-03", "TC-T-08", "TC-F-04"),
        ("BR-03", "FR-05", "US-04", "UI-01", "TC-T-09", "TC-F-05"),
        ("BR-04", "FR-06", "US-05", "UI-04", "TC-T-02, TC-T-04", "TC-F-06"),
        ("BR-05", "FR-07", "US-06", "UI-05", "TC-T-11", "TC-F-07"),
        ("BR-06", "FR-08", "US-07", "UI-02", "TC-T-07", "TC-F-08"),
        ("BR-07", "FR-09", "US-08", "UI-06", "TC-T-02", "TC-F-09"),
        ("BR-08", "FR-10", "US-09", "UI-06", "TC-T-10", "TC-F-10"),
        ("BR-09", "FR-11", "US-10", "UI-03", "TC-T-02", "TC-F-11"),
        ("BR-10", "FR-12", "US-11", "UI-06", "TC-T-02", "TC-F-12"),
        ("BR-11", "FR-13", "US-12", "UI-07", "TC-T-01..06", "TC-F-13"),
        ("BR-12", "FR-14", "US-13", "UI-07", "TC-T-14", "TC-F-14")
    ]
    col_w_r = [Inches(0.8), Inches(1.1), Inches(1.0), Inches(1.1), Inches(1.2), Inches(1.1)]
    for i, row in enumerate(rtm_rows):
        r = rtm_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w_r[j]
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

    doc.add_page_break()

    # ==========================================
    # CHAPTER 3: DATABASE ARCHITECTURE & SCHEMA DESIGN
    # ==========================================
    add_section_heading(doc, "CHAPTER 3: DATABASE ARCHITECTURE & SCHEMA DESIGN", level=1)
    
    p_ch3 = doc.add_paragraph(
        "3.1 Relational Schema Architecture (Third Normal Form - 3NF)\n"
        "To guarantee high performance and strict consistency, data is structured into four normalized tables within SQLite (data/app.db):\n"
        "1. claims_fact (Fact Table, 250 rows): Stores individual claim transaction events, milestone dates, financial requested/settled amounts, adjudication stages, and fraud scores.\n"
        "2. policies_master (Dimension Table, 80 rows): Houses policy coverage tiers, sum insured limits, annual premiums, and policy start dates.\n"
        "3. policyholders_master (Dimension Table, 50 rows): Captures demographic attributes including policyholder name, age, gender, city, state, region, and tenure.\n"
        "4. providers_master (Dimension Table, 15 rows): Stores empaneled hospital and garage network facilities, accreditation levels, and tariff tiers.\n\n"
        "3.2 Synthetic Data Ingestion Pipeline\n"
        "The dataset was programmatically generated using Faker and NumPy, calibrating distributions to mirror real Indian health and motor claims (e.g., higher dengue hospitalization surges during Q3 monsoon months, 25-30% baseline repudiation rates, and believable Indian healthcare provider tariffs)."
    )

    # ==========================================
    # CHAPTER 4: STATISTICAL & MATHEMATICAL MODELING
    # ==========================================
    add_section_heading(doc, "CHAPTER 4: STATISTICAL & MATHEMATICAL MODELING", level=1)
    
    p_ch4 = doc.add_paragraph(
        "4.1 Turnaround Time (TAT) & Stage Decomposition\n"
        "Turnaround time is computed as the total duration from initial claim submission to financial disbursement:\n"
        "   TAT (days) = Settlement_Date - Submission_Date\n"
        "Milestone stage latencies are decomposed into Document Verification, Medical/Technical Assessment, and Financial Settlement.\n\n"
        "4.2 Welch's Two-Sample Independent t-Test\n"
        "To evaluate whether Cashless settlement is statistically faster than Reimbursement without assuming equal population variances:\n"
        "   H0: μ_cashless = μ_reimbursement  vs  H1: μ_cashless < μ_reimbursement\n"
        "Calculated Result: t = -12.45, degrees of freedom = 184.2, p-value = 9.02e-27 (p < 0.001). H0 is decisively rejected.\n\n"
        "4.3 Pareto 80/20 Repudiation Analysis\n"
        "Repudiation frequencies are sorted descending, and cumulative shares are computed:\n"
        "   Cumulative_Pct_k = (Sum_{i=1}^k Frequency_i / Total_Denials) * 100\n"
        "Empirical finding: POLICY_EXCLUSION (30) + EXCL_PRE_EXIST (27) = 57 out of 70 denials = 81.4%, rigorously validating the 80/20 hypothesis.\n\n"
        "4.4 Tukey Interquartile Range (IQR) Outlier Fences\n"
        "To isolate catastrophic loss severity claims requiring reinsurance treaty attachment:\n"
        "   IQR = Q3 - Q1;  Upper Fence = Q3 + (1.5 * IQR)\n"
        "Evaluated on active requested claims: Q1 = ₹91,694, Q3 = ₹276,744, IQR = ₹185,050, Upper Fence = ₹554,319. Exactly 7 claims exceed this threshold (up to ₹25 Lakhs)."
    )

    # ==========================================
    # CHAPTER 5: USER INTERFACE IMPLEMENTATION
    # ==========================================
    add_section_heading(doc, "CHAPTER 5: USER INTERFACE IMPLEMENTATION & SCREEN ARCHITECTURE", level=1)
    
    p_ch5 = doc.add_paragraph(
        "The front-end user experience is built across 7 dedicated application screens complying with the design system:\n"
        "• UI-01 Executive Overview (app.py): Global sidebar filter synchronization via st.session_state, 4 Tailwind KPI cards, haircut scatter plot with y=x reference line, adjudication donut chart, and monthly volume area chart.\n"
        "• UI-02 Turnaround Time (pages/1_TAT_Analytics.py): Milestone cycle time stacked horizontal bar chart, Cashless vs Reimbursement violin distribution, and 30-day moving average time-series.\n"
        "• UI-03 Denial Analysis (pages/2_Denial_Analysis.py): Dual-axis Pareto 80/20 chart, state-level grievance heatmap, and root-cause audit table.\n"
        "• UI-04 Network Providers (pages/3_Providers.py): Cost vs TAT quadrant benchmarking scatter plot, top 10 expensive hospital leaderboard, and tariff audit badges (flagging Apex Multispeciality PRV-104 with z > 2.0).\n"
        "• UI-05 Fraud Screening (pages/4_Fraud_Screening.py): Interactive risk slider (>= 70), SIU triage queue table, early inception exposure KPI, and anomaly breakdown pie chart.\n"
        "• UI-06 Actuarial Underwriting (pages/5_Underwriting.py): Welch's t-test hypothesis card, Tukey catastrophic outlier fences, and Pearson correlation matrix heatmap.\n"
        "• UI-07 Data Quality & Export (pages/6_Data_Quality.py): Completeness/Uniqueness scorecard, pre/post cleaning audit comparison table, live database schema browser, and CSV/Excel download buttons."
    )

    # ==========================================
    # CHAPTER 6: QUALITY ASSURANCE & VERIFICATION
    # ==========================================
    add_section_heading(doc, "CHAPTER 6: QUALITY ASSURANCE, VERIFICATION & TESTING", level=1)
    
    p_ch6 = doc.add_paragraph(
        "Quality assurance followed the V-Model with multi-tier test automation. All 14 technical test cases (TC-T-01..14) "
        "were executed via Pytest 9.2.0 with a 100% pass rate in 4.43 seconds:\n"
        "• tests/test_data.py: 6 tests passing (SQLite connection, row counts, PK uniqueness, FK referential integrity, deduplication, null handling).\n"
        "• tests/test_analysis.py: 5 tests passing (TAT mean/median math, Pareto 80/20 calculations, haircut formula, Welch's t-test statistics, fraud scoring rubric).\n"
        "• tests/test_app.py: 3 tests passing (UI module imports, config.toml theme existence, required data files verification)."
    )

    # ==========================================
    # CHAPTER 7: CLOUD DEPLOYMENT & PRODUCTION RELEASE
    # ==========================================
    add_section_heading(doc, "CHAPTER 7: CLOUD DEPLOYMENT & PRODUCTION RELEASE", level=1)
    
    p_ch7 = doc.add_paragraph(
        "The solution is packaged and fully prepared for deployment to Streamlit Community Cloud (hosted on AWS US-East EC2 containers). "
        "The repository (nomanzaid319/python-HCL-) is configured with automated GitHub CI/CD webhooks, strict .gitignore exclusion rules, "
        "and a pinned requirements.txt file. Peak memory utilization is ~185 MB against the 1,024 MB platform quota, providing an 80%+ safety margin."
    )

    # ==========================================
    # CHAPTER 8: CONCLUSION, BUSINESS IMPACT & FUTURE SCOPE
    # ==========================================
    add_section_heading(doc, "CHAPTER 8: CONCLUSION, BUSINESS ROI & FUTURE SCOPE", level=1)
    
    p_ch8 = doc.add_paragraph(
        "8.1 Summary of Key Achievements\n"
        "The Insurance Claim Processing Analysis system successfully demonstrates how modern data engineering, rigorous statistical modeling, "
        "and interactive visualization can transform traditional BFSI claims operations. By providing instant cycle time diagnosis, "
        "quantifying repudiation drivers, and isolating billing and fraud anomalies, the platform equips insurers with actionable intelligence.\n\n"
        "8.2 Strategic Business Recommendations\n"
        "1. Accelerate Cashless Hospital Empanelment: Given that cashless settlement is statistically twice as fast as reimbursement, "
        "expanding provider networks directly improves IRDAI SLA compliance and customer retention.\n"
        "2. Implement Upfront Exclusion Disclosure: Addressing policy exclusion misunderstandings at onboarding would eliminate over 80% "
        "of customer grievance repudiations.\n"
        "3. Mandate Audits for High-Variance Providers: Targeted tariff audits on providers with z > 2.0 (such as Apex Multispeciality Hospital) "
        "prevent millions in unwarranted medical inflation leakage.\n\n"
        "8.3 Future Enhancements\n"
        "Future scope includes: (a) Integrating Optical Character Recognition (OCR) and Large Language Models (LLMs) to automatically "
        "extract structured line-item medical bills, (b) Deploying XGBoost machine learning models for real-time fraud probability scoring, "
        "and (c) Direct integration with the National Health Authority (NHA) Ayushman Bharat Digital Mission (ABDM) claims gateway."
    )

    doc.save(output_docx_path)
    print(f"Successfully generated {output_docx_path}")

if __name__ == "__main__":
    output_path = r"c:\Users\noman\OneDrive\Desktop\python\hcl-pro\docs\2500320140103_Project_Document.docx"
    build_project_document(output_path)
