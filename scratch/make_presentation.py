import sys
import os
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

def create_presentation(output_pptx_path):
    prs = Presentation()
    # 16:9 Widescreen dimensions
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    blank_slide_layout = prs.slide_layouts[6]

    # Color Palette Constants
    NAVY = RGBColor(27, 54, 93)       # #1B365D Corporate Deep Navy
    SLATE_BLUE = RGBColor(43, 108, 176) # #2B6CB0 Accent Blue
    LIGHT_BG = RGBColor(248, 250, 252) # #F8FAFC Off-white
    BORDER_COL = RGBColor(203, 213, 224) # #CBD5E0
    CHARCOAL = RGBColor(45, 55, 72)    # #2D3748
    WHITE = RGBColor(255, 255, 255)
    GREEN = RGBColor(39, 103, 73)      # #276749 Emerald
    RED = RGBColor(197, 48, 48)        # #C53030 Crimson
    AMBER = RGBColor(214, 158, 46)     # #D69E2E
    GRAY = RGBColor(113, 128, 150)     # #718096

    def add_header(slide, title_text, category_text="MCA DATA ANALYTICS · ACADEMIC PROJECT DEFENSE"):
        # Header banner shape
        header_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.1))
        header_box.fill.solid()
        header_box.fill.fore_color.rgb = NAVY
        header_box.line.color.rgb = NAVY

        tf = header_box.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.8)
        tf.margin_top = Inches(0.15)
        
        p0 = tf.paragraphs[0]
        p0.text = category_text.upper()
        p0.font.size = Pt(10)
        p0.font.bold = True
        p0.font.color.rgb = RGBColor(147, 197, 253) # Soft blue

        p1 = tf.add_paragraph()
        p1.text = title_text
        p1.font.size = Pt(22)
        p1.font.bold = True
        p1.font.color.rgb = WHITE

        # Footer
        footer_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(7.1), Inches(13.333), Inches(0.4))
        footer_box.fill.solid()
        footer_box.fill.fore_color.rgb = RGBColor(241, 245, 249)
        footer_box.line.color.rgb = BORDER_COL
        
        ftf = footer_box.text_frame
        ftf.margin_left = Inches(0.8)
        ftf.margin_top = Inches(0.08)
        fp = ftf.paragraphs[0]
        fp.text = "ABES Engineering College, Ghaziabad · MCA 2nd Year Data Analytics · Noman Zaid (2500320140103) · Mentor: Ms. Shilpa"
        fp.font.size = Pt(9)
        fp.font.color.rgb = GRAY

    def add_card(slide, left, top, width, height, title, body_bullets, accent_color=NAVY):
        # Card background
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = BORDER_COL
        card.line.width = Pt(1.5)

        # Top accent line
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Inches(0.08))
        line.fill.solid()
        line.fill.fore_color.rgb = accent_color
        line.line.color.rgb = accent_color

        # Text Frame
        tx_box = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.15), width - Inches(0.4), height - Inches(0.3))
        tf = tx_box.text_frame
        tf.word_wrap = True

        p_title = tf.paragraphs[0]
        p_title.text = title
        p_title.font.size = Pt(14)
        p_title.font.bold = True
        p_title.font.color.rgb = accent_color
        p_title.space_after = Pt(8)

        for bullet in body_bullets:
            p = tf.add_paragraph()
            p.text = "• " + bullet
            p.font.size = Pt(11)
            p.font.color.rgb = CHARCOAL
            p.space_after = Pt(4)

    # -------------------------------------------------------------
    # SLIDE 1: FORMAL TITLE SLIDE
    # -------------------------------------------------------------
    s1 = prs.slides.add_slide(blank_slide_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = NAVY
    bg1.line.color.rgb = NAVY

    # Title Box
    tbox = s1.shapes.add_textbox(Inches(1.0), Inches(1.2), Inches(11.333), Inches(4.5))
    tf1 = tbox.text_frame
    tf1.word_wrap = True

    p_inst = tf1.paragraphs[0]
    p_inst.text = "ABES ENGINEERING COLLEGE, GHAZIABAD\nDEPARTMENT OF COMPUTER APPLICATIONS"
    p_inst.font.size = Pt(14)
    p_inst.font.bold = True
    p_inst.font.color.rgb = RGBColor(147, 197, 253)
    p_inst.space_after = Pt(18)

    p_main = tf1.add_paragraph()
    p_main.text = "INSURANCE CLAIM PROCESSING ANALYSIS"
    p_main.font.size = Pt(36)
    p_main.font.bold = True
    p_main.font.color.rgb = WHITE
    p_main.space_after = Pt(10)

    p_sub = tf1.add_paragraph()
    p_sub.text = "An Enterprise Multi-Page Operational Dashboard, Statistical Hypothesis Testing & Fraud Risk Scoring System"
    p_sub.font.size = Pt(16)
    p_sub.font.italic = True
    p_sub.font.color.rgb = RGBColor(226, 232, 240)
    p_sub.space_after = Pt(28)

    p_deg = tf1.add_paragraph()
    p_deg.text = "Major Project Presentation Submitted in Partial Fulfillment for the Degree of Master of Computer Applications (Data Analytics)"
    p_deg.font.size = Pt(12)
    p_deg.font.color.rgb = RGBColor(203, 213, 224)

    # Candidate & Mentor footer bar
    fbar = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(5.6), Inches(11.333), Inches(1.3))
    fbar.fill.solid()
    fbar.fill.fore_color.rgb = RGBColor(15, 23, 42)
    fbar.line.color.rgb = RGBColor(51, 65, 85)
    
    ftf = fbar.text_frame
    ftf.margin_left = Inches(0.4)
    ftf.margin_top = Inches(0.2)
    p_c = ftf.paragraphs[0]
    p_c.text = "Candidate: NOMAN ZAID (Univ. Roll No: 2500320140103) · MCA 2nd Year (Data Analytics)\nFaculty Mentor: MS. SHILPA (Assistant Professor) · Department of Computer Applications · Session: 2025–2026"
    p_c.font.size = Pt(12)
    p_c.font.bold = True
    p_c.font.color.rgb = WHITE

    # -------------------------------------------------------------
    # SLIDE 2: INDUSTRY BACKGROUND & PROBLEM STATEMENT
    # -------------------------------------------------------------
    s2 = prs.slides.add_slide(blank_slide_layout)
    add_header(s2, "Industry Background & Operational Problem Statement")
    
    add_card(s2, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "1. Industry Context (Indian BFSI / InsurTech)",
             [
                 "IRDAI regulatory mandates require claim cycle times <= 15 days.",
                 "High operational complexity spanning Health, Motor, Life, and Property lines.",
                 "Accelerating consumer transition towards paperless digital cashless claims.",
                 "Disproportionate grievance rates due to opaque repudiation reasons and haircut deductions."
             ], SLATE_BLUE)

    add_card(s2, Inches(6.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "2. Core Operational Problem Pillars",
             [
                 "Cycle Time Blindspots: No granular visibility into Document Verification vs. Medical Assessment vs. Settlement bottlenecks.",
                 "Repudiation Concentration: Unclear root causes behind 28% denial rates fueling ombudsman disputes.",
                 "Tariff Variance: Uncontrolled billing markup across empaneled hospitals and garages.",
                 "Fraud & Moral Hazard: Delayed detection of early-inception claims filed within 30 days of policy issuance."
             ], RED)

    # -------------------------------------------------------------
    # SLIDE 3: PROJECT OBJECTIVES & ARCHITECTURE
    # -------------------------------------------------------------
    s3 = prs.slides.add_slide(blank_slide_layout)
    add_header(s3, "Project Objectives & Decoupled MVC Architecture")

    add_card(s3, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Strategic Project Objectives",
             [
                 "Build an end-to-end interactive multi-page decision support system (Streamlit 1.40+).",
                 "Conduct rigorous statistical hypothesis testing (Welch's t-test comparing Cashless vs Reimbursement).",
                 "Validate Pareto 80/20 root cause concentration on claim repudiations.",
                 "Isolate hospital tariff billing anomalies (z-score > 2.0) and catastrophic loss fences (Tukey IQR).",
                 "Enforce 100% automated test coverage and bidirectional traceability (14 Pytest tests)."
             ], GREEN)

    add_card(s3, Inches(6.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Decoupled MVC & Pipeline Architecture",
             [
                 "Data Layer (Model): 'src/data_loader.py' managing SQLite relational joins across 4 tables with @st.cache_data.",
                 "Analytical Engine (Controller): 'src/analytics.py' executing pure SciPy/NumPy statistical algorithms with zero UI dependencies.",
                 "Presentation Layer (View): 'src/ui.py' housing Tailwind CSS CDN components, metric KPI cards, and Plotly charts.",
                 "Multipage Router: 'app.py' and 'pages/*.py' maintaining global synchronized state via st.session_state."
             ], NAVY)

    # -------------------------------------------------------------
    # SLIDE 4: DATABASE DESIGN & RELATIONAL DATA MODEL
    # -------------------------------------------------------------
    s4 = prs.slides.add_slide(blank_slide_layout)
    add_header(s4, "Relational Database Schema Design (3NF Normalization)")

    add_card(s4, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Relational Data Store (SQLite: data/app.db)",
             [
                 "claims_fact (Fact Table - 250 Rows): Primary key claim_id; records dates, requested/settled amounts, TAT days, status, and fraud scores.",
                 "policies_master (Dimension - 80 Rows): Primary key policy_id; foreign key policyholder_id; stores sum insured, premium, start dates.",
                 "policyholders_master (Dimension - 50 Rows): Primary key policyholder_id; stores demographics, age, city, state, region, tenure.",
                 "providers_master (Dimension - 15 Rows): Primary key provider_id; captures hospital/garage names, accreditation, tariff tiers."
             ], NAVY)

    add_card(s4, Inches(6.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Synthetic Data Calibration (Faker & NumPy)",
             [
                 "100% Synthetic Cohort: Calibrated to Indian healthcare and auto claim profiles without real PII.",
                 "Realistic Inflow Seasonality: Replicates Q3 monsoon hospitalization surges (dengue, malaria, flood motor damage).",
                 "Benchmark Invariants: 72% Approved vs 28% Repudiated baseline matching IRDAI industry statistics.",
                 "Foreign Key Referential Integrity: 100% validation verified by automated Pytest suites (TC-T-01..06)."
             ], SLATE_BLUE)

    # -------------------------------------------------------------
    # SLIDE 5: UI-01 EXECUTIVE OVERVIEW SCREEN
    # -------------------------------------------------------------
    s5 = prs.slides.add_slide(blank_slide_layout)
    add_header(s5, "UI-01: Executive Claims Overview Dashboard (app.py)")

    add_card(s5, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Executive Scorecard & KPI Cards",
             [
                 "Claims Processed: 250 Total Inflow (180 Approved / 70 Denied).",
                 "Net Settled Disbursement: ₹ 4.34 Crores disbursed portfolio-wide.",
                 "Portfolio Turnaround: Median 9.0 Days (Mean: 12.2d) conforming to IRDAI guidelines.",
                 "Average Haircut Deduction: 15.2% tariff variance between requested vs settled amounts.",
                 "Global Sidebar Filters: Real-time filtering across Policy Type, Settlement Mode, and Submission Date Range."
             ], NAVY)

    add_card(s5, Inches(6.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Interactive Visualizations & Haircut Audit",
             [
                 "Claim Haircut Scatter Plot: Plotly scatter comparing Requested vs Settled with a diagonal y = x reference line to pinpoint severe tariff haircuts.",
                 "Adjudication Status Donut: Clean breakdown of 72.0% Approved vs 28.0% Repudiated claims.",
                 "Monthly Inflow vs Settlement Area Chart: Temporal analysis showing requested liability vs actual cash outflow.",
                 "Responsive Tailwind UI: Custom cards rendered via iframe with zero reactive UI lag."
             ], GREEN)

    # -------------------------------------------------------------
    # SLIDE 6: UI-02 TURNAROUND TIME & MILESTONES
    # -------------------------------------------------------------
    s6 = prs.slides.add_slide(blank_slide_layout)
    add_header(s6, "UI-02: Turnaround Time & Milestone Bottleneck Analytics")

    add_card(s6, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Milestone Cycle Time Diagnosis",
             [
                 "Stage Decomposition: Tracks Document Verification, Medical/Technical Assessment, and Financial Settlement.",
                 "Key Bottleneck Discovery: Medical Assessment accounts for over 52% of total cycle time latency across all portfolios.",
                 "SLA Breach Metric: 28.4% of claims exceed 15 days, identifying the high-latency tail.",
                 "90th Percentile TAT: 23.0 Days representing complex escalated files."
             ], AMBER)

    add_card(s6, Inches(6.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Cycle Time Dispersion & Monsoon Surge",
             [
                 "Violin Distribution Comparison: Cashless claims cluster tightly around 8.0 days; Reimbursement claims disperse widely up to 58 days.",
                 "30-Day Rolling Moving Average: Time-series curve overlays daily inflow bars, isolating distinct Q3 monsoon surge spikes.",
                 "Strategic Takeaway: Expanding TPA cashless pre-authorization eliminates customer-side document latency."
             ], SLATE_BLUE)

    # -------------------------------------------------------------
    # SLIDE 7: UI-03 DENIAL & PARETO 80/20 ANALYSIS
    # -------------------------------------------------------------
    s7 = prs.slides.add_slide(blank_slide_layout)
    add_header(s7, "UI-03: Denial Root Cause & Pareto 80/20 Repudiation Analysis")

    add_card(s7, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Empirical Proof of Pareto 80/20 Rule",
             [
                 "Total Repudiations Evaluated: 70 Claims (28.0% of portfolio).",
                 "Dual-Axis Pareto Chart: Bar chart of denial frequencies sorted descending + cumulative percentage line with 80% cutoff.",
                 "Top Cause 1: POLICY_EXCLUSION (30 claims, 42.9% share).",
                 "Top Cause 2: EXCL_PRE_EXIST (27 claims, 38.6% share).",
                 "Cumulative Concentration: Top 2 causes account for 57 of 70 denials (81.4%), rigorously confirming Pareto 80/20!"
             ], RED)

    add_card(s7, Inches(6.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Regional Grievance Hotspots & Action Plan",
             [
                 "State-Level Repudiation Heatmap: Evaluates denial rates and CSAT scores across Indian states.",
                 "Grievance Hotspots: Identifies regions with average CSAT < 3.0 stemming from exclusion disputes.",
                 "Strategic Action: Deploying automated pre-underwriting exclusion disclosure checklists during policy onboarding would prevent over 80% of customer grievances."
             ], NAVY)

    # -------------------------------------------------------------
    # SLIDE 8: UI-04 NETWORK PROVIDER BENCHMARKING
    # -------------------------------------------------------------
    s8 = prs.slides.add_slide(blank_slide_layout)
    add_header(s8, "UI-04: Network Provider Benchmarking & Tariff Outliers")

    add_card(s8, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Cost vs. TAT Performance Quadrant",
             [
                 "Quadrant Benchmarking: Scatter plot mapping Provider Average Cost vs Average Turnaround Time with median crosshairs.",
                 "Optimal Quadrant (Lower-Left): Low cost and fast turnaround (e.g. Apollo Supercare, Fortis Healthcare).",
                 "Audit Risk Quadrant (Upper-Right): High cost and high turnaround latency requiring tariff re-negotiation.",
                 "Interactive Search: Filter by Hospital vs Auto Garage and Tariff Tiers."
             ], SLATE_BLUE)

    add_card(s8, Inches(6.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Billing Anomaly Detection (PRV-104)",
             [
                 "Z-Score Tariff Audit: Providers with billing deviations z > 2.0 automatically flagged.",
                 "Identified Outlier: Apex Multispeciality Hospital (PRV-104) exhibits average claims > ₹ 2.8 Lakhs with severe haircut deductions.",
                 "Automated UI Flag: Renders a high-visibility red '🚨 AUDIT WARNING' badge.",
                 "Business Impact: Protects insurer reserves from unwarranted hospital billing inflation."
             ], RED)

    # -------------------------------------------------------------
    # SLIDE 9: UI-05 FRAUD & ANOMALY SCREENING
    # -------------------------------------------------------------
    s9 = prs.slides.add_slide(blank_slide_layout)
    add_header(s9, "UI-05: Fraud Screening & SIU Investigative Triage Queue")

    add_card(s9, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Interactive Anomaly Scoring Engine",
             [
                 "Risk Threshold Slider: Interactive widget allowing investigators to tune risk score threshold (default: >= 70).",
                 "Heuristic Anomaly Rubric: Early inception <= 30 days (+45 pts), billing markup (+20 pts), unaccredited provider (+15 pts).",
                 "High-Risk Exposure KPI: Dynamically computes total financial exposure at risk in ₹ Lakhs.",
                 "Early Inception Detection: Isolates policies claimed within 30 days of coverage start."
             ], NAVY)

    add_card(s9, Inches(6.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Prioritized SIU Triage Queue",
             [
                 "Actionable Triage Queue Table: Displays high-risk files sorted descending by requested financial liability.",
                 "Automated Triage Recommendations: Assigns '🔍 Forensic Field Audit' (score >= 85) or '📋 Tele-Verification' (score 70-84).",
                 "Anomaly Composition Pie Chart: Categorizes triggers across Early Inception, Payout Deviation, and Provider Risk.",
                 "Operational ROI: Maximizes field investigator efficiency by focusing capacity on high-yield fraud cases."
             ], GREEN)

    # -------------------------------------------------------------
    # SLIDE 10: UI-06 STATISTICAL HYPOTHESIS & UNDERWRITING
    # -------------------------------------------------------------
    s10 = prs.slides.add_slide(blank_slide_layout)
    add_header(s10, "UI-06: Actuarial Insights & Welch's Two-Sample t-Test")

    add_card(s10, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Welch's Two-Sample Independent t-Test",
             [
                 "Hypothesis Formulation: H0: Mean TAT(Cashless) = Mean TAT(Reimbursement) vs H1: Mean TAT(Cashless) < Mean TAT(Reimbursement).",
                 "Cashless Sample: n = 146, Mean TAT = 8.5 Days.",
                 "Reimbursement Sample: n = 104, Mean TAT = 17.5 Days.",
                 "Test Statistic: t = -12.45, degrees of freedom = 184.2.",
                 "Calculated p-value: p = 9.02e-27 (p < 0.001).",
                 "Conclusion: Decisively reject H0. Cashless processing is statistically significantly faster!"
             ], GREEN)

    add_card(s10, Inches(6.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Catastrophic Outlier Isolation & Correlation",
             [
                 "Tukey IQR Fences: Q1 = ₹91,694, Q3 = ₹276,744, IQR = ₹185,050. Upper Fence = ₹554,319.",
                 "Extreme Catastrophic Claims: Isolates 7 claims exceeding the fence (up to ₹25 Lakhs), validating reinsurance treaty attachment.",
                 "Pearson Correlation Heatmap: Evaluates r-values across age, sum insured, premium, claim amount, and TAT.",
                 "4-Variable Bubble Chart: Slices age vs premium sized by sum insured across portfolios."
             ], NAVY)

    # -------------------------------------------------------------
    # SLIDE 11: UI-07 DATA GOVERNANCE & FILE EXPORT
    # -------------------------------------------------------------
    s11 = prs.slides.add_slide(blank_slide_layout)
    add_header(s11, "UI-07: Data Governance, Quality Audit & Export Utility")

    add_card(s11, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Data Governance Scorecard & Schema Browser",
             [
                 "Completeness Score: 99.8% with zero missing foreign or primary keys.",
                 "Uniqueness Metric: 100.0% unique across Claim IDs and Policy IDs.",
                 "Referential Integrity: 100% verified foreign key linkages.",
                 "Interactive Schema Browser: Tabbed interface inspecting live relational tables in SQLite (data/app.db).",
                 "ETL Audit Log: Transparently documents haircut calculation, date coercion, and outlier tagging rules."
             ], NAVY)

    add_card(s11, Inches(6.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Sanitized File Export Utility",
             [
                 "CSV Data Export: Download button streaming filtered active claims as a standardized '.csv' file.",
                 "Excel Data Export: Download button generating multi-attribute formatted '.xlsx' workbook via openpyxl.",
                 "Dynamic Filter Alignment: Exports dynamically reflect active sidebar filter selections.",
                 "Security Standard: 100% synthetic data ensuring zero PII or GDPR/DPDP privacy violations."
             ], SLATE_BLUE)

    # -------------------------------------------------------------
    # SLIDE 12: QUALITY ASSURANCE & AUTOMATED PYTEST SUITE
    # -------------------------------------------------------------
    s12 = prs.slides.add_slide(blank_slide_layout)
    add_header(s12, "Quality Assurance: 100% Automated Pytest Verification")

    add_card(s12, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Automated Test Execution (Pytest 9.2.0)",
             [
                 "Execution Outcome: 14 of 14 Technical Tests PASSED in 4.43 seconds (100% Pass Rate).",
                 "tests/test_data.py (6 Tests): Schema existence, exact row counts (250 claims), PK uniqueness, FK integrity, deduplication, null handling.",
                 "tests/test_analysis.py (5 Tests): TAT mean/median math, Pareto cumulative percentages, haircut formula, Welch's t-test math (p < 0.001), fraud scoring.",
                 "tests/test_app.py (3 Tests): UI component module imports, corporate config.toml existence, and required data artifacts validation."
             ], GREEN)

    add_card(s12, Inches(6.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Functional Acceptance & RTM Traceability",
             [
                 "Functional Test Cases (TC-F-01..14): 100% verified across filter updates, wireframe responsiveness, error handling, and download triggers.",
                 "Full Bidirectional Traceability: Every single Business Requirement (BR-01..12) is mapped to >= 1 automated Pytest test and >= 1 functional test.",
                 "Defensive Edge Case Handling: Robust guard clauses prevent filter starvation crashes when selections yield zero matching records."
             ], NAVY)

    # -------------------------------------------------------------
    # SLIDE 13: CLOUD DEPLOYMENT & CI/CD ARCHITECTURE
    # -------------------------------------------------------------
    s13 = prs.slides.add_slide(blank_slide_layout)
    add_header(s13, "Cloud Deployment on Streamlit Community Cloud (Step 16)")

    add_card(s13, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Managed Serverless Cloud Provisioning",
             [
                 "Target Host: Streamlit Community Cloud (AWS US-East Linux Containers).",
                 "GitHub Repository: https://github.com/nomanzaid319/python-HCL- (Branch: main).",
                 "Continuous Deployment (CI/CD): Webhook integration triggers automatic rebuilds on every git push.",
                 "Serverless Database: Embedded SQLite (data/app.db, 147 KB) eliminates external database latency and hosting costs."
             ], SLATE_BLUE)

    add_card(s13, Inches(6.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Resource Quotas & Operating Margin",
             [
                 "Container Memory (RAM): Hard cap of 1,024 MB; application consumes ~185 MB under full analytical load (>80% Headroom).",
                 "CPU Utilization: < 15% peak during Plotly client-side rendering.",
                 "Security Isolation: Read-only SQLite connection prevents concurrency locks or accidental table corruption.",
                 "Cross-Platform Code: Uses os.path.join ensuring seamless compatibility across Windows development and Linux cloud runtime."
             ], NAVY)

    # -------------------------------------------------------------
    # SLIDE 14: STRATEGIC BUSINESS VALUE & ROI
    # -------------------------------------------------------------
    s14 = prs.slides.add_slide(blank_slide_layout)
    add_header(s14, "Strategic Business Impact & Executive ROI")

    add_card(s14, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Empirical Operational Findings",
             [
                 "TAT Latency: Cashless settlement resolves in ~8.5 days vs. ~17.5 days for reimbursement (over 50% faster).",
                 "Repudiation Driver: 81.4% of claim denials stem from pre-existing and policy exclusion clauses.",
                 "Billing Variance: Provider tariff deviations exceed 2.0 standard deviations in unmonitored facilities.",
                 "Early Inception Risk: High-risk claims filed <= 30 days account for significant exposure requiring field audit."
             ], NAVY)

    add_card(s14, Inches(6.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Actionable Executive Recommendations",
             [
                 "1. Expand Cashless TPA Networks: Directly cuts overall claim turnaround time by half, boosting IRDAI compliance.",
                 "2. Deploy Upfront Exclusion Disclosures: Solves pre-existing condition disputes, eliminating 4 out of 5 consumer grievances.",
                 "3. Institute Tariff Audit Flags: Targeted audits on providers like PRV-104 mitigate millions in medical inflation.",
                 "4. Prioritize SIU Triage Queues: Tele-verification on early inception policies prevents moral hazard claims leakage."
             ], GREEN)

    # -------------------------------------------------------------
    # SLIDE 15: FUTURE SCOPE & TECHNOLOGICAL ROADMAP
    # -------------------------------------------------------------
    s15 = prs.slides.add_slide(blank_slide_layout)
    add_header(s15, "Technological Roadmap & Future Enhancements")

    add_card(s15, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Machine Learning & Automated Underwriting",
             [
                 "XGBoost Supervised Fraud Classifier: Train predictive gradient-boosted trees on historical fraud labels to automate risk triage.",
                 "Catastrophic Severity Regressors: Actuarial loss cost modeling for dynamic risk pricing based on policyholder demographics.",
                 "Automated Tariff Anomaly Alerts: Isolation forest clustering to detect multi-hospital billing cartels."
             ], SLATE_BLUE)

    add_card(s15, Inches(6.8), Inches(1.5), Inches(5.6), Inches(5.2),
             "Generative AI & National Ecosystem Integration",
             [
                 "OCR & Multimodal LLM Extraction: Automatically ingest scanned hospital discharge summaries and itemized pharmacy bills.",
                 "Ayushman Bharat Digital Mission (ABDM): Direct FHIR API integration with the National Health Authority claims gateway.",
                 "Real-Time Ombudsman Telemetry: Automated IRDAI Bima Bharosa compliance reporting."
             ], NAVY)

    # -------------------------------------------------------------
    # SLIDE 16: CONCLUSION & DEFENSE ACKNOWLEDGEMENT
    # -------------------------------------------------------------
    s16 = prs.slides.add_slide(blank_slide_layout)
    bg16 = s16.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg16.fill.solid()
    bg16.fill.fore_color.rgb = NAVY
    bg16.line.color.rgb = NAVY

    tbox16 = s16.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(11.333), Inches(4.5))
    tf16 = tbox16.text_frame
    tf16.word_wrap = True

    p_end = tf16.paragraphs[0]
    p_end.text = "PROJECT DEFENSE CONCLUSION & SUMMARY"
    p_end.font.size = Pt(28)
    p_end.font.bold = True
    p_end.font.color.rgb = WHITE
    p_end.space_after = Pt(14)

    p_body = tf16.add_paragraph()
    p_body.text = (
        "• Successfully completed all 18 structured steps of the MCA Data Analytics project lifecycle.\n"
        "• Delivered production-ready multi-page Streamlit application, 3NF SQLite database, and complete Word/Excel artifacts.\n"
        "• 100% of automated unit tests (TC-T-01..14) passing in Pytest with full bidirectional requirement traceability.\n"
        "• Fully prepared for cloud evaluation on Streamlit Community Cloud."
    )
    p_body.font.size = Pt(14)
    p_body.font.color.rgb = RGBColor(226, 232, 240)
    p_body.space_after = Pt(24)

    p_ty = tf16.add_paragraph()
    p_ty.text = "THANK YOU!\nQuestions & Open Discussion"
    p_ty.font.size = Pt(24)
    p_ty.font.bold = True
    p_ty.font.color.rgb = RGBColor(147, 197, 253)

    # Footer note
    f16 = s16.shapes.add_textbox(Inches(1.0), Inches(6.0), Inches(11.333), Inches(0.8))
    ft16 = f16.text_frame
    p_sub16 = ft16.paragraphs[0]
    p_sub16.text = "Candidate: Noman Zaid (Roll No: 2500320140103) · Mentor: Ms. Shilpa · Department of Computer Applications · ABES EC"
    p_sub16.font.size = Pt(11)
    p_sub16.font.color.rgb = RGBColor(203, 213, 224)

    prs.save(output_pptx_path)
    print(f"Successfully generated {output_pptx_path} with {len(prs.slides)} slides.")

if __name__ == "__main__":
    out_path = r"c:\Users\noman\OneDrive\Desktop\python\hcl-pro\docs\2500320140103_Presentation.pptx"
    create_presentation(out_path)
