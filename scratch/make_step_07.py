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

def generate_trd():
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
    run_title = p_title.add_run("Step 7: Technical Requirements Document (TRD)")
    run_title.font.name = 'Calibri'
    run_title.font.size = Pt(24)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(16)
    run_sub = p_sub.add_run("Engineering Specifications, System Architecture, Statistical Pipelines, and Cloud Deployment Blueprint")
    run_sub.font.size = Pt(13)
    run_sub.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)

    # Document Control
    ctrl_table = doc.add_table(rows=6, cols=2)
    ctrl_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(ctrl_table, "CBD5E0")

    ctrl_data = [
        ("Document Title", "Technical Requirements Document (TRD) - Insurance Claim Analytics"),
        ("Version & Status", "Version 1.0 (Officially Baselined)"),
        ("Lead Software Architect / Author", "Noman Zaid (Roll No: 2500320140103, MCA 2nd Year DA)"),
        ("Faculty Mentor", "Ms. Shilpa, Department of Computer Applications, ABES EC"),
        ("Upstream Baselines", "docs/05_BRD.docx (BR-01 to 12), docs/06_PRD.docx (F-01 to 12)"),
        ("Target Technology Stack", "Python 3.12, Pandas, NumPy, SciPy, Statsmodels, Plotly, Streamlit, Tailwind, SQLite, pytest")
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

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # Section 1: System Architecture
    h1 = doc.add_heading(level=1)
    r = h1.add_run("1. System Architecture & Data Flow")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "The application is engineered as a decoupled, layered analytical system following modern data science engineering best practices. The high-level data flow pipeline progresses unidirectionally: Data Layer -> Loading & Caching -> Cleaning & Validation -> Statistical Processing -> Interactive UI Delivery:"
    )

    # Architecture Box with Mermaid Code
    arch_box = doc.add_table(rows=1, cols=1)
    arch_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = arch_box.rows[0].cells[0]
    c.width = Inches(6.5)
    set_cell_background(c, "F7FAFC")
    set_cell_margins(c, 100, 100, 140, 140)
    set_table_borders(arch_box, "CBD5E0")
    cp = c.paragraphs[0]
    cp.paragraph_format.space_after = Pt(0)
    cr_title = cp.add_run("System Architecture Data Pipeline Flow (Mermaid Diagram):\n\n")
    cr_title.font.bold = True
    cr_title.font.size = Pt(10)
    cr_title.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    mermaid_txt = (
        "flowchart TD\n"
        "    subgraph Storage [Persistent Storage Layer]\n"
        "        DB[(SQLite data/app.db)]\n"
        "        CSV[(Static CSV Fallbacks)]\n"
        "    end\n\n"
        "    subgraph Ingestion [Data Ingestion & Caching Layer]\n"
        "        DL[src/data_loader.py]\n"
        "        CACHE[@st.cache_data In-Memory Buffer]\n"
        "    end\n\n"
        "    subgraph Quality [Data Quality & Transformation Layer]\n"
        "        DC[src/data_cleaner.py]\n"
        "        LOG[Audit & Anomaly Logging]\n"
        "    end\n\n"
        "    subgraph Analytics [Analytical & Statistical Engine]\n"
        "        AN[src/analytics.py: Aggregations & KPIs]\n"
        "        ST[src/statistics.py: Hypothesis & KDE]\n"
        "    end\n\n"
        "    subgraph Presentation [Streamlit Multipage UI Layer]\n"
        "        APP[app.py: Executive Overview]\n"
        "        P1[pages/1_TAT_Analytics.py]\n"
        "        P2[pages/2_Denial_Analysis.py]\n"
        "        P3[pages/3_Providers.py]\n"
        "        P4[pages/4_Fraud_Screening.py]\n"
        "        P5[pages/5_Underwriting.py]\n"
        "        P6[pages/6_Data_Quality.py]\n"
        "        TW[Tailwind CSS Cards via st.iframe]\n"
        "        PLT[Interactive Plotly Engine]\n"
        "    end\n\n"
        "    Storage --> DL --> CACHE --> DC --> LOG\n"
        "    DC --> AN & ST\n"
        "    AN & ST --> APP & P1 & P2 & P3 & P4 & P5 & P6\n"
        "    APP & P1 & P2 & P3 & P4 & P5 & P6 --> TW & PLT\n"
    )
    cr_code = cp.add_run(mermaid_txt)
    cr_code.font.name = "Consolas"
    cr_code.font.size = Pt(8.5)
    cr_code.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 2: Folder Structure
    h1 = doc.add_heading(level=1)
    r = h1.add_run("2. Repository & Folder Structure")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "The project strictly adheres to professional Python application standards, isolating configuration, documentation, source modules, tests, and database storage:"
    )

    folder_box = doc.add_table(rows=1, cols=1)
    folder_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = folder_box.rows[0].cells[0]
    c.width = Inches(6.5)
    set_cell_background(c, "2D3748") # Dark slate
    set_cell_margins(c, 100, 100, 140, 140)
    set_table_borders(folder_box, "1A202C")
    cp = c.paragraphs[0]
    cp.paragraph_format.space_after = Pt(0)
    folder_txt = (
        "hcl-pro/\n"
        "├── .streamlit/\n"
        "│   └── config.toml             # Theme settings (primary, background, fonts)\n"
        "├── app.py                      # Main entry point & Executive Overview (UI-01)\n"
        "├── pages/                      # Multipage application screens\n"
        "│   ├── 1_TAT_Analytics.py      # Stage-wise cycle times & bottlenecks (UI-02)\n"
        "│   ├── 2_Denial_Analysis.py    # Pareto repudiation analysis (UI-03)\n"
        "│   ├── 3_Providers.py          # Hospital/Garage benchmarking (UI-04)\n"
        "│   ├── 4_Fraud_Screening.py    # Anomaly scoring & exposure (UI-05)\n"
        "│   ├── 5_Underwriting.py       # Demographics & Hypothesis testing (UI-06)\n"
        "│   └── 6_Data_Quality.py       # Data cleaning log & export (UI-07)\n"
        "├── src/                        # Core backend application modules\n"
        "│   ├── __init__.py\n"
        "│   ├── data_loader.py          # SQLite connection & cached queries\n"
        "│   ├── data_cleaner.py         # Data validation, nulls, outliers, log\n"
        "│   ├── analytics.py            # KPI calculations, aggregations, Pareto\n"
        "│   ├── statistics.py           # Two-sample t-test, KDE, correlations\n"
        "│   ├── ui_components.py        # Tailwind KPI cards & iframe renderers\n"
        "│   └── generate_data.py        # Step 10 synthetic generator script\n"
        "├── data/                       # Local data storage (SQLite & CSV)\n"
        "│   ├── app.db                  # Production SQLite 3 database\n"
        "│   └── insurance_claims_data.xlsx\n"
        "├── tests/                      # Automated pytest test suite\n"
        "│   ├── __init__.py\n"
        "│   ├── test_data.py            # SQLite schema, PK/FK, null tests\n"
        "│   ├── test_analysis.py        # KPI math, Pareto, t-test tests\n"
        "│   └── test_app.py             # Streamlit AppTest page render checks\n"
        "├── docs/                       # Project artifacts (Word, Excel, PPTX)\n"
        "├── requirements.txt            # Pinned Python package dependencies\n"
        "├── BUILD_PLAN.md               # Step 14 machine-readable build steps\n"
        "└── README.md                   # Setup guide, live URL, and documentation\n"
    )
    cr_f = cp.add_run(folder_txt)
    cr_f.font.name = "Consolas"
    cr_f.font.size = Pt(8.5)
    cr_f.font.color.rgb = RGBColor(0xED, 0xF2, 0xF7)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 3: Functional Requirements (FR-01 to FR-12)
    h1 = doc.add_heading(level=1)
    r = h1.add_run("3. Technical Functional Requirements (FR-01 to FR-12)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "Technical functional requirements define the exact backend Python modules, function signatures, and execution complexity corresponding to each product feature:"
    )

    fr_table = doc.add_table(rows=13, cols=5)
    fr_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(fr_table, "CBD5E0")

    fr_headers = ["FR ID", "Technical Specification", "Linked BR / Feature", "Target Module & Function", "Complexity"]
    for j, h in enumerate(fr_headers):
        c = fr_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 70, 70)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    fr_data = [
        ("FR-01", "Compute stage-wise duration deltas (days) across submission, verification, assessment, and settlement milestones.", "BR-01 / F-01", "src/analytics.py :: compute_stage_tat()", "O(N)"),
        ("FR-02", "Calculate frequency distribution and cumulative percentage curve for claim repudiations.", "BR-02 / F-02", "src/analytics.py :: compute_pareto_denials()", "O(K log K)"),
        ("FR-03", "Compute payout haircut ratios ((Claimed - Settled) / Claimed) and identify deductions > 20%.", "BR-03 / F-03", "src/analytics.py :: compute_haircut_metrics()", "O(N)"),
        ("FR-04", "Standardize provider average settled amounts using Z-scores and evaluate cost vs. TAT quadrants.", "BR-04 / F-04", "src/analytics.py :: benchmark_providers()", "O(M)"),
        ("FR-05", "Evaluate early-inception (<30d), duplicate incidents, and high billing to output composite fraud scores.", "BR-05 / F-05", "src/analytics.py :: detect_claim_anomalies()", "O(N)"),
        ("FR-06", "Aggregate claims inflow by month/quarter and calculate 30-day rolling moving averages and MoM growth.", "BR-06 / F-06", "src/analytics.py :: compute_seasonal_trends()", "O(N)"),
        ("FR-07", "Compute parametric stats, Gaussian KDE density, and Tukey IQR upper fences (Q3 + 1.5*IQR).", "BR-07 / F-07", "src/statistics.py :: compute_outlier_distribution()", "O(N log N)"),
        ("FR-08", "Execute two-sample independent Welch's t-test comparing Cashless vs. Reimbursement TAT (alpha=0.05).", "BR-08 / F-08", "src/statistics.py :: test_settlement_modes_ttest()", "O(N)"),
        ("FR-09", "Aggregate customer dispute rates and compute weighted average CSAT scores across geographic territories.", "BR-09 / F-09", "src/analytics.py :: compute_geographic_grievances()", "O(N)"),
        ("FR-10", "Compute Pearson and Spearman correlation matrices across tenure, age, premium, and claim amounts.", "BR-10 / F-10", "src/statistics.py :: compute_demographic_correlations()", "O(N * V^2)"),
        ("FR-11", "Execute data quality audit: scan nulls, check duplicate PKs, validate FK integrity, log transformations.", "BR-11 / F-11", "src/data_cleaner.py :: validate_and_clean_data()", "O(N)"),
        ("FR-12", "Serialize filtered pandas dataframes to in-memory bytes buffers for CSV and Excel download.", "BR-12 / F-12", "src/ui_components.py :: export_dataframe_bytes()", "O(N)")
    ]

    col_w_fr = [Inches(0.7), Inches(2.3), Inches(1.1), Inches(1.8), Inches(0.6)]
    for i, row in enumerate(fr_data):
        r = fr_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w_fr[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 70, 70, 70, 70)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(8)
            if j == 0:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 4: Data Layer Architecture
    h1 = doc.add_heading(level=1)
    r = h1.add_run("4. Data Layer Architecture & Caching Strategy")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "4.1 SQLite Connection Strategy:\n"
        "The application utilizes SQLite 3 as an embedded relational datastore (`data/app.db`). To ensure cloud stability and prevent file lock conflicts:\n"
        "• Connections are established using Python's standard `sqlite3` module with `check_same_thread=False`.\n"
        "• Queries are read-optimized using parameterized SQL statements joined across the main fact table and three dimension masters.\n"
        "• SQLite Write-Ahead Logging (WAL) mode is enabled for rapid concurrent reads."
    )
    p = doc.add_paragraph(
        "4.2 Streamlit Caching Pipeline:\n"
        "All database queries and computationally heavy operations are cached in-memory using Streamlit's `@st.cache_data` decorator:\n"
        "```python\n"
        "@st.cache_data(ttl=3600, show_spinner=False)\n"
        "def load_claims_data(db_path: str) -> pd.DataFrame:\n"
        "    conn = sqlite3.connect(db_path, check_same_thread=False)\n"
        "    query = '''\n"
        "        SELECT c.*, p.policy_type, p.sum_insured, p.premium_amount,\n"
        "               ph.age, ph.gender, ph.city, ph.state, ph.tenure_months,\n"
        "               pr.provider_name, pr.provider_type, pr.empanelment_status\n"
        "        FROM claims_fact c\n"
        "        JOIN policies_master p ON c.policy_id = p.policy_id\n"
        "        JOIN policyholders_master ph ON p.policyholder_id = ph.policyholder_id\n"
        "        JOIN providers_master pr ON c.provider_id = pr.provider_id\n"
        "    '''\n"
        "    df = pd.read_sql_query(query, conn)\n"
        "    conn.close()\n"
        "    return df\n"
        "```\n"
        "This guarantees sub-second dashboard rendering upon repeated user filter interactions."
    )

    # Section 5: Statistical Methods for Each OQ
    h1 = doc.add_heading(level=1)
    r = h1.add_run("5. Mathematical & Statistical Methods (OQ-01 to OQ-10)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "Every Operational Question (OQ) is mapped to explicit mathematical formulations implemented via SciPy, Statsmodels, and NumPy:"
    )

    stat_table = doc.add_table(rows=11, cols=4)
    stat_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(stat_table, "CBD5E0")

    stat_headers = ["OQ ID", "Analytical Objective", "Mathematical Formulation / Method", "Python Implementation Library"]
    for j, h in enumerate(stat_headers):
        c = stat_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 80, 80)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    stat_methods = [
        ("OQ-01", "Stage-wise Turnaround Time", "Cycle Time Delta: Delta_T = t_milestone2 - t_milestone1; Aggregate: Mean(Delta_T) & Median(Delta_T)", "pandas, numpy"),
        ("OQ-02", "Pareto Repudiation Decomposition", "Cumulative Percentage: C_k = (Sum_{i=1}^k x_i / Sum_{i=1}^N x_i) * 100%; Rank order desc.", "pandas, plotly"),
        ("OQ-03", "Settlement Payout Haircut", "Haircut % = ((Claimed - Settled) / Claimed) * 100; Payout Ratio = Settled / Claimed", "pandas, numpy"),
        ("OQ-04", "Provider Cost Standard Score", "Standardized Z-Score: z = (x - mu) / sigma; Flagging z > 2.0 (abnormal inflation)", "scipy.stats.zscore"),
        ("OQ-05", "Multi-factor Anomaly Scoring", "Fraud Score = w1*EarlyInception(30d) + w2*Duplicate + w3*CostVariance; Exposure = Sum(Amt)", "src/analytics.py"),
        ("OQ-06", "Surge & Trend Time Series", "Rolling 30-Day Moving Average: MA_t = (1/k) * Sum_{i=0}^{k-1} x_{t-i}; MoM Growth %", "pandas.Series.rolling"),
        ("OQ-07", "Non-parametric Outlier Modeling", "Tukey IQR Fence: Upper = Q3 + 1.5*(Q3 - Q1); Gaussian Kernel Density Estimation (KDE)", "scipy.stats.gaussian_kde"),
        ("OQ-08", "Cashless vs. Reimbursement Test", "Independent Two-Sample Welch's t-test: t = (X1 - X2) / sqrt(s1^2/n1 + s2^2/n2), p-value, Cohen's d", "scipy.stats.ttest_ind"),
        ("OQ-09", "Geographic Grievance Rate", "Dispute Rate = (Escalations_region / Claims_region) * 100%; CSAT Weighted Mean", "pandas.DataFrame.groupby"),
        ("OQ-10", "Demographic Risk Correlation", "Pearson Correlation: r = Cov(X,Y) / (sigma_X * sigma_Y); Spearman Rank Correlation rho", "pandas.DataFrame.corr")
    ]

    col_w_st = [Inches(0.7), Inches(1.8), Inches(2.7), Inches(1.3)]
    for i, row in enumerate(stat_methods):
        r = stat_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w_st[j]
            set_cell_background(c, bg)
            set_cell_margins(c, 70, 70, 70, 70)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            run = p.add_run(val)
            run.font.size = Pt(8)
            if j == 0:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 6: UI & Tailwind Framework
    h1 = doc.add_heading(level=1)
    r = h1.add_run("6. UI Framework & Tailwind CSS Integration")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "6.1 Tailwind CSS via Streamlit Component Iframe:\n"
        "Streamlit native markdown lacks enterprise UI styling flexibility. To achieve polished executive cards, headers, and badges, the application injects Tailwind CSS via `streamlit.components.v1.html`:\n"
        "```python\n"
        "import streamlit.components.v1 as components\n\n"
        "def render_tailwind_kpi(title: str, value: str, subtext: str, bg_color: str = 'bg-blue-50'):\n"
        "    html_code = f'''\n"
        "    <script src=\"https://cdn.tailwindcss.com\"></script>\n"
        "    <div class=\"p-4 rounded-xl border border-slate-200 {bg_color} shadow-sm font-sans\">\n"
        "        <p class=\"text-xs font-semibold uppercase tracking-wider text-slate-500\">{title}</p>\n"
        "        <h3 class=\"text-2xl font-bold text-slate-800 my-1\">{value}</h3>\n"
        "        <p class=\"text-xs text-slate-500\">{subtext}</p>\n"
        "    </div>\n"
        "    '''\n"
        "    components.html(html_code, height=110)\n"
        "```\n"
        "6.2 Global Color Theme (.streamlit/config.toml):\n"
        "• Primary Accent: Corporate Navy `#1B365D`\n"
        "• Background Color: Soft Slate Gray `#F8FAFC`\n"
        "• Secondary Background (Sidebar): `#FFFFFF`\n"
        "• Text Color: Charcoal `#1A202C`\n"
        "• Font: Sans-serif (Inter / System Font)"
    )

    # Section 7: Performance, Error Handling & Security
    h1 = doc.add_heading(level=1)
    r = h1.add_run("7. Performance, Security & Error Handling")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "• Performance Ceilings: Dataset memory footprint is capped at under 20 MB in memory; categorical columns (`policy_type`, `claim_status`, `state`) are downcasted to `category` dtypes.\n"
        "• Defensive Error Handling: All analytical functions wrap calculations in try-except blocks, falling back to clean user-facing warnings (`st.warning(\"Insufficient filtered data points to compute correlation\")`) without unhandled stack traces.\n"
        "• Security & Privacy: No secrets or credentials are hardcoded. Database connections are local and read-only. Zero personal identifiable information (PII) is included in synthetic generation."
    )

    # Section 8: Pinned Dependencies
    h1 = doc.add_heading(level=1)
    r = h1.add_run("8. Pinned Package Dependencies (requirements.txt)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "To ensure deterministic, reproducible builds across local developer machines and Streamlit Community Cloud, package versions are strictly pinned:"
    )

    req_box = doc.add_table(rows=1, cols=1)
    req_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = req_box.rows[0].cells[0]
    c.width = Inches(6.5)
    set_cell_background(c, "F7FAFC")
    set_cell_margins(c, 80, 80, 120, 120)
    set_table_borders(req_box, "CBD5E0")
    cp = c.paragraphs[0]
    cp.paragraph_format.space_after = Pt(0)
    req_txt = (
        "streamlit==1.40.2\n"
        "pandas==2.2.3\n"
        "numpy==2.1.3\n"
        "scipy==1.14.1\n"
        "statsmodels==0.14.4\n"
        "plotly==5.24.1\n"
        "openpyxl==3.1.5\n"
        "python-docx==1.1.2\n"
        "pytest==8.3.4\n"
        "pytest-mock==3.14.0\n"
        "faker==33.1.0\n"
    )
    cr_r = cp.add_run(req_txt)
    cr_r.font.name = "Consolas"
    cr_r.font.size = Pt(9)
    cr_r.font.color.rgb = RGBColor(0x1A, 0x20, 0x2C)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 9: Deployment Requirements
    h1 = doc.add_heading(level=1)
    r = h1.add_run("9. Cloud Deployment Specifications")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "• Hosting Target: Streamlit Community Cloud (connected directly to GitHub repository).\n"
        "• Repository Structure: Root contains `app.py`, `requirements.txt`, and `.streamlit/config.toml`.\n"
        "• Path Resolution: All internal paths use dynamic relative paths: `Path(__file__).parent / 'data' / 'app.db'` to prevent Windows/Linux path separator errors.\n"
        "• Cloud Resource Management: Application maintains memory footprint well within the 1 GB container limit."
    )

    # Section 10: Technical Risks & Tradeoffs
    h1 = doc.add_heading(level=1)
    r = h1.add_run("10. Technical Risks & Architectural Tradeoffs")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "1. SQLite vs. Distributed Database: SQLite provides zero-configuration simplicity and complete portability within a git repository. The tradeoff is limited multi-threaded write capability, which is acceptable since the app is read-heavy.\n"
        "2. Client-Side Rendering vs. Server-Side: Plotly renders vector visualizations client-side in the user's browser, minimizing server CPU load on Streamlit Cloud.\n"
        "3. In-Memory Pandas vs. SQL Engine: For 200-500 claims records, pandas vectorized operations in RAM are orders of magnitude faster than repeatedly querying an external SQL server."
    )

    # Section 11: Sign-off
    h1 = doc.add_heading(level=1)
    r = h1.add_run("11. Technical Architecture Sign-off")
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
        ("Lead Software Architect", "Noman Zaid", "MCA 2nd Year Data Analytics", "APPROVED (Baselined)", "20-Nov-2026"),
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

    output_path = r"c:\Users\noman\OneDrive\Desktop\python\hcl-pro\docs\07_TRD.docx"
    doc.save(output_path)
    print(f"Successfully generated {output_path}")

if __name__ == "__main__":
    generate_trd()
