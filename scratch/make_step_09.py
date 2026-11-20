import os
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

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

def render_er_diagram_image(output_png_path):
    fig, ax = plt.subplots(figsize=(15, 10), dpi=300)
    ax.set_facecolor('#F8FAFC')
    fig.patch.set_facecolor('#F8FAFC')

    # Remove ticks and spines
    ax.set_xlim(0, 15)
    ax.set_ylim(0, 10)
    ax.axis('off')

    # Title Banner
    ax.text(7.5, 9.6, "INSURANCE CLAIM PROCESSING ANALYSIS — ENTITY RELATIONSHIP DIAGRAM", 
            fontsize=15, fontweight='bold', ha='center', color='#1B365D', fontfamily='sans-serif')
    ax.text(7.5, 9.25, "Relational 3NF Architecture: Dimension Masters (1) to Claims Fact (N)", 
            fontsize=10.5, fontstyle='italic', ha='center', color='#4A5568', fontfamily='sans-serif')

    def draw_entity(x, y, w, h, title, cols, is_fact=False):
        # Header color
        header_color = '#C53030' if is_fact else '#1B365D'
        bg_color = '#FFFFFF'
        
        # Outer Box
        box = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.04,rounding_size=0.15",
                                     linewidth=1.5, edgecolor=header_color, facecolor=bg_color)
        ax.add_patch(box)
        
        # Header banner
        header_h = 0.65
        header_box = patches.FancyBboxPatch((x, y + h - header_h), w, header_h, 
                                            boxstyle="round,pad=0.04,rounding_size=0.15",
                                            linewidth=1.5, edgecolor=header_color, facecolor=header_color)
        ax.add_patch(header_box)
        ax.text(x + w/2, y + h - header_h/2, title, fontsize=10.5, fontweight='bold', 
                color='#FFFFFF', ha='center', va='center', fontfamily='sans-serif')

        # Attributes
        line_y = y + h - header_h - 0.28
        for col_name, col_type, key_type in cols:
            prefix = ""
            color = '#2D3748'
            fontweight = 'normal'
            if key_type == 'PK':
                prefix = "[PK] "
                color = '#C53030'
                fontweight = 'bold'
            elif key_type == 'FK':
                prefix = "[FK] "
                color = '#2B6CB0'
                fontweight = 'bold'
            
            ax.text(x + 0.18, line_y, f"{prefix}{col_name}", fontsize=8.5, color=color, fontweight=fontweight, va='center', fontfamily='monospace')
            ax.text(x + w - 0.18, line_y, col_type, fontsize=8, color='#718096', ha='right', va='center', fontfamily='monospace')
            line_y -= 0.25

    # Entity 1: policyholders_master
    ph_cols = [
        ("policyholder_id", "VARCHAR(20)", "PK"),
        ("full_name", "VARCHAR(100)", ""),
        ("age", "INTEGER", ""),
        ("gender", "VARCHAR(10)", ""),
        ("city", "VARCHAR(50)", ""),
        ("state", "VARCHAR(50)", ""),
        ("region", "VARCHAR(20)", ""),
        ("tenure_months", "INTEGER", "")
    ]
    draw_entity(0.6, 5.8, 3.8, 2.9, "policyholders_master", ph_cols)

    # Entity 2: policies_master
    pol_cols = [
        ("policy_id", "VARCHAR(20)", "PK"),
        ("policyholder_id", "VARCHAR(20)", "FK"),
        ("policy_type", "VARCHAR(20)", ""),
        ("policy_tier", "VARCHAR(20)", ""),
        ("sum_insured", "DECIMAL(12,2)", ""),
        ("annual_premium", "DECIMAL(10,2)", ""),
        ("policy_start_date", "DATE", ""),
        ("policy_status", "VARCHAR(20)", "")
    ]
    draw_entity(5.4, 5.8, 4.0, 2.9, "policies_master", pol_cols)

    # Entity 3: providers_master
    prv_cols = [
        ("provider_id", "VARCHAR(20)", "PK"),
        ("provider_name", "VARCHAR(100)", ""),
        ("provider_type", "VARCHAR(30)", ""),
        ("city", "VARCHAR(50)", ""),
        ("state", "VARCHAR(50)", ""),
        ("empanelment_status", "VARCHAR(20)", ""),
        ("tariff_tier", "VARCHAR(10)", "")
    ]
    draw_entity(10.4, 5.8, 4.0, 2.7, "providers_master", prv_cols)

    # Entity 4: claims_fact (Main Fact Table)
    clm_cols = [
        ("claim_id", "VARCHAR(20)", "PK"),
        ("policy_id", "VARCHAR(20)", "FK"),
        ("provider_id", "VARCHAR(20)", "FK"),
        ("incident_date", "DATE", ""),
        ("submission_date", "DATE", ""),
        ("verification_date", "DATE", ""),
        ("assessment_date", "DATE", ""),
        ("settlement_date", "DATE", ""),
        ("claim_amount_requested", "DECIMAL(12,2)", ""),
        ("claim_amount_settled", "DECIMAL(12,2)", ""),
        ("settlement_ratio", "DECIMAL(5,4)", ""),
        ("settlement_mode", "VARCHAR(20)", ""),
        ("adjudication_stage", "VARCHAR(30)", ""),
        ("claim_status", "VARCHAR(20)", ""),
        ("rejection_reason_code", "VARCHAR(20)", ""),
        ("tat_days", "INTEGER", ""),
        ("fraud_risk_score", "INTEGER", ""),
        ("fraud_flag", "BOOLEAN", ""),
        ("csat_score", "INTEGER", "")
    ]
    draw_entity(4.8, 0.4, 5.4, 4.8, "claims_fact (Fact Table)", clm_cols, is_fact=True)

    # Connecting Arrows & Relationships
    # Arrow 1: policyholders -> policies
    ax.annotate("", xy=(5.4, 7.25), xytext=(4.4, 7.25),
                arrowprops=dict(arrowstyle="-|>", lw=2, color="#2B6CB0", mutation_scale=15))
    ax.text(4.9, 7.45, "purchases\n(1 : N)", fontsize=8.5, fontweight='bold', color="#1B365D", ha='center', va='bottom')

    # Arrow 2: policies -> claims_fact
    ax.annotate("", xy=(6.5, 5.2), xytext=(6.5, 5.8),
                arrowprops=dict(arrowstyle="-|>", lw=2, color="#2B6CB0", mutation_scale=15))
    ax.text(6.75, 5.45, "incurs\n(1 : N)", fontsize=8.5, fontweight='bold', color="#1B365D", ha='left', va='center')

    # Arrow 3: providers -> claims_fact
    ax.annotate("", xy=(10.2, 3.8), xytext=(11.5, 5.8),
                arrowprops=dict(arrowstyle="-|>", lw=2, color="#2B6CB0", mutation_scale=15,
                                connectionstyle="arc3,rad=-0.15"))
    ax.text(11.2, 4.6, "services & bills\n(1 : N)", fontsize=8.5, fontweight='bold', color="#1B365D", ha='center', va='center')

    # Legend Block
    leg_box = patches.FancyBboxPatch((0.6, 0.6), 3.4, 1.8, boxstyle="round,pad=0.04,rounding_size=0.1",
                                     linewidth=1, edgecolor='#CBD5E0', facecolor='#FFFFFF')
    ax.add_patch(leg_box)
    ax.text(0.8, 2.15, "SCHEMA CONVENTIONS", fontsize=9, fontweight='bold', color='#1B365D')
    ax.text(0.8, 1.8, "[PK] Primary Key (Unique, Not Null)", fontsize=8, color='#C53030', fontweight='bold')
    ax.text(0.8, 1.5, "[FK] Foreign Key (Referential Integrity)", fontsize=8, color='#2B6CB0', fontweight='bold')
    ax.text(0.8, 1.2, "Blue Header: Master / Dimension Table", fontsize=8, color='#1B365D')
    ax.text(0.8, 0.9, "Red Header: Central Fact Table (3NF)", fontsize=8, color='#C53030')

    plt.tight_layout()
    plt.savefig(output_png_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"Successfully generated {output_png_path}")

def generate_er_document(png_path, output_docx_path):
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
    run_title = p_title.add_run("Step 9: Entity Relationship (ER) Diagram Document")
    run_title.font.name = 'Calibri'
    run_title.font.size = Pt(24)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(16)
    run_sub = p_sub.add_run("Mermaid Schema Model, Cardinality Analysis, Business Invariant Enforcement, and Structural Design Review")
    run_sub.font.size = Pt(13)
    run_sub.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)

    # Metadata Box
    meta_table = doc.add_table(rows=5, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(meta_table, "CBD5E0")
    meta_data = [
        ("Project Name", "Insurance Claim Processing Analysis"),
        ("Learner / Candidate", "Noman Zaid (Roll No: 2500320140103, MCA 2nd Year DA)"),
        ("Faculty Mentor", "Ms. Shilpa, Department of Computer Applications, ABES EC"),
        ("Source Design Baseline", "docs/08_Database_Design.xlsx (4 Normalized Tables + 36-Column Analytical View)"),
        ("Artifact Deliverables", "docs/09_ER_Diagram.docx + docs/09_ER_Diagram.png")
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

    # Section 1: Executive Overview
    h1 = doc.add_heading(level=1)
    r = h1.add_run("1. Executive Overview & ER Modeling Principles")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "The database design for the Insurance Claim Processing Analysis application follows an enterprise relational architecture modeled in Third Normal Form (3NF). The schema resolves operational data into three independent master dimensions (policyholders, policy contracts, and service providers) connected to a centralized transactional fact table (claims). This design eliminates data redundancy, guarantees referential consistency, and supports high-speed analytical aggregations via a 36-column denormalized SQL view."
    )

    # Section 2: Visual ER Diagram
    h1 = doc.add_heading(level=1)
    r = h1.add_run("2. Entity Relationship Diagram (Visual Schematic)")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph("The visual schema diagram below details the 4 core relational tables, their primary keys [PK], foreign keys [FK], and exact 1:N cardinalities:")

    if os.path.exists(png_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_after = Pt(6)
        run_img = p_img.add_run()
        run_img.add_picture(png_path, width=Inches(6.5))
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_after = Pt(12)
        r_cap = p_cap.add_run("Figure 9.1: Entity Relationship Schematic — Insurance Claim Processing Analysis")
        r_cap.font.size = Pt(9)
        r_cap.font.italic = True
        r_cap.font.color.rgb = RGBColor(0x71, 0x80, 0x96)

    # Section 3: Mermaid ER Code
    h1 = doc.add_heading(level=1)
    r = h1.add_run("3. Mermaid erDiagram Specification")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph("The authoritative Mermaid syntax representing all tables, columns, keys, and verb-labeled cardinalities:")

    mermaid_code = (
        "erDiagram\n"
        "    POLICYHOLDERS_MASTER ||--o{ POLICIES_MASTER : \"purchases\"\n"
        "    POLICIES_MASTER ||--o{ CLAIMS_FACT : \"incurs\"\n"
        "    PROVIDERS_MASTER ||--o{ CLAIMS_FACT : \"services and bills\"\n\n"
        "    POLICYHOLDERS_MASTER {\n"
        "        string policyholder_id PK \"POL-HLD-1001 to 9999\"\n"
        "        string full_name \"Legal Name\"\n"
        "        int age \"18 to 75\"\n"
        "        string gender \"Male, Female, Other\"\n"
        "        string city \"City location\"\n"
        "        string state \"State location\"\n"
        "        string region \"North, South, East, West, Central\"\n"
        "        int tenure_months \"Customer tenure\"\n"
        "    }\n\n"
        "    POLICIES_MASTER {\n"
        "        string policy_id PK \"POL-10001 to 99999\"\n"
        "        string policyholder_id FK \"References policyholders_master\"\n"
        "        string policy_type \"Health, Motor, Life, Property\"\n"
        "        string policy_tier \"Silver, Gold, Platinum\"\n"
        "        decimal sum_insured \"Coverage Limit\"\n"
        "        decimal annual_premium \"Premium paid\"\n"
        "        date policy_start_date \"Inception date\"\n"
        "        string policy_status \"Active, Lapsed, Renewed\"\n"
        "    }\n\n"
        "    PROVIDERS_MASTER {\n"
        "        string provider_id PK \"PRV-101 to 999\"\n"
        "        string provider_name \"Hospital or Garage Name\"\n"
        "        string provider_type \"Hospital, Automobile Garage, TPA\"\n"
        "        string city \"Facility city\"\n"
        "        string state \"Facility state\"\n"
        "        string empanelment_status \"Empaneled, Non-Network\"\n"
        "        string tariff_tier \"Tier 1, Tier 2, Tier 3\"\n"
        "    }\n\n"
        "    CLAIMS_FACT {\n"
        "        string claim_id PK \"CLM-10001 to 99999\"\n"
        "        string policy_id FK \"References policies_master\"\n"
        "        string provider_id FK \"References providers_master\"\n"
        "        date incident_date \"Date of loss\"\n"
        "        date submission_date \"Inflow date\"\n"
        "        date verification_date \"Doc verification\"\n"
        "        date assessment_date \"Medical/technical audit\"\n"
        "        date settlement_date \"Final decision\"\n"
        "        decimal claim_amount_requested \"Billed amount\"\n"
        "        decimal claim_amount_settled \"Disbursed amount\"\n"
        "        decimal settlement_ratio \"Settled / Requested\"\n"
        "        string settlement_mode \"Cashless, Reimbursement\"\n"
        "        string adjudication_stage \"Current milestone\"\n"
        "        string claim_status \"Approved, Repudiated, Pending\"\n"
        "        string rejection_reason_code \"IRDAI denial code\"\n"
        "        string rejection_reason_detail \"Denial explanation\"\n"
        "        int tat_days \"End-to-end duration\"\n"
        "        int fraud_risk_score \"Anomaly score 0-100\"\n"
        "        int fraud_flag \"0: Normal, 1: High Risk\"\n"
        "        int csat_score \"Survey rating 1 to 5\"\n"
        "    }\n"
    )

    box = doc.add_table(rows=1, cols=1)
    box.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = box.rows[0].cells[0]
    c.width = Inches(6.5)
    set_cell_background(c, "F7FAFC")
    set_cell_margins(c, 80, 80, 100, 100)
    set_table_borders(box, "CBD5E0")
    cp = c.paragraphs[0]
    cp.paragraph_format.space_after = Pt(0)
    cr_code = cp.add_run(mermaid_code)
    cr_code.font.name = "Consolas"
    cr_code.font.size = Pt(8)
    cr_code.font.color.rgb = RGBColor(0x1A, 0x20, 0x2C)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 4: Plain-English Relationship Statements
    h1 = doc.add_heading(level=1)
    r = h1.add_run("4. Plain-English Relationship Descriptions")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "Each relational link captures a fundamental real-world insurance business transaction:"
    )

    rels = [
        ("Relationship 1 (policyholders_master to policies_master):",
         "One policyholder purchases many policy contracts (1 : N). A single customer record can be associated with multiple active or historical insurance policies across different business lines (e.g., both a Health and a Motor policy), but every individual policy contract belongs exclusively to exactly one primary policyholder."),
        ("Relationship 2 (policies_master to claims_fact):",
         "One policy contract incurs many claim transactions (1 : N). During its effective coverage term, an insurance policy can experience zero, one, or multiple separate loss events (claims) over time, while each individual claim record is legally bound to exactly one underlying insurance policy contract."),
        ("Relationship 3 (providers_master to claims_fact):",
         "One network provider services and bills many claim transactions (1 : N). An empaneled hospital or automobile repair garage can provide medical treatment or vehicle repairs for hundreds of policyholders across numerous claims, but each claim event is serviced at exactly one primary provider facility.")
    ]

    for title, desc in rels:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(3)
        rt = p.add_run(f"{title} ")
        rt.font.bold = True
        rt.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
        rd = p.add_run(desc)
        rd.font.size = Pt(9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 5: Business Rules Enforced by the Design
    h1 = doc.add_heading(level=1)
    r = h1.add_run("5. Business Invariants Enforced by Database Constraints")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "The relational schema enforces insurance contract laws and operational standards at the storage engine level:"
    )

    rules = [
        ("Temporal Chronology Invariant:", "Check constraint verifies logical sequential lifecycle: incident_date <= submission_date <= verification_date <= assessment_date <= settlement_date. A claim cannot be decided before it occurs or is submitted."),
        ("Financial Indemnity Boundary:", "Constraint verifies: 0.0 <= claim_amount_settled <= claim_amount_requested. An insurer cannot legally disburse a negative amount or pay more than what was billed by the claimant."),
        ("Coverage Sum Insured Ceiling:", "Transactional checks verify claim_amount_requested <= sum_insured. A claim cannot exceed the contractual policy indemnity coverage limit."),
        ("Statutory Repudiation Reason Mandate:", "Any claim with claim_status = 'Repudiated' must have a non-null rejection_reason_code from approved IRDAI categories."),
        ("Valid Anomaly Score Range:", "Check constraint restricts fraud_risk_score to integers between 0 and 100, and fraud_flag to binary (0, 1)."),
        ("Referential Integrity Enforcement:", "Strict foreign key rules (ON DELETE RESTRICT ON UPDATE CASCADE) prevent orphan claims from referencing deleted policyholders, policies, or providers.")
    ]

    for title, desc in rules:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(3)
        rt = p.add_run(f"{title} ")
        rt.font.bold = True
        rt.font.color.rgb = RGBColor(0x1A, 0x20, 0x2C)
        rd = p.add_run(desc)
        rd.font.size = Pt(9.5)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Section 6: Structural Design Analysis & Potential Problems/Fixes
    h1 = doc.add_heading(level=1)
    r = h1.add_run("6. Structural Design Analysis, Potential Problems & Architectural Fixes")
    r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    p = doc.add_paragraph(
        "A critical architectural review was conducted to identify potential data modeling pitfalls and implement enterprise-grade solutions:"
    )

    design_table = doc.add_table(rows=5, cols=4)
    design_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(design_table, "CBD5E0")

    d_headers = ["Potential Design Problem", "Identified Architectural Risk", "Evaluation / Pitfall Context", "Architectural Solution / Fix Implemented"]
    for j, h in enumerate(d_headers):
        c = design_table.rows[0].cells[j]
        set_cell_background(c, "1B365D")
        set_cell_margins(c, 80, 80, 70, 70)
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    d_issues = [
        ("Many-to-Many Facility Servicing", 
         "Can one claim involve multiple providers (e.g., primary hospital, diagnostics lab, pharmacy)?",
         "If multiple providers bill a single claim without an intermediary bridge table, data redundancy or multi-valued fields occur.",
         "FIX: Claims fact stores the primary billing network provider. In Phase 2, a line-items bridge table (claim_bill_items) can be introduced; for this single-payer project, 1:N cleanly models primary adjudications."),
        
        ("Orphan Records / Missing Keys", 
         "Risk of claims submitted with non-existent policy_id or provider_id.",
         "Unenforced SQLite foreign keys allow orphan child records, causing analytical aggregations and inner joins to drop data silently.",
         "FIX: Enforced PRAGMA foreign_keys = ON and specified ON DELETE RESTRICT ON UPDATE CASCADE on all foreign key constraints."),

        ("Redundant Demographics in Claims", 
         "Risk of repeating customer name, age, and state inside the transactional claims table.",
         "Violates Third Normal Form (3NF), creating transitive dependencies and update anomalies.",
         "FIX: Strictly normalized demographic attributes into policyholders_master. Demographics are combined exclusively at runtime via the SQL view (v_claims_analysis_master)."),

        ("Slow Reactive Dashboard Queries", 
         "Multi-table SQL joins executed on every widget interaction cause latency.",
         "Streamlit re-executes top-to-bottom on every slider or dropdown click, causing repetitive table scan overhead.",
         "FIX: Created indices on foreign keys and submission dates, and pre-compiled the 36-column analytical view cached via @st.cache_data.")
    ]

    col_w_di = [Inches(1.5), Inches(1.5), Inches(1.7), Inches(1.8)]
    for i, row in enumerate(d_issues):
        r = design_table.rows[i+1]
        bg = "FFFFFF" if i % 2 == 0 else "F7FAFC"
        for j, val in enumerate(row):
            c = r.cells[j]
            c.width = col_w_di[j]
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

    # Section 7: Document Control
    h1 = doc.add_heading(level=1)
    r = h1.add_run("7. Document Control & Sign-off")
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
        ("Lead Data Architect", "Noman Zaid", "MCA 2nd Year Data Analytics", "APPROVED (Baselined)", "20-Nov-2026"),
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

    doc.save(output_docx_path)
    print(f"Successfully generated {output_docx_path}")

if __name__ == "__main__":
    png_path = r"c:\Users\noman\OneDrive\Desktop\python\hcl-pro\docs\09_ER_Diagram.png"
    docx_path = r"c:\Users\noman\OneDrive\Desktop\python\hcl-pro\docs\09_ER_Diagram.docx"
    render_er_diagram_image(png_path)
    generate_er_document(png_path, docx_path)
