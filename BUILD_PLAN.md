# System Build Plan & Engineering Implementation Roadmap
**Project:** Insurance Claim Processing Analysis  
**Academic Affiliation:** MCA 2nd Year (Data Analytics), ABES Engineering College  
**Author:** Noman Zaid (Roll No: 2500320140103)  
**Faculty Mentor:** Ms. Shilpa  
**Date:** November 20, 2026  
**Artifact ID:** STEP-14 (`BUILD_PLAN.md` & `docs/14_BUILD_PLAN.docx`)

---

## 1. Executive Summary & Build Strategy

### 1.1 Objective
The purpose of this build plan is to provide a complete, rigorous, and phased software engineering blueprint for developing the **Insurance Claim Processing Analysis System**. Following the completion of the Test Cases Specification and QA Test Suite (Step 13), this document translates the functional requirements (FR-01 to FR-14), user stories (US-01 to US-14), database schemas (Step 8), UI design system (UI-01 to UI-07), and test suites (TC-F-01..14, TC-T-01..14) into concrete, modular implementation tasks.

### 1.2 Technology Stack & Libraries
| Component | Technology | Version / Specification | Justification |
| :--- | :--- | :--- | :--- |
| **Language** | Python | 3.12+ | Native support for modern typing, dataclasses, and performance |
| **Web Framework** | Streamlit | 1.40+ | Native multi-page application support (`pages/`), reactive caching |
| **Component UI** | Tailwind CSS CDN | 3.4+ via iframe | Modern typography, gradient cards, flexbox layouts, corporate styling |
| **Visualization** | Plotly Express / Graph Objects | 5.24+ | Interactive client-side charting (zoom, hover tooltips, responsive reflow) |
| **Statistical Engine** | SciPy & NumPy | 1.14+ / 2.0+ | Welch's two-sample t-test, Tukey IQR outlier fences, correlation matrix |
| **Data Engine** | Pandas & SQLite3 | 2.2+ / SQLite 3.45+ | Fast tabular manipulations, SQL joins, deterministic relational querying |
| **Testing Engine** | Pytest | 9.2+ | Automated unit tests, invariant checks, relational integrity validation |

### 1.3 Key Architectural Principles
1. **Layered Decoupling (MVC Pattern):**
   - **Data Layer (`src/data_loader.py`):** Encapsulates SQLite querying, `@st.cache_data` caching, data cleaning pipelines, and dimension joins.
   - **Analytical Engine (`src/analytics.py`):** Pure Python mathematical functions with zero Streamlit UI dependencies, ensuring testability via pytest.
   - **UI & Presentation Layer (`src/ui.py`):** Reusable HTML/Tailwind CSS components (headers, KPI cards, insight callouts, badge ribbons).
   - **Application Pages (`app.py`, `pages/*.py`):** Streamlit page controllers orchestrating filters, analytics, charts, and tables.
2. **Deterministic State Management:**
   - Global filters (Date range, Policy Type, Settlement Mode) reside in `st.sidebar` and sync across multipage views via `st.session_state`.
   - Clear and graceful handling of empty query results with user-friendly warning messages.
3. **Strict Adherence to Quality Gates:**
   - 100% pass rate across the 14 automated pytest test cases (`TC-T-01` to `TC-T-14`).
   - Every single functional requirement (`FR-01` to `FR-14`) and business requirement (`BR-01` to `BR-12`) is verified against its corresponding acceptance test.

---

## 2. Target File & Module Structure

```
hcl-pro/
├── app.py                          # UI-01: Executive Portfolio Overview & Navigation Hub
├── pages/
│   ├── 1_TAT_Analytics.py         # UI-02: Turnaround Time, Milestone Delays & Surge Analysis
│   ├── 2_Denial_Analysis.py        # UI-03: Repudiation Root Cause, Pareto 80/20 & Regional Heatmap
│   ├── 3_Providers.py              # UI-04: Network Hospital & Auto Garage Benchmarking
│   ├── 4_Fraud_Screening.py        # UI-05: Anomaly Triage Queue & Early-Inception Exposure
│   ├── 5_Underwriting.py           # UI-06: Welch's t-test, Catastrophic Outliers & Correlations
│   └── 6_Data_Quality.py           # UI-07: ETL Transformation Audits, Null Handling & File Exports
├── src/
│   ├── __init__.py
│   ├── data_loader.py              # SQLite DB connection, @st.cache_data, joins & filter queries
│   ├── analytics.py               # Pure math functions: TAT, Pareto, Haircut, Welch's t-test, Fraud score
│   ├── ui.py                      # Tailwind CSS iframe components, KPI cards, banners, insight callouts
│   └── generate_data.py           # Synthetic data generation script (Step 10 baseline)
├── data/
│   ├── app.db                     # SQLite database with 4 relational tables (claims, policies, policyholders, providers)
│   ├── claims_fact.csv
│   ├── policies_master.csv
│   ├── policyholders_master.csv
│   ├── providers_master.csv
│   └── insurance_claims_data.xlsx
├── tests/
│   ├── test_data.py               # TC-T-01..06: Database schema, row counts, PK/FK & data cleaner tests
│   ├── test_analysis.py           # TC-T-07..11: TAT KPI, Pareto math, Haircut %, Welch's t-test, Fraud scoring
│   ├── test_app.py                # TC-T-12..14: UI module imports, config.toml & data artifact smoke tests
│   └── __init__.py
├── .streamlit/
│   └── config.toml                # Corporate navy theme (#1B365D) and headless server config
├── requirements.txt               # Pinned Python package dependencies
├── CLAUDE.md                      # Academic guidelines and 18-step execution rules
└── BUILD_PLAN.md                  # This implementation roadmap
```

---

## 3. Phased Implementation Roadmap

### Phase 1: Core Foundation & Analytical Engine (`src/`)
* **Goal:** Build the backend data connector, caching layer, and purely functional analytics engine.
* **Component 1.1: `src/data_loader.py`**
  - Implement `get_database_connection()` pointing to `data/app.db`.
  - Implement `@st.cache_data(ttl=3600) def load_master_claims_data()`:
    - Executes SQL joining `claims_fact` (c), `policies_master` (p), `policyholders_master` (h), and `providers_master` (pr).
    - Computes derived attributes: `tat_days = (settlement_date - submission_date).days`, `haircut_pct = ((claim_amount_requested - claim_amount_settled) / claim_amount_requested) * 100`, `days_from_inception = (claim_date - policy_start_date).days`.
    - Handles null values for repudiated/pending claims gracefully without throwing NaT/NaN runtime errors.
  - Implement filter utility: `filter_claims(df, policy_types, settlement_modes, date_range, region)`.
* **Component 1.2: `src/analytics.py`**
  - `compute_kpi_summary(df)`: Calculates Total Claims, Total Settled Amount (₹), Mean TAT, Median TAT, Approval Rate (%), and Net Haircut (%).
  - `compute_pareto_repudiation(df)`: Aggregates claim rejections by denial reason, computes frequency, cumulative sums, and percentage thresholds (verifying the 80/20 rule).
  - `perform_welchs_ttest(cashless_tat, reimbursement_tat)`: Runs `scipy.stats.ttest_ind(..., equal_var=False)`, returning sample means, standard errors, degrees of freedom, t-statistic, and p-value.
  - `calculate_fraud_score(row)`: Applies rule-based scoring (early inception <= 30 days: +45 pts, tariff deviation > 20%: +30 pts, repeated claim: +25 pts), isolating scores >= 70 as High Risk.
  - `calculate_tukey_fences(series, k=1.5)`: Returns `Q1`, `Q3`, `IQR`, Lower Bound, and Upper Bound (`Q3 + 1.5 * IQR`) to isolate catastrophic loss claims.
* **Verification Gate:**
  - Execute `pytest tests/test_data.py tests/test_analysis.py -v`.
  - All 11 unit tests must pass.

---

### Phase 2: Executive Overview & Multi-Page Shell (`app.py`, `src/ui.py`)
* **Goal:** Create the main application entry point, sidebar filter controls, corporate branding, and UI-01 Executive Dashboard.
* **Component 2.1: Multi-Page Configuration & Session State**
  - Set `st.set_page_config(page_title="Insurance Claim Processing Analysis", page_icon="🛡️", layout="wide", initial_sidebar_state="expanded")`.
  - Implement persistent session state keys in `st.sidebar`:
    - Policy Type Multiselect (`Health`, `Motor`, `Life`).
    - Settlement Mode Multiselect (`Cashless`, `Reimbursement`).
    - Claim Submission Date Range (`date_input`).
* **Component 2.2: UI-01 Executive Overview Screen**
  - Header banner via `ui.render_header("Executive Claims Overview", "Portfolio-wide operational scorecard and claim haircut monitoring")`.
  - 4 Tailwind KPI Cards via `ui.render_kpi_row()`:
    1. Total Claims Processed (Count & YoY Trend)
    2. Total Net Payout (₹ Crores & Haircut Deduction %)
    3. Median Processing TAT (Days & SLA Compliance %)
    4. Claim Approval Rate (% & Repudiation Share)
  - Visualization 1: Claim Haircut Scatter Plot (`Requested Amount` vs `Settled Amount` with diagonal equality line `y = x` and color-coded haircut severity).
  - Visualization 2: Claim Adjudication Status Donut Chart (Approved, Rejected, Under Review, Escalated).
  - Visualization 3: Monthly Inflow vs Settled Volume Area Chart.
* **Verification Gate:**
  - Verify UI-01 renders without console errors.
  - Test sidebar filter interactions: changing policy type immediately updates KPI cards and Plotly charts within 1.0 second.

---

### Phase 3: Operational Cycle Time & Provider Benchmarking (`pages/1_`, `pages/3_`)
* **Goal:** Implement deep-dive operational cycle time analytics (UI-02) and empaneled provider benchmarking (UI-04).
* **Component 3.1: `pages/1_TAT_Analytics.py` (UI-02)**
  - Header & Metric ribbon: Mean TAT, Median TAT, SLA Breach Count (>15 days), 90th Percentile TAT.
  - Visualization 1: Milestone Cycle Time Breakdown (Stacked horizontal bar chart comparing Document Review, Medical Assessment, and Financial Settlement across Policy Types).
  - Visualization 2: Distribution Comparison (Violin / Box Plot comparing Cashless vs Reimbursement turnaround distributions).
  - Visualization 3: 30-Day Rolling Moving Average of Daily Submissions (highlighting Q3 monsoon surge).
* **Component 3.2: `pages/3_Providers.py` (UI-04)**
  - Header & KPI ribbon: Active Providers, Average Cost per Admission/Repair, Flagged Tariff Outliers, Mean Provider TAT.
  - Visualization 1: Cost vs. TAT Performance Quadrant (Scatter plot mapping Provider Average Cost vs Provider Average TAT with quadrant benchmark crosshairs).
  - Visualization 2: Top 10 High-Cost Provider Leaderboard with tariff audit badges (z-score > 2.0).
  - Tabular Provider Audit Directory with interactive search and accreditation filter (NABH/JCI vs Non-accredited).
* **Verification Gate:**
  - Verify that milestone breakdown sums to total turnaround time.
  - Verify that Apex Multispeciality (PRV-104) is correctly flagged with a red audit badge in the provider quadrant.

---

### Phase 4: Repudiation, Fraud & Underwriting Analytics (`pages/2_`, `pages/4_`, `pages/5_`)
* **Goal:** Deliver actionable risk mitigation screens covering claim denials (UI-03), fraud screening (UI-05), and actuarial hypothesis testing (UI-06).
* **Component 4.1: `pages/2_Denial_Analysis.py` (UI-03)**
  - Header & Metric ribbon: Repudiation Count, Repudiation Rate (%), Total Denied Exposure (₹), Average Grievance Score.
  - Visualization 1: Dual-Axis Pareto 80/20 Chart (Bar chart of rejection reasons sorted descending + cumulative percentage line with 80% benchmark).
  - Visualization 2: Regional Repudiation & Grievance Heatmap (State-level denial rate and CSAT metrics).
  - Detailed Rejection Drill-down Table with IRDAI dispute classifications.
* **Component 4.2: `pages/4_Fraud_Screening.py` (UI-05)**
  - Interactive Risk Threshold Slider (`st.slider("High Risk Score Threshold", 50, 95, 70)`).
  - Metric ribbon: Flagged High-Risk Claims, Total At-Risk Financial Exposure (₹ Lakhs), Early Inception Ratio (<30 days).
  - High-Risk Triage Queue: Filtered data table displaying suspicious claims with risk score, anomaly flags, and recommended action (Field Investigation / Audit).
  - Visualization 1: Anomaly Type Breakdown Pie Chart.
  - Visualization 2: Risk Score Distribution Histogram.
* **Component 4.3: `pages/5_Underwriting.py` (UI-06)**
  - Actuarial Hypothesis Testing Card:
    - Displays two-sample Welch's t-test results: Cashless vs Reimbursement TAT.
    - Shows t-stat, p-value (< 0.001), 95% Confidence Interval, and formal conclusion rejecting the null hypothesis ($H_0$).
  - Visualization 1: Catastrophic Claim Outlier Detection (Box plot with Tukey upper fence isolating extreme losses > ₹20 Lakhs).
  - Visualization 2: Pearson Correlation Matrix Heatmap (Age, Sum Insured, Claim Amount, Premium, TAT).
  - Visualization 3: 4-Variable Bubble Chart (Age vs Premium vs Claim Amount sized by Sum Insured).
* **Verification Gate:**
  - Confirm Pareto 80/20 correctly shows `POLICY_EXCLUSION` and `EXCL_PRE_EXIST` accounting for ~80% of denials.
  - Verify Welch's t-test output matches pytest benchmark (`p < 0.001`).

---

### Phase 5: Data Quality, Auditing & Export (`pages/6_Data_Quality.py`)
* **Goal:** Implement the data provenance screen (UI-07) allowing data governance validation and filtered artifact exports.
* **Component 5.1: `pages/6_Data_Quality.py` (UI-07)**
  - Data Governance Scorecard: Completeness (99.8%), Uniqueness (100%), Referential Integrity (100%), Valid Dates (100%).
  - Pre- vs Post-Cleaning Audit Log: Details deduction cleaning, date parsing, and schema type casting.
  - Relational Schema Table Browser: Interactive tabbed view of `claims_fact`, `policies_master`, `policyholders_master`, and `providers_master`.
  - Export Utility:
    - `st.download_button` for Filtered Claims Data (`.csv`).
    - `st.download_button` for Filtered Claims Data (`.xlsx`).
* **Verification Gate:**
  - Verify file downloads produce valid CSV and Excel files.
  - Verify that no sensitive PII is exposed.

---

### Phase 6: System Integration, Test Verification & Hardening
* **Goal:** Perform full regression testing, smoke test all routes, verify edge cases, and ensure clean production readiness.
* **Tasks:**
  1. Run complete automated test suite (`pytest -v`).
  2. Verify all 14 Functional Test Cases (`TC-F-01` to `TC-F-14`).
  3. Validate performance SLAs: Initial page load < 2.0s, filtered reflow < 1.0s.
  4. Test edge cases: Empty filter selections, extreme date ranges, missing provider records.
  5. Code formatting and cleanup (PEP8 compliance, docstrings, type annotations).

---

## 4. Requirements & Test Traceability Matrix (Build Phase)

| Business Req | Functional Req | User Story | UI Screen | Target Code Module | Pytest Verification | Functional Test |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **BR-01** | FR-01, FR-02 | US-01, US-14 | UI-01, UI-02 | `app.py`, `pages/1_TAT_Analytics.py` | `TC-T-07`, `TC-T-12` | `TC-F-01`, `TC-F-02`, `TC-F-03` |
| **BR-02** | FR-03, FR-04 | US-03 | UI-03 | `pages/2_Denial_Analysis.py` | `TC-T-08` | `TC-F-04` |
| **BR-03** | FR-05 | US-04 | UI-01 | `app.py`, `src/analytics.py` | `TC-T-09` | `TC-F-05` |
| **BR-04** | FR-06 | US-05 | UI-04 | `pages/3_Providers.py` | `TC-T-02`, `TC-T-04` | `TC-F-06` |
| **BR-05** | FR-07 | US-06 | UI-05 | `pages/4_Fraud_Screening.py` | `TC-T-11` | `TC-F-07` |
| **BR-06** | FR-08 | US-07 | UI-02 | `pages/1_TAT_Analytics.py` | `TC-T-07` | `TC-F-08` |
| **BR-07** | FR-09 | US-08 | UI-06 | `pages/5_Underwriting.py` | `TC-T-02` | `TC-F-09` |
| **BR-08** | FR-10 | US-09 | UI-06 | `pages/5_Underwriting.py`, `src/analytics.py` | `TC-T-10` | `TC-F-10` |
| **BR-09** | FR-11 | US-10 | UI-03 | `pages/2_Denial_Analysis.py` | `TC-T-02` | `TC-F-11` |
| **BR-10** | FR-12 | US-11 | UI-06 | `pages/5_Underwriting.py` | `TC-T-02` | `TC-F-12` |
| **BR-11** | FR-13 | US-12 | UI-07 | `pages/6_Data_Quality.py`, `src/data_loader.py` | `TC-T-01..06` | `TC-F-13` |
| **BR-12** | FR-14 | US-13 | UI-07 | `pages/6_Data_Quality.py` | `TC-T-14` | `TC-F-14` |

---

## 5. Risk Assessment, Fallbacks & Contingency Plan

| Risk Event | Severity | Probability | Impact | Mitigation Strategy |
| :--- | :---: | :---: | :--- | :--- |
| **Data Filter Starvation** (Filter returns 0 rows) | Medium | Medium | App renders empty/broken Plotly charts or crashes | Wrap chart functions in `if df.empty:` check; display `st.warning("No records match the active filter criteria. Please broaden your selections.")`. |
| **Streamlit Page Switching State Loss** | Medium | Low | User selections reset when navigating between pages | Store all filter widgets in `st.session_state` keys (`selected_policy_types`, etc.) rather than local page variables. |
| **Plotly CDN / CDN Tailwind Lag** | Low | Low | Blank styling or delayed rendering on slow networks | Tailwind is embedded via lightweight HTML iframe; Plotly uses offline Python bundled JS (`plotly.graph_objects`). |
| **SQLite Concurrency Lock** | Low | Very Low | Streamlit re-runs trigger database lock | Database access is read-only (`query = "SELECT ..."`) wrapped with `@st.cache_data`. |

---

## 6. Execution Checklist for Step 15

- [ ] Build `src/data_loader.py` with SQL join and `@st.cache_data` caching
- [ ] Build `src/analytics.py` with pure mathematical calculation routines
- [ ] Implement `app.py` (Executive Overview UI-01) with sidebar filters and Tailwind KPI cards
- [ ] Implement `pages/1_TAT_Analytics.py` (UI-02) with milestone cycle time breakdown
- [ ] Implement `pages/2_Denial_Analysis.py` (UI-03) with dual-axis Pareto 80/20 chart
- [ ] Implement `pages/3_Providers.py` (UI-04) with cost vs TAT quadrant and audit badges
- [ ] Implement `pages/4_Fraud_Screening.py` (UI-05) with interactive risk slider and triage table
- [ ] Implement `pages/5_Underwriting.py` (UI-06) with Welch's t-test card and Tukey outlier fences
- [ ] Implement `pages/6_Data_Quality.py` (UI-07) with health scorecard and CSV/Excel exports
- [ ] Run full automated test suite (`pytest -v`) ensuring 100% pass rate
- [ ] Conduct multi-page browser smoke test
