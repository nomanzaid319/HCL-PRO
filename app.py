"""
Insurance Claim Processing Analysis - Main Application Entrypoint & Executive Overview (UI-01)
Author: Noman Zaid (Roll No: 2500320140103, MCA 2nd Year Data Analytics, ABES)
Faculty Mentor: Ms. Shilpa
Linked Requirements: BR-01, BR-03, US-01, US-04, US-14, UI-01, TC-F-01, TC-F-02, TC-F-05
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from src.data_loader import load_master_claims_data, filter_claims_data
from src.analytics import compute_kpi_summary
import src.ui as ui

# Page configuration
st.set_page_config(
    page_title="Insurance Claim Analytics",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Load master claims dataset
df_master = load_master_claims_data()

# Global Sidebar Navigation & Filters
st.sidebar.image("https://img.icons8.com/fluency/96/shield.png", width=64)
st.sidebar.title("Claim Analysis Hub")
st.sidebar.markdown("**MCA Academic Project · ABES EC**\n*Noman Zaid (2500320140103)*")
st.sidebar.markdown("---")
st.sidebar.subheader("🔍 Global Filters")

# Policy Type Filter
all_policy_types = sorted(list(df_master["policy_type"].dropna().unique()))
if "selected_policy_types" not in st.session_state:
    st.session_state["selected_policy_types"] = all_policy_types

selected_policy_types = st.sidebar.multiselect(
    "Policy Type",
    options=all_policy_types,
    default=st.session_state["selected_policy_types"],
    help="Filter claims across Health, Motor, Life, and Property portfolios."
)
st.session_state["selected_policy_types"] = selected_policy_types

# Settlement Mode Filter
all_modes = sorted(list(df_master["settlement_mode"].dropna().unique()))
if "selected_modes" not in st.session_state:
    st.session_state["selected_modes"] = all_modes

selected_modes = st.sidebar.multiselect(
    "Settlement Mode",
    options=all_modes,
    default=st.session_state["selected_modes"],
    help="Toggle Cashless vs Reimbursement channel claims."
)
st.session_state["selected_modes"] = selected_modes

# Date Range Filter
min_date = df_master["submission_date"].min().date()
max_date = df_master["submission_date"].max().date()
if "date_range" not in st.session_state:
    st.session_state["date_range"] = (min_date, max_date)

date_range = st.sidebar.date_input(
    "Submission Period",
    value=st.session_state["date_range"],
    min_value=min_date,
    max_value=max_date,
    help="Filter by claim submission date."
)
if isinstance(date_range, tuple) and len(date_range) == 2:
    st.session_state["date_range"] = date_range

# Reset Filters Action
if st.sidebar.button("↺ Reset All Filters", use_container_width=True):
    st.session_state["selected_policy_types"] = all_policy_types
    st.session_state["selected_modes"] = all_modes
    st.session_state["date_range"] = (min_date, max_date)
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.caption("System v1.0 · SQLite Ingestion · Tailwind UI")

# Apply Filters
df = filter_claims_data(
    df_master,
    policy_types=st.session_state["selected_policy_types"],
    settlement_modes=st.session_state["selected_modes"],
    date_range=st.session_state["date_range"]
)

# Header Component
ui.render_header(
    title="Executive Claims Overview",
    subtitle="Enterprise operational scorecard, claims payout audit, and haircut analysis",
    badge_text="EXECUTIVE DASHBOARD · UI-01"
)

# Defensive Check for Filter Starvation
if df.empty:
    st.warning("⚠️ No claim records match the active filter selections. Please broaden your sidebar filters to inspect data.")
    st.stop()

# Compute KPI Metrics
kpis = compute_kpi_summary(df)

kpi_data = [
    {
        "title": "Claims Processed",
        "value": f"{kpis['total_claims']:,}",
        "subtext": f"{len(df[df['claim_status'] == 'Approved'])} Approved · {len(df[df['claim_status'] == 'Repudiated'])} Denied",
        "trend": "100% Inflow",
        "color": "blue"
    },
    {
        "title": "Net Settled Payout",
        "value": f"₹ {kpis['total_settled_cr']:.2f} Cr",
        "subtext": f"Haircut: {kpis['average_haircut']:.1f}% Avg Deduction",
        "trend": "Audit Compliant",
        "color": "emerald"
    },
    {
        "title": "Median Turnaround",
        "value": f"{kpis['median_tat']} Days",
        "subtext": f"Mean: {kpis['mean_tat']}d · SLA Breach: {kpis['sla_breach_rate']}%",
        "trend": "IRDAI Compliant",
        "color": "amber"
    },
    {
        "title": "Claim Approval Rate",
        "value": f"{kpis['approval_rate']:.1f}%",
        "subtext": f"Repudiation: {100 - kpis['approval_rate']:.1f}% Share",
        "trend": "Standard Benchmark",
        "color": "blue"
    }
]

ui.render_kpi_row(kpi_data)
st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

# Row 1: Haircut Scatter Plot & Adjudication Status Donut
col1, col2 = st.columns([7, 5])

with col1:
    st.subheader("Claim Haircut & Payout Variance")
    st.caption("Requested vs. Settled amounts. Points below the diagonal line represent claims subjected to tariff deductions.")
    
    # Plotly Haircut Scatter
    fig_scatter = px.scatter(
        df,
        x="claim_amount_requested",
        y="claim_amount_settled",
        color="policy_type",
        size="tat_days",
        hover_data=["claim_id", "policyholder_name", "provider_name", "haircut_pct"],
        labels={
            "claim_amount_requested": "Claim Amount Requested (₹)",
            "claim_amount_settled": "Claim Amount Settled (₹)",
            "policy_type": "Policy Portfolio",
            "tat_days": "TAT (Days)"
        },
        color_discrete_sequence=["#1B365D", "#2B6CB0", "#319795", "#D69E2E", "#E53E3E"]
    )
    
    # Add 1:1 diagonal reference line (y = x)
    max_val = max(df["claim_amount_requested"].max(), df["claim_amount_settled"].max())
    fig_scatter.add_trace(
        go.Scatter(
            x=[0, max_val],
            y=[0, max_val],
            mode="lines",
            name="Full Payout (Zero Haircut)",
            line=dict(color="#E53E3E", dash="dash", width=1.5)
        )
    )
    fig_scatter.update_layout(
        template="plotly_white",
        height=380,
        margin=dict(l=20, r=20, t=30, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    st.plotly_chart(fig_scatter, use_container_width=True)

with col2:
    st.subheader("Adjudication Status Distribution")
    st.caption("Proportion of claims approved vs repudiated across active filters.")
    
    status_counts = df["claim_status"].value_counts().reset_index()
    status_counts.columns = ["Status", "Count"]
    
    fig_donut = px.pie(
        status_counts,
        values="Count",
        names="Status",
        hole=0.55,
        color="Status",
        color_discrete_map={"Approved": "#276749", "Repudiated": "#9B2C2C"}
    )
    fig_donut.update_traces(
        textposition="inside",
        textinfo="percent+label",
        marker=dict(line=dict(color="#FFFFFF", width=2))
    )
    fig_donut.update_layout(
        template="plotly_white",
        height=380,
        margin=dict(l=20, r=20, t=30, b=20),
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5)
    )
    st.plotly_chart(fig_donut, use_container_width=True)

# Row 2: Monthly Volume & Payout Timeline
st.subheader("Monthly Submission Volume & Payout Trends")
st.caption("Temporal progression of total requested claim volume versus actual settled disbursement.")

monthly_agg = df.groupby("submission_month").agg(
    total_requested=("claim_amount_requested", "sum"),
    total_settled=("claim_amount_settled", "sum"),
    claim_count=("claim_id", "count")
).reset_index().sort_values("submission_month")

fig_area = go.Figure()
fig_area.add_trace(go.Scatter(
    x=monthly_agg["submission_month"],
    y=monthly_agg["total_requested"],
    mode="lines+markers",
    name="Requested Volume (₹)",
    line=dict(color="#3182CE", width=2.5),
    fill="tozeroy",
    fillcolor="rgba(49, 130, 206, 0.1)"
))
fig_area.add_trace(go.Scatter(
    x=monthly_agg["submission_month"],
    y=monthly_agg["total_settled"],
    mode="lines+markers",
    name="Settled Disbursement (₹)",
    line=dict(color="#2F855A", width=2.5),
    fill="tozeroy",
    fillcolor="rgba(47, 133, 90, 0.15)"
))
fig_area.update_layout(
    template="plotly_white",
    height=320,
    margin=dict(l=20, r=20, t=20, b=20),
    xaxis_title="Submission Month",
    yaxis_title="Amount (₹)",
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
)
st.plotly_chart(fig_area, use_container_width=True)

# Insight Callout
ui.render_insight_box(
    title="Portfolio Financial Exposure Summary",
    text=f"Out of {kpis['total_claims']} active claims analyzed, total settled disbursement is ₹{kpis['total_settled_cr']:.2f} Crores with an average haircut deduction of {kpis['average_haircut']:.1f}%. The median turnaround time is {kpis['median_tat']} days, well within the IRDAI 15-day regulatory SLA benchmark.",
    category="Executive Briefing"
)
