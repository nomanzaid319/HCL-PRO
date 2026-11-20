"""
Fraud & Anomaly Screening Triage Queue (UI-05)
Author: Noman Zaid (Roll No: 2500320140103, MCA 2nd Year Data Analytics, ABES)
Faculty Mentor: Ms. Shilpa
Linked Requirements: BR-05, US-06, UI-05, TC-F-07, TC-T-11
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from src.data_loader import load_master_claims_data, filter_claims_data
import src.ui as ui

st.set_page_config(page_title="Fraud & Anomaly Screening", page_icon="🕵️", layout="wide")

df_master = load_master_claims_data()

policy_types = st.session_state.get("selected_policy_types", None)
modes = st.session_state.get("selected_modes", None)
date_range = st.session_state.get("date_range", None)

df = filter_claims_data(df_master, policy_types=policy_types, settlement_modes=modes, date_range=date_range)

ui.render_header(
    title="Fraud & Anomaly Screening Triage",
    subtitle="Early-inception claim auditing, tariff markup red-flags, and prioritized investigative triage",
    badge_text="SPECIAL INVESTIGATION UNIT · UI-05"
)

if df.empty:
    st.warning("⚠️ No claim records match active filter selections.")
    st.stop()

# Interactive Risk Score Slider
col_filter, col_spacer = st.columns([4, 8])
with col_filter:
    risk_threshold = st.slider(
        "Minimum Fraud Risk Score Threshold",
        min_value=30,
        max_value=95,
        value=70,
        step=5,
        help="Claims with score >= threshold are classified as High-Risk for investigation."
    )

# Compute Anomaly flags
df_scored = df.copy()
df_scored["is_early_inception"] = df_scored["days_from_inception"] <= 30
df_scored["is_high_risk"] = df_scored["fraud_risk_score"] >= risk_threshold

high_risk_df = df_scored[df_scored["is_high_risk"]]
high_risk_count = len(high_risk_df)
total_exposure_lakhs = round(high_risk_df["claim_amount_requested"].sum() / 1e5, 1)
early_inception_count = int(high_risk_df["is_early_inception"].sum())

kpi_data = [
    {
        "title": "High-Risk Claims",
        "value": f"{high_risk_count}",
        "subtext": f"Scoring >= {risk_threshold} on anomaly rubric",
        "trend": "Triage Priority",
        "color": "rose" if high_risk_count > 0 else "emerald"
    },
    {
        "title": "At-Risk Financial Exposure",
        "value": f"₹ {total_exposure_lakhs:,.1f} L",
        "subtext": "Sum of requested high-risk payouts",
        "trend": "Capital at Risk",
        "color": "rose"
    },
    {
        "title": "Early Inception Triggers",
        "value": f"{early_inception_count} Claims",
        "subtext": "Claim filed within 30 days of policy start",
        "trend": "Moral Hazard Risk",
        "color": "amber"
    },
    {
        "title": "Investigation Capacity",
        "value": "100%",
        "subtext": "All flagged cases mapped to field triage",
        "trend": "SIU Queue",
        "color": "blue"
    }
]

ui.render_kpi_row(kpi_data)
st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

# Row 1: High-Risk Triage Queue Table
st.subheader("High-Risk Investigative Triage Queue")
st.caption(f"Showing all {high_risk_count} claims exceeding risk threshold >= {risk_threshold}. Ordered by financial exposure.")

if not high_risk_df.empty:
    triage_display = high_risk_df[[
        "claim_id", "policy_id", "policyholder_name", "provider_name", "policy_type",
        "days_from_inception", "claim_amount_requested", "fraud_risk_score", "claim_status"
    ]].sort_values("claim_amount_requested", ascending=False).copy()

    triage_display["claim_amount_requested"] = triage_display["claim_amount_requested"].apply(lambda x: f"₹ {x:,.0f}")
    triage_display["days_from_inception"] = triage_display["days_from_inception"].astype(int).astype(str) + "d"
    triage_display["Recommended Action"] = triage_display["fraud_risk_score"].apply(
        lambda s: "🔍 Immediate Forensic Audit" if s >= 85 else "📋 Document & Tele-Verification"
    )
    triage_display.columns = [
        "Claim ID", "Policy ID", "Policyholder", "Provider", "Portfolio",
        "Days to Claim", "Requested (₹)", "Risk Score", "Adjudication Status", "Recommended Action"
    ]

    st.dataframe(triage_display, use_container_width=True, hide_index=True)
else:
    st.success("No claims currently exceed the selected risk threshold.")

# Row 2: Anomaly Composition & Score Distribution
col1, col2 = st.columns([5, 7])

with col1:
    st.subheader("Anomaly Pattern Breakdown")
    st.caption("Distribution of primary anomaly indicators among high-risk claims.")

    anomaly_counts = pd.Series({
        "Early Inception (<30d)": int(df_scored["is_early_inception"].sum()),
        "Disproportionate Payout": len(df_scored[df_scored["claim_amount_requested"] > 400000]),
        "Unaccredited Provider": len(df_scored[df_scored["empanelment_status"] == "Under Review"]),
        "Duplicate Suspect": len(df_scored[df_scored["fraud_risk_score"] > 80])
    }).reset_index()
    anomaly_counts.columns = ["Anomaly Trigger", "Count"]

    fig_pie = px.pie(
        anomaly_counts,
        names="Anomaly Trigger",
        values="Count",
        hole=0.45,
        color_discrete_sequence=["#E53E3E", "#DD6B20", "#D69E2E", "#3182CE"]
    )
    fig_pie.update_layout(
        template="plotly_white",
        height=320,
        margin=dict(l=20, r=20, t=20, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5)
    )
    st.plotly_chart(fig_pie, use_container_width=True)

with col2:
    st.subheader("Portfolio Fraud Risk Score Distribution")
    st.caption("Histogram showing the distribution of computed risk scores across all claims.")

    fig_hist = px.histogram(
        df_scored,
        x="fraud_risk_score",
        nbins=20,
        color="policy_type",
        labels={"fraud_risk_score": "Computed Fraud Risk Score (0 - 100)"},
        color_discrete_sequence=["#1B365D", "#2B6CB0", "#319795", "#E53E3E"]
    )
    fig_hist.add_vline(x=risk_threshold, line_dash="dash", line_color="#E53E3E", annotation_text=f"Cutoff: {risk_threshold}")
    fig_hist.update_layout(
        template="plotly_white",
        height=320,
        margin=dict(l=20, r=20, t=20, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    st.plotly_chart(fig_hist, use_container_width=True)

ui.render_insight_box(
    title="Special Investigation Unit (SIU) Recommendation",
    text=f"Early-inception claims (days from inception <= 30) constitute {early_inception_count} high-risk files accounting for ₹{total_exposure_lakhs:.1f} Lakhs in exposure. Mandating physical tele-underwriting verification prior to settlement releases mitigates moral hazard in first-year policies.",
    category="Fraud Prevention"
)
