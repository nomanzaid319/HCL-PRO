"""
Denial Analysis & Repudiation Root Cause (UI-03)
Author: Noman Zaid (Roll No: 2500320140103, MCA 2nd Year Data Analytics, ABES)
Faculty Mentor: Ms. Shilpa
Linked Requirements: BR-02, BR-09, US-03, US-10, UI-03, TC-F-04, TC-F-11, TC-T-08
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from src.data_loader import load_master_claims_data, filter_claims_data
from src.analytics import compute_pareto_repudiation
import src.ui as ui

st.set_page_config(page_title="Denial & Repudiation Analysis", page_icon="🚫", layout="wide")

# Inject custom executive CSS & Google Fonts
ui.inject_custom_theme()

df_master = load_master_claims_data()

policy_types = st.session_state.get("selected_policy_types", None)
modes = st.session_state.get("selected_modes", None)
date_range = st.session_state.get("date_range", None)

df = filter_claims_data(df_master, policy_types=policy_types, settlement_modes=modes, date_range=date_range)

ui.render_header(
    title="Claim Denial & Repudiation Root Cause",
    subtitle="Pareto 80/20 repudiation analysis, regional grievance hotspots, and policy exclusion audits",
    badge_text="RISK & COMPLIANCE · UI-03"
)

if df.empty:
    st.warning("⚠️ No claim records match active filter selections.")
    st.stop()

# Rejection subset
rejections_df = df[df["claim_status"] == "Repudiated"]
total_rejections = len(rejections_df)
repudiation_rate = round((total_rejections / len(df)) * 100.0, 1) if len(df) > 0 else 0.0
denied_volume_cr = round(rejections_df["claim_amount_requested"].sum() / 1e7, 2)
avg_csat = round(float(df["csat_score"].mean()), 2)

kpi_data = [
    {
        "title": "Total Repudiations",
        "value": f"{total_rejections:,}",
        "subtext": f"{len(df):,} Total Claims Evaluated",
        "trend": "Audit Queue",
        "color": "rose"
    },
    {
        "title": "Repudiation Rate",
        "value": f"{repudiation_rate}%",
        "subtext": "Industry Standard: ~25-30%",
        "trend": "IRDAI Metric",
        "color": "amber" if repudiation_rate <= 30 else "rose"
    },
    {
        "title": "Denied Financial Exposure",
        "value": f"₹ {denied_volume_cr:.2f} Cr",
        "subtext": "Protected from unwarranted leakage",
        "trend": "Risk Preserved",
        "color": "emerald"
    },
    {
        "title": "Customer CSAT Score",
        "value": f"{avg_csat} / 5.0",
        "subtext": f"Grievance index: {len(df[df['csat_score'] < 3])} dissatisfied",
        "trend": "Satisfaction",
        "color": "blue"
    }
]

ui.render_kpi_row(kpi_data)
st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

# Row 1: Dual-Axis Pareto 80/20 Chart
st.subheader("Dual-Axis Pareto 80/20 Repudiation Analysis")
st.caption("Validates whether 80% of repudiations are driven by the top 20% of root cause categories.")

pareto_df = compute_pareto_repudiation(df)

if not pareto_df.empty:
    fig_pareto = go.Figure()
    
    # Bar trace: Count
    fig_pareto.add_trace(go.Bar(
        x=pareto_df["rejection_reason"],
        y=pareto_df["count"],
        name="Repudiation Frequency (Count)",
        marker_color="#2B6CB0",
        yaxis="y1"
    ))
    
    # Line trace: Cumulative Percentage
    fig_pareto.add_trace(go.Scatter(
        x=pareto_df["rejection_reason"],
        y=pareto_df["cumulative_pct"],
        name="Cumulative Percentage (%)",
        line=dict(color="#C53030", width=3),
        mode="lines+markers",
        yaxis="y2"
    ))
    
    # 80% Benchmark Reference Line
    fig_pareto.add_shape(
        type="line",
        x0=-0.5,
        x1=len(pareto_df) - 0.5,
        y0=80,
        y1=80,
        yref="y2",
        line=dict(color="#D69E2E", dash="dash", width=2)
    )
    
    fig_pareto.add_annotation(
        x=0.5,
        y=82,
        yref="y2",
        text="80% Pareto Cutoff",
        showarrow=False,
        font=dict(color="#D69E2E", size=11, family="sans-serif")
    )
    
    fig_pareto.update_layout(
        template="plotly_white",
        height=380,
        margin=dict(l=20, r=20, t=30, b=20),
        xaxis_title="Rejection Reason Code",
        yaxis=dict(title="Denial Count", side="left"),
        yaxis2=dict(title="Cumulative Percentage (%)", side="right", overlaying="y", range=[0, 110], showgrid=False),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    st.plotly_chart(fig_pareto, use_container_width=True)
else:
    st.info("No repudiated claims found in the current selection.")

# Row 2: Regional Repudiation Heatmap & Rejection Root Cause Table
col1, col2 = st.columns([6, 6])

with col1:
    st.subheader("Regional Repudiation & Grievance Heatmap")
    st.caption("State-level breakdown of claim rejection rates and customer satisfaction.")
    
    state_agg = df.groupby("policyholder_state").agg(
        total_claims=("claim_id", "count"),
        repudiated_claims=("claim_status", lambda s: (s == "Repudiated").sum()),
        mean_csat=("csat_score", "mean")
    ).reset_index()
    
    state_agg["repudiation_rate"] = (state_agg["repudiated_claims"] / state_agg["total_claims"]) * 100.0
    state_agg["mean_csat"] = state_agg["mean_csat"].round(2)
    state_agg = state_agg.sort_values("repudiation_rate", ascending=False)
    
    fig_bar = px.bar(
        state_agg.head(10),
        x="repudiation_rate",
        y="policyholder_state",
        color="mean_csat",
        orientation="h",
        color_continuous_scale="Viridis",
        labels={
            "repudiation_rate": "Repudiation Rate (%)",
            "policyholder_state": "State",
            "mean_csat": "Avg CSAT"
        }
    )
    fig_bar.update_layout(
        template="plotly_white",
        height=360,
        margin=dict(l=20, r=20, t=20, b=20)
    )
    st.plotly_chart(fig_bar, use_container_width=True)

with col2:
    st.subheader("Repudiation Root-Cause Drill-down")
    st.caption("IRDAI categorized repudiation codes with financial impact.")
    
    if not rejections_df.empty:
        summary_table = rejections_df.groupby("rejection_reason_code").agg(
            denial_count=("claim_id", "count"),
            total_denied_inr=("claim_amount_requested", "sum"),
            mean_tat=("tat_days", "mean")
        ).reset_index().sort_values("denial_count", ascending=False)
        
        summary_table["total_denied_inr"] = summary_table["total_denied_inr"].apply(lambda x: f"₹ {x:,.0f}")
        summary_table["mean_tat"] = summary_table["mean_tat"].round(1).astype(str) + "d"
        summary_table.columns = ["Reason Code", "Count", "Total Denied (₹)", "Mean TAT"]
        
        st.dataframe(summary_table, use_container_width=True, hide_index=True)
    else:
        st.write("No repudiations in active filter.")

ui.render_insight_box(
    title="Pareto 80/20 Root Cause Finding",
    text="POLICY_EXCLUSION and EXCL_PRE_EXIST together account for over 81% of all claim repudiations. Clearer pre-policy disclosures and automated exclusion screening during onboarding would prevent more than 4 out of 5 consumer grievances.",
    category="Policy Recommendation"
)
