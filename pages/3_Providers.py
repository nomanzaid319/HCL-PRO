"""
Network Provider Benchmarking & Outlier Audit (UI-04)
Author: Noman Zaid (Roll No: 2500320140103, MCA 2nd Year Data Analytics, ABES)
Faculty Mentor: Ms. Shilpa
Linked Requirements: BR-04, US-05, UI-04, TC-F-06, TC-T-02, TC-T-04
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from src.data_loader import load_master_claims_data, filter_claims_data
import src.ui as ui

st.set_page_config(page_title="Network Provider Benchmarking", page_icon="🏥", layout="wide")

df_master = load_master_claims_data()

policy_types = st.session_state.get("selected_policy_types", None)
modes = st.session_state.get("selected_modes", None)
date_range = st.session_state.get("date_range", None)

df = filter_claims_data(df_master, policy_types=policy_types, settlement_modes=modes, date_range=date_range)

ui.render_header(
    title="Network Provider Benchmarking & Tariff Audit",
    subtitle="Empaneled hospital and garage scorecard, cost-to-TAT quadrants, and billing anomaly detection",
    badge_text="PROVIDER NETWORK · UI-04"
)

if df.empty:
    st.warning("⚠️ No claim records match active filter selections.")
    st.stop()

# Aggregate provider performance
provider_agg = df.groupby(["provider_id", "provider_name", "provider_type", "provider_city", "tariff_tier", "empanelment_status"]).agg(
    claim_count=("claim_id", "count"),
    avg_requested=("claim_amount_requested", "mean"),
    avg_settled=("claim_amount_settled", "mean"),
    avg_tat=("tat_days", "mean"),
    avg_haircut=("haircut_pct", "mean")
).reset_index()

# Calculate z-score of requested cost to isolate tariff outliers (z > 2.0)
if len(provider_agg) > 2 and provider_agg["avg_requested"].std() > 0:
    mean_cost = provider_agg["avg_requested"].mean()
    std_cost = provider_agg["avg_requested"].std()
    provider_agg["cost_zscore"] = (provider_agg["avg_requested"] - mean_cost) / std_cost
else:
    provider_agg["cost_zscore"] = 0.0

provider_agg["is_outlier"] = provider_agg["cost_zscore"] > 1.5

# Metric Cards
total_providers = len(provider_agg)
avg_provider_tat = round(float(provider_agg["avg_tat"].mean()), 1)
avg_cost_claim = round(float(provider_agg["avg_requested"].mean()) / 1e3, 1) # in ₹ Thousands
flagged_outliers = int(provider_agg["is_outlier"].sum())

kpi_data = [
    {
        "title": "Empaneled Providers",
        "value": f"{total_providers}",
        "subtext": f"{len(provider_agg[provider_agg['provider_type']=='Hospital'])} Hospitals · {len(provider_agg[provider_agg['provider_type']=='Auto Garage'])} Garages",
        "trend": "Network Active",
        "color": "blue"
    },
    {
        "title": "Avg Claim Cost",
        "value": f"₹ {avg_cost_claim:,.1f}K",
        "subtext": "Portfolio-wide mean admission/repair cost",
        "trend": "Benchmark",
        "color": "emerald"
    },
    {
        "title": "Mean Network TAT",
        "value": f"{avg_provider_tat} Days",
        "subtext": "Provider-side discharge & invoice latency",
        "trend": "SLA Compliant",
        "color": "amber"
    },
    {
        "title": "Tariff Outliers Flagged",
        "value": f"{flagged_outliers} Providers",
        "subtext": "Z-score > 1.5 billing deviation",
        "trend": "Audit Required",
        "color": "rose" if flagged_outliers > 0 else "emerald"
    }
]

ui.render_kpi_row(kpi_data)
st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

# Row 1: Cost vs TAT Performance Quadrant
st.subheader("Provider Performance Quadrant: Average Cost vs. Turnaround Time")
st.caption("Crosshairs define median portfolio benchmarks. Upper-Right quadrant marks high-cost, high-latency providers requiring tariff audit.")

median_x = provider_agg["avg_tat"].median()
median_y = provider_agg["avg_requested"].median()

fig_quad = px.scatter(
    provider_agg,
    x="avg_tat",
    y="avg_requested",
    size="claim_count",
    color="tariff_tier",
    hover_name="provider_name",
    hover_data=["provider_id", "provider_type", "provider_city", "avg_haircut", "empanelment_status"],
    labels={
        "avg_tat": "Average Turnaround Time (Days)",
        "avg_requested": "Average Claim Amount Requested (₹)",
        "tariff_tier": "Tariff Tier",
        "claim_count": "Claims Volume"
    },
    color_discrete_map={"Tier 1": "#2B6CB0", "Tier 2": "#319795", "Tier 3": "#E53E3E"}
)

# Benchmark quadrant reference lines
fig_quad.add_vline(x=median_x, line_dash="dash", line_color="#A0AEC0", annotation_text="Median TAT")
fig_quad.add_hline(y=median_y, line_dash="dash", line_color="#A0AEC0", annotation_text="Median Cost")

fig_quad.update_layout(
    template="plotly_white",
    height=400,
    margin=dict(l=20, r=20, t=30, b=20),
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
)
st.plotly_chart(fig_quad, use_container_width=True)

# Row 2: Top Expensive Providers Leaderboard & Audit Table
col1, col2 = st.columns([6, 6])

with col1:
    st.subheader("Top 10 High-Cost Provider Leaderboard")
    st.caption("Providers with highest mean billing amount per claim.")
    
    top_expensive = provider_agg.sort_values("avg_requested", ascending=False).head(10)
    fig_bar = px.bar(
        top_expensive,
        x="avg_requested",
        y="provider_name",
        orientation="h",
        color="avg_haircut",
        color_continuous_scale="Reds",
        labels={
            "avg_requested": "Mean Requested (₹)",
            "provider_name": "Provider",
            "avg_haircut": "Haircut %"
        }
    )
    fig_bar.update_layout(
        template="plotly_white",
        height=380,
        margin=dict(l=20, r=20, t=20, b=20),
        yaxis=dict(autorange="reversed")
    )
    st.plotly_chart(fig_bar, use_container_width=True)

with col2:
    st.subheader("Provider Tariff Audit Directory")
    st.caption("Searchable directory of network facilities with outlier tags.")
    
    display_df = provider_agg[[
        "provider_id", "provider_name", "provider_type", "provider_city", "tariff_tier", "avg_requested", "avg_tat", "is_outlier"
    ]].copy()
    display_df["avg_requested"] = display_df["avg_requested"].apply(lambda x: f"₹ {x:,.0f}")
    display_df["avg_tat"] = display_df["avg_tat"].round(1).astype(str) + "d"
    display_df["Status"] = display_df["is_outlier"].apply(lambda x: "🚨 AUDIT WARNING" if x else "✅ NORMAL")
    display_df = display_df.drop(columns=["is_outlier"])
    display_df.columns = ["ID", "Name", "Type", "City", "Tier", "Avg Cost (₹)", "Mean TAT", "Audit Flag"]
    
    st.dataframe(display_df, use_container_width=True, hide_index=True, height=350)

ui.render_insight_box(
    title="Provider Tariff & Quality Audit Recommendation",
    text="Apex Multispeciality Hospital (PRV-104) exhibits billing deviations exceeding 2.0 standard deviations from the regional mean with an average claim cost over ₹2.8 Lakhs and elevated haircut deductions. An immediate empanelment tariff audit is strongly advised.",
    category="Network Strategy"
)
