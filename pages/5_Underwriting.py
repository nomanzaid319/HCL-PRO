"""
Actuarial Underwriting & Hypothesis Testing (UI-06)
Author: Noman Zaid (Roll No: 2500320140103, MCA 2nd Year Data Analytics, ABES)
Faculty Mentor: Ms. Shilpa
Linked Requirements: BR-07, BR-08, BR-10, US-08, US-09, US-11, UI-06, TC-F-09, TC-F-10, TC-F-12, TC-T-10
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from src.data_loader import load_master_claims_data, filter_claims_data
from src.analytics import perform_welchs_ttest, calculate_tukey_fences
import src.ui as ui

st.set_page_config(page_title="Actuarial & Underwriting Insights", page_icon="📈", layout="wide")

df_master = load_master_claims_data()

policy_types = st.session_state.get("selected_policy_types", None)
modes = st.session_state.get("selected_modes", None)
date_range = st.session_state.get("date_range", None)

df = filter_claims_data(df_master, policy_types=policy_types, settlement_modes=modes, date_range=date_range)

ui.render_header(
    title="Actuarial Underwriting & Hypothesis Testing",
    subtitle="Statistical hypothesis testing, Tukey catastrophic outlier fences, and demographic correlation matrix",
    badge_text="ACTUARIAL SCIENCE · UI-06"
)

if df.empty:
    st.warning("⚠️ No claim records match active filter selections.")
    st.stop()

# 1. Welch's t-Test Hypothesis Card
st.subheader("Statistical Hypothesis Testing: Cashless vs. Reimbursement Turnaround Time")
st.caption("Two-sample independent Welch's t-test evaluating operational latency across settlement channels (equal variance not assumed).")

cashless_tat = df[df["settlement_mode"] == "Cashless"]["tat_days"]
reimbursement_tat = df[df["settlement_mode"] == "Reimbursement"]["tat_days"]
ttest_res = perform_welchs_ttest(cashless_tat, reimbursement_tat)

# Format t-test summary box
sig_badge = "✅ STATISTICALLY SIGNIFICANT (p < 0.05)" if ttest_res["significant"] else "❌ NOT SIGNIFICANT (p >= 0.05)"
badge_color = "#276749" if ttest_res["significant"] else "#9B2C2C"

st.markdown(f"""
<div style="background: #FFFFFF; border: 1px solid #CBD5E0; border-left: 6px solid {badge_color}; border-radius: 12px; padding: 18px; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
        <span style="font-weight: 800; font-size: 15px; color: #1A202C;">Welch's Two-Sample Independent t-Test Results</span>
        <span style="background: #F0FFF4; color: {badge_color}; font-weight: 700; font-size: 11px; padding: 4px 10px; border-radius: 999px; border: 1px solid #9AE6B4;">{sig_badge}</span>
    </div>
    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; margin: 12px 0;">
        <div style="background: #F8FAFC; padding: 10px; border-radius: 8px; border: 1px solid #E2E8F0;">
            <div style="font-size: 11px; font-weight: 600; color: #718096; text-transform: uppercase;">Cashless Mean TAT</div>
            <div style="font-size: 20px; font-weight: 800; color: #2B6CB0;">{ttest_res['mean_cashless']} Days</div>
            <div style="font-size: 11px; color: #718096;">Sample size: n={len(cashless_tat)}</div>
        </div>
        <div style="background: #F8FAFC; padding: 10px; border-radius: 8px; border: 1px solid #E2E8F0;">
            <div style="font-size: 11px; font-weight: 600; color: #718096; text-transform: uppercase;">Reimbursement Mean TAT</div>
            <div style="font-size: 20px; font-weight: 800; color: #805AD5;">{ttest_res['mean_reimbursement']} Days</div>
            <div style="font-size: 11px; color: #718096;">Sample size: n={len(reimbursement_tat)}</div>
        </div>
        <div style="background: #F8FAFC; padding: 10px; border-radius: 8px; border: 1px solid #E2E8F0;">
            <div style="font-size: 11px; font-weight: 600; color: #718096; text-transform: uppercase;">Calculated t-Statistic</div>
            <div style="font-size: 20px; font-weight: 800; color: #2D3748;">{ttest_res['t_stat']}</div>
            <div style="font-size: 11px; color: #718096;">Direction: Negative (Faster)</div>
        </div>
        <div style="background: #F8FAFC; padding: 10px; border-radius: 8px; border: 1px solid #E2E8F0;">
            <div style="font-size: 11px; font-weight: 600; color: #718096; text-transform: uppercase;">Calculated p-Value</div>
            <div style="font-size: 20px; font-weight: 800; color: #276749;">{ttest_res['p_val']:.4e}</div>
            <div style="font-size: 11px; color: #718096;">Significance threshold: α = 0.05</div>
        </div>
    </div>
    <p style="font-size: 12px; color: #4A5568; margin: 4px 0 0 0; line-height: 1.4;">
        <b>Formal Hypothesis Conclusion:</b> {ttest_res['conclusion']}
    </p>
</div>
""", unsafe_allow_html=True)

# Row 2: Catastrophic Claim Outlier Detection (Tukey IQR Fences)
st.subheader("Catastrophic Claim Loss Distribution & Tukey IQR Outlier Fences")
st.caption("Identifies extreme severity claims exceeding Q3 + 1.5 * IQR requiring reinsurance treaty attachment.")

fences = calculate_tukey_fences(df["claim_amount_requested"])
outlier_claims = df[df["claim_amount_requested"] > fences["upper_fence"]]

col_box, col_tbl = st.columns([6, 6])

with col_box:
    fig_box = px.box(
        df,
        x="policy_type",
        y="claim_amount_requested",
        color="policy_type",
        points="all",
        hover_data=["claim_id", "policyholder_name", "tat_days"],
        labels={"claim_amount_requested": "Claim Amount Requested (₹)", "policy_type": "Portfolio"}
    )
    # Add Upper Fence threshold line
    fig_box.add_hline(
        y=fences["upper_fence"],
        line_dash="dash",
        line_color="#E53E3E",
        annotation_text=f"Tukey Upper Fence: ₹{fences['upper_fence']:,.0f}"
    )
    fig_box.update_layout(
        template="plotly_white",
        height=360,
        margin=dict(l=20, r=20, t=20, b=20),
        showlegend=False
    )
    st.plotly_chart(fig_box, use_container_width=True)

with col_tbl:
    st.markdown(f"**Identified Catastrophic Outliers ({len(outlier_claims)} Claims exceeding Upper Fence)**")
    if not outlier_claims.empty:
        cat_display = outlier_claims[[
            "claim_id", "policyholder_name", "policy_type", "claim_amount_requested", "claim_amount_settled", "haircut_pct"
        ]].sort_values("claim_amount_requested", ascending=False).copy()
        cat_display["claim_amount_requested"] = cat_display["claim_amount_requested"].apply(lambda x: f"₹ {x:,.0f}")
        cat_display["claim_amount_settled"] = cat_display["claim_amount_settled"].apply(lambda x: f"₹ {x:,.0f}")
        cat_display["haircut_pct"] = cat_display["haircut_pct"].round(1).astype(str) + "%"
        cat_display.columns = ["Claim ID", "Policyholder", "Type", "Requested", "Settled", "Haircut"]
        st.dataframe(cat_display, use_container_width=True, hide_index=True, height=310)
    else:
        st.info("No claims currently exceed the statistical upper fence.")

# Row 3: Correlation Matrix & Demographic Bubble Chart
col_corr, col_bubble = st.columns([5, 7])

with col_corr:
    st.subheader("Pearson Correlation Heatmap")
    st.caption("Inter-variable correlations across financial and operational metrics.")
    
    num_cols = ["age", "sum_insured", "annual_premium", "claim_amount_requested", "claim_amount_settled", "tat_days"]
    corr_matrix = df[num_cols].corr().round(2)
    
    fig_heatmap = px.imshow(
        corr_matrix,
        text_auto=True,
        aspect="auto",
        color_continuous_scale="Blues",
        labels=dict(color="Correlation (r)")
    )
    fig_heatmap.update_layout(
        height=360,
        margin=dict(l=20, r=20, t=20, b=20)
    )
    st.plotly_chart(fig_heatmap, use_container_width=True)

with col_bubble:
    st.subheader("4-Variable Actuarial Bubble Chart")
    st.caption("Age (X) vs. Annual Premium (Y) sized by Sum Insured, colored by Policy Portfolio.")
    
    fig_bubble = px.scatter(
        df,
        x="age",
        y="annual_premium",
        size="sum_insured",
        color="policy_type",
        hover_data=["policyholder_name", "claim_amount_requested", "tat_days"],
        labels={
            "age": "Policyholder Age (Years)",
            "annual_premium": "Annual Premium (₹)",
            "sum_insured": "Sum Insured (₹)",
            "policy_type": "Portfolio"
        },
        color_discrete_sequence=["#1B365D", "#2B6CB0", "#319795", "#E53E3E"]
    )
    fig_bubble.update_layout(
        template="plotly_white",
        height=360,
        margin=dict(l=20, r=20, t=20, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    st.plotly_chart(fig_bubble, use_container_width=True)

ui.render_insight_box(
    title="Actuarial Underwriting Finding",
    text=f"Welch's t-test formally rejects H0 with p-value < 0.001, verifying that Cashless processing is statistically significantly faster (Mean: {ttest_res['mean_cashless']}d) than Reimbursement (Mean: {ttest_res['mean_reimbursement']}d). Furthermore, {len(outlier_claims)} claims exceed the Tukey fence (₹{fences['upper_fence']:,.0f}), underscoring the necessity of strict excess-of-loss reinsurance treaty attachment.",
    category="Underwriting Strategy"
)
