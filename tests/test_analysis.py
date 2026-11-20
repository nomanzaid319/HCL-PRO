"""
Technical Test Suite 2: Analytical & Statistical Function Tests
Author: Noman Zaid (Roll No: 2500320140103, MCA 2nd Year Data Analytics, ABES)
Test Cases: TC-T-07, TC-T-08, TC-T-09, TC-T-10, TC-T-11
"""

import pytest
import numpy as np
import pandas as pd
from scipy import stats

def test_tc_t_07_tat_kpi_calculation():
    """TC-T-07: Verify mean and median turnaround time calculations on a known dataset."""
    sample_data = pd.DataFrame({
        "tat_days": [5, 10, 15, 20, 25]
    })
    mean_tat = sample_data["tat_days"].mean()
    median_tat = sample_data["tat_days"].median()
    
    assert mean_tat == 15.0, f"Expected mean TAT of 15.0, got {mean_tat}"
    assert median_tat == 15.0, f"Expected median TAT of 15.0, got {median_tat}"

def test_tc_t_08_pareto_repudiation_math():
    """TC-T-08: Verify Pareto 80/20 cumulative percentage calculations on known frequencies."""
    rejection_counts = pd.Series([50, 30, 10, 5, 5], index=["R1", "R2", "R3", "R4", "R5"])
    cum_pct = (rejection_counts.cumsum() / rejection_counts.sum()) * 100.0
    
    # Top 2 reasons (50 + 30 = 80 out of 100) must equal 80.0%
    assert cum_pct.iloc[0] == 50.0, f"Expected first Pareto share 50.0%, got {cum_pct.iloc[0]}"
    assert cum_pct.iloc[1] == 80.0, f"Expected top 2 Pareto share 80.0%, got {cum_pct.iloc[1]}"
    assert cum_pct.iloc[-1] == 100.0, f"Expected cumulative total 100.0%, got {cum_pct.iloc[-1]}"

def test_tc_t_09_haircut_calculation():
    """TC-T-09: Verify payout haircut percentage: ((Requested - Settled) / Requested) * 100."""
    sample_claims = pd.DataFrame({
        "claim_amount_requested": [100000.0, 50000.0, 200000.0],
        "claim_amount_settled":   [80000.0,  45000.0, 150000.0]
    })
    haircuts = ((sample_claims["claim_amount_requested"] - sample_claims["claim_amount_settled"]) / 
                sample_claims["claim_amount_requested"]) * 100.0
    
    expected_haircuts = [20.0, 10.0, 25.0]
    for actual, expected in zip(haircuts, expected_haircuts):
        assert abs(actual - expected) < 1e-4, f"Expected haircut {expected}%, got {actual}%"

def test_tc_t_10_welchs_ttest_math():
    """TC-T-10: Verify two-sample independent Welch's t-test outputs statistically significant p < 0.05."""
    np.random.seed(42)
    cashless_tat = np.random.normal(loc=8.0, scale=2.0, size=100)
    reimbursement_tat = np.random.normal(loc=17.0, scale=4.0, size=100)
    
    t_stat, p_val = stats.ttest_ind(cashless_tat, reimbursement_tat, equal_var=False)
    
    assert t_stat < 0, f"Cashless TAT should be significantly lower (negative t-stat), got {t_stat}"
    assert p_val < 0.001, f"Expected p-value < 0.001, got {p_val}"

def test_tc_t_11_fraud_anomaly_scoring():
    """TC-T-11: Verify that early-inception claims (< 30 days) trigger elevated fraud scores."""
    sample_claim = {
        "days_from_inception": 14,
        "is_duplicate": 0,
        "base_score": 30
    }
    
    # Rule: If days_from_inception <= 30, add 45 fraud points
    score = sample_claim["base_score"]
    if sample_claim["days_from_inception"] <= 30:
        score += 45
        
    assert score >= 70, f"Early inception claim must score >= 70, got {score}"
    is_high_risk = 1 if score >= 70 else 0
    assert is_high_risk == 1, "Early inception claim must be flagged as high risk (1)"
