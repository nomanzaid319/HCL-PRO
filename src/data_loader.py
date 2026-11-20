"""
Data Loader & Query Ingestion Module
Author: Noman Zaid (Roll No: 2500320140103, MCA 2nd Year Data Analytics, ABES)
Linked Requirements: BR-01..12, FR-01..14, TC-T-01..06, TC-F-01..02
"""

import os
import sqlite3
import pandas as pd
import numpy as np
import streamlit as st

ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
DB_PATH = os.path.join(ROOT_DIR, "data", "app.db")

def get_connection(db_path: str = DB_PATH) -> sqlite3.Connection:
    """
    Establishes and returns a connection to the SQLite database.
    """
    if not os.path.exists(db_path):
        raise FileNotFoundError(f"Database file not found at {db_path}")
    return sqlite3.connect(db_path, check_same_thread=False)

@st.cache_data(ttl=3600)
def load_master_claims_data(db_path: str = DB_PATH) -> pd.DataFrame:
    """
    Executes a multi-table SQL join linking claims_fact with policies_master,
    policyholders_master, and providers_master. Computes derived fields and returns a clean DataFrame.
    """
    conn = get_connection(db_path)
    sql = """
    SELECT 
        c.claim_id,
        c.policy_id,
        c.provider_id,
        c.incident_date,
        c.submission_date,
        c.verification_date,
        c.assessment_date,
        c.settlement_date,
        c.claim_amount_requested,
        c.claim_amount_settled,
        c.settlement_ratio,
        c.settlement_mode,
        c.adjudication_stage,
        c.claim_status,
        c.rejection_reason_code,
        c.rejection_reason_detail,
        c.tat_days,
        c.fraud_risk_score,
        c.fraud_flag,
        c.csat_score,
        p.policyholder_id,
        p.policy_type,
        p.policy_tier,
        p.sum_insured,
        p.annual_premium,
        p.policy_start_date,
        p.policy_status,
        h.full_name AS policyholder_name,
        h.age,
        h.gender,
        h.city AS policyholder_city,
        h.state AS policyholder_state,
        h.region,
        h.tenure_months,
        pr.provider_name,
        pr.provider_type,
        pr.city AS provider_city,
        pr.state AS provider_state,
        pr.empanelment_status,
        pr.tariff_tier
    FROM claims_fact c
    LEFT JOIN policies_master p ON c.policy_id = p.policy_id
    LEFT JOIN policyholders_master h ON p.policyholder_id = h.policyholder_id
    LEFT JOIN providers_master pr ON c.provider_id = pr.provider_id
    """
    df = pd.read_sql_query(sql, conn)
    conn.close()

    # Date parsing
    date_cols = ["incident_date", "submission_date", "verification_date", "assessment_date", "settlement_date", "policy_start_date"]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # Numeric guarantees
    df["claim_amount_requested"] = pd.to_numeric(df["claim_amount_requested"], errors="coerce").fillna(0.0)
    df["claim_amount_settled"] = pd.to_numeric(df["claim_amount_settled"], errors="coerce").fillna(0.0)
    df["tat_days"] = pd.to_numeric(df["tat_days"], errors="coerce").fillna(0.0)
    df["fraud_risk_score"] = pd.to_numeric(df["fraud_risk_score"], errors="coerce").fillna(0.0)
    df["csat_score"] = pd.to_numeric(df["csat_score"], errors="coerce").fillna(0.0)

    # Derived analytical columns
    # 1. Haircut Percentage: ((Requested - Settled) / Requested) * 100
    df["haircut_pct"] = np.where(
        df["claim_amount_requested"] > 0,
        ((df["claim_amount_requested"] - df["claim_amount_settled"]) / df["claim_amount_requested"]) * 100.0,
        0.0
    )
    df["haircut_pct"] = df["haircut_pct"].clip(lower=0.0, upper=100.0)

    # 2. Days from policy inception to claim
    df["days_from_inception"] = (df["submission_date"] - df["policy_start_date"]).dt.days.fillna(999)

    # 3. Stage cycle times (in days)
    df["stage_doc_days"] = (df["verification_date"] - df["submission_date"]).dt.days.clip(lower=0).fillna(0)
    df["stage_assess_days"] = (df["assessment_date"] - df["verification_date"]).dt.days.clip(lower=0).fillna(0)
    df["stage_settle_days"] = (df["settlement_date"] - df["assessment_date"]).dt.days.clip(lower=0).fillna(0)

    # 4. Submission Month string for time-series aggregation (YYYY-MM)
    df["submission_month"] = df["submission_date"].dt.to_period("M").astype(str)

    return df

def filter_claims_data(
    df: pd.DataFrame,
    policy_types: list = None,
    settlement_modes: list = None,
    date_range: tuple = None,
    claim_status: list = None
) -> pd.DataFrame:
    """
    Applies global filter criteria dynamically to the master claims dataset.
    """
    filtered = df.copy()

    if policy_types:
        filtered = filtered[filtered["policy_type"].isin(policy_types)]

    if settlement_modes:
        filtered = filtered[filtered["settlement_mode"].isin(settlement_modes)]

    if claim_status:
        filtered = filtered[filtered["claim_status"].isin(claim_status)]

    if date_range and len(date_range) == 2 and date_range[0] and date_range[1]:
        start_date = pd.to_datetime(date_range[0])
        end_date = pd.to_datetime(date_range[1])
        filtered = filtered[(filtered["submission_date"] >= start_date) & (filtered["submission_date"] <= end_date)]

    return filtered
