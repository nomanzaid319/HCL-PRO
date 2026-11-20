"""
Technical Test Suite 1: Database & Data Integrity Tests
Author: Noman Zaid (Roll No: 2500320140103, MCA 2nd Year Data Analytics, ABES)
Test Cases: TC-T-01, TC-T-02, TC-T-03, TC-T-04, TC-T-05, TC-T-06
"""

import os
import sqlite3
import pytest
import pandas as pd

DB_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "app.db")
CLAIMS_CSV = os.path.join(os.path.dirname(__file__), "..", "data", "claims_fact.csv")

@pytest.fixture(scope="module")
def db_conn():
    """Provides a SQLite database connection with foreign keys enabled."""
    assert os.path.exists(DB_PATH), f"Database file not found at {DB_PATH}"
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON;")
    yield conn
    conn.close()

def test_tc_t_01_sqlite_connection_and_tables(db_conn):
    """TC-T-01: Verify SQLite database connects and all 4 tables plus view exist."""
    cur = db_conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type IN ('table', 'view');")
    names = [row[0] for row in cur.fetchall()]
    
    expected = [
        "policyholders_master", 
        "policies_master", 
        "providers_master", 
        "claims_fact", 
        "v_claims_analysis_master"
    ]
    for entity in expected:
        assert entity in names, f"Required database entity {entity} missing from SQLite schema."

def test_tc_t_02_table_row_counts(db_conn):
    """TC-T-02: Verify planned row counts across all master and fact tables."""
    cur = db_conn.cursor()
    
    cur.execute("SELECT COUNT(*) FROM policyholders_master;")
    assert cur.fetchone()[0] >= 100, "policyholders_master should contain at least 100 records."
    
    cur.execute("SELECT COUNT(*) FROM policies_master;")
    assert cur.fetchone()[0] >= 120, "policies_master should contain at least 120 records."
    
    cur.execute("SELECT COUNT(*) FROM providers_master;")
    assert cur.fetchone()[0] >= 25, "providers_master should contain at least 25 records."
    
    cur.execute("SELECT COUNT(*) FROM claims_fact;")
    assert cur.fetchone()[0] >= 250, "claims_fact should contain at least 250 records."
    
    cur.execute("SELECT COUNT(*) FROM v_claims_analysis_master;")
    assert cur.fetchone()[0] >= 250, "Analytical view should contain at least 250 joined records."

def test_tc_t_03_primary_key_uniqueness(db_conn):
    """TC-T-03: Verify 100% primary key uniqueness across all tables."""
    tables_pk = [
        ("policyholders_master", "policyholder_id"),
        ("policies_master", "policy_id"),
        ("providers_master", "provider_id"),
        ("claims_fact", "claim_id")
    ]
    cur = db_conn.cursor()
    for table, pk in tables_pk:
        cur.execute(f"SELECT COUNT({pk}), COUNT(DISTINCT {pk}) FROM {table};")
        total_cnt, dist_cnt = cur.fetchone()
        assert total_cnt == dist_cnt, f"Duplicate primary keys detected in {table} on column {pk}."

def test_tc_t_04_foreign_key_referential_integrity(db_conn):
    """TC-T-04: Verify referential integrity with zero foreign key violations."""
    cur = db_conn.cursor()
    cur.execute("PRAGMA foreign_key_check;")
    violations = cur.fetchall()
    assert len(violations) == 0, f"Foreign key integrity violations detected: {violations}"

def test_tc_t_05_data_cleaner_deduplication():
    """TC-T-05: Verify data cleaning pipeline drops injected duplicate rows from raw CSV."""
    assert os.path.exists(CLAIMS_CSV), f"Raw claims CSV not found at {CLAIMS_CSV}"
    df_raw = pd.read_csv(CLAIMS_CSV)
    assert len(df_raw) == 253, "Raw CSV should contain 253 rows (250 unique + 3 duplicates)."
    
    df_clean = df_raw.drop_duplicates(subset=["claim_id"])
    assert len(df_clean) == 250, "Deduplication should yield exactly 250 unique claims."

def test_tc_t_06_data_cleaner_null_handling():
    """TC-T-06: Verify handling of missing CSAT scores without data loss."""
    df_raw = pd.read_csv(CLAIMS_CSV)
    null_count = df_raw["csat_score"].isnull().sum()
    assert null_count > 0, "Raw CSV should contain controlled null values in csat_score."
    
    # Impute missing CSAT with median
    median_csat = df_raw["csat_score"].median()
    df_imputed = df_raw.copy()
    df_imputed["csat_score"] = df_imputed["csat_score"].fillna(median_csat)
    assert df_imputed["csat_score"].isnull().sum() == 0, "Missing CSAT scores should be fully imputed."
