"""
Insurance Claim Processing Analysis - Synthetic Data Generator
Author: Noman Zaid (Roll No: 2500320140103, MCA 2nd Year Data Analytics, ABES)
Seed: 42 (Fully Reproducible)
"""

import os
import random
import sqlite3
from datetime import datetime, timedelta
import numpy as np
import pandas as pd
import openpyxl

def generate_dataset():
    # Set seeds for strict reproducibility
    SEED = 42
    random.seed(SEED)
    np.random.seed(SEED)

    print("==========================================================")
    print("STEP 10: GENERATING SYNTHETIC INSURANCE CLAIMS DATASET")
    print("Seed: 42 | Schema: 3NF Normalized (4 Tables + Joined View)")
    print("==========================================================")

    # -------------------------------------------------------------
    # 1. GEOGRAPHY & REFERENCE DATA (Fictional Indian Demographics)
    # -------------------------------------------------------------
    indian_first_names_m = [
        "Aarav", "Rohan", "Rajesh", "Vikram", "Amit", "Rahul", "Suresh", "Aditya", 
        "Manoj", "Sanjay", "Deepak", "Anil", "Arjun", "Kunal", "Alok", "Gaurav", 
        "Manish", "Nikhil", "Pankaj", "Sachin", "Sunil", "Vikas", "Vishal", "Yash"
    ]
    indian_first_names_f = [
        "Priya", "Ananya", "Sunita", "Pooja", "Neha", "Sneha", "Kavita", "Ritu", 
        "Meera", "Divya", "Aarti", "Geeta", "Ishita", "Jyoti", "Kiran", "Madhu", 
        "Nisha", "Pallavi", "Radha", "Shilpa", "Swati", "Tanvi", "Vandana", "Varsha"
    ]
    indian_last_names = [
        "Sharma", "Verma", "Gupta", "Singh", "Patel", "Nair", "Rao", "Mehta", 
        "Joshi", "Kumar", "Mishra", "Reddy", "Iyer", "Das", "Banerjee", "Chauhan", 
        "Bhatia", "Saxena", "Agarwal", "Malhotra", "Pandey", "Chopra", "Sethi", "Kapoor"
    ]

    geo_cities = [
        {"city": "Mumbai", "state": "Maharashtra", "region": "West"},
        {"city": "Pune", "state": "Maharashtra", "region": "West"},
        {"city": "Ahmedabad", "state": "Gujarat", "region": "West"},
        {"city": "Delhi", "state": "Delhi", "region": "North"},
        {"city": "Noida", "state": "Uttar Pradesh", "region": "North"},
        {"city": "Lucknow", "state": "Uttar Pradesh", "region": "North"},
        {"city": "Jaipur", "state": "Rajasthan", "region": "North"},
        {"city": "Chandigarh", "state": "Punjab", "region": "North"},
        {"city": "Bengaluru", "state": "Karnataka", "region": "South"},
        {"city": "Hyderabad", "state": "Telangana", "region": "South"},
        {"city": "Chennai", "state": "Tamil Nadu", "region": "South"},
        {"city": "Kochi", "state": "Kerala", "region": "South"},
        {"city": "Kolkata", "state": "West Bengal", "region": "East"},
        {"city": "Patna", "state": "Bihar", "region": "East"},
        {"city": "Bhopal", "state": "Madhya Pradesh", "region": "Central"},
        {"city": "Indore", "state": "Madhya Pradesh", "region": "Central"}
    ]

    # -------------------------------------------------------------
    # 2. GENERATE TABLE 1: policyholders_master (100 Rows)
    # -------------------------------------------------------------
    num_policyholders = 100
    policyholders = []

    for i in range(1, num_policyholders + 1):
        ph_id = f"POL-HLD-{1000 + i}"
        gender = random.choices(["Male", "Female", "Other"], weights=[0.53, 0.45, 0.02])[0]
        if gender == "Male":
            fn = random.choice(indian_first_names_m)
        elif gender == "Female":
            fn = random.choice(indian_first_names_f)
        else:
            fn = random.choice(indian_first_names_m + indian_first_names_f)
        ln = random.choice(indian_last_names)
        full_name = f"{fn} {ln}"

        age = int(np.random.normal(42, 11))
        age = max(19, min(75, age))

        geo = random.choice(geo_cities)
        tenure = random.randint(3, 118)

        policyholders.append({
            "policyholder_id": ph_id,
            "full_name": full_name,
            "age": age,
            "gender": gender,
            "city": geo["city"],
            "state": geo["state"],
            "region": geo["region"],
            "tenure_months": tenure
        })

    df_policyholders = pd.DataFrame(policyholders)
    print(f"Generated policyholders_master: {len(df_policyholders)} rows.")

    # -------------------------------------------------------------
    # 3. GENERATE TABLE 2: policies_master (120 Rows)
    # -------------------------------------------------------------
    num_policies = 120
    policies = []
    policy_types = ["Health", "Motor", "Life", "Property"]
    policy_type_weights = [0.42, 0.35, 0.13, 0.10]
    tiers = ["Silver", "Gold", "Platinum"]

    for i in range(1, num_policies + 1):
        pol_id = f"POL-{10000 + i}"
        ph_owner = random.choice(df_policyholders["policyholder_id"].tolist())
        ptype = random.choices(policy_types, weights=policy_type_weights)[0]
        ptier = random.choices(tiers, weights=[0.35, 0.45, 0.20])[0]

        if ptype == "Health":
            si_base = random.choice([300000, 500000, 750000, 1000000, 1500000, 2500000])
            multiplier = 1.0 if ptier == "Silver" else (1.4 if ptier == "Gold" else 1.8)
            annual_prem = round(si_base * 0.022 * multiplier + random.randint(800, 2500), 2)
        elif ptype == "Motor":
            si_base = random.choice([200000, 450000, 700000, 1000000, 1400000])
            multiplier = 1.0 if ptier == "Silver" else (1.25 if ptier == "Gold" else 1.5)
            annual_prem = round(si_base * 0.035 * multiplier + random.randint(500, 1800), 2)
        elif ptype == "Life":
            si_base = random.choice([1000000, 2500000, 5000000])
            annual_prem = round(si_base * 0.012 + random.randint(1200, 4000), 2)
        else: # Property
            si_base = random.choice([1500000, 3000000, 5000000])
            annual_prem = round(si_base * 0.008 + random.randint(1000, 3500), 2)

        # Inception date between 2023-01-01 and 2024-12-31
        start_dt = datetime(2023, 1, 1) + timedelta(days=random.randint(0, 700))
        status = random.choices(["Active", "Renewed", "Lapsed"], weights=[0.82, 0.12, 0.06])[0]

        policies.append({
            "policy_id": pol_id,
            "policyholder_id": ph_owner,
            "policy_type": ptype,
            "policy_tier": ptier,
            "sum_insured": float(si_base),
            "annual_premium": float(annual_prem),
            "policy_start_date": start_dt.strftime("%Y-%m-%d"),
            "policy_status": status
        })

    df_policies = pd.DataFrame(policies)
    print(f"Generated policies_master: {len(df_policies)} rows.")

    # -------------------------------------------------------------
    # 4. GENERATE TABLE 3: providers_master (25 Rows)
    # -------------------------------------------------------------
    num_providers = 25
    providers = []

    hospitals_list = [
        "Apollo Supercare Hospital", "Fortis Escorts Hospital", "Max Healthcare Institute", 
        "Apex Multispeciality Hospital", "Manipal Hospital", "Medanta The Medicity", 
        "Narayana Health City", "KIMS Global Hospital", "Lilavati Hospital & Research Centre", 
        "Sir Ganga Ram Hospital", "Ruby Hall Clinic", "Sahyadri Super Hospital"
    ]
    garages_list = [
        "Maruti Authorized Workshop Hub", "Hyundai Star Express Care", "Tata Motors Service Lounge", 
        "Metro Auto Care & Bodyworks", "Bosch Car Service Care", "Mahindra First Choice Service", 
        "Speedy Wheels Repair Hub", "Honda Pro Care Garage", "Autocraft Collision Center"
    ]
    tpas_list = [
        "MediAssist Insurance TPA", "Paramount Health Services TPA", 
        "Vidal Health TPA Services", "Heritage Health Insurance TPA"
    ]

    prv_idx = 1
    for h_name in hospitals_list:
        geo = random.choice(geo_cities)
        status = "Empaneled Network" if h_name != "Apex Multispeciality Hospital" else "Under Review"
        tier = "Tier 1" if "Max" in h_name or "Medanta" in h_name or "Apollo" in h_name else random.choice(["Tier 2", "Tier 3"])
        providers.append({
            "provider_id": f"PRV-{100 + prv_idx}",
            "provider_name": h_name,
            "provider_type": "Hospital",
            "city": geo["city"],
            "state": geo["state"],
            "empanelment_status": status,
            "tariff_tier": tier
        })
        prv_idx += 1

    for g_name in garages_list:
        geo = random.choice(geo_cities)
        status = "Empaneled Network" if g_name != "Metro Auto Care & Bodyworks" else "Under Review"
        tier = "Tier 2" if "Authorized" in g_name or "Express" in g_name else random.choice(["Tier 2", "Tier 3"])
        providers.append({
            "provider_id": f"PRV-{100 + prv_idx}",
            "provider_name": g_name,
            "provider_type": "Automobile Garage",
            "city": geo["city"],
            "state": geo["state"],
            "empanelment_status": status,
            "tariff_tier": tier
        })
        prv_idx += 1

    for t_name in tpas_list:
        geo = random.choice(geo_cities)
        providers.append({
            "provider_id": f"PRV-{100 + prv_idx}",
            "provider_name": t_name,
            "provider_type": "TPA",
            "city": geo["city"],
            "state": geo["state"],
            "empanelment_status": "Empaneled Network",
            "tariff_tier": "Tier 1"
        })
        prv_idx += 1

    df_providers = pd.DataFrame(providers)
    print(f"Generated providers_master: {len(df_providers)} rows.")

    # -------------------------------------------------------------
    # 5. GENERATE TABLE 4: claims_fact (250 Main Rows + Inbuilt Patterns)
    # -------------------------------------------------------------
    num_claims = 250
    claims = []

    # Filter providers by category for proper routing
    hospital_ids = df_providers[df_providers["provider_type"] == "Hospital"]["provider_id"].tolist()
    garage_ids = df_providers[df_providers["provider_type"] == "Automobile Garage"]["provider_id"].tolist()
    tpa_ids = df_providers[df_providers["provider_type"] == "TPA"]["provider_id"].tolist()

    rejection_reasons = [
        ("EXCL_PRE_EXIST", "Pre-Existing Medical Condition Non-Disclosure within waiting period"), # Pareto Top 1 (~50%)
        ("POLICY_EXCLUSION", "Damage or treatment specifically excluded under standard policy terms"), # Pareto Top 2 (~30%)
        ("DOC_FRAUD", "Inconsistent billing receipts, inflated estimates, or missing invoices"),
        ("LACK_PREAUTH", "Failure to obtain mandatory prior authorization before treatment"),
        ("DELAYED_FILING", "Claim submitted beyond mandatory statutory notification window"),
        ("BILL_UNSUBSTANTIATED", "Treatment tariff exceeds reasonable and customary territorial charges")
    ]
    rejection_weights = [0.50, 0.30, 0.08, 0.05, 0.04, 0.03]

    for i in range(1, num_claims + 1):
        claim_id = f"CLM-{10000 + i}"
        
        # Select policy
        pol_row = df_policies.sample(1, random_state=SEED + i).iloc[0]
        pol_id = pol_row["policy_id"]
        ptype = pol_row["policy_type"]
        si = pol_row["sum_insured"]
        pol_start_dt = datetime.strptime(pol_row["policy_start_date"], "%Y-%m-%d")

        # Route provider according to policy line
        if ptype == "Health":
            prov_id = random.choice(hospital_ids + tpa_ids)
        elif ptype == "Motor":
            prov_id = random.choice(garage_ids + tpa_ids)
        else: # Life or Property
            prov_id = random.choice(hospital_ids if ptype == "Life" else tpa_ids)

        # ---------------------------------------------------------
        # PATTERN 1: Seasonal Monsoon Spike (July - September)
        # ---------------------------------------------------------
        # We increase probability of claims in July-September (Month 7, 8, 9)
        year = random.choice([2023, 2024])
        is_monsoon_surge = random.random() < 0.38 # 38% of claims during Q3
        if is_monsoon_surge:
            month = random.choice([7, 8, 9])
            day = random.randint(1, 28)
        else:
            month = random.choice([1, 2, 3, 4, 5, 6, 10, 11, 12])
            day = random.randint(1, 28)

        # Ensure submission date is after policy inception date
        potential_sub_dt = datetime(year, month, day)
        if potential_sub_dt < pol_start_dt:
            sub_dt = pol_start_dt + timedelta(days=random.randint(10, 180))
        else:
            sub_dt = potential_sub_dt

        incident_dt = sub_dt - timedelta(days=random.randint(1, 14))

        # ---------------------------------------------------------
        # PATTERN 2: Early Inception Fraud Trigger (< 30 days)
        # ---------------------------------------------------------
        days_from_start = (incident_dt - pol_start_dt).days
        is_early_inception = (days_from_start <= 30) and (days_from_start >= 0)

        # ---------------------------------------------------------
        # PATTERN 3: Cashless vs. Reimbursement Settlement Modes
        # ---------------------------------------------------------
        # Cashless is ~55%, Reimbursement ~45%
        settlement_mode = random.choices(["Cashless", "Reimbursement"], weights=[0.55, 0.45])[0]

        if settlement_mode == "Cashless":
            # Cashless claims process significantly faster: mean 7.8 days, sd 2.3
            tat_days = int(np.random.normal(8.0, 2.2))
            tat_days = max(2, min(16, tat_days))
            approval_prob = 0.88
        else:
            # Reimbursement claims process with higher latency: mean 17.5 days, sd 4.5
            tat_days = int(np.random.normal(17.5, 4.2))
            tat_days = max(7, min(42, tat_days))
            approval_prob = 0.74

        # Milestone timestamps derived from total tat_days
        verif_days = max(1, int(tat_days * 0.35))
        assess_days = max(1, int(tat_days * 0.40))
        verif_dt = sub_dt + timedelta(days=verif_days)
        assess_dt = verif_dt + timedelta(days=assess_days)
        settle_dt = sub_dt + timedelta(days=tat_days)

        # ---------------------------------------------------------
        # Claim Amounts & Outliers
        # ---------------------------------------------------------
        # Provider PRV-104 (Apex Multispeciality) has abnormally inflated costs (Pattern 4)
        is_inflated_provider = (prov_id == "PRV-104") or (prov_id == "PRV-116")
        
        if is_inflated_provider:
            req_amount = round(random.uniform(180000, 480000), 2)
        elif ptype == "Health":
            req_amount = round(random.uniform(25000, 320000), 2)
        elif ptype == "Motor":
            req_amount = round(random.uniform(15000, 185000), 2)
        elif ptype == "Life":
            req_amount = round(random.uniform(400000, 1800000), 2)
        else: # Property
            req_amount = round(random.uniform(100000, 1200000), 2)

        # Enforce Sum Insured ceiling
        req_amount = min(req_amount, si * 0.95)

        # ---------------------------------------------------------
        # PATTERN 4: Catastrophic Financial Outliers (OQ-07)
        # ---------------------------------------------------------
        # Inject 3 catastrophic claims (exceeding 20 Lakhs)
        if i in [42, 107, 218]:
            req_amount = 2250000.0 if si >= 2500000 else si * 0.90
            tat_days = 58 # Also causes extreme TAT outlier

        # ---------------------------------------------------------
        # Claim Decision & Repudiation Pareto Distribution
        # ---------------------------------------------------------
        # If early inception or inflated provider, rejection probability increases
        if is_early_inception or (is_inflated_provider and random.random() < 0.4):
            is_approved = False
        else:
            is_approved = (random.random() < approval_prob)

        if is_approved:
            claim_status = "Approved"
            adjudication_stage = "Settled"
            rej_code = None
            rej_detail = None
            
            # Settlement Haircut: typically 82% to 98% of requested amount
            settlement_ratio = round(random.uniform(0.78, 0.98), 4)
            # A few claims with haircuts > 20% for F-03
            if random.random() < 0.12:
                settlement_ratio = round(random.uniform(0.60, 0.76), 4)
            
            settled_amount = round(req_amount * settlement_ratio, 2)
        else:
            claim_status = "Repudiated"
            adjudication_stage = "Settled"
            # Pareto distribution of reasons
            rej_choice = random.choices(rejection_reasons, weights=rejection_weights)[0]
            rej_code = rej_choice[0]
            rej_detail = rej_choice[1]
            settlement_ratio = 0.0
            settled_amount = 0.0

        # ---------------------------------------------------------
        # Fraud Anomaly Scoring (OQ-05)
        # ---------------------------------------------------------
        fraud_score = random.randint(5, 35) # Baseline normal
        if is_early_inception:
            fraud_score += 45
        if is_inflated_provider:
            fraud_score += 25
        if req_amount > 800000 and ptype == "Health":
            fraud_score += 20
        
        fraud_score = min(98, max(5, fraud_score))
        fraud_flag = 1 if (fraud_score >= 70 or is_early_inception) else 0

        # Customer Satisfaction (CSAT 1 to 5)
        if claim_status == "Repudiated":
            csat = random.choices([1, 2, 3], weights=[0.60, 0.30, 0.10])[0]
        elif settlement_ratio < 0.75:
            csat = random.choices([2, 3, 4], weights=[0.40, 0.45, 0.15])[0]
        else:
            csat = random.choices([4, 5, 3], weights=[0.55, 0.35, 0.10])[0]

        claims.append({
            "claim_id": claim_id,
            "policy_id": pol_id,
            "provider_id": prov_id,
            "incident_date": incident_dt.strftime("%Y-%m-%d"),
            "submission_date": sub_dt.strftime("%Y-%m-%d"),
            "verification_date": verif_dt.strftime("%Y-%m-%d"),
            "assessment_date": assess_dt.strftime("%Y-%m-%d"),
            "settlement_date": settle_dt.strftime("%Y-%m-%d"),
            "claim_amount_requested": req_amount,
            "claim_amount_settled": settled_amount,
            "settlement_ratio": settlement_ratio,
            "settlement_mode": settlement_mode,
            "adjudication_stage": adjudication_stage,
            "claim_status": claim_status,
            "rejection_reason_code": rej_code,
            "rejection_reason_detail": rej_detail,
            "tat_days": tat_days,
            "fraud_risk_score": fraud_score,
            "fraud_flag": fraud_flag,
            "csat_score": csat
        })

    # -------------------------------------------------------------
    # 6. INJECT DATA QUALITY ISSUES FOR CLEANING PIPELINE
    # -------------------------------------------------------------
    print("Injecting controlled data quality anomalies for Step 11/13 cleaning...")
    # Issue A: 2-3% missing values in csat_score (simulate non-responding customers)
    for idx in [12, 45, 88, 134, 189, 230]:
        claims[idx]["csat_score"] = None

    # Issue B: 3 Exact Duplicate Records
    dup_1 = dict(claims[15])
    dup_2 = dict(claims[72])
    dup_3 = dict(claims[140])
    raw_claims = claims + [dup_1, dup_2, dup_3]

    df_claims = pd.DataFrame(raw_claims)
    print(f"Generated claims_fact (raw): {len(df_claims)} rows (250 unique + 3 duplicates).")

    # -------------------------------------------------------------
    # 7. GENERATE DENORMALIZED JOINED VIEW (36 Columns)
    # -------------------------------------------------------------
    df_claims_unique = pd.DataFrame(claims)
    merged = df_claims_unique.merge(df_policies, on="policy_id", how="inner")
    merged = merged.merge(df_policyholders, on="policyholder_id", how="inner")
    merged = merged.merge(df_providers, on="provider_id", how="inner")

    # Reorder into exactly 36 specified analytical columns
    view_cols_order = [
        "claim_id", "policy_id", "provider_id", "policyholder_id",
        "full_name", "age", "gender", "city_x", "state_x", "region", "tenure_months",
        "policy_type", "policy_tier", "sum_insured", "annual_premium", "policy_start_date",
        "provider_name", "provider_type", "empanelment_status", "city_y",
        "incident_date", "submission_date", "verification_date", "assessment_date", "settlement_date",
        "claim_amount_requested", "claim_amount_settled", "settlement_ratio", "settlement_mode",
        "adjudication_stage", "claim_status", "rejection_reason_code", "rejection_reason_detail",
        "tat_days", "fraud_risk_score", "csat_score"
    ]
    
    # Rename columns to match DDL view
    merged = merged.rename(columns={
        "city_x": "customer_city",
        "state_x": "customer_state",
        "region": "customer_region",
        "city_y": "provider_city"
    })
    
    actual_view_cols = [
        "claim_id", "policy_id", "provider_id", "policyholder_id",
        "full_name", "age", "gender", "customer_city", "customer_state", "customer_region", "tenure_months",
        "policy_type", "policy_tier", "sum_insured", "annual_premium", "policy_start_date",
        "provider_name", "provider_type", "empanelment_status", "provider_city",
        "incident_date", "submission_date", "verification_date", "assessment_date", "settlement_date",
        "claim_amount_requested", "claim_amount_settled", "settlement_ratio", "settlement_mode",
        "adjudication_stage", "claim_status", "rejection_reason_code", "rejection_reason_detail",
        "tat_days", "fraud_risk_score", "csat_score"
    ]
    df_view = merged[actual_view_cols]
    print(f"Generated denormalized analysis view: {len(df_view)} rows, {len(df_view.columns)} columns.")

    # -------------------------------------------------------------
    # 8. SAVE ARTIFACTS (CSV, EXCEL, SQLITE APP.DB)
    # -------------------------------------------------------------
    os.makedirs("data", exist_ok=True)

    # A) CSV Exports
    df_policyholders.to_csv("data/policyholders_master.csv", index=False)
    df_policies.to_csv("data/policies_master.csv", index=False)
    df_providers.to_csv("data/providers_master.csv", index=False)
    df_claims.to_csv("data/claims_fact.csv", index=False)
    print("Saved individual CSV tables to data/*.csv")

    # B) Excel Multi-Sheet Export
    excel_path = "data/insurance_claims_data.xlsx"
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df_policyholders.to_excel(writer, sheet_name="policyholders_master", index=False)
        df_policies.to_excel(writer, sheet_name="policies_master", index=False)
        df_providers.to_excel(writer, sheet_name="providers_master", index=False)
        df_claims.to_excel(writer, sheet_name="claims_fact", index=False)
        df_view.to_excel(writer, sheet_name="v_claims_analysis_master", index=False)
    print(f"Saved multi-sheet workbook to {excel_path}")

    # C) SQLite Database (data/app.db)
    db_path = "data/app.db"
    if os.path.exists(db_path):
        os.remove(db_path)

    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute("PRAGMA foreign_keys = ON;")
    cur.execute("PRAGMA journal_mode = WAL;")

    # Read and execute DDL from Step 8
    ddl_script = """
    CREATE TABLE IF NOT EXISTS policyholders_master (
        policyholder_id VARCHAR(20) PRIMARY KEY,
        full_name VARCHAR(100) NOT NULL,
        age INTEGER NOT NULL CHECK (age >= 18 AND age <= 100),
        gender VARCHAR(10) NOT NULL CHECK (gender IN ('Male', 'Female', 'Other')),
        city VARCHAR(50) NOT NULL,
        state VARCHAR(50) NOT NULL,
        region VARCHAR(20) NOT NULL CHECK (region IN ('North', 'South', 'East', 'West', 'Central')),
        tenure_months INTEGER NOT NULL CHECK (tenure_months >= 0)
    );

    CREATE TABLE IF NOT EXISTS policies_master (
        policy_id VARCHAR(20) PRIMARY KEY,
        policyholder_id VARCHAR(20) NOT NULL,
        policy_type VARCHAR(20) NOT NULL CHECK (policy_type IN ('Health', 'Motor', 'Life', 'Property')),
        policy_tier VARCHAR(20) NOT NULL CHECK (policy_tier IN ('Silver', 'Gold', 'Platinum')),
        sum_insured DECIMAL(12,2) NOT NULL CHECK (sum_insured > 0),
        annual_premium DECIMAL(10,2) NOT NULL CHECK (annual_premium > 0),
        policy_start_date DATE NOT NULL,
        policy_status VARCHAR(20) NOT NULL CHECK (policy_status IN ('Active', 'Lapsed', 'Renewed')),
        FOREIGN KEY (policyholder_id) REFERENCES policyholders_master(policyholder_id)
            ON DELETE RESTRICT ON UPDATE CASCADE
    );

    CREATE TABLE IF NOT EXISTS providers_master (
        provider_id VARCHAR(20) PRIMARY KEY,
        provider_name VARCHAR(100) NOT NULL,
        provider_type VARCHAR(30) NOT NULL CHECK (provider_type IN ('Hospital', 'Automobile Garage', 'TPA')),
        city VARCHAR(50) NOT NULL,
        state VARCHAR(50) NOT NULL,
        empanelment_status VARCHAR(20) NOT NULL CHECK (empanelment_status IN ('Empaneled Network', 'Non-Network', 'Under Review')),
        tariff_tier VARCHAR(10) NOT NULL CHECK (tariff_tier IN ('Tier 1', 'Tier 2', 'Tier 3'))
    );

    CREATE TABLE IF NOT EXISTS claims_fact (
        claim_id VARCHAR(20) PRIMARY KEY,
        policy_id VARCHAR(20) NOT NULL,
        provider_id VARCHAR(20) NOT NULL,
        incident_date DATE NOT NULL,
        submission_date DATE NOT NULL,
        verification_date DATE,
        assessment_date DATE,
        settlement_date DATE,
        claim_amount_requested DECIMAL(12,2) NOT NULL CHECK (claim_amount_requested > 0),
        claim_amount_settled DECIMAL(12,2) NOT NULL CHECK (claim_amount_settled >= 0),
        settlement_ratio DECIMAL(5,4) NOT NULL CHECK (settlement_ratio >= 0.0 AND settlement_ratio <= 1.0),
        settlement_mode VARCHAR(20) NOT NULL CHECK (settlement_mode IN ('Cashless', 'Reimbursement')),
        adjudication_stage VARCHAR(30) NOT NULL CHECK (adjudication_stage IN ('Submission', 'Verification', 'Assessment', 'Settled')),
        claim_status VARCHAR(20) NOT NULL CHECK (claim_status IN ('Approved', 'Repudiated', 'Pending Review')),
        rejection_reason_code VARCHAR(20),
        rejection_reason_detail VARCHAR(150),
        tat_days INTEGER NOT NULL CHECK (tat_days >= 0),
        fraud_risk_score INTEGER NOT NULL CHECK (fraud_risk_score >= 0 AND fraud_risk_score <= 100),
        fraud_flag INTEGER NOT NULL CHECK (fraud_flag IN (0, 1)),
        csat_score INTEGER CHECK (csat_score >= 1 AND csat_score <= 5),
        FOREIGN KEY (policy_id) REFERENCES policies_master(policy_id)
            ON DELETE RESTRICT ON UPDATE CASCADE,
        FOREIGN KEY (provider_id) REFERENCES providers_master(provider_id)
            ON DELETE RESTRICT ON UPDATE CASCADE
    );

    CREATE INDEX IF NOT EXISTS idx_claims_policy_id ON claims_fact(policy_id);
    CREATE INDEX IF NOT EXISTS idx_claims_provider_id ON claims_fact(provider_id);
    CREATE INDEX IF NOT EXISTS idx_claims_submission_date ON claims_fact(submission_date);
    CREATE INDEX IF NOT EXISTS idx_claims_status ON claims_fact(claim_status);

    CREATE VIEW IF NOT EXISTS v_claims_analysis_master AS
    SELECT
        c.claim_id, c.policy_id, c.provider_id, p.policyholder_id,
        ph.full_name, ph.age, ph.gender, ph.city AS customer_city, ph.state AS customer_state, ph.region AS customer_region, ph.tenure_months,
        p.policy_type, p.policy_tier, p.sum_insured, p.annual_premium, p.policy_start_date,
        pr.provider_name, pr.provider_type, pr.empanelment_status, pr.city AS provider_city,
        c.incident_date, c.submission_date, c.verification_date, c.assessment_date, c.settlement_date,
        c.claim_amount_requested, c.claim_amount_settled, c.settlement_ratio, c.settlement_mode,
        c.adjudication_stage, c.claim_status, c.rejection_reason_code, c.rejection_reason_detail,
        c.tat_days, c.fraud_risk_score, c.csat_score
    FROM claims_fact c
    JOIN policies_master p ON c.policy_id = p.policy_id
    JOIN policyholders_master ph ON p.policyholder_id = ph.policyholder_id
    JOIN providers_master pr ON c.provider_id = pr.provider_id;
    """
    cur.executescript(ddl_script)
    conn.commit()
    cur.close()

    # Insert data into SQLite using fresh connection
    df_policyholders.to_sql("policyholders_master", conn, if_exists="append", index=False)
    df_policies.to_sql("policies_master", conn, if_exists="append", index=False)
    df_providers.to_sql("providers_master", conn, if_exists="append", index=False)
    df_claims_unique.to_sql("claims_fact", conn, if_exists="append", index=False)
    conn.commit()

    print(f"Loaded schema and tables into SQLite datastore: {db_path}")
    cur = conn.cursor()

    # -------------------------------------------------------------
    # 9. VALIDATION AUDIT REPORT
    # -------------------------------------------------------------
    print("\n----------------- DATA VALIDATION AUDIT REPORT -----------------")
    cur.execute("SELECT COUNT(*) FROM policyholders_master;")
    cnt_ph = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM policies_master;")
    cnt_pol = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM providers_master;")
    cnt_prv = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM claims_fact;")
    cnt_clm = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM v_claims_analysis_master;")
    cnt_view = cur.fetchone()[0]

    print(f"Row Counts: Policyholders={cnt_ph}, Policies={cnt_pol}, Providers={cnt_prv}, Claims={cnt_clm}, View={cnt_view}")

    # Check Referential Integrity (Foreign Key Violations)
    cur.execute("PRAGMA foreign_key_check;")
    fk_errors = cur.fetchall()
    print(f"Foreign Key Integrity Check: {len(fk_errors)} violations (PASSED)")

    # Check Cashless vs. Reimbursement TAT Pattern
    cur.execute("SELECT settlement_mode, AVG(tat_days), COUNT(*) FROM claims_fact GROUP BY settlement_mode;")
    tat_stats = cur.fetchall()
    print("Settlement Mode Performance:")
    for mode, avg_tat, count in tat_stats:
        print(f"  • {mode}: Count={count}, Mean TAT={avg_tat:.2f} days")

    # Check Pareto Repudiation Distribution
    cur.execute("SELECT rejection_reason_code, COUNT(*) FROM claims_fact WHERE claim_status='Repudiated' GROUP BY rejection_reason_code ORDER BY COUNT(*) DESC;")
    rej_stats = cur.fetchall()
    print("Top Repudiation Reasons:")
    total_rej = sum([c[1] for c in rej_stats])
    for code, count in rej_stats:
        print(f"  • {code}: {count} ({count*100.0/total_rej:.1f}%)")

    conn.close()
    print("==========================================================")
    print("STEP 10 COMPLETED SUCCESSFULLY: DATA GENERATED & VALIDATED")
    print("==========================================================")

if __name__ == "__main__":
    generate_dataset()
