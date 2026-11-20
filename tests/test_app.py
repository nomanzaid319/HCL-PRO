"""
Technical Test Suite 3: UI, Configuration & Smoke Tests
Author: Noman Zaid (Roll No: 2500320140103, MCA 2nd Year Data Analytics, ABES)
Test Cases: TC-T-12, TC-T-13, TC-T-14
"""

import os
import pytest

ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

def test_tc_t_12_ui_module_imports():
    """TC-T-12: Verify that the UI component module imports cleanly and exposes render functions."""
    import src.ui as ui
    assert hasattr(ui, "render_header"), "ui module must have render_header"
    assert hasattr(ui, "render_kpi_row"), "ui module must have render_kpi_row"
    assert hasattr(ui, "render_insight_box"), "ui module must have render_insight_box"

def test_tc_t_13_config_toml_exists():
    """TC-T-13: Verify that .streamlit/config.toml exists with corporate theme colors."""
    config_path = os.path.join(ROOT_DIR, ".streamlit", "config.toml")
    assert os.path.exists(config_path), f"Config file not found at {config_path}"
    
    with open(config_path, "r", encoding="utf-8") as f:
        content = f.read()
    assert "primaryColor" in content, "config.toml must specify primaryColor"
    assert "#1B365D" in content, "config.toml must use corporate navy #1B365D"

def test_tc_t_14_data_files_exist():
    """TC-T-14: Verify that all required database, CSV, and Excel data files exist in data/."""
    required_files = [
        "app.db",
        "claims_fact.csv",
        "policies_master.csv",
        "policyholders_master.csv",
        "providers_master.csv",
        "insurance_claims_data.xlsx"
    ]
    for filename in required_files:
        filepath = os.path.join(ROOT_DIR, "data", filename)
        assert os.path.exists(filepath), f"Required data artifact missing: {filepath}"
        assert os.path.getsize(filepath) > 0, f"Data file {filename} is empty."
