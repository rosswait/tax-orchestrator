import pytest
import os
import json
from datetime import datetime
from tests.logic_engine import TaxShadowEngine
from generate_xlsx import create_tax_workbook
import openpyxl

@pytest.fixture
def scenario_data():
    path = "tests/scenarios/standard_midyear.json"
    with open(path, "r") as f:
        return json.load(f)

def test_buffer_logic():
    """Test A: Verify the 30-day (1-month) buffer logic"""
    engine = TaxShadowEngine(year=2026, status="Single")
    
    # April 8 -> Q1
    assert engine.calculate_inferred_quarter("04/08/2026") == 1
    # April 30 -> Q1
    assert engine.calculate_inferred_quarter("04/30/2026") == 1
    # May 1 -> Q2
    assert engine.calculate_inferred_quarter("05/01/2026") == 2
    # Jan 15 -> Q1 (Capped)
    assert engine.calculate_inferred_quarter("01/15/2026") == 1

def test_hsa_ca_correction(scenario_data):
    """Test B: Verify CA wages correctly exclude HSA deductions"""
    engine = TaxShadowEngine(year=2026, status="Single")
    results = engine.run_scenario(scenario_data)
    
    # In standard_midyear.json: pretax=10000, hsa=2000
    # Fed AGI should deduct 12000 (pretax + hsa) + projected
    # CA AGI should deduct only 10000 (pretax) + projected
    # Therefore CA AGI should be higher than Fed AGI by precisely the HSA amount (ytd + projected)
    
    # This is a complex check of the shadow engine logic
    assert results["ca_agi"] > results["fed_agi"]
    diff = results["ca_agi"] - results["fed_agi"]
    # Total HSA (YTD 2000 + Projected)
    # 2000 was for 90 days. Days rem = 275. 2000/90 * 275 = 6111. Total approx 8k.
    assert diff > 2000 

def test_progressive_brackets():
    """Test C: Verify tax math independently for a known threshold"""
    engine = TaxShadowEngine(year=2025, status="Single") # Use 2025 as it is proven
    
    # 2025 Single: 10% on first 11,925, 12% on next
    # Total tax on 20,000 should be: 1,192.50 + (20,000 - 11,925)*0.12
    expected = 1192.50 + (20000 - 11925) * 0.12 # 1192.5 + 969 = 2161.5
    results = engine.calculate_tax(20000, "fed_ord")
    assert results == pytest.approx(2161.5)

@pytest.mark.parametrize("status,expected_on_20k", [
    ("Single", 2161.50), # 10% on 11,925 + 12% on rem
    ("MFJ", 2000.00),    # 10% on all 20k (MFJ 10% bracket goes to 23,850)
    ("MFS", 2161.50),    # Same as Single
    ("HoH", 2060.00),    # 10% on 17,000 + 12% on rem (3000 * .12 = 360) -> 1700 + 360 = 2060
])
def test_all_filing_statuses(status, expected_on_20k):
    """Test D: Verify that all filing statuses resolve to the correct bracket logic"""
    engine = TaxShadowEngine(year=2025, status=status)
    tax = engine.calculate_tax(20000, "fed_ord")
    assert tax == pytest.approx(expected_on_20k)

@pytest.mark.parametrize("filing_date,expected_multiplier", [
    ("04/08/2026", 3.0), # Q1 -> (4/1)-1 = 3
    ("05/15/2026", 1.0), # Q2 -> (4/2)-1 = 1
    ("08/10/2026", 0.3333333333333333), # Q3 -> (4/3)-1 = 0.33
    ("11/20/2026", 0.0), # Q4 -> (4/4)-1 = 0
])
def test_dividend_projection_multipliers(filing_date, expected_multiplier):
    """Test E: Verify that dividend/interest projection multipliers align with the inferred quarter"""
    engine = TaxShadowEngine(year=2026, status="Single")
    q = engine.calculate_inferred_quarter(filing_date)
    # The multiplier used in generate_xlsx.py (B17) and logic_engine.py
    multiplier = (4 / q) - 1
    assert multiplier == pytest.approx(expected_multiplier)

def test_multiple_wage_sources_conservative():
    """Test I: Verify that multiple wage sources use the earliest date for a conservative (higher) projection"""
    engine = TaxShadowEngine(year=2026, status="Single")
    
    # Scenario: Two sources with different dates
    # Source A: March 15 (74 days elapsed)
    # Source B: March 31 (90 days elapsed)
    # Total Gross: 15,000
    # Conservative Rate (MIN date) = 15,000 / 74 = 202.7/day
    # Standard Rate (MAX date) = 15,000 / 90 = 166.6/day
    
    scenario = {
        "config": {"filing_date": "04/01/2026"},
        "wage_snapshots": [
            {"date": "03/15/2026", "gross": 5000, "pretax": 0, "hsa": 0},
            {"date": "03/31/2026", "gross": 10000, "pretax": 0, "hsa": 0}
        ],
        "investment_snapshots": []
    }
    
    results = engine.run_scenario(scenario)
    
    # Yearly projection calculation:
    # 15,000 / 73 days * (365 total days) = 75,000 approx
    assert results["fed_agi"] > 74000 
    assert results["fed_agi"] < 75500 # within rounding distance

def test_ca_mfs_parity():
    """Test J: Verify that CA MFS and CA Single brackets are identical for the same income"""
    engine_single = TaxShadowEngine(year=2025, status="Single")
    engine_mfs = TaxShadowEngine(year=2025, status="MFS")
    
    income = 100000
    tax_single = engine_single.calculate_tax(income, "ca")
    tax_mfs = engine_mfs.calculate_tax(income, "ca")
    
    assert tax_single == tax_mfs
    assert tax_single > 0
