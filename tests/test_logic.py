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

def test_spreadsheet_formula_integrity():
    """Verify that the generated Excel template contains the correct formula targets after refactors"""
    filename = "tests/data/integrity_check.xlsx"
    # Ensure constants folder exists for generation
    create_tax_workbook(status="Single", dependents=0, year=2026)
    os.rename("Estimated_Tax_Calculator_Template.xlsx", filename)
    
    wb = openpyxl.load_workbook(filename)
    ds = wb["Dashboard"]
    
    # Verification of total liability ref (B48 in latest refactor)
    assert ds["A48"].value == "Total Federal Liability"
    formula = ds["B48"].value
    # It should include Fed Ord Tax (B31), CG Tax (B32), NIIT (B43), Addl Med (B45), minus CTC (B47)
    assert "B31" in formula
    assert "B32" in formula
    assert "B43" in formula
    assert "B45" in formula
    assert "B47" in formula
    
    # Cleanup
    os.remove(filename)
