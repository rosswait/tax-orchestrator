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

def test_spreadsheet_formula_integrity():
    """Verify that the generated Excel template contains the correct formula targets after refactors"""
    filename = "tests/data/integrity_check.xlsx"
    # Use a dedicated test filename to avoid clobbering the user's main xlsx
    test_output = "Estimated_Tax_Calculator_TEST.xlsx"
    create_tax_workbook(status="Single", dependents=0, year=2026, filename=test_output)
    os.rename(test_output, filename)
    
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

def test_filing_date_validity_check():
    """Test F: Verify the Active Warning for Filing Date out-of-range detection"""
    filename = "tests/data/validity_test.xlsx"
    # Generate for 2026
    test_output = "Estimated_Tax_Calculator_TEST.xlsx"
    create_tax_workbook(status="Single", dependents=0, year=2026, filename=test_output)
    os.rename(test_output, filename)
    
    wb = openpyxl.load_workbook(filename)
    ds = wb["Dashboard"]
    
    # 1. Formula Presence
    assert ds["I32"].value == "Filing Date Validity:"
    assert ds["J32"].value.startswith("=IF(OR(B9 < DATE(B8,1,1)")
    
    # Clean up
    os.remove(filename)

def test_hoh_dependent_warning():
    """Test G: Verify the Active Warning for HoH with 0 dependents"""
    filename = "tests/data/hoh_test.xlsx"
    test_output = "Estimated_Tax_Calculator_TEST.xlsx"
    create_tax_workbook(status="HoH", dependents=0, year=2026, filename=test_output)
    os.rename(test_output, filename)
    
    wb = openpyxl.load_workbook(filename)
    ds = wb["Dashboard"]
    
    # 1. Formula Presence
    assert ds["I33"].value == "HOH Dependents:"
    assert "B2=\"HoH\"" in ds["J33"].value
    assert "B3=0" in ds["J33"].value
    
    # Clean up
    os.remove(filename)
def test_federal_formula_parity():
    """Test H: Verify that Federal Tax formulas are identical in both CA and Fed-only modes"""
    ca_file = "tests/data/parity_ca.xlsx"
    fed_file = "tests/data/parity_fed.xlsx"
    
    # Generate both versions
    create_tax_workbook(status="Single", dependents=1, year=2026, fed_only=False, filename=ca_file)
    create_tax_workbook(status="Single", dependents=1, year=2026, fed_only=True, filename=fed_file)
    
    wb_ca = openpyxl.load_workbook(ca_file)
    wb_fed = openpyxl.load_workbook(fed_file)
    ds_ca = wb_ca["Dashboard"]
    ds_fed = wb_fed["Dashboard"]
    
    # Key Federal Cells that must be identical
    federal_cells = [
        "B21", # Fed W-2 Wages
        "B25", # Fed AGI
        "B31", # Fed Ord Tax
        "B48", # Total Fed Liability
        "B51", # Fed Target
        "J2",  # Fed Due Now
        "J6",  # Fed Status
    ]
    
    for coord in federal_cells:
        formula_ca = ds_ca[coord].value
        formula_fed = ds_fed[coord].value
        assert formula_ca == formula_fed, f"Federal formula mismatch at {coord}: CA='{formula_ca}', Fed='{formula_fed}'"
    
    # Cleanup
    os.remove(ca_file)
    os.remove(fed_file)
