import json
import os
from datetime import datetime
import math

class TaxShadowEngine:
    def __init__(self, year=2026, status="Single"):
        self.year = year
        self.status = status
        self.constants = self._load_constants(year)
        
    def _load_constants(self, target_year):
        base_dir = "constants"
        available_years = sorted([int(d) for d in os.listdir(base_dir) if d.isdigit()], reverse=True)
        logic_year = target_year if target_year in available_years else available_years[0]
        data_path = os.path.join(base_dir, str(logic_year))
        
        data = {}
        for key, filename in [("fed_ord", "federal_ord.json"), ("fed_cg", "federal_cg.json"), 
                             ("ca", "ca_brackets.json"), ("surtaxes", "surtaxes.json")]:
            with open(os.path.join(data_path, filename), "r") as f:
                data[key] = json.load(f)
        return data

    def calculate_inferred_quarter(self, filing_date_str):
        dt = datetime.strptime(filing_date_str, "%m/%d/%Y")
        # 30-day buffer logic: ROUNDUP((MONTH-1)/3, 0)
        q = math.ceil((dt.month - 1) / 3)
        return max(1, min(4, q))

    def get_bracket(self, income, table_key):
        # Tables are: [Status, Floor, Base, Rate, ...]
        # Filter by status
        relevant = [r for r in self.constants[table_key] if r[0] == self.status]
        # Sort by floor descending
        relevant.sort(key=lambda x: x[1], reverse=True)
        for row in relevant:
            if income >= row[1]:
                return row
        return relevant[-1] # fallback

    def calculate_tax(self, taxable_income, table_key):
        bracket = self.get_bracket(taxable_income, table_key)
        floor = bracket[1]
        base_tax = bracket[2]
        rate = bracket[3]
        return base_tax + (taxable_income - floor) * rate

    def run_scenario(self, scenario):
        config = scenario["config"]
        filing_date = config["filing_date"]
        q_inferred = self.calculate_inferred_quarter(filing_date)
        
        # 1. Projections
        snapshot_dates = [datetime.strptime(s["date"], "%m/%d/%Y") for s in scenario["wage_snapshots"]]
        dt_conservative = min(snapshot_dates)
        days_elapsed = (dt_conservative - datetime(self.year, 1, 1)).days
        days_total = 365 
        days_remaining = (datetime(self.year, 12, 31) - dt_conservative).days
        
        ytd_gross = sum(s["gross"] for s in scenario["wage_snapshots"])
        ytd_pretax = sum(s["pretax"] for s in scenario["wage_snapshots"])
        ytd_hsa = sum(s["hsa"] for s in scenario["wage_snapshots"])
        
        daily_rate = ytd_gross / max(1, days_elapsed)
        rem_wage = daily_rate * days_remaining
        
        total_projected_wage = ytd_gross + (rem_wage * config.get("future_weight", 1.0)) + config.get("manual_offset", 0)
        
        # Deductions
        daily_deduct = (ytd_pretax + ytd_hsa) / max(1, days_elapsed)
        rem_deduct = daily_deduct * days_remaining
        
        # Fed AGI
        fed_w2_state = total_projected_wage - (ytd_pretax + ytd_hsa + rem_deduct * config.get("future_weight", 1.0))
        
        ytd_div = sum(s["dividends_interest"] for s in scenario["investment_snapshots"])
        proj_div = ytd_div * ((4 / q_inferred) - 1)
        ytd_stg = sum(s["short_term_gains"] for s in scenario["investment_snapshots"])
        ytd_ltg = sum(s["long_term_gains"] for s in scenario["investment_snapshots"])
        
        inv_ord = ytd_div + proj_div + ytd_stg
        fed_agi = fed_w2_state + inv_ord + ytd_ltg
        
        # Fed Tax
        fed_std = [r for r in self.constants["fed_ord"] if r[0] == self.status][0][4]
        fed_deduction = max(fed_std, config.get("itemized_fed", 0))
        fed_taxable_ord = max(0, fed_agi - fed_deduction - ytd_ltg)
        
        fed_ord_tax = self.calculate_tax(fed_taxable_ord, "fed_ord")
        
        # CG Tax (Blunt lookup for simplicity)
        cg_table = [r for r in self.constants["fed_cg"] if r[0] == self.status]
        cg_table.sort(key=lambda x: x[1], reverse=True)
        cg_rate = 0.15 # default
        for r in cg_table:
            if (fed_taxable_ord + ytd_ltg) >= r[1]:
                cg_rate = r[2]
                break
        fed_cg_tax = ytd_ltg * cg_rate
        
        # CA Tax
        # CA Wage = Total Wage - Fed/State Deductions (excluding HSA)
        # Note: B21 formula in Excel is: =B19 - SUM('Wage Snapshots'!C:C) - (SUM('Wage Snapshots'!C:C)/MAX(1, SUM('Wage Snapshots'!C:D))) * B15 * B13
        # Simplification: CA Wages = Total Wage - (YTD Pretax + Remaining Pretax)
        ca_rem_pretax = (ytd_pretax / max(1, days_elapsed)) * days_remaining
        ca_w2_state = total_projected_wage - (ytd_pretax + ca_rem_pretax * config.get("future_weight", 1.0))
        ca_agi = ca_w2_state + inv_ord + ytd_ltg
        
        ca_std = [r for r in self.constants["ca"] if r[0] == self.status][0][4]
        ca_deduction = max(ca_std, config.get("itemized_ca", 0))
        ca_taxable = max(0, ca_agi - ca_deduction)
        ca_tax = self.calculate_tax(ca_taxable, "ca")
        
        return {
            "inferred_quarter": q_inferred,
            "fed_agi": fed_agi,
            "ca_agi": ca_agi,
            "fed_liability": fed_ord_tax + fed_cg_tax, # simplistic for surtax check
            "ca_liability": ca_tax
        }

if __name__ == "__main__":
    # Self-test with the boilerplate
    with open("tests/scenarios/standard_midyear.json", "r") as f:
        scen = json.load(f)
    engine = TaxShadowEngine(year=scen["config"]["year"], status=scen["config"]["status"])
    results = engine.run_scenario(scen)
    print(json.dumps(results, indent=2))
