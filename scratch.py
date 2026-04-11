import json
from tests.logic_engine import TaxShadowEngine

scenario = {
    "config": {
        "filing_date": "04/03/2026"
    },
    "wage_snapshots": [
        {"date": "04/03/2026", "gross": 50000, "pretax": 2000, "hsa": 400},
        {"date": "04/03/2026", "gross": 12500, "pretax": 500, "hsa": 100}
    ],
    "investment_snapshots": []
}

engine = TaxShadowEngine(year=2026, status="Single")
results = engine.run_scenario(scenario)
print(json.dumps(results, indent=2))
