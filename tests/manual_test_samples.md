# Manual Verification Samples

You can use the samples below to verify your Estimated Tax Calculator. Copy the markdown tables and paste them into the corresponding tabs of your generated spreadsheet.

---

## Scenario 1: Standard Mid-Year (Single Filer)
**Configuration**:
- Filing Status: `Single`
- Filing (Current) Date: `04/08/2026` (Inferred Quarter: **1**)

### Wage Snapshots
*Copy and paste into row 2 of 'Wage Snapshots'*

| Date | Gross W-2 Income | Pre-tax Deductions | HSA Contributions | Fed Tax Withheld | CA Tax Withheld | FICA/Med/SDI |
|---|---|---|---|---|---|---|
| 03/31/2026 | 75000.00 | 10000.00 | 2000.00 | 12000.00 | 4000.00 | 5700.00 |

### Investment Income Snapshots
*Copy and paste into row 2 of 'Investment Income Snapshots'*

| Quarter | Entity | Dividends & Interest | Short-Term Gains | Long-Term Gains |
|---|---|---|---|---|
| Q1 | Primary Brokerage | 1000.00 | 500.00 | 2000.00 |

---

## Expected Results (Summary)
If you enter the data above, you should see approximately the following on your **Dashboard**:

- **B10 (Inferred Quarter)**: `1`
- **B17 (Remaining Year Int/Div)**: `$3,000.00`
- **B15 (Remaining Year Wages)**: `~$228,000.00` (Assumes ~90 days elapsed)
- **I28 (HSA Verification)**: `✅ HSA Corrected (CA)`

---

## How to Test
1. Run `./venv/bin/python generate_xlsx.py --status Single`.
2. Open the file in Excel.
3. Paste the table data above into the appropriate sheets.
4. Verify the Dashboard matches the expected summary.
