# 📊 Estimated Tax Calculator

The **Estimated Tax Calculator** is an automated Federal and California tax projection engine. It generates a Excel workbook that allows you to track your estimated tax liability and payments throughout the year. 

It also includes markdown instructions for an LLM agent to translate paystub and brokerage statements (or snapshots) into the required format for the Excel workbook.

---

## 🚀 Key Features

### 1. Smart Status Dashboard
The workbook features a dedicated **Action Center** and **Diagnostics Panel** that answers:
- **"How much do I owe today?"** - Dynamic detection of the next quarterly deadline (Apr 15, Jun 15, Sep 15, Jan 15).
- **"Am I safe?"** - Automatic calculation of Safe Harbor (110% Prior Year) vs. 90% Forecast targets.
- **"What is my tax health?"** - Real-time Effective Tax Rates, Marginal Brackets, and Itemization vs. Standard audits.

### 2. Year-Agnostic Engine
Supports any tax year via a simple CLI parameter. All internal formulas pivot from a single "Tax Year" configuration cell, correctly handling year-over-year date shifts (including Q4 January deadlines).

### 3. Progressive Tax Logic
- **Federal**: 2026 Ordinary and Capital Gains brackets for all filing statuses (Single, MFJ, MFS, HoH).
- **California**: CA FTB progressive logic, including the 1% Mental Health Services Surcharge for high-income earners.
- **Surtaxes**: Integrated Federal NIIT, Additional Medicare, and Child Tax Credit phase-outs.

### 4. Pro-Active Warning System
The dashboard automatically flags:
- **Stale Snapshots**: Alerts if your latest paystub is >30 days old.
- **Prior Year Missing**: Warns when safe-harbor targets are unverified.
- **HSA Audit**: Verification of state-tax deduction exclusions.

---

## 🛠️ Usage

### 1. Generation
Run the Python script to generate a clean, formatted `.xlsx` template pre-configured for your situation.

```bash
# Generate for 2026, Single Filer, 0 Dependents
./venv/bin/python generate_xlsx.py --year 2026 --status Single --dependents 0
```

### 2. Data Entry (Snapshots)
Instead of entering every transaction, use the **Snapshots** tabs:
- **Wage Snapshots**: Enter your latest paystub total gross, deductions, and withholdings. The engine will pro-rate the rest of the year automatically based on days elapsed.
- **Investment Snapshots**: Enter your YTD statements for dividends and capital gains. 

### 3. AI Enrichment (Optional)
Use the included `gem_instructions.md` with Gemini to extract snapshot-ready data directly from your PDF paystubs or brokerage statements.

---

## 📋 Sheet Structure
- **Dashboard**: High-level status board, Action Center, Diagnostics, and full tax engine.
- **Wage Snapshots**: Input ledger for W-2 income records.
- **Investment Income Snapshots**: Input ledger for brokerage/capital gains records.
- **Tax Constants**: The source-of-truth for all Federal and CA tax brackets.

---

## 🏛️ Requirements
- Python 3.9+
- `openpyxl`: For Excel workbook generation and formatting.
