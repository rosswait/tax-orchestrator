# 📊 Estimated Tax Calculator

The **Estimated Tax Calculator** is an automated Federal and California tax projection engine. It generates an Excel workbook that allows you to track your estimated tax liability and payments throughout the year with minimal manual entry.

It also includes integrated instructions for an LLM agent or skill to translate paystub and brokerage statements directly into the required format for the workbook.

---

## 🚀 Key Features

### 1. Smart Status Dashboard
The workbook features a dedicated **Action Center** and **Diagnostics Panel** that answers:
- **"How much do I owe today?"** - Dynamic detection of the next quarterly deadline (Apr 15, Jun 15, Sep 15, Jan 15) and required payments.
- **"Am I safe?"** - Automatic calculation of Safe Harbor (110% Prior Year) vs. 90% Forecast targets, prioritizing the most conservative baseline.
- **"What is my tax health?"** - Real-time Effective Tax Rates, Marginal Brackets, and Itemization vs. Standard audits.

### 2. Automated Tax Year Logic
The engine automatically infers the target tax year based on the current date:
- **January Buffer**: Correctly defaults to the previous tax year during the Jan 1st - 30th "Q4 Hangover" period, ensuring accurate finalization of the previous year's taxes.
- **Bracket Staleness Handling**: If no data exists for the projected year, the engine uses the most recent available constants as a safe, conservative proxy and triggers an alert.

### 3. Decoupled Tax Constants
Tax laws are stored in the `constants/` directory, organized by year (e.g., `constants/2025/`). 
- **Federal**: Ordinary and Capital Gains brackets for all filing statuses (Single, MFJ, MFS, HoH).
- **California**: progressive logic, including the 1% Mental Health Services Surcharge.
- **Surtaxes**: Integrated NIIT, Additional Medicare, and CTC phase-outs.

### 4. Pro-Active Warning System
The dashboard automatically flags:
- **Stale Snapshots**: Alerts if your latest paystub is >30 days old.
- **Bracket Staleness**: Specific Federal and CA alerts when using historical data as a proxy for the next year.
- **HSA Audit**: Verification that state-tax deductions are correctly excluded for California.

---

## 🛠️ Usage

### 1. Generation
Run the Python script to generate a clean, formatted `.xlsx` template.

```bash
# Generate for current inferred year (default), Single Filer, 0 Dependents
./venv/bin/python generate_xlsx.py --status Single --dependents 0
```

### 2. Data Entry (Snapshots)
Instead of entering every transaction, use the **Snapshots** methodology:
- **Wage Snapshots**: Enter your latest YTD paystub totals. The engine will pro-rate the rest of the year automatically based on days elapsed.
- **Investment Snapshots**: Enter your cumulative YTD statements. 

### 3. Agent-Assisted Parsing
Use the included `parsing_agent_instructions.md` (also found in the Excel tab "Parsing Instructions for Agents") with your preferred LLM to extract snapshot-ready data directly from screenshots or statements.  This can be imlemented as a skill, prompt, or preconfigured prompt (eg. Gemini Gem).

---

## 📋 Sheet Structure
- **Dashboard**: High-level status board, Action Center, and calculation engine.
- **Wage Snapshots**: Input ledger for W-2 income records.
- **Investment Income Snapshots**: Input ledger for brokerage records.
- **Tax Constants**: Self-documenting table of the Federal and CA data used in the simulation.
- **Parsing Instructions for Agents**: Embedded prompt for use with LLMs.

---

## 🏛️ Requirements
- Python 3.9+
- `openpyxl`: For Excel workbook generation and formatting.
