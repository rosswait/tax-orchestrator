import argparse
import openpyxl
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from datetime import datetime

def create_tax_workbook(status="Single", dependents=0, year=2026):
    wb = Workbook()
    
    # Define common format strings for Google Sheets parsing
    FORMAT_CURRENCY = '"$"#,##0.00'
    FORMAT_DATE = 'mm/dd/yyyy'
    FORMAT_PERCENT = '0.00%'

    # --- 0. Shared Tax Logic Constants ---
    fed_ord_brackets = [
        ["Single", 0, 0, 0.10, 15000], ["Single", 11925, 1192.50, 0.12, 15000], ["Single", 48475, 5578.50, 0.22, 15000], ["Single", 103350, 17651, 0.24, 15000], ["Single", 197300, 40199, 0.32, 15000], ["Single", 250525, 57231, 0.35, 15000], ["Single", 626350, 188770, 0.37, 15000],
        ["MFJ", 0, 0, 0.10, 30000], ["MFJ", 23850, 2385, 0.12, 30000], ["MFJ", 96950, 11157, 0.22, 30000], ["MFJ", 206700, 35302, 0.24, 30000], ["MFJ", 394600, 80398, 0.32, 30000], ["MFJ", 501050, 114462, 0.35, 30000], ["MFJ", 751600, 202154.50, 0.37, 30000],
        ["MFS", 0, 0, 0.10, 15000], ["MFS", 11925, 1192.50, 0.12, 15000], ["MFS", 48475, 5578.50, 0.22, 15000], ["MFS", 103350, 17651, 0.24, 15000], ["MFS", 197300, 40199, 0.32, 15000], ["MFS", 250525, 57231, 0.35, 15000], ["MFS", 375800, 101077.25, 0.37, 15000],
        ["HoH", 0, 0, 0.10, 22500], ["HoH", 17000, 1700, 0.12, 22500], ["HoH", 64850, 7442, 0.22, 22500], ["HoH", 103350, 15912, 0.24, 22500], ["HoH", 197300, 38460, 0.32, 22500], ["HoH", 250500, 55484, 0.35, 22500], ["HoH", 626350, 187031.50, 0.37, 22500]
    ]
    fed_cg_brackets = [
        ["Single", 0, 0.0], ["Single", 47025, 0.15], ["Single", 518900, 0.20],
        ["MFJ", 0, 0.0], ["MFJ", 96700, 0.15], ["MFJ", 600050, 0.20],
        ["MFS", 0, 0.0], ["MFS", 48350, 0.15], ["MFS", 300000, 0.20],
        ["HoH", 0, 0.0], ["HoH", 64750, 0.15], ["HoH", 566700, 0.20]
    ]
    ca_brackets = [
        ["Single", 0, 0, 0.01, 5706, 1000000], ["Single", 11079, 110.79, 0.02, 5706, 1000000], ["Single", 26264, 414.49, 0.04, 5706, 1000000], ["Single", 41452, 1022.01, 0.06, 5706, 1000000], ["Single", 57542, 1987.41, 0.08, 5706, 1000000], ["Single", 72724, 3201.97, 0.093, 5706, 1000000], ["Single", 371479, 31036.19, 0.103, 5706, 1000000], ["Single", 445771, 38688.19, 0.113, 5706, 1000000], ["Single", 742953, 72269.75, 0.123, 5706, 1000000],
        ["MFJ", 0, 0, 0.01, 11412, 1000000], ["MFJ", 22158, 221.58, 0.02, 11412, 1000000], ["MFJ", 52528, 828.98, 0.04, 11412, 1000000], ["MFJ", 82904, 2044.02, 0.06, 11412, 1000000], ["MFJ", 115084, 3974.82, 0.08, 11412, 1000000], ["MFJ", 145448, 6403.94, 0.093, 11412, 1000000], ["MFJ", 742958, 61972.37, 0.103, 11412, 1000000], ["MFJ", 891542, 77276.52, 0.113, 11412, 1000000], ["MFJ", 1485906, 144439.65, 0.123, 11412, 1000000],
        ["HoH", 0, 0, 0.01, 11412, 1000000], ["HoH", 22173, 221.73, 0.02, 11412, 1000000], ["HoH", 52530, 828.87, 0.04, 11412, 1000000], ["HoH", 67716, 1436.31, 0.06, 11412, 1000000], ["HoH", 83804, 2401.59, 0.08, 11412, 1000000], ["HoH", 98990, 3616.47, 0.093, 11412, 1000000], ["HoH", 505208, 41394.74, 0.103, 11412, 1000000], ["HoH", 606251, 51802.17, 0.113, 11412, 1000000], ["HoH", 1010417, 97472.93, 0.123, 11412, 1000000]
    ]
    surtaxes = [["Single", 200000, 200000, 200000], ["MFJ", 250000, 250000, 400000], ["MFS", 125000, 125000, 200000], ["HoH", 200000, 200000, 200000]]

    # --- 1. Instructions Tab (First) ---
    ws_instr = wb.active; ws_instr.title = "Instructions"
    ws_instr.append(["Quick Start Guide"])
    ws_instr.append([""])
    ws_instr.append(["Step 1: Enter your latest YTD paystub details in the 'Wage Snapshots' tab."])
    ws_instr.append(["Step 2: Enter your latest YTD brokerage totals in 'Investment Income Snapshots'."])
    ws_instr.append(["Step 3: Enter your 'Prior Year Tax' liability in the Dashboard to enable Safe Harbor targets."])
    ws_instr.append(["Step 4: Check the 'PAYMENT ACTION CENTER' on the Dashboard for immediate payment requirements."])
    ws_instr.append([""])
    ws_instr.append(["Important Notes"])
    ws_instr.append(["1. Investment Income: All dividends and interest are combined and taxed as Ordinary Income for a safe, conservative projection."])
    ws_instr.append(["2. HSA (California): HSA contributions are automatically added back to CA state income as they are not deductible in California."])
    ws_instr.append(["3. Income Projections: The engine pro-rates your remaining annual income based on the days elapsed since Jan 1. You can adjust these projections by entering values into 'Manual Income Offset' or 'Future Income Weight' in Section 1.5 of the Dashboard."])
    ws_instr.append(["4. YTD Methodology: Always use Year-to-Date (YTD) totals from your statements. Overwrite existing rows as new statements arrive."])

    # --- 2. Dashboard Tab (Second Position) ---
    ws_ds = wb.create_sheet("Dashboard")
    ws_ds["A1"] = "Configuration"
    ws_ds["A2"] = f"Filing Status ({status})"; ws_ds["B2"] = status
    ws_ds["A3"] = "Dependents (Under 17)"; ws_ds["B3"] = dependents
    ws_ds["A4"] = "Prior Year Fed Tax (Safe Harbor)"; ws_ds["B4"] = 0
    ws_ds["A5"] = "Prior Year CA Tax (Safe Harbor)"; ws_ds["B5"] = 0
    ws_ds["A6"] = "Estimated Fed Itemized Deduction"; ws_ds["B6"] = 0
    ws_ds["A7"] = "Estimated CA Itemized Deduction"; ws_ds["B7"] = 0
    ws_ds["A8"] = "Tax Year"; ws_ds["B8"] = year
    ws_ds["A9"] = "Status Date"; ws_ds["B9"] = "=TODAY()"

    ws_ds["A11"] = "Projection & Assumptions"
    ws_ds["A12"] = "Manual Income Offset ($)"; ws_ds["B12"] = 0
    ws_ds["A13"] = "Future Income Weight (%)"; ws_ds["B13"] = 1.0
    ws_ds["A14"] = "Remaining Year Wage Income"; ws_ds["B14"] = "=IF(MAX('Wage Snapshots'!A:A)=0, 0, (SUM('Wage Snapshots'!B:B) / MAX(1, MAX('Wage Snapshots'!A:A) - DATE(B8,1,1))) * (DATE(B8,12,31) - MAX('Wage Snapshots'!A:A)))"
    ws_ds["A15"] = "Remaining Year Deductions"; ws_ds["B15"] = "=IF(MAX('Wage Snapshots'!A:A)=0, 0, (SUM('Wage Snapshots'!C:D) / MAX(1, MAX('Wage Snapshots'!A:A) - DATE(B8,1,1))) * (DATE(B8,12,31) - MAX('Wage Snapshots'!A:A)))"

    ws_ds["A17"] = "Consolidated Income Projection"
    ws_ds["A18"] = "Total Projected Wage Income"; ws_ds["B18"] = "=SUM('Wage Snapshots'!B:B) + (B14 * B13) + B12"
    ws_ds["A19"] = "Federal W-2 State Wages"; ws_ds["B19"] = "=B18 - SUM('Wage Snapshots'!C:D) - (B15 * B13)"
    ws_ds["A20"] = "CA W-2 State Wages"; ws_ds["B20"] = "=B18 - SUM('Wage Snapshots'!C:C) - (SUM('Wage Snapshots'!C:C)/MAX(1, SUM('Wage Snapshots'!C:D))) * B15 * B13"
    ws_ds["A21"] = "Investment Ordinary (Div/Int + STG)"; ws_ds["B21"] = "=SUM('Investment Income Snapshots'!C:C) + SUM('Investment Income Snapshots'!D:D)"
    ws_ds["A22"] = "Investment Preferential (LTG Only)"; ws_ds["B22"] = "=SUM('Investment Income Snapshots'!E:E)"
    ws_ds["A23"] = "Total Projected Federal AGI"; ws_ds["B23"] = "=B19 + B21 + B22"
    ws_ds["A24"] = "Total Projected CA AGI"; ws_ds["B24"] = "=B20 + B21 + B22"

    ws_ds["A26"] = "Federal Tax Calculation"
    ws_ds["A27"] = "Deduction Applied (Max Std/Item)"; ws_ds["B27"] = "=MAX(XLOOKUP(B2, 'Tax Constants'!A3:A30, 'Tax Constants'!E3:E30, 0), B6)"
    ws_ds["A28"] = "Ordinary Taxable Income"; ws_ds["B28"] = "=MAX(0, B23 - B27 - B22)"
    ws_ds["A29"] = "Ordinary Income Tax"; ws_ds["B29"] = "=XLOOKUP(B28, FILTER('Tax Constants'!B3:B30, 'Tax Constants'!A3:A30=B2), FILTER('Tax Constants'!C3:C30, 'Tax Constants'!A3:A30=B2), 0, -1) + (B28 - XLOOKUP(B28, FILTER('Tax Constants'!B3:B30, 'Tax Constants'!A3:A30=B2), FILTER('Tax Constants'!B3:B30, 'Tax Constants'!A3:A30=B2), 0, -1)) * XLOOKUP(B28, FILTER('Tax Constants'!B3:B30, 'Tax Constants'!A3:A30=B2), FILTER('Tax Constants'!D3:D30, 'Tax Constants'!A3:A30=B2), 0, -1)"
    ws_ds["A30"] = "Capital Gains Tax"; ws_ds["B30"] = "=IF(B22>0, B22 * XLOOKUP(B28+B22, FILTER('Tax Constants'!B34:B45, 'Tax Constants'!A34:A45=B2), FILTER('Tax Constants'!C34:C45, 'Tax Constants'!A34:A45=B2), 0, -1), 0)"

    ws_ds["A32"] = "CA Tax Calculation"
    ws_ds["A33"] = "CA Deduction Applied"; ws_ds["B33"] = "=MAX(XLOOKUP(B2, 'Tax Constants'!A49:A84, 'Tax Constants'!E49:E84, 0), B7)"
    ws_ds["A34"] = "CA Taxable Income"; ws_ds["B34"] = "=MAX(0, B24 - B33)"
    ws_ds["A35"] = "CA Regular Tax"; ws_ds["B35"] = "=XLOOKUP(B34, FILTER('Tax Constants'!B49:B84, 'Tax Constants'!A49:A84=B2), FILTER('Tax Constants'!C49:C84, 'Tax Constants'!A49:A84=B2), 0, -1) + (B34 - XLOOKUP(B34, FILTER('Tax Constants'!B49:B84, 'Tax Constants'!A49:A84=B2), FILTER('Tax Constants'!B49:B84, 'Tax Constants'!A49:A84=B2), 0, -1)) * XLOOKUP(B34, FILTER('Tax Constants'!B49:B84, 'Tax Constants'!A49:A84=B2), FILTER('Tax Constants'!D49:D84, 'Tax Constants'!A49:A84=B2), 0, -1)"
    ws_ds["A36"] = "CA MH Surcharge (1%)"; ws_ds["B36"] = "=IF(B34 > 1000000, (B34 - 1000000) * 0.01, 0)"
    ws_ds["A37"] = "Total CA Liability"; ws_ds["B37"] = "=B35 + B36"

    ws_ds["A39"] = "Final Liability & Surtaxes"
    ws_ds["A40"] = "NIIT Threshold"; ws_ds["B40"] = "=XLOOKUP(B2, 'Tax Constants'!A88:A91, 'Tax Constants'!B88:B91, 200000)"
    ws_ds["A41"] = "NIIT (Fed)"; ws_ds["B41"] = "=IF(B23 > B40, 0.038 * MIN(B23-B40, B21+B22), 0)"
    ws_ds["A42"] = "Addl Medicare Threshold"; ws_ds["B42"] = "=XLOOKUP(B2, 'Tax Constants'!A88:A91, 'Tax Constants'!C88:C91, 200000)"
    ws_ds["A43"] = "Addl Medicare (Fed)"; ws_ds["B43"] = "=IF(B18 > B42, 0.009 * (B18-B42), 0)"
    ws_ds["A44"] = "CTC Phaseout Start"; ws_ds["B44"] = "=XLOOKUP(B2, 'Tax Constants'!A88:A91, 'Tax Constants'!D88:D91, 200000)"
    ws_ds["A45"] = "Child Tax Credit"; ws_ds["B45"] = "=IF(B23 > B44, MAX(0, (B3*2000)-((B23-B44)/1000)*50), B3*2000)"
    ws_ds["A46"] = "Total Federal Liability"; ws_ds["B46"] = "=B29 + B30 + B41 + B43 - B45"

    ws_ds["D1"] = "Estimated Tax Payments Ledger"
    ws_ds["D2"] = "Date"; ws_ds["E2"] = "Calculated Estimate (Optional)"; ws_ds["F2"] = "Actual Payment Made (Required)"; ws_ds["G2"] = "Note"
    ws_ds["D3"]=f"04/15/{year}"; ws_ds["G3"]="Fed Q1"; ws_ds["D4"]=f"04/15/{year}"; ws_ds["G4"]="CA Q1"
    ws_ds["D5"]=f"06/15/{year}"; ws_ds["G5"]="Fed Q2"; ws_ds["D6"]=f"06/15/{year}"; ws_ds["G6"]="CA Q2"
    ws_ds["D7"]=f"09/15/{year}"; ws_ds["G7"]="Fed Q3"; ws_ds["D8"]=f"09/15/{year}"; ws_ds["G8"]="CA Q3"
    ws_ds["D9"]=f"01/15/{year+1}"; ws_ds["G9"]="Fed Q4"; ws_ds["D10"]=f"01/15/{year+1}"; ws_ds["G10"]="CA Q4"

    ws_ds["A48"] = "Payment Requirements"
    ws_ds["A49"] = "Fed Target"; ws_ds["B49"] = "=IF(B4=0, B46 * 0.9, MIN(B46 * 0.9, B4 * 1.1))"
    ws_ds["A50"] = "Total Fed Payments YTD"; ws_ds["B50"] = "=SUM('Wage Snapshots'!E:E) + SUMIFS(F3:F10, G3:G10, \"Fed*\")"
    ws_ds["A51"] = "CA Target"; ws_ds["B51"] = "=IF(B5=0, B37 * 0.8, MIN(B37 * 0.8, B5 * 1.1))"
    ws_ds["A52"] = "Total CA Payments YTD"; ws_ds["B52"] = "=SUM('Wage Snapshots'!F:F) + SUMIFS(F3:F10, G3:G10, \"CA*\")"

    for i, q in enumerate(["Q1 (Apr 15)", "Q2 (Jun 15)", "Q3 (Sep 15)", "Q4 (Jan 15)"], 1):
        ws_ds[f"A{54+i}"] = q; ws_ds[f"B{54+i}"] = f"=B49 * {0.25*i}"; ws_ds[f"C{54+i}"] = f"=MAX(0, B{54+i} - B50)"
        rate = [0.3, 0.7, 0.7, 1.0][i-1]
        ws_ds[f"A{60+i}"] = q; ws_ds[f"B{60+i}"] = f"=B51 * {rate}"; ws_ds[f"C{60+i}"] = f"=MAX(0, B{60+i} - B52)"

    # --- Right Side Status Panel ---
    ws_ds["I1"] = "PAYMENT ACTION CENTER"
    ws_ds["I2"] = "FED DUE NOW:"; ws_ds["J2"] = "=IFS(B9<=DATE(B8,4,15), C55, B9<=DATE(B8,6,15), C56, B9<=DATE(B8,9,15), C57, TRUE, C58)"
    ws_ds["I3"] = "BY DEADLINE:"; ws_ds["J3"] = "=IFS(B9<=DATE(B8,4,15), \"04/15/\"&B8, B9<=DATE(B8,6,15), \"06/15/\"&B8, B9<=DATE(B8,9,15), \"09/15/\"&B8, TRUE, \"01/15/\"&(B8+1))"
    ws_ds["I4"] = "Next FED Target:"; ws_ds["J4"] = "=IFS(B9<=DATE(B8,4,15), B55, B9<=DATE(B8,6,15), B56, B9<=DATE(B8,9,15), B57, TRUE, B58)"
    ws_ds["I5"] = "FED Payments YTD:"; ws_ds["J5"] = "=B50"
    ws_ds["I6"] = "FED Status:"; ws_ds["J6"] = "=IF(J5>=J4, \"✅ Met (Surplus: \" & TEXT(J5-J4, \"$#,##0\") & \")\", \"🔴 Shortfall (\" & TEXT(J4-J5, \"$#,##0\") & \")\")"
    
    ws_ds["I8"] = "CA DUE NOW:"; ws_ds["J8"] = "=IFS(B9<=DATE(B8,4,15), C61, B9<=DATE(B8,6,15), C62, B9<=DATE(B8,9,15), C63, TRUE, C64)"
    ws_ds["I9"] = "BY DEADLINE:"; ws_ds["J9"] = "=IFS(B9<=DATE(B8,4,15), \"04/15/\"&B8, B9<=DATE(B8,6,15), \"06/15/\"&B8, B9<=DATE(B8,9,15), \"09/15/\"&B8, TRUE, \"01/15/\"&(B8+1))"
    ws_ds["I10"] = "Next CA Target:"; ws_ds["J10"] = "=IFS(B9<=DATE(B8,4,15), B61, B9<=DATE(B8,6,15), B62, B9<=DATE(B8,9,15), B63, TRUE, B64)"
    ws_ds["I11"] = "CA Payments YTD:"; ws_ds["J11"] = "=B52"
    ws_ds["I12"] = "CA Status:"; ws_ds["J12"] = "=IF(J11>=J10, \"✅ Met (Surplus: \" & TEXT(J11-J10, \"$#,##0\") & \")\", \"🔴 Shortfall (\" & TEXT(J10-J11, \"$#,##0\") & \")\")"
    
    ws_ds["I15"] = "TAX DIAGNOSTICS"
    ws_ds["I16"] = "Fed Target Method:"; ws_ds["J16"] = "=IF(B4=0, \"90% Forecast\", IF(B46*0.9 < B4*1.1, \"90% Forecast\", \"110% Safe Harbor\"))"
    ws_ds["I17"] = "CA Target Method:"; ws_ds["J17"] = "=IF(B5=0, \"80% Forecast\", IF(B37*0.8 < B5*1.1, \"80% Forecast\", \"110% Safe Harbor\"))"
    ws_ds["I18"] = "Effective Fed Rate:"; ws_ds["J18"] = "=B46 / MAX(1, B23)"
    ws_ds["I19"] = "Effective CA Rate:"; ws_ds["J19"] = "=B37 / MAX(1, B23)"
    ws_ds["I20"] = "Marginal Fed Bracket:"; ws_ds["J20"] = "=XLOOKUP(B28, FILTER('Tax Constants'!B3:B30, 'Tax Constants'!A3:A30=B2), FILTER('Tax Constants'!D3:D30, 'Tax Constants'!A3:A30=B2), 0, -1)"
    ws_ds["I21"] = "Marginal CA Bracket:"; ws_ds["J21"] = "=XLOOKUP(B34, FILTER('Tax Constants'!B49:B84, 'Tax Constants'!A49:A84=B2), FILTER('Tax Constants'!D49:D84, 'Tax Constants'!A49:A84=B2), 0, -1)"
    ws_ds["I22"] = "Deduction Applied:"; ws_ds["J22"] = "=IF(B6>XLOOKUP(B2, 'Tax Constants'!A3:A30, 'Tax Constants'!E3:E30, 0), \"ITEMIZED\", \"STANDARD\")"
    
    ws_ds["I25"] = "ACTIVE WARNINGS"
    ws_ds["I26"] = "Stale Snapshots:"; ws_ds["J26"] = "=IF(OR(MAX('Wage Snapshots'!A:A)=0, (B9 - MAX('Wage Snapshots'!A:A)) > 30), \"🔴 !!! 30+ DAYS OLD !!!\", \"OK\")"
    ws_ds["I27"] = "Prior Year Data:"; ws_ds["J27"] = "=IF(B4=0, \"🔴 WARNING: FED MISSING\", \"OK\")"
    ws_ds["I28"] = "HSA Audit (CA):"; ws_ds["J28"] = "=IF(B20=B19, \"🔴 ERR: HSA NOT ADDED TO CA\", \"✅ HSA Corrected (CA)\")"

    # --- 5. Data & Constants Tabs ---
    ws_inv = wb.create_sheet("Investment Income Snapshots")
    ws_inv.append(["Quarter", "Broker", "Dividends & Interest", "Short-Term Gains", "Long-Term Gains"])
    ws_wage = wb.create_sheet("Wage Snapshots")
    ws_wage.append(["Date", "Gross W-2 Income", "Pre-tax Deductions", "HSA Contributions", "Fed Tax Withheld", "CA Tax Withheld", "FICA/Med/SDI"])
    ws_const = wb.create_sheet("Tax Constants")
    ws_const.append(["Table A: Federal Brackets (Ordinary Income)"])
    ws_const.append(["Status", "Bracket Floor", "Base Tax", "Marginal Rate", "Standard Deduction"])
    for row in fed_ord_brackets: ws_const.append(row)
    ws_const.append([]); ws_const.append(["Table B: Federal Capital Gains Brackets"])
    ws_const.append(["Status", "Bracket Floor", "LTCG Rate"])
    for row in fed_cg_brackets: ws_const.append(row)
    ws_const.append([]); ws_const.append(["Table C: California FTB Brackets"])
    ws_const.append(["Status", "Bracket Floor", "Base Tax", "Marginal Rate", "Standard Deduction", "MH Surcharge Floor"])
    for row in ca_brackets: ws_const.append(row)
    ws_const.append([]); ws_const.append(["Table D: Surtaxes & Phaseouts"])
    ws_const.append(["Status", "NIIT Threshold", "Addl Medicare", "CTC Phaseout Start"])
    for row in surtaxes: ws_const.append(row)

    # --- Extensive Annotations (Refined: No Headers in text) ---
    ANN = {
        "D1": "Record your quarterly estimated tax payments made directly to the IRS or FTB here. This is the primary source of truth for tracking payments made outside of payroll withholding.",
        "E2": "An automated guide for this deadline based on your current year-to-date data. You can use this as a guide for your payment, or override it by entering your actual payment in the next column.",
        "F2": "ENTER PAYMENTS HERE. This field is REQUIRED to accurately calculate your 'DUE NOW' totals in the Action Center. This ensures the system recognizes your progress and doesn't double-count required taxes.",
        "G2": "Track payment confirmation numbers, specific quarterly intent, or voucher details here.",
        "B4": "Enter your total tax from last year's Federal Form 1040 (typically Line 24 minus Line 19).",
        "B5": "Enter your total tax from last year's California Form 540.",
        "B12": "Adjust total income for one-off events like bonuses or unpaid leave that aren't recurring in your payroll snapshots.",
        "B13": "1.0 = normal earnings. Use < 1.0 if you expect to stop working. Use > 1.0 if you expect a year-end windfall.",
        "I1": "The high-visibility focal point for immediate tax obligations. Values here update automatically based on today's date and your entered payments.",
        "I15": "A real-time health check on your tax situation. Verify your Effective Rates and Brackets to ensure the model matches your expectations.",
        "I16": "Specifies if the system is currently targeting 110% of last year's tax (Safe Harbor) or 90% of your current forecast (Forecast), prioritizing the lower baseline for your safety.",
        "I18": "Total projected Federal Tax divided by Federal AGI. (Standardized against Federal AGI for comparability).",
        "I19": "Total projected California Tax divided by Federal AGI. (Standardized against Federal AGI for comparability).",
        "I21": "The highest tax rate applied to your last dollar of California income.",
        "I28": "Confirms that HSA contributions are successfully added back to California income (as they are not deductible at the state level)."
    }

    # --- Premium Formatting Engine ---
    st_sec = (Font(bold=True, size=11, color="FFFFFF"), PatternFill(start_color="2F75B5", end_color="2F75B5", fill_type="solid"))
    st_lbl = Font(bold=True); st_in = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    st_ac_bg = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    st_crit = Font(bold=True, color="FF0000"); st_calc = Font(bold=True)
    
    side = Side(border_style="thin", color="000000")
    st_border = Border(top=side, left=side, right=side, bottom=side)
    
    sec_k = ["Configuration", "Notes", "Quick Start", "Assumptions", "Projection", "Calculation", "Requirements", "Ledger", "SCHEDULE", "CENTER", "DIAGNOSTICS", "WARNINGS"]
    in_c = ["B2", "B3", "B4", "B5", "B6", "B7", "B8", "B9", "B12", "B13", "B14", "B15"]

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                coord = cell.coordinate; val = str(cell.value)
                
                # Apply Comments (Refined)
                if ws.title == "Dashboard" and coord in ANN:
                    cell.comment = Comment(ANN[coord], "Ross Wait")

                # Main Headers (Blue background)
                if any(k in val for k in sec_k) or (ws.title != "Dashboard" and cell.row == 1): cell.font, cell.fill = st_sec
                # Labels (Bold)
                elif (cell.column in [1, 4, 9] and cell.row > 1): cell.font = st_lbl
                
                if ws.title == "Dashboard":
                    if coord in in_c or (cell.column in [4,5,6,7] and 3 <= cell.row <= 10): cell.fill = st_in
                    if cell.column in [5, 6, 7] and 2 <= cell.row <= 10: cell.font = st_calc
                    
                    # Action Center Style (Border + Background)
                    if (9 <= cell.column <= 10 and 1 <= cell.row <= 13):
                        cell.fill = st_ac_bg
                        cell.border = st_border
                    
                    if cell.coordinate in ["I2", "J2", "I8", "J8"]: cell.font = st_crit
                    
                    if cell.column == 2:
                        if cell.row in [4,5,6,7,12,14,15,18,19,20,21,22,23,24,27,28,29,30,33,34,35,36,37,40,41,42,43,44,45,46,49,50,51,52]: cell.number_format = FORMAT_CURRENCY
                        elif cell.row == 13: cell.number_format = FORMAT_PERCENT
                        elif cell.row == 9: cell.number_format = FORMAT_DATE
                    if cell.column == 3 and cell.row >= 54: cell.number_format = FORMAT_CURRENCY
                    if cell.column == 10:
                        if cell.row in [2,4,5,8,10,11]: cell.number_format = FORMAT_CURRENCY
                        elif cell.row in [18,19,20,21]: cell.number_format = FORMAT_PERCENT
                    if cell.column == 4 and 3 <= cell.row <= 10: cell.number_format = FORMAT_DATE
                    if cell.column in [5,6] and 3 <= cell.row <= 10: cell.number_format = FORMAT_CURRENCY
                elif ws.title in ["Wage Snapshots", "Investment Income Snapshots"] and cell.row > 1:
                    if ws.title == "Wage Snapshots" and cell.column == 1: cell.number_format = FORMAT_DATE
                    else: cell.number_format = FORMAT_CURRENCY
        # Column Widths
        for col in ws.columns:
            l = max([len(str(c.value)) for c in col if c.value and not str(c.value).startswith('=')] + [10])
            ws.column_dimensions[col[0].column_letter].width = min(l + 2, 28)

    wb.save("Tax_Orchestrator_Template.xlsx")
    print(f"Workbook polished and generated successfully.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(); parser.add_argument("--status", default="Single"); parser.add_argument("--dependents", type=int, default=0); parser.add_argument("--year", type=int, default=datetime.now().year)
    args = parser.parse_args(); create_tax_workbook(args.status, args.dependents, args.year)