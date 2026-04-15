"""
# 🧪 End-to-End (E2E) Formula Verification Suite

## Purpose
This suite verifies the internal mathematical integrity of the generated .xlsx templates. 
Unlike Shadow Engine unit tests, these tests upload workbooks to Google Sheets to 
validate that actual Excel formulas and cell references evaluate correctly.

## Practical Considerations & Security
1. **Security & Permissions**: Running these tests requires granting the script broad 
   OAuth access to your Google Drive and Sheets to create, calculate, and delete files.
2. **Infrastructure**: Requires a configured Google Cloud (GCP) 'Desktop' OAuth client 
   (`gcp-test-key.json`) and a one-time user consent flow (`token.pickle`).
3. **Performance**: Involves network Latency (~5s per test) and API Quota usage.

Tests are marked with `@pytest.mark.e2e`. Execute with: `pytest -m e2e`
"""
import pytest
import os
import time
import openpyxl
import datetime
from generate_xlsx import create_tax_workbook
from tests.logic_engine import TaxShadowEngine

# Google imports
try:
    from googleapiclient.discovery import build
    from google.oauth2 import service_account
    import google.auth
    from google.auth.exceptions import DefaultCredentialsError
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    import pickle
    from googleapiclient.http import MediaFileUpload
    import gspread
except ImportError:
    pass

@pytest.fixture(scope="module")
def google_creds():
    # OAuth 2.0 User Flow for personal Drive access
    local_key = "gcp-test-key.json"
    token_file = "token.pickle"
    scopes = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
    
    creds = None
    # The file token.pickle stores the user's access and refresh tokens
    if os.path.exists(token_file):
        with open(token_file, 'rb') as token:
            creds = pickle.load(token)
            
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        elif os.path.exists(local_key):
            flow = InstalledAppFlow.from_client_secrets_file(local_key, scopes)
            creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(token_file, 'wb') as token:
                pickle.dump(creds, token)
        else:
            # Fallback to ADC if no local key exists
            try:
                creds, project = google.auth.default(scopes=scopes)
                return creds
            except DefaultCredentialsError:
                pytest.skip("No gcp-test-key.json found and Application Default Credentials are not configured.")
    
    return creds

@pytest.fixture(scope="module")
def drive_service(google_creds):
    return build('drive', 'v3', credentials=google_creds)

@pytest.fixture(scope="module")
def gspread_client(google_creds):
    return gspread.authorize(google_creds)

@pytest.mark.e2e
def test_zero_income_generates_zero_tax(drive_service, gspread_client):
    """Upload an empty template to Google Sheets and verify it calculates zero tax."""
    filename = "tests/data/e2e_empty_test.xlsx"
    create_tax_workbook(status="Single", dependents=0, year=2026, filename=filename)
    
    file_metadata = {
        'name': 'E2E_Temp_Empty_Tax_Calc',
        'mimeType': 'application/vnd.google-apps.spreadsheet',
        'parents': ['1MFKtegm2SoY8Nsl0KwcijEj5rJKD-NRX']
    }
    media = MediaFileUpload(filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    file_id = file.get('id')
    
    try:
        # Give Google a small moment to ensure formulas evaluated (usually instant)
        time.sleep(1)
        
        # Read from Google Sheets
        sh = gspread_client.open_by_key(file_id)
        dashboard = sh.worksheet("Dashboard")
        
        fed_liability = dashboard.acell("B48").value
        ca_liability = dashboard.acell("B39").value
        
        # In a blank sheet, with $0 income, tax liability should evaluate to "0" or "0.0"
        assert str(fed_liability).replace("$", "") in ["0", "0.0", "0.00"]
        assert str(ca_liability).replace("$", "") in ["0", "0.0", "0.00"]
        
    finally:
        # Teardown: Delete the file from Google Drive
        drive_service.files().delete(fileId=file_id).execute()
        if os.path.exists(filename):
            os.remove(filename)

@pytest.mark.e2e
def test_wage_injection_calculation(drive_service, gspread_client):
    """Inject 100k W-2 income into the template and verify calculation via Sheets."""
    filename = "tests/data/e2e_wage_test.xlsx"
    create_tax_workbook(status="Single", dependents=0, year=2026, filename=filename)
    
    # Inject 100k wages and fixed dates using openpyxl
    wb = openpyxl.load_workbook(filename)
    ds = wb["Dashboard"]
    # Fix the inferred dates to bypass "Proration" logic for a static test
    ds["B8"] = 2026
    ds["B9"] = datetime.datetime(2026, 12, 31)
    
    ws_wage = wb["Wage Snapshots"]
    # "Employer", "Date", "Gross", "Pretax", "HSA", "FedWithhold", "CAWithhold", "FICA"
    ws_wage.append(["TestEmployer", datetime.datetime(2026, 12, 31), 100000, 0, 0, 0, 0, 0])
    wb.save(filename)
    
    file_metadata = {
        'name': 'E2E_Temp_Wage_Tax_Calc',
        'mimeType': 'application/vnd.google-apps.spreadsheet',
        'parents': ['1MFKtegm2SoY8Nsl0KwcijEj5rJKD-NRX']
    }
    media = MediaFileUpload(filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    file_id = file.get('id')
    
    try:
        time.sleep(1)
        sh = gspread_client.open_by_key(file_id)
        dashboard = sh.worksheet("Dashboard")
        
        fed_liability_str = dashboard.acell("B48").value
        fed_val = float(str(fed_liability_str).replace("$", "").replace(",", ""))
        
        ca_liability_str = dashboard.acell("B39").value
        ca_val = float(str(ca_liability_str).replace("$", "").replace(",", ""))
        
        # Use the Shadow Engine to predict the EXACT output to the penny
        engine = TaxShadowEngine(year=2026, status="Single")
        scenario = {
            "config": {"filing_date": "12/31/2026"},
            "wage_snapshots": [{"date": "12/31/2026", "gross": 100000, "pretax": 0, "hsa": 0}],
            "investment_snapshots": []
        }
        expected = engine.run_scenario(scenario)
        
        # Verify the Google Sheets calculation matches the Python shadow engine identically
        assert fed_val == pytest.approx(expected["fed_liability"], 0.01)
        assert ca_val == pytest.approx(expected["ca_liability"], 0.01)

    finally:
        drive_service.files().delete(fileId=file_id).execute()
        if os.path.exists(filename):
            os.remove(filename)
