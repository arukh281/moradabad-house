import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import time
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Load data from Excel
def load_excel(file_path):
    df = pd.read_excel(file_path)
    return df

# Authenticate and create a service
def authenticate():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/userinfo.email','https://www.googleapis.com/auth/drive.readonly','https://www.googleapis.com/auth/drive.readonly','https://www.googleapis.com/auth/drive.readonly','https://www.googleapis.com/auth/drive.readonly','https://www.googleapis.com/auth/drive.readonly','https://www.googleapis.com/auth/drive.readonly']

    # Get credentials file path from environment variable
    credentials_file = os.getenv('GOOGLE_SHEETS_CREDENTIALS_FILE')
    if not credentials_file:
        raise ValueError("GOOGLE_SHEETS_CREDENTIALS_FILE environment variable is not set")

    creds = Credentials.from_service_account_file(credentials_file, scopes=SCOPES)
    return creds

# Process data and upload to Google Sheets
def upload_to_google_sheets(df, sheet_name, creds):
    client = gspread.authorize(creds)
    total_operations = 0
    total_entries = 0

    try:
        # Open the Google Sheet
        spreadsheet = client.open(sheet_name)
        # Process each particular
        for particular in df['Particulars'].unique():
            # Filter data for the particular and exclude the header row
            try:
                worksheet = spreadsheet.worksheet(particular)
                print(f"worksheet {particular} found")
                total_operations+=1
            except gspread.WorksheetNotFound:
                print(f"{particular} not found making worksheet")
                total_operations += 10  # Account for header, balance, and table headers
                if total_operations >= 35:
                    print(f"Pausing for 60 seconds after {total_operations} write operations...")
                    time.sleep(60)
                    total_operations = 0

                # Add new worksheet and format
                
                worksheet = spreadsheet.add_worksheet(title=particular, rows="1000", cols="20")
                header_format = {
                    "horizontalAlignment": "CENTER",
                    "textFormat": {"bold": True}
                }
                worksheet.format("A1:D1", header_format)
                worksheet.merge_cells('A1:D1')
                worksheet.update_cell(1, 1, particular.upper())
                
                # Add balance formula
                worksheet.update_cell(2, 1, "BALANCE:")
                total_operations = 0  # Reset count after adding header and balance
                
                # Add table headers
                headers = ['Date', 'Ref No', 'Credit', 'Debit']
                worksheet.append_row(headers)
                total_operations += 1  # Count for table headers
                # Update balance formula dynamically
                values = worksheet.get_all_values()
                data_length = len(values)
                if data_length > 3:  # Check if there are data rows to sum
                    worksheet.update_cell(2, 2, f'=(SUM(C4:C{data_length})-SUM(D4:D{data_length}))')

    except gspread.exceptions.APIError as e:
        print(f" API Error:{e.response.text}")
    except Exception as e:
        print(f" Unexpected error: {e}")  

if __name__ == "__main__":
    # Define paths and sheet name
    excel_file_path = 'C:/Users/arukh/Desktop/moradabadhouse/invoice upload/purchase list/twentyfour-five/fwdpurchaselist/PUR JAN 25.xlsx'
    google_sheet_name = '2024-2025'

    try:
        # Load, process, and upload data
        df = load_excel(excel_file_path)
        creds = authenticate()
        upload_to_google_sheets(df, google_sheet_name, creds)
    except Exception as e:
        print(f"Failed to upload data to Google Sheets: {e}")