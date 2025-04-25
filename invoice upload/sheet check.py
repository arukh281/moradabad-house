import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
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
    try:
        # Open the Google Sheet
        spreadsheet = client.open(sheet_name)

        # Process each particular
        for particular in df['Particulars'].unique():
            try:
                worksheet = spreadsheet.worksheet(particular)
                # print(f"{particular} found")
            except gspread.WorksheetNotFound:
                print(f"❌ {particular} NOT FOUND")

    except gspread.exceptions.APIError as e:
        print(f"API Error: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")

if __name__ == "__main__":
    # Define paths and sheet name
    excel_file_path = "C:/Users/arukh/Desktop/moradabadhouse/invoice upload/purchase list/twentyfour-five/fwdpurchaselist/PUR FEB 2025.xlsx"
    google_sheet_name = '2024-2025'

    try:
        # Load, process, and upload data
        df = load_excel(excel_file_path)
        creds = authenticate()
        upload_to_google_sheets(df, google_sheet_name, creds)
    except Exception as e:
        print(f"Failed to upload data to Google Sheets: {e}")
