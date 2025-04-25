import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import time
import random
from mapping import name_mapping, account_mapping
import traceback
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Load data from Excel
def load_excel(file_path):
    df = pd.read_excel(file_path)
    return df

# Authenticate and create a service
def authenticate(json_key_file):
    SCOPES = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/userinfo.email',
        'https://www.googleapis.com/auth/drive.readonly'
    ]
    creds = Credentials.from_service_account_file(json_key_file, scopes=SCOPES)
    return creds

# Process data to map beneficiary names and account numbers
def process_data(df):
    # Ensure required columns are present
    required_columns = ['Transfer Amount', 'Payment Date', 'Credit A/c No', 'Beneciary Name', 'Reference No.']
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Column '{col}' is missing in the Excel file.")

    # Process Transfer Amount
    df['Transfer Amount'] = df['Transfer Amount'].astype(float)

    # Process Payment Date
    df['Payment Date'] = pd.to_datetime(df['Payment Date'].str.strip(), format='%d/%m/%Y', errors='coerce')
    df['Payment Date'] = df['Payment Date'].dt.strftime('%d-%m-%Y')

    # Process Credit A/c No
    df['Credit A/c No'] = df['Credit A/c No'].astype(str).str.strip()

    # Map Beneciary Names
    for i, row in df.iterrows():
        ben_name = row['Beneciary Name']
        credit_no = row['Credit A/c No'].strip()

        if isinstance(ben_name, (int, float)) or ben_name.strip().isdigit():
            df.at[i, 'Beneciary Name'] = account_mapping.get(credit_no, credit_no)
        else:
            df.at[i, 'Beneciary Name'] = name_mapping.get(ben_name.strip(), ben_name.strip())
    return df

# Exponential backoff for retry mechanism
def exponential_backoff(retry_count):
    wait_time = min(60, (2 ** retry_count) + random.uniform(0, 1))
    time.sleep(wait_time)

# Upload to Google Sheets
def upload_to_google_sheets(df, sheet_name, creds):
    client = gspread.authorize(creds)
    total_operations = 0
    total_entries = 0

    try:
        spreadsheet = client.open(sheet_name)

        for ben_name in df['Beneciary Name'].unique():
            ben_df = df[df['Beneciary Name'] == ben_name].copy()
            ben_df['Date'] = ben_df['Payment Date']
            ben_df['Ref No'] = ben_df['Reference No.']
            ben_df['Credit'] = 0
            ben_df['Debit'] = ben_df['Transfer Amount']

            try:
                worksheet = spreadsheet.worksheet(ben_name)
            except gspread.WorksheetNotFound:
                try:
                    worksheet = spreadsheet.add_worksheet(title=ben_name, rows="1000", cols="20")
                except Exception as e:
                    print(f"Failed to create worksheet for {ben_name}: {e}")
                    continue
                worksheet.update('A1:D1', [[ben_name.upper()]])
                worksheet.update('A2:D2', [['BALANCE:', '', '', '']])
                worksheet.update('A3:D3', [['Date', 'Ref No', 'Credit', 'Debit']])

            data = ben_df[['Date', 'Ref No', 'Credit', 'Debit']].values.tolist()
            for row in data:
                retry = True
                retry_count = 0
                while retry:
                    try:
                        total_operations += 1
                        if total_operations >= 25:
                            time.sleep(60)
                            total_operations = 0
                        worksheet.append_row(row)
                        total_entries += 1
                        retry = False
                    except gspread.exceptions.APIError as e:
                        if 'quota' in str(e).lower():
                            retry_count += 1
                            exponential_backoff(retry_count)
                        else:
                            print(f"API Error: {e}")
                            retry = False

            values = worksheet.get_all_values()
            data_length = len(values)
            if data_length > 3:
                data_start_row = 4  # Adjust this if headers change
                worksheet.update_cell(2, 2, f'=SUM(C{data_start_row}:C{data_length})-SUM(D{data_start_row}:D{data_length})')

    except gspread.exceptions.APIError as e:
        print(f"API Error: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")

if __name__ == "__main__":
    excel_file_path = 'C:/Users/arukh/Desktop/moradabadhouse/statement upload/new_statement/excel payment/jan 25.xlsx'
    google_sheet_name = 'Copy of 2024-2025'
    json_key_file = os.getenv('GOOGLE_SHEETS_CREDENTIALS_FILE')
    if not json_key_file:
        raise ValueError("GOOGLE_SHEETS_CREDENTIALS_FILE environment variable is not set")

    try:
        # Load and process the Excel file
        df = load_excel(excel_file_path)
        df = process_data(df)

        # Authenticate with Google Sheets API
        creds = authenticate(json_key_file)

        # Upload data to Google Sheets
        upload_to_google_sheets(df, google_sheet_name, creds)
    except Exception as e:
        print(f"Failed to upload data to Google Sheets: {e}")
        traceback.print_exc()