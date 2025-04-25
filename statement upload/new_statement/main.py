import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import time
import os
from dotenv import load_dotenv
from mapping import name_mapping, account_mapping

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
    # Ensure Transfer Amount column is correctly handled
    if 'Transfer Amount' in df.columns:
        df['Transfer Amount'] = df['Transfer Amount'].astype(float)
    else:
        raise ValueError("Transfer Amount column is missing in the Excel file.")
    
    if 'Payment Date' in df.columns:
        df['Payment Date'] = df['Payment Date'].str.strip()  # Strip leading and trailing spaces
        df['Payment Date'] = pd.to_datetime(df['Payment Date'], format='%d/%m/%Y').dt.strftime('%d-%m-%Y')
    else:
        raise ValueError("Payment Date column is missing in the Excel file.")
    # Iterate over the rows to map beneficiary names and handle exceptions
    if 'Credit A/c No' in df.columns:
        df['Credit A/c No'] = df['Credit A/c No'].astype(str).str.strip()
        # print(df['Credit A/c No'])
    else:
        raise ValueError("Credit A/c No column is missing in the Excel file.")
    
    # Iterate over the rows to map beneficiary names and handle exceptions
    for i, row in df.iterrows():
        ben_name = row['Beneciary Name']
        credit_no = row['Credit A/c No'].strip()  # Ensure trailing spaces are removed
        
        # Check if the beneficiary name is numeric or contains leading/trailing spaces
        if isinstance(ben_name, (int, float)) or ben_name.strip().isdigit():
            df.at[i, 'Beneciary Name'] = account_mapping.get(credit_no, credit_no)
        else:
            # Map beneficiary name if exists in the dictionary
            df.at[i, 'Beneciary Name'] = name_mapping.get(ben_name.strip(), ben_name.strip())
    return df

# Upload to Google Sheets
def upload_to_google_sheets(df, sheet_name, creds):
    client = gspread.authorize(creds)
    total_operations = 0
    total_entries = 0

    try:
        # Open the Google Sheet
        spreadsheet = client.open(sheet_name)

        for ben_name in df['Beneciary Name'].unique():
            # print(f"Searching for worksheet: {ben_name}")
            ben_df = df[df['Beneciary Name'] == ben_name].copy()
            ben_df['Date'] = ben_df['Payment Date'].apply(lambda x: x.strftime('%d-%m-%Y') if isinstance(x, pd.Timestamp) else str(x))
            ben_df['Ref No'] = ben_df['Reference No.']
            ben_df['Credit'] = 0
            ben_df['Debit'] = ben_df['Transfer Amount']

            # Select or create the sheet for the beneficiary
            try:
                worksheet = spreadsheet.worksheet(ben_name)
                print(f"Worksheet found: {ben_name}")
            except gspread.WorksheetNotFound:
                print(f"Worksheet not found: {ben_name}")
                total_operations += 10  # Account for header, balance, and table headers
                if total_operations >= 25:
                    print(f"Pausing for 60 seconds after {total_operations} write operations...")
                    time.sleep(60)
                    total_operations = 0

                # Add new worksheet and format
                worksheet = spreadsheet.add_worksheet(title=ben_name, rows="1000", cols="20")
                header_format = {
                    "horizontalAlignment": "CENTER",
                    "textFormat": {"bold": True}
                }
                worksheet.format("A1:D1", header_format)
                worksheet.merge_cells('A1:D1')
                worksheet.update_cell(1, 1, ben_name.upper())
                
                # Add balance formula
                worksheet.update_cell(2, 1, "BALANCE:")
                total_operations = 0  # Reset count after adding header and balance
                
                # Add table headers
                headers = ['Date', 'Ref No', 'Credit', 'Debit']
                worksheet.append_row(headers)
                total_operations += 1  # Count for table headers

            # Prepare data to upload
            data = ben_df[['Date', 'Ref No', 'Credit', 'Debit']].values.tolist()

            # Upload data rows
            for row in data:
                retry = True
                while retry:
                    try:
                        total_operations += 1
                        if total_operations >= 25:
                            print(f"Pausing for 60 seconds after {total_operations} write operations...")
                            time.sleep(60)
                            total_operations = 0  # Reset count after pause
                        worksheet.append_row(row)
                        total_entries += 1
                        retry = False  # Successful, so no need to retry
                        
                    except gspread.exceptions.APIError as e:
                        if 'quota' in str(e).lower():
                            print(f"Quota exceeded. Retrying in 60 seconds...")
                            time.sleep(60)
                        else:
                            print(f"API Error: {e}")
                            retry = False  # Exit retry loop on other errors
            print(f"Total entries uploaded: {total_entries}")
            
            # Update balance formula dynamically
            values = worksheet.get_all_values()
            data_length = len(values)
            if data_length > 3:  # Check if there are data rows to sum
                worksheet.update_cell(2, 2, f'=(SUM(C4:C{data_length})-SUM(D4:D{data_length}))')

    except gspread.exceptions.APIError as e:
        print(f"API Error: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")

if __name__ == "__main__":
    excel_file_path = 'C:/Users/arukh/Desktop/moradabadhouse/statement upload/new_statement/excel payment/feb 25.xlsx'
    google_sheet_name = 'Copy of 2024-2025'
    json_key_file = os.getenv('GOOGLE_SHEETS_CREDENTIALS_FILE')
    if not json_key_file:
        raise ValueError("GOOGLE_SHEETS_CREDENTIALS_FILE environment variable is not set")

    try:
        df = load_excel(excel_file_path)
        df = process_data(df)
        creds = authenticate(json_key_file)
        upload_to_google_sheets(df, google_sheet_name, creds)
    except Exception as e:
        print(f"Failed to upload data to Google Sheets: {e}")