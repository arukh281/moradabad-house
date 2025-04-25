import gspread
from google.oauth2.service_account import Credentials
import time
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

def calculate_balance_in_sheets(sheet_name, json_key_file):
    # Authenticate
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets',
              'https://www.googleapis.com/auth/userinfo.email',
              'https://www.googleapis.com/auth/drive.readonly']
    creds = Credentials.from_service_account_file(json_key_file, scopes=SCOPES)
    client = gspread.authorize(creds)

    try:
        # Open the Google Sheet
        spreadsheet = client.open(sheet_name)
        total_requests = 0  # Track the number of API requests
        
        # Process all worksheets
        for worksheet in spreadsheet.worksheets():
            if worksheet.title.lower() == 'index':
                print(f"Skipping sheet: {worksheet.title}")
                continue
                
            print(f"Processing sheet: {worksheet.title}")
            
            # Merge cells A1 to E1 without changing the content
            sheet_id = worksheet.id
            body = {
                "requests": [
                    {
                        "mergeCells": {
                            "range": {
                                "sheetId": sheet_id,
                                "startRowIndex": 0,
                                "endRowIndex": 1,
                                "startColumnIndex": 0,  # Column A (0-based index)
                                "endColumnIndex": 5   # Column E (exclusive, 0-based index)
                            },
                            "mergeType": "MERGE_ALL"
                        }
                    }
                ]
            }
            spreadsheet.batch_update(body)
            total_requests += 1
            print(f"Merged cells A1 to E1 in sheet: {worksheet.title}")
            
            # Get all values
            values = worksheet.get_all_values()
            if len(values) < 4:  # Skip if sheet is too small
                print(f"Sheet {worksheet.title} is too small, skipping...")
                continue
                
            header_row = 3  # Headers are in row 3 (1-based index)
            headers = values[header_row - 1]  # Adjust for 1-based index
            
            # Check if "Balance" column exists
            balance_col = None
            for idx, header in enumerate(headers, start=1):  # Start index from 1
                if header == 'Balance':
                    balance_col = idx
                    break
            
            # If "Balance" column doesn't exist, create it at column E (index 5)
            if balance_col is None:
                balance_col = 5  # Column E (1-based index is 5)
                worksheet.update_cell(header_row, balance_col, 'Balance')  # Add "Balance" header to E3
                total_requests += 1
                print(f"Added 'Balance' column at E3.")
            
            # Apply formulas for the "Balance" column
            last_row = len(values)
            for row_idx in range(header_row + 1, last_row + 1):  # Start from row 4 (1-based index)
                if row_idx == header_row + 1:  # First row (E4)
                    formula = f"=C{row_idx}-D{row_idx}"
                else:  # Subsequent rows (E5, E6, ...)
                    formula = f"=E{row_idx - 1}+C{row_idx}-D{row_idx}"
                worksheet.update_cell(row_idx, balance_col, formula)
                total_requests += 1
                print(f"Added formula to row {row_idx}: {formula}")
                
                # Pause if requests exceed 25
                if total_requests >= 25:
                    print("Pausing for 60 seconds to avoid API quota limits...")
                    time.sleep(60)
                    total_requests = 0  # Reset the request counter

            print(f"Completed processing sheet: {worksheet.title}")

    except Exception as e:
        print(f"Error processing sheet: {str(e)}")

if __name__ == "__main__":
    SHEET_NAME = "2024-2025"
    # Get credentials file path from environment variable
    JSON_KEY_FILE = os.getenv('GOOGLE_SHEETS_CREDENTIALS_FILE')
    if not JSON_KEY_FILE:
        raise ValueError("GOOGLE_SHEETS_CREDENTIALS_FILE environment variable is not set")
    
    calculate_balance_in_sheets(SHEET_NAME, JSON_KEY_FILE)