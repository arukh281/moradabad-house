import gspread
from google.oauth2.service_account import Credentials
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Define the scopes required for Google Sheets API
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

# Get credentials file path from environment variable
credentials_file = os.getenv('GOOGLE_SHEETS_CREDENTIALS_FILE')
if not credentials_file:
    raise ValueError("GOOGLE_SHEETS_CREDENTIALS_FILE environment variable is not set")

# Load credentials from the service account JSON file
creds = Credentials.from_service_account_file(credentials_file, scopes=SCOPES)

# Authorize gspread client
client = gspread.authorize(creds)

# Open the Google Sheet by name
spreadsheet = client.open('2024-2025')

# Count the number of worksheets
worksheet_count = len(spreadsheet.worksheets())

print(f"The number of worksheets in the spreadsheet is: {worksheet_count}")