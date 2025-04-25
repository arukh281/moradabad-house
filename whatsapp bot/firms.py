import gspread
from google.oauth2.service_account import Credentials

# Define the scopes required for Google Sheets API
SCOPES = ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/userinfo.email','https://www.googleapis.com/auth/drive.readonly','https://www.googleapis.com/auth/drive.readonly','https://www.googleapis.com/auth/drive.readonly','https://www.googleapis.com/auth/drive.readonly','https://www.googleapis.com/auth/drive.readonly','https://www.googleapis.com/auth/drive.readonly']

# Load credentials from the service account JSON file
creds = Credentials.from_service_account_file('moradabad-house1.json', scopes=SCOPES)

# Authorize gspread client
client = gspread.authorize(creds)

# Open the 'mbh' spreadsheet
spreadsheet = client.open('2024-2025')

# Access the 'INDEX' worksheet
index_sheet = spreadsheet.worksheet('INDEX')

# Function to retrieve firm names from column A (excluding header)
def get_firm_names():
    firms = index_sheet.col_values(1)[1:]  # Get all values in column A starting from row 2
    return firms

# Get the list of firm names
firms = get_firm_names()

# Join firm names into a single string with each name on a new line
firm_list = "\n".join(firms)

# Print or use firm_list as needed
print(firm_list)
