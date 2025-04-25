from flask import Flask, request, session
from twilio.twiml.messaging_response import MessagingResponse
from twilio.rest import Client
import gspread
from google.oauth2.service_account import Credentials
from gspread.exceptions import WorksheetNotFound
from fuzzywuzzy import process
from fuzzywuzzy import fuzz
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from datetime import datetime
from fpdf import FPDF
import os
from babel.numbers import format_decimal
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv('FLASK_SECRET_KEY')


SCOPES = ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/userinfo.email','https://www.googleapis.com/auth/drive.readonly']

# Get credentials file path from environment variable
credentials_file = os.getenv('GOOGLE_SHEETS_CREDENTIALS_FILE')
if not credentials_file:
    raise ValueError("GOOGLE_SHEETS_CREDENTIALS_FILE environment variable is not set")

creds = Credentials.from_service_account_file(credentials_file, scopes=SCOPES)
client = gspread.authorize(creds)
spreadsheet = client.open('2024-2025')

# Function to retrieve firm names from column A (excluding header)
def get_firm_names():
    index_sheet = spreadsheet.worksheet('INDEX')
    firms = index_sheet.col_values(1)[1:]  # Get all values in column A starting from row 2
    return firms

def get_balance(firm_name):
    try:
        sheet = spreadsheet.worksheet(firm_name)
        balance = sheet.cell(2, 2).value  # Assuming balance is in cell B2
        return balance
    except WorksheetNotFound as e:
        print(f"Worksheet '{firm_name}' not found: {e}")
        return None

def generate_pdf(firm_name, month_year):
    try:
        # Convert month input to numeric format
        month_text, year_input = month_year.split()
        year = f"20{year_input}"
        month_mapping = {
            'january': '01', 'february': '02', 'march': '03', 'april': '04',
            'may': '05', 'june': '06', 'july': '07', 'august': '08',
            'september': '09', 'october': '10', 'november': '11', 'december': '12'
        }
        month_numeric = month_mapping[month_text.lower()]
        period = f"{month_numeric}-{year}"
    except Exception as e:
        return None, f"Invalid format for month and year: {e}"

    try:
        # Fetch the worksheet by firm_name
        statement_sheet = spreadsheet.worksheet(firm_name)
        all_records = statement_sheet.get_all_values()

        # Filter records for the specified month and year
        filtered_records = []
        for record in all_records[3:]:  # Skip header row
            if len(record) >= 3:
                # Assuming date is in format (YYYY-MM-DD)
                date_str = record[0].strip('()')
                record_date = datetime.strptime(date_str, "%d-%m-%Y")
                if record_date.strftime("%m-%Y") == month_numeric + '-' + year:
                    filtered_records.append(record)

        if not filtered_records:
            return None, f"No data found for {firm_name} in {month_text.capitalize()} {year}"
    except WorksheetNotFound as e:
        return None, f"Worksheet '{firm_name}' not found: {e}"

    # Prepare PDF content
    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=14, style='B')

        pdf.cell(200, 10, txt=f"Statement for {firm_name} of {month_text.capitalize()} {year}", ln=True, align='C')

        # Calculate the starting position to center the table horizontally
        page_width = pdf.w
        table_width = 4 * 40  # Assuming each cell is 40 units wide (adjust as necessary)
        start_x = (page_width - table_width) / 2

        # Add table headers
        headers = ["Date", "Ref No", "Credit", "Debit"]
        pdf.set_font("Arial", size=12, style='B')
        pdf.set_x(start_x)  # Set starting position for headers
        for header in headers:
            pdf.cell(40, 10, header, border=1, align='C', ln=False)
        pdf.ln()

        total_credit = 0.00
        total_debit = 0.00
        # Add table rows from filtered_records
        pdf.set_font("Arial", size=10)
        for record in filtered_records:
            pdf.set_x(start_x)  # Set starting position for each row
            date_str = record[0].strip('()')
            pdf.cell(40, 7, date_str, border=1, ln=False)
            pdf.cell(40, 7, record[1], border=1, ln=False)
            try:
                credit = float(record[2])
                debit = float(record[3])
            except ValueError:
                credit = 0.00
                debit = 0.00
            total_credit += credit
            total_debit += debit
            pdf.cell(40, 7, format_decimal(credit, format='#,##,##0.00', locale='en_IN'), border=1, align='R', ln=False)
            
            pdf.cell(40, 7, format_decimal(debit, format='#,##,##0.00', locale='en_IN'), border=1, align='R')
            pdf.ln()

        # Add totals to the PDF
        pdf.set_x(start_x)
        pdf.cell(80, 7, "Total:", border=1, ln=False)
        pdf.cell(40, 7, format_decimal(total_credit, format='#,##,##0.00', locale='en_IN'), border=1, ln=False,align='R')
        pdf.cell(40, 7, format_decimal(total_debit, format='#,##,##0.00', locale='en_IN'), border=1,align='R')
        pdf.ln()

        balance = total_credit - total_debit
        pdf.set_x(start_x-1)
        if balance > 0:
            pdf.cell(200, 10, txt=f"Net Payable:  {format_decimal(balance, format='#,##,##0.00', locale='en_IN')}", ln=True, align='L')
        else:
            pdf.cell(200, 10, txt=f"Net Receivable:  {format_decimal(-balance, format='#,##,##0.00', locale='en_IN')}", ln=True, align='L')

        # Save the PDF to a file
        pdf_output_path = f"static/{firm_name.replace(' ', '_')}_{period}.pdf"
        if os.path.exists(pdf_output_path):
            os.remove(pdf_output_path)

        pdf.output(pdf_output_path)
        pdf.output(pdf_output_path)

        return pdf_output_path, None
    except Exception as e:
        return None, f"Failed to generate PDF: {e}"

# Twilio API setup
account_sid = os.getenv('TWILIO_ACCOUNT_SID')
if not account_sid:
    raise ValueError("TWILIO_ACCOUNT_SID environment variable is not set")

auth_token = os.getenv('TWILIO_AUTH_TOKEN')
if not auth_token:
    raise ValueError("TWILIO_AUTH_TOKEN environment variable is not set")

twilio_client = Client(account_sid, auth_token)

# Flask route to handle incoming WhatsApp messages
@app.route("/whatsapp", methods=['POST'])
def whatsapp_bot():
    incoming_msg = request.values.get('Body', '').strip().lower()
    resp = MessagingResponse()
    msg = resp.message()

    if incoming_msg == 'firms':
        firms = get_firm_names()
        firm_list = "\n".join(firms)
        msg.body(f"Please choose a firm:\n{firm_list}")
    
    elif incoming_msg.startswith('balance '):
        firm_name_input = incoming_msg[8:].strip().upper()
        firms = get_firm_names()
        matches = [firm for firm in firms if firm_name_input in firm]

        if len(matches) == 1:
            firm_name = matches[0]
            balance = get_balance(firm_name)
            if balance:
                msg.body(f"The balance for {firm_name} is {balance}")
            else:
                msg.body(f"Sorry, no balance information available for '{firm_name}'.")
        elif len(matches) > 1:
            options = "\n".join([f"{i+1}) {matches[i]}" for i in range(len(matches))])
            msg.body(f"These are the probable results with '{firm_name_input}':\n{options}\nChoose the number which you want to get balance for, or type 'none' to write a new firm name.")
            session['options'] = matches
            session['request_type'] = 'balance'
        else:
            msg.body("No firms found with that name. Please try again.")

    elif incoming_msg.startswith('statement '):
        firm_name_input = incoming_msg[10:].strip().upper()
        firms = get_firm_names()
        matches = [firm for firm in firms if firm_name_input in firm]

        if len(matches) == 1:
            firm_name = matches[0]
            msg.body(f"Please provide the month and year (e.g., JAN 24) for the statement of '{firm_name}'.")
            session['firm_name'] = firm_name
        elif len(matches) > 1:
            options = "\n".join([f"{i+1}) {matches[i]}" for i in range(len(matches))])
            msg.body(f"These are the probable results with '{firm_name_input}':\n{options}\nChoose the number which you want to get statement for, or type 'none' to write a new firm name.")
            session['options'] = matches
            session['request_type'] = 'statement'
        else:
            msg.body("No firms found with that name. Please try again.")

    elif incoming_msg.isdigit() and 'options' in session:
        choice = int(incoming_msg) - 1
        options = session['options']
        if 0 <= choice < len(options):
            firm_name = options[choice]
            if session['request_type'] == 'balance':
                balance = get_balance(firm_name)
                if balance:
                    msg.body(f"The balance for {firm_name} is {balance}")
                else:
                    msg.body(f"Sorry, no balance information available for '{firm_name}'.")
            elif session['request_type'] == 'statement':
                msg.body(f"Please provide the month and year (e.g., JAN 24) for the statement of '{firm_name}'.")
                session['firm_name'] = firm_name
            session.pop('options', None)
            session.pop('request_type', None)
        else:
            msg.body("Invalid choice. Please try again.")

    elif 'firm_name' in session and (len(incoming_msg.split()) == 2):
        try:
            month, year = incoming_msg.split()
            month_year = f"{month.upper()} {year}"  # Format like "MAY 2024"
            firm_name = session['firm_name']
            msg.body(f"Generating PDF statement for {firm_name} of {month.capitalize()} 20{year}.")
            pdf_file_path, error = generate_pdf(firm_name, month_year)
            
            if pdf_file_path:
                # Send the PDF via WhatsApp using Twilio
                media_url = f"https://{request.host}/{pdf_file_path}"
                msg.media(media_url)  # Send the PDF as a media message attachment
                session.pop('firm_name', None)
            else:
                msg.body(f"Failed to generate the statement for '{firm_name}'. {error}")

        except ValueError:
            msg.body("Invalid format. Please provide the month and year in the format 'MMM YY' (e.g., MAY 24).")

    else:
        msg.body("Welcome to the MBH's kingdom. Please type 'FIRMS' to get started.")

    return str(resp)

if __name__ == "__main__":
    app.run(debug=True)