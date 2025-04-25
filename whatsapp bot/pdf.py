from datetime import datetime
from fpdf import FPDF
import gspread
from google.oauth2.service_account import Credentials
from gspread.exceptions import WorksheetNotFound
from babel.numbers import format_decimal
import os


# Initialize Google Sheets credentials and client
SCOPES = ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/userinfo.email','https://www.googleapis.com/auth/drive.readonly']
creds = Credentials.from_service_account_file('moradabad-house1.json', scopes=SCOPES)
client = gspread.authorize(creds)
spreadsheet = client.open('mbh')

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
                record_date = datetime.strptime(date_str, "%Y-%m-%d")
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

# Example usage
firm_name = "SHRI SAI AGENCIES RKE"
month_year = "MAY 24"
result = generate_pdf(firm_name, month_year)
print(result)
