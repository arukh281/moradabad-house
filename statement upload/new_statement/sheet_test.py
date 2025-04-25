import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Beneficiary name mappings
name_mapping = {
    "orchid green Mumbai": "ORCHIDS INTERNATIONAL MUMBAI",
    "kapoor plastic new": "KAPOOR PLASTICS SRE",
    "a k trading": "A.K. TRADING CO. SRE",
    "Sanjeev Distributors CB": "SANJEEV DISTRIBUTORS JWP",
    "Shree Balaji Enterprises Dehradun": "SHREE BALAJI ENTERPRISES D.DUN",
    "Shri Laxmi Bhawani Agencies": "SHRI LAXMI BHAWANI AGENCIES RKE",
    "TARANG AGENCIES JWP": "TARANG AGENCIES JWP",
    "Shree Jain Radios Saharanpur": "SHREE JAIN RADIOS SRE",
    "Kaushal enterprises rke": "KAUSHAL ENTERPRISES RKE",
    "malik trading company ok": "MALIK TRADING CO. DELHI",
    "Shri Krishna Engineering Works DLH": "SHRI KRISHNA ENGG. WORKS DELHI",
    "Shri Krishna Engineering works hdfc": "SHRI KRISHNA ENGG. WORKS DELHI",
    "Kumar and  sons": "KUMAR & SONS SRE",
    "NOOTAN WARE HOUSE Moradabad": "NOOTAN WARE HOUSE MRD",
    "Ena Enterprises Jwalapur": "ENA ENTERPRISES JWP",
    "Shri Sai Agencies Rke": "SHRI SAI AGENCIES RKE",
    "royal traders hrw": "ROYAL TRADERS HRD",
    "royal enterprises": "ROYAL ENTERPRISES RKE",
    "archive international jwp": "ACHIEVE INTERNATIONAL JWP",
    "shri adinath traders ok": "SHREE ADINATH TRADERS BAJAJ RKE",
    "t p traders dl ok": "TP TRADERS DELHI",
    "jain brother dlh napkin": "JAIN BROTHERS DELHI",
    "Osho distributor RK": "OSHO DISTRIBUTORS RKE",
    "shri Krishna inalai jwp": "SHRI KRISHNA ENTERPRISES HRD",
    "Manchanda Trading Co Rishikesh": "MANCHANDA TRADING CORPORATION RISHIKESH",
    "Atc rsk": "AJAY TRADECORPORATION P LTD. RISHIKESH",
    "KB Sales Associates jwl": "KB SLES AND ASSOCIATES JWP",
    "dehradun marketing networ": "DEHRADUN MARKETING NETWORK D.DUN",
    "Arora Brother Roorkee": "ARORA BROTHERS RKE",
    "H S TRADING CO SRE": "H.S. TRADING COMPANY SRE",
    "Gandhi Plastic Rishikesh": "GANDHI PLASTIC RISHIKESH",
    "garg traders roorkee": "GARG TRADERS RKE",
    "Kumar Distributor Jwalapur": "KUMAR DISTRIBUTORS JWP",
    "Ganga Plastics Center": "GANGA PLASTIC CENTER HRD",
    "dhruve enterprises jwp": "DHRUV TRADING CO. HRD",
    "KB Small Trade Solutions": "K.B. TRADE SOLUTIONS HRD",
    "om agencies rke": "OM AGENCIES RKE",
    "Bharat Traders Jagadhri": "BHARAT TRADERS JAG",
    "p b khanna and co ok": "P.B. KHANNA & COMPANY DELHI",
    "Rishabh ent jwp ok" : "RISHABH ENTERPRISES JWP",
    "ANIL BROTHERS":"ANIL BROTHRES AGRA",
    "aradhana trading ddn":"ARADHNA TRADING CO. D.DUN",
    "Delhi Traders":"DELHI TRADERS DELHI",
    "Chawla Brothers Old Shop Jagadri":"CHAWLA KITCHEN WARE JAG",
    # "Online Marketing Roorkee":,
    "Manya Enterprises Delhi":"MANYA ENTERPRISES DELHI",
    "Chawla Kitchenware Jagadhri":"CHAWLA KITCHEN WARE JAG",
    "Gulshan dlh": "V.K. MEHRA  & SONS DELHI",
    "maharani mixer ishan": "ISHAAN HOME APPLIANCES DELHI",
    # "spacevogs furniture pvt l":,
    "asha enterprise jdt":"ASHA ENTERPRISES JAG",
    "K K Narang Trading Co Roorkee":"K.K. NARANG TRADING CO. RKE",
    "OK Industries Delh": "O.K. INDUSTRIES DELHI",
    "Bhawani metal industries":"BHAWANI METAL INDUSTRIES MBD",
    "mehra dons dlh":"V.K. MEHRA  & SONS DELHI", #doubt
    # "MANCHANDA TRADING CO TAG":"MANCHANDA TRADING CORPORATION RISHIKESH",
    "dhasmesh trades":"DASHMESH TRADERS RKE",
    "Online Marketing Roorkee":"ON LINE MARKETING RKE",
    "Mulakh Raj  Sons Delhi":"R.S.MULKH RAJ & SONS DELHI",
    "natraj agency":"NATRAJ AGENCY SRE",
    "spacevogs furniture pvt l":"SPACE VEGUES FURNITURES P LTD  RKE",
    "Ritu KH hdfcc":"PERSONAL",
    "manoj kumar":"PERSONAL",
    "Radhika Enterprises Delhi":"RADHIKA ENTERPRISES DELHI",
    "sachdeva store DDN":"SACHDEVA STORE D.DUN",
    "shri krisna trading co rk":"SRI KRISHNA TRADING CO. TANSHIPUR",
    "vikas home decor":"VIKAS HOME DECOR RKE",
    "Ajay traders shamli":"AJAY TRADERS SHAMLI",
    "KISHORI LAL BALWANT RAI Jagadhri":"KISHORI LAL BALWANT RAI JAG",
    "SACHDEVA AGENCY DDN":"SACHDEVA AGENCIES D DUN",
    "kapoor Chand pawan Kumar":"KAPUR CHAND PAWAN KUMAR CHD",
    "SumerChand Ambika Prasad":"SUMER CHAND AMBIKA PRASHAD JAG",
    "m k industry":"M.K. INDUSTRIES SONIPAT",
    
    # new
    "BHAWANI SHANKAR AND SONS Delhi":"BHAWANI SHANKAR & SONS DELHI",
    "Hamnami Store":"HAMNANI STORE DELHI",
    "shree plastic sre":"SHREE PLASTIC SRE",
    
    "Shyam Brassware Shamli":"SHYAM BRASS WARE SHAMLI",
    "shri krisna enterpriseJWP":"SHRI KRISHNA ENTERPRISES HRD",
    "shri krishna traders dlh":"SHRI KRISHNA TRADERS DELHI",
    "shree ji enterprise jwp":"SHREE JI ENTERPRISES JWP",
    "Cee Aar Products Jagadhri":"CEE AAR PRODUCTS JAGD.",
    "a k trading co sre":"A.K. TRADING CO. SRE",
    "ARORA PLASTIC":"ARORA PLASTICS DELHI",
    
    "Ajay Machinery and Mill Store SRE":"AJAY MACHINERY & MILL STORE SRE",
    "Vardhman Enterprises":"VARDHMAN ENTERPRISES D.DUN",
    "Chaman Lal Rajinder Kumar":"CHAMAN LA; RAJENDER KUMAR JAG",
    "chaman lal Rajender Kumar":"CHAMAN LA; RAJENDER KUMAR JAG",
    "Rangoli khandelwal" :"PERSONAL",
    "dev bhoomi enterprise ddn":"DEV BHOOMI ENTERPRISE U.K",
    "vijay machinery store":"VIJAY MACHINARY STORE SRE",
    "om bartan bhandar new": "OM AGENCIES RKE",
    

    
    # no invoice till no:
    "rama agency rke ok" : "RAMA AGENCY RKE",
    "exident enterprises": "EXIDENT ENTERPRISES",
    "Ram Narain Ashok Kumar Jagadhri":"RAM NARAIN ASHOK KUMAR JAG",
    
    # feb new
    "paras enterprise ddn":"PARAS ENTERPRISES D.DUN",
    "Manjit Traders Jagadhri":"MANJEET TRADERS  JAG",
    "nootan wear house mb":"NOOTAN WARE HOUSE MRD",
    "bhagwati crocery sre":"BHAGWATI CROCKERY STORE SRE"
}

account_mapping={
    "00000032303114485":"SHIVAM TRADERS SRE",
    "32303114485":"SHIVAM TRADERS SRE",
    "00000038573343921":"PABREJA BARTAN BHANDAR  MANGLORE",
    "38573343921":"PABREJA BARTAN BHANDAR  MANGLORE",
    # "00000039675596335":"rsm enterprises mohali",
    # "00000061339160386":"oms enterprises",
    # "00000035178834069":"m a enterprises ddn",
    # "00000034903220554":"goyel super Store",
    # "00000061327202991":"Appliances World",
    "00000036560517489":"KANISHK PLASTIC & CO. SRE",
    "36560517489":"KANISHK PLASTIC & CO. SRE",
    "30385154595":"DESHRI DISTRIBUTORS 24-25 SRE",
    "00000030385154595":"DESHRI DISTRIBUTORS 24-25 SRE",
    "35275862303": "ARIHANT PLASTICS D.DUN",
    "00000035275862303": "ARIHANT PLASTICS D.DUN",
    "00000065249946451": "SUGAN CHAND SOHAN LAL SHAMLI",
    "65249946451": "SUGAN CHAND SOHAN LAL SHAMLI",
    "41449085865":"LAKSHAY BALIYAN ENTERPRISES HRD"
}

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
    if 'Credit A/c No' in df.columns:
        df['Credit A/c No'] = df['Credit A/c No'].astype(str).str.strip()
        print(df['Credit A/c No'])
    else:
        raise ValueError("Credit A/c No column is missing in the Excel file.")
    
    # Iterate over the rows to map beneficiary names and handle exceptions
    for i, row in df.iterrows():
        ben_name = row['Beneciary Name']
        credit_no = row['Credit A/c No']  # Ensure trailing spaces are removed
        
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

    try:
        # Open the Google Sheet
        spreadsheet = client.open(sheet_name)
        count = 1
        for ben_name in df['Beneciary Name'].unique():
            try:
                worksheet = spreadsheet.worksheet(ben_name)
                print(f"{count} ✅ Worksheet found: {ben_name}")
                count += 1
            except gspread.WorksheetNotFound:
                print(f"{count} ❌ Worksheet not found: {ben_name}")
                count += 1

    except gspread.exceptions.APIError as e:
        print(f"API Error: {e}")
    except gspread.exceptions.GSpreadException as e:
        print(f"GSpread Exception: {e}")
        # Log more details about the response
        if hasattr(e, 'response'):
            print(f"Response content: {e.response.content}")
            print(f"Response headers: {e.response.headers}")
    except Exception as e:
        print(f"Unexpected error: {e}")

if __name__ == "__main__":
    excel_file_path = 'C:/Users/arukh/Desktop/moradabadhouse/statement upload/new_statement/excel payment/feb 25.xlsx'
    google_sheet_name = '2024-2025'
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