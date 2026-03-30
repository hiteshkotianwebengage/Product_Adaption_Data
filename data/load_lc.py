import gspread
from google.oauth2.service_account import Credentials

def load_licence_codes():

    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]

    creds = Credentials.from_service_account_file(
        "/Users/admin/Desktop/Code Directory/Product_Adaption_Data/Credential File/mycred-googlesheet.json",
        scopes=scopes
    )

    client = gspread.authorize(creds)

    sheet = client.open_by_key(
        "1o5QRUGQYptkwe1NdsZcfgD44fQCjjkfY2D_DSftSYa4"
    ).worksheet("Philip")

    data = sheet.get_all_records()

    result = []

    regions = ["India", "Global", "KSA"]

    for row in data:
        for region in regions:
            val = row.get(region)
            # Check if value exists and isn't just spaces
            if val and str(val).strip():
                result.append((str(val).strip(), region.upper()))

    return result