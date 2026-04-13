import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
from utils.logger import logger

def get_gsheet_client():

    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    
    creds = Credentials.from_service_account_file(
        "/Users/admin/Desktop/Code Directory/Product_Adaption_Data/Credential File/mycred-googlesheet.json",
        scopes=scopes
    )

    return gspread.authorize(creds)

# Monhtly_Folder_ID
MONTH_FOLDER_ID = "1iSeNsHi69H3lddZVyfmUNRmiUy-vOWLQ"

# Spreadsheet_Overview & Channel
SPREADSHEET_ID_O_C = "1RtP7vtmkNEHxH5gvoicxzrZNcNLALNa95ir0cfVXHrA"

# Spreadsheet_Dashboard & MAU
SPREADSHEET_ID_D_M = "1jO0Ojw_DuulUo4Dy31iSYh2r7eu05ppPGqEqeYX88sY"

# Removing this create spreadsheet as im facing some error so directly using SPREADSHEET_ID = "..." but now we are working with monthly_folder_id

# def get_month_sheet_name(start_label):
#     dt = datetime.strptime(start_label, "%B %d, %Y")
#     return dt.strftime("%b'%y")
# def get_or_create_spreadsheet(client, folder_id, sheet_name):
#     # 1. Search for the file specifically in the folder to avoid conflicts
#     query = f"name = '{sheet_name}' and '{folder_id}' in parents and trashed = false"
    
#     try:
#         # Check if the file already exists in that folder
#         files = client.list_spreadsheet_files(query=query)
#         if files:
#             return client.open_by_key(files[0]['id'])
#     except Exception as e:
#         logger.warning(f"Search failed, attempting creation: {e}")

#     # If sheet doesnt exist then create new
#     logger.info(f"✨ Creating new spreadsheet: {sheet_name}")
#     sh = client.create(sheet_name)

#     # Move to folder
#     drive = client.auth
#     drive.request(
#         "post",
#         f"https://www.googleapis.com/drive/v3/files/{sh.id}/parents",
#         params={"addParents": folder_id, "removeParents": "root"}
#     )

#     return sh

def get_or_create_worksheet(sh, tab_name, header):

    try:
        ws = sh.worksheet(tab_name)
    except:
        ws = sh.add_worksheet(title=tab_name, rows=1000, cols=len(header) + 1)
        ws.append_row(header)

    return ws

def push_rows(worksheet, rows_to_add):
    if not rows_to_add:
        return

    try:
        # Modern gspread automatically expands the grid to fit 'rows_to_add'
        worksheet.append_rows(rows_to_add, value_input_option='USER_ENTERED')
    except Exception as e:
        logger.error(f"❌ GSheet Push Failed: {e}")