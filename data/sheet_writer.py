import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
from utils.logger import logger
from config.settings import Credential_File

def get_gsheet_client():

    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    
    creds = Credentials.from_service_account_file(Credential_File, scopes=scopes)

    return gspread.authorize(creds)

# We have to only change the "MONTH_FOLDER_ID" and links of All the spreadsheet after creating new spreadsheet in the new month and moving to the folder and then we can use the same code for pushing data to the sheet without any change in the code.

# Monhtly_Folder_ID
MONTH_FOLDER_ID = "12we1Z9QLYugag3jPPGVFe5wRdZoF0YW6"

# Spreadsheet Overview & Channel
# Jan
# SPREADSHEET_ID_O_C = "1YYa9vwR2y4-_FD9-ZcMX5aOdmCh3o65Mu07cAceWhic"
# Feb
# SPREADSHEET_ID_O_C = "1jcaG62jrBcL5XJ9YitnJiobCdEIaNr4tQyrZVss8ies"
# April
SPREADSHEET_ID_O_C = "1w0VfmtukcWxgngKK_3PTWV7Kk1kTA1oiQq8vK4OfPUA"
# May
# SPREADSHEET_ID_O_C = "1YAd5S4U67Kegm1VCbhdh-SCadFxDVKlZD9oj48Gj71A"

# Spreadsheet Dashboard & MAU, Funnel
# SPREADSHEET_ID_D_M_F = "1oq3MUvYvvzfFykh-xk_B1UU1TL2rF1zG_FJkQZPCSFU"
# Feb
SPREADSHEET_ID_D_M_F = "1LSP5qW0Pra4ssafnrC1mAIImkNznSx-HF7uZSTErugY"

# Spreadsheet Custom Event, Revenue Mapping, Alert
# SPREADSHEET_ID_C_R_A = "1otFDRib02J6JQvDDkXf4jp1y6mFb2NE_vMgROW2ensc"
# Feb
SPREADSHEET_ID_C_R_A = "1otFDRib02J6JQvDDkXf4jp1y6mFb2NE_vMgROW2ensc"

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
        # -------------------------------------------------
        # PRE-EXPAND SHEET
        # -------------------------------------------------
        current_rows = worksheet.row_count
        
        # Performance optimization: get length directly instead of pulling down all cell data
        all_values = worksheet.get_all_values()
        existing_rows_count = len(all_values)

        required_rows = len(rows_to_add) + existing_rows_count + 100

        if required_rows > current_rows:
            worksheet.resize(rows=required_rows)
            logger.info(f"📈 Worksheet Expanded | {current_rows} -> {required_rows}")

        # -------------------------------------------------
        # APPEND (Directly writes the ordered lists passed from parser)
        # -------------------------------------------------
        worksheet.append_rows(rows_to_add, value_input_option='USER_ENTERED')

    except Exception as e:
        logger.error(f"❌ GSheet Push Failed: {repr(e)}")
        raise