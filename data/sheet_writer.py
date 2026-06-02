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
MONTH_FOLDER_ID = "1atWKSTnnQ-z8zMj8oOUqrvePHm8bOmlB"

# Spreadsheet Overview & Channel
SPREADSHEET_ID_O_C = "1YAd5S4U67Kegm1VCbhdh-SCadFxDVKlZD9oj48Gj71A"

# Spreadsheet Dashboard & MAU, Funnel
SPREADSHEET_ID_D_M_F = "1Aas9lFYVet6P9poJ8Km7prijH6wCIxjfa6RuMziH2Ug"

# Spreadsheet Custom Event, Revenue Mapping, Alert
SPREADSHEET_ID_C_R_A = "1RwTri8R2ZzoCHm5yWAx1I5yAP7I7sEatEgivjLltg_8"

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

        required_rows = (
            len(rows_to_add)
            + len(worksheet.get_all_values())
            + 100
        )

        if required_rows > current_rows:

            worksheet.resize(
                rows=required_rows
            )

            logger.info(
                f"📈 Worksheet Expanded "
                f"| {current_rows} -> {required_rows}"
            )

        # -------------------------------------------------
        # APPEND
        # -------------------------------------------------

        worksheet.append_rows(
            rows_to_add,
            value_input_option='USER_ENTERED'
        )

    except Exception as e:

        logger.error(
            f"❌ GSheet Push Failed: {repr(e)}"
        )

        raise