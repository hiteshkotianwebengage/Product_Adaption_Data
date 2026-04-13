from auth.login import init_driver
from auth.cookies import get_session_cookies
from access.request_access import request_access
from data.fetch_overview import fetch_overview
from data.parser_overview import parse_overview
from config.settings import PRE_BASE_URLS, BASE_URLS, ROLE_IDS
import pandas as pd
import time, random
from config.settings import get_month_info
from utils.logger import logger
from data.load_lc import load_licence_codes
import argparse
import os
from data.sheet_writer import (
    get_gsheet_client, get_or_create_worksheet, push_rows)

SPREADSHEET_ID = "1SR1LceLDom9Oy4eCUh65nAJEf7Q2qMUXXRvriq1x_cI"

parser = argparse.ArgumentParser()
parser.add_argument("--region", required=True, help="Region to run: INDIA/GLOBAL/KSA")
args = parser.parse_args()

REGION = args.region.upper()

start_date, end_date, start_label, end_label = get_month_info()

# Load LC
lc_region_list = load_licence_codes()

# 1. LOGIN
driver = init_driver("selenium_profile")
driver.get(f"{PRE_BASE_URLS[REGION]}/admin")

input("👉 Login and press ENTER...")

# 2. GET 
cookies = get_session_cookies(driver)

# 3. LOOP
all_rows = []

payload = {
    "campaignStatType": "CHANNEL_STATS_OVERVIEW",
    "splitBy": "CAMPAIGN_CHANNEL",
    "startTime": start_date,
    "endTime": end_date,
    "channels": [
        "OVERALL","PUSH_NOTIFICATION","IN_APP_NOTIFICATION",
        "SMS","ON_SITE_NOTIFICATION","WEB_PUSH",
        "EMAIL","WHATSAPP","WEB_PERSONALIZATION"
    ],
    "containerTypes": ["ALL"],
    "isFunnelView": False,
    "tags": []
}

# Google sheet initializing

client = get_gsheet_client()

folder_id = "1ofGHkTIOYcYvJe_IMuuLo8SSW3P69WUT"

# sheet_name = get_month_sheet_name(start_label)

# Open the existing sheet
spreadsheet = client.open_by_key(SPREADSHEET_ID)

tab_name = f"Overview {REGION}"

header = ["License", "Channel", "Users", "Campaigns", 
          "Deliveries", "CTR", "CVR", "Revenue", 
          "Start Date", "End Date"]

worksheet = get_or_create_worksheet(spreadsheet, tab_name, header)

license_codes = [lc for lc, r in lc_region_list if r == REGION]

batch_rows = [] # This will help to make sure google dont cry while adding data so we add to our local list

for i, lc in enumerate(license_codes):

    time.sleep(random.uniform(1.5, 3.5))

    res = fetch_overview(lc, REGION, cookies, BASE_URLS, payload)

    if res.status_code == 403:
        logger.info(f"🔒 Requesting access → {lc}")
        status = request_access(lc, REGION, cookies, BASE_URLS, ROLE_IDS)
        time.sleep(random.uniform(5, 10))

        if status == 200:
            res = fetch_overview(lc, REGION, cookies, BASE_URLS, payload)
        else:
            logger.error("❌ Access failed")
            continue

    if res.status_code == 200:
        rows = parse_overview(res.json(), lc, start_label, end_label)
        # Append to Google sheet
        if rows:
            batch_rows.extend(rows)
            logger.info(f"✅ Processed {lc}")
        else:
            logger.warning(f"⚠️ No data rows for {lc}")
    # Now we push to the Google sheet safely
    if len(batch_rows) >= 5:
        push_rows(worksheet, batch_rows)
        batch_rows = [] #Clear the batch
        time.sleep(1)
        
        # Cooldown
    if i % 20 == 0 and i != 0:
        logger.info("😴 Cooling down batch ....")
        time.sleep(random.uniform(15,30))

if batch_rows:
    push_rows(worksheet, batch_rows)
    logger.info("📦 Pushed final remaining rows to Google Sheet.")
logger.info(f"🚀 All licenses processed. Check Google sheet")