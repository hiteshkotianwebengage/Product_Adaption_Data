import argparse
import os
import time
import random
import pandas as pd
import json
from utils.logger import logger

# Project Imports
from auth.login import init_driver
from auth.cookies import get_session_cookies
from access.request_access import request_access
from data.fetch_overview import fetch_overview
from data.parser_overview import parse_overview
from data.load_lc import load_licence_codes
from data.sheet_writer import (
    get_gsheet_client, 
    get_or_create_worksheet, 
    push_rows
)
from config.settings import (
    PRE_BASE_URLS, 
    BASE_URLS, 
    ROLE_IDS, 
    get_backfill_months,
    get_month_info
)

# Progress file
PROGRESS_FILE = "progress_overview.json"

# Constants
SPREADSHEET_ID = "1NlmL3UiBT8b1mirlyf09iLoRB2cmtR0eNhnraXt9Bu8"

# ----------------------
# RESUME HELPERS
# ----------------------

def load_progress():
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, "r") as f:
            return json.load(f)
    return {}

def save_progress(progress):
    with open(PROGRESS_FILE, "w") as f:
        json.dump(progress, f, indent=4)

def mark_done(progress, region, lc):
    progress.setdefault(region, [])
    if lc not in progress[region]:
        progress[region].append(lc)

def main():
    # 1. Setup Arguments
    parser = argparse.ArgumentParser()
    parser.add_argument("--region", required=True, help="Region to run: INDIA/GLOBAL/KSA")
    args = parser.parse_args()
    REGION = args.region.upper()

    # 2. Initialization
    backfill_months = get_backfill_months()
    lc_region_list = load_licence_codes()
    license_codes = [lc for lc, r in lc_region_list if r == REGION]
    
    if not license_codes:
        logger.error(f"❌ No license codes found for region: {REGION}")
        return

    # 3. Browser & GSheet Setup
    driver = init_driver("selenium_profile")
    driver.get(f"{PRE_BASE_URLS[REGION]}/admin")
    input("👉 Please login and press ENTER...")

    cookies = get_session_cookies(driver)
    client = get_gsheet_client()
    spreadsheet = client.open_by_key(SPREADSHEET_ID)

    header = ["License", "Channel", "Users", "Campaigns", "Deliveries", "CTR", "CVR", "Revenue", "Start Date", "End Date"]

    # Pre-initialize Worksheets for all months to avoid API calls inside the inner loop
    worksheets = {}
    for m in backfill_months:
        tab_name = f"Overview {REGION} {m['month_name']}"
        worksheets[m['month_name']] = get_or_create_worksheet(spreadsheet, tab_name, header)

    start_date, end_date, start_label, end_label = get_month_info()

    progress = load_progress()

    # 4. DATA COLLECTION LOOP: One License at a time
    for i, lc in enumerate(license_codes):
        if lc in progress.get(REGION, []):
            logger.info(f"⏭️ Skipping {lc}")
            continue

        logger.info(f"🔍 [{i+1}/{len(license_codes)}] {lc}")
        
        # We store monthly batches locally so we can push to GSheets at the end of each license
        license_data_per_month = {m['month_name']: [] for m in backfill_months}

        # Handle Access ONCE per License
        # We do a quick "Check" fetch using the first month's dates
        check_payload = {
            "campaignStatType": "CHANNEL_STATS_OVERVIEW",
            "startTime": backfill_months[0]["start_api"],
            "endTime": backfill_months[0]["end_api"],
            "channels": ["OVERALL"],
            "containerTypes": ["ALL"]
        }
        
        res = fetch_overview(lc, REGION, cookies, BASE_URLS, check_payload)
        
        if res.status_code in [401, 403]:
            logger.info(f"🔒 Access/Session issue → {lc}")
            status = request_access(lc, REGION, cookies, BASE_URLS, ROLE_IDS)

            # If first attempt fails, try a session refresh
            if status != 200:
                logger.warning("⚠️ Access request failed → Refreshing Session...")
                driver.get(f"{PRE_BASE_URLS[REGION]}/admin")
                time.sleep(3) # Give SSO time to settle
                cookies = get_session_cookies(driver)

                # Try requesting access one last time with fresh cookies
                status = request_access(lc, REGION, cookies, BASE_URLS, ROLE_IDS)

            # Check if we finally got access (either 1st or 2nd try)
            if status == 200:
                logger.info(f"✅ Access granted for {lc}. Waiting for sync...")
                # CRITICAL: WebEngage needs a moment to propagate permissions
                time.sleep(10)
                
            else:
                logger.error(f"❌ Permanent Access failure for {lc}. Skipping to next License.")
                continue

        # Now pull data for EVERY month for this license
        for month_data in backfill_months:
            m_name = month_data["month_name"]
            
            payload = {
                "campaignStatType": "CHANNEL_STATS_OVERVIEW",
                "splitBy": "CAMPAIGN_CHANNEL",
                "startTime": month_data["start_api"],
                "endTime": month_data["end_api"],
                "channels": ["OVERALL","PUSH_NOTIFICATION","IN_APP_NOTIFICATION","SMS","ON_SITE_NOTIFICATION","WEB_PUSH","EMAIL","WHATSAPP","WEB_PERSONALIZATION"],
                "containerTypes": ["ALL"],
                "isFunnelView": False, "tags": []
            }

            time.sleep(random.uniform(1.2, 2.2)) # Smaller jitter between months of the same LC
            # Inside the month loop
            res = fetch_overview(lc, REGION, cookies, BASE_URLS, payload)

            # 🔄 One-time retry for non-200 responses (like 500, 502, or timeouts)
            if res.status_code != 200:
                logger.warning(f"⚠️ Retry {lc} for {m_name} (Status: {res.status_code})")
                time.sleep(3)
                res = fetch_overview(lc, REGION, cookies, BASE_URLS, payload)

            if res.status_code == 200:
                parsed_rows = parse_overview(res.json(), lc, month_data["start_label"], month_data["end_label"])
                license_data_per_month[m_name].extend(parsed_rows)
            else:
                logger.error(f"❌ Failed permanently {lc} for {m_name} after retry")

        # Add this in your main loop instead:
        if not any(license_data_per_month.values()):
            logger.info(f"📭 No data for {lc}")

        # 5. Push this specific license's data to all relevant tabs
        for m_name, rows in license_data_per_month.items():
            if rows:
                push_rows(worksheets[m_name], rows)
        
        # ✅ PLACE PROGRESS LOGIC HERE
        mark_done(progress, REGION, lc)
        save_progress(progress)
        
        logger.info(f"✅ Finished all months for {lc}")
        time.sleep(random.uniform(2,4))

        # Big cooldown every 15 licenses to prevent session blocking
        if (i + 1) % 15 == 0:
            logger.info("😴 Cooling down...")
            time.sleep(random.uniform(20, 40))

    logger.info("✨ DONE! All months for all licenses processed.")
    driver.quit()

if __name__ == "__main__":
    main()