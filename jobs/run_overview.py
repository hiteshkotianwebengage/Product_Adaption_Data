import argparse
import os
import time
import random
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
    SPREADSHEET_ID_O_C,
    get_or_create_worksheet,
    push_rows
)

from config.settings import (
    PRE_BASE_URLS,
    BASE_URLS,
    ROLE_IDS,
    get_month_info
)

from config.headers import OVERVIEW_HEADER

# Progress file
PROGRESS_FILE = "progress_overview.json"


# ----------------------
# PROGRESS HELPERS
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


# ----------------------
# MAIN FUNCTION
# ----------------------

def run_overview():

    parser = argparse.ArgumentParser()
    parser.add_argument("--region", required=True)
    args = parser.parse_args()

    REGION = args.region.upper()

    # ----------------------
    # INIT
    # ----------------------

    start_date, end_date, start_label, end_label = get_month_info()

    lc_region_list = load_licence_codes()
    license_codes = [lc for lc, r in lc_region_list if r == REGION]

    if not license_codes:
        logger.error(f"❌ No license codes for {REGION}")
        return

    # ----------------------
    # LOGIN FLOW (CORRECT)
    # ----------------------

    driver = init_driver("selenium_profile")

    # Step 1: Login
    driver.get(f"{PRE_BASE_URLS[REGION]}/admin")
    input("👉 Login and press ENTER...")

    # Step 2: Open publisher page (CRITICAL)
    publisher_list_url = f"{PRE_BASE_URLS[REGION]}/admin/publisher.html?action=list"
    driver.get(publisher_list_url)
    time.sleep(6)

    # Step 3: Get cookies
    cookies = get_session_cookies(driver)

    # ----------------------
    # GOOGLE SHEETS
    # ----------------------

    client = get_gsheet_client()
    spreadsheet = client.open_by_key(SPREADSHEET_ID_O_C)

    tab_name = f"Overview {REGION}"
    worksheet = get_or_create_worksheet(spreadsheet, tab_name, OVERVIEW_HEADER)

    progress = load_progress()

    # ----------------------
    # STATIC PAYLOAD
    # ----------------------

    payload = {
        "campaignStatType": "CHANNEL_STATS_OVERVIEW",
        "splitBy": "CAMPAIGN_CHANNEL",
        "startTime": start_date,
        "endTime": end_date,
        "channels": [
            "OVERALL", "PUSH_NOTIFICATION", "IN_APP_NOTIFICATION",
            "SMS", "ON_SITE_NOTIFICATION", "WEB_PUSH",
            "EMAIL", "WHATSAPP", "WEB_PERSONALIZATION", "RCS"
        ],
        "containerTypes": ["ALL"],
        "isFunnelView": False,
        "tags": []
    }

    # ----------------------
    # MAIN LOOP
    # ----------------------

    for i, lc in enumerate(license_codes):

        if lc in progress.get(REGION, []):
            logger.info(f"⏭️ Skipping {lc}")
            continue

        logger.info(f"🔍 [{i+1}/{len(license_codes)}] {lc}")

        # ----------------------
        # STEP 1: REQUEST ACCESS FIRST (🔥 FIX)
        # ----------------------
        status = request_access(lc, REGION, cookies, BASE_URLS, ROLE_IDS)

        if status == 200:
            logger.info(f"✅ Access granted → {lc} (waiting for sync...)")
            time.sleep(5)   # 🔥 CRITICAL WAIT
        else:
            logger.warning(f"⚠️ Access unclear → {lc} (continuing anyway)")

        # ----------------------
        # STEP 2: FETCH DATA
        # ----------------------
        res = fetch_overview(lc, REGION, cookies, BASE_URLS, payload)

        # ----------------------
        # STEP 3: SESSION EXPIRED
        # ----------------------
        if not res or res.status_code == 401:
            logger.warning(f"🔄 Session expired → refreshing")

            driver.get(f"{PRE_BASE_URLS[REGION]}/admin")
            time.sleep(8)

            publisher_list_url = f"{PRE_BASE_URLS[REGION]}/admin/publisher.html?action=list"
            driver.get(publisher_list_url)
            time.sleep(6)

            cookies = get_session_cookies(driver)

            # 🔁 Retry access again after refresh
            status = request_access(lc, REGION, cookies, BASE_URLS, ROLE_IDS)
            if status == 200:
                time.sleep(10)

            res = fetch_overview(lc, REGION, cookies, BASE_URLS, payload)

        # ----------------------
        # STEP 4: STILL 403 → ACCESS AGAIN
        # ----------------------
        if res and res.status_code == 403:
            logger.info(f"🔒 Retry access → {lc}")

            status = request_access(lc, REGION, cookies, BASE_URLS, ROLE_IDS)

            if status == 200:
                time.sleep(12)
                res = fetch_overview(lc, REGION, cookies, BASE_URLS, payload)
            else:
                logger.error(f"❌ Access failed → {lc}")
                continue

        # ----------------------
        # STEP 5: FINAL PROCESS
        # ----------------------
        if res and res.status_code == 200:
            try:
                rows = parse_overview(res.json(), lc, start_label, end_label)
                
                sheet_push_successful = True # Track GSheet status
                
                if rows:
                    try:
                        push_rows(worksheet, rows)
                        logger.info(f"✅ Data pushed ({len(rows)} rows) → {lc}")
                    except Exception as sheet_err:
                        logger.error(f"❌ GSheet Push Failed for {lc}: {sheet_err}")
                        sheet_push_successful = False
                else:
                    logger.info(f"📭 No data found → {lc}")

                # 🔥 ONLY SAVE PROGRESS IF EVERYTHING WORKED
                if sheet_push_successful:
                    mark_done(progress, REGION, lc)
                    save_progress(progress)
                    logger.info(f"💾 Completed and Saved → {lc}")
                else:
                    logger.warning(f"⚠️ Progress NOT saved for {lc} due to Sheet error.")

            except Exception as e:
                logger.error(f"❌ Parsing error for {lc}: {e}")

        else:
            logger.error(f"❌ Permanent Failure for {lc} | Status: {res.status_code if res else 'None'}")

        # ----------------------
        # COOL DOWN
        # ----------------------
        if (i + 1) % 10 == 0:
            logger.info("😴 Cooling down...")
            time.sleep(20)

    logger.info("✨ DONE!")
    driver.quit()


# ----------------------
# ENTRY POINT
# ----------------------

if __name__ == "__main__":
    run_overview()