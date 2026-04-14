import argparse
import os
import time
import random
import json
import requests
from requests.exceptions import ReadTimeout, ConnectionError, ChunkedEncodingError

from utils.logger import logger

from auth.login import init_driver
from auth.cookies import get_session_cookies
from access.request_access import request_access

from data.fetch_rev_map import fetch_revenue_mapping
from data.parser_rev_map import parse_revenue_mapping
from data.load_lc import load_licence_codes
from data.sheet_writer import (
    get_gsheet_client,
    SPREADSHEET_ID_C_R_A,
    get_or_create_worksheet,
    push_rows
)

from config.settings import (
    PRE_BASE_URLS,
    BASE_URLS,
    ROLE_IDS
)

PROGRESS_FILE = "progress_revenue.json"


# ----------------------
# PROGRESS
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
# MAIN
# ----------------------

def run_revenue():

    parser = argparse.ArgumentParser()
    parser.add_argument("--region", required=True)
    args = parser.parse_args()

    REGION = args.region.upper()

    lc_region_list = load_licence_codes()
    license_codes = [lc for lc, r in lc_region_list if r == REGION]

    if not license_codes:
        logger.error(f"❌ No license codes for {REGION}")
        return

    # ----------------------
    # LOGIN
    # ----------------------

    driver = init_driver("selenium_profile")

    driver.get(f"{PRE_BASE_URLS[REGION]}/admin")
    input("👉 Login and press ENTER...")

    publisher_list_url = f"{PRE_BASE_URLS[REGION]}/admin/publisher.html?action=list"
    driver.get(publisher_list_url)
    time.sleep(5)

    cookies = get_session_cookies(driver)

    # ----------------------
    # SHEET
    # ----------------------

    client = get_gsheet_client()
    spreadsheet = client.open_by_key(SPREADSHEET_ID_C_R_A)

    header = [
        "License","Mapping ID","Event Name","Revenue Attribute","Active"
    ]

    worksheet = get_or_create_worksheet(
        spreadsheet,
        f"Revenue Mapping {REGION}",
        header
    )

    progress = load_progress()

    # ----------------------
    # LOOP
    # ----------------------

    for i, lc in enumerate(license_codes):

        if lc in progress.get(REGION, []):
            logger.info(f"⏭️ Skipping {lc}")
            continue

        logger.info(f"🔍 [{i+1}/{len(license_codes)}] {lc}")

        # --- STEP 1: REQUEST ACCESS ---
        try:
            request_access(lc, REGION, cookies, BASE_URLS, ROLE_IDS)
            time.sleep(8) # Standardized wait for backend sync
        except Exception as e:
            logger.error(f"❌ Access Request hard-failed for {lc}: {e}")
            continue

        # --- STEP 2: FETCH + SELF HEAL ---
        res = None
        fetch_success = False

        try:
            res = fetch_revenue_mapping(lc, REGION, cookies, BASE_URLS)

            if not res or res.status_code in [401, 403]:
                raise requests.exceptions.RequestException()
            fetch_success = True

        except (ReadTimeout, ConnectionError, ChunkedEncodingError, requests.exceptions.RequestException):

            logger.warning(f"🔄 Recovery → {lc}")

            driver.get(f"{PRE_BASE_URLS[REGION]}/admin")
            time.sleep(5)

            driver.get(f"{BASE_URLS[REGION]}/accounts/{lc}/engagement/overview/all")
            time.sleep(6)

            cookies = get_session_cookies(driver)

            request_access(lc, REGION, cookies, BASE_URLS, ROLE_IDS)
            time.sleep(6)

            try:
                res = fetch_revenue_mapping(lc, REGION, cookies, BASE_URLS)
                if res and res.status_code == 200:
                    fetch_success = True
            except:
                res = None

        # --- STEP 3: FINAL PROCESS & GSHEET PUSH ---
        if fetch_success and res and res.status_code == 200:
            rows = parse_revenue_mapping(res.json(), lc)
            
            sheet_push_successful = True # Track GSheet status

            if rows:
                try:
                    push_rows(worksheet, rows)
                    logger.info(f"✅ Data pushed ({len(rows)} rows) → {lc}")
                except Exception as e:
                    logger.error(f"❌ GSheet Push Failed for {lc}: {e}")
                    sheet_push_successful = False
            else:
                logger.info(f"📭 No mapping found → {lc}")

            # 🔥 ONLY mark done if Sheet worked OR if no mapping existed
            if sheet_push_successful:
                mark_done(progress, REGION, lc)
                save_progress(progress)
                logger.info(f"💾 Completed and Saved → {lc}")
            else:
                logger.warning(f"⚠️ Progress NOT saved for {lc} due to GSheet error.")

        else:
            logger.error(f"❌ Failed to fetch revenue mapping for {lc}")

        # Cooling
        time.sleep(random.uniform(2, 3))

    logger.info("✨ REVENUE MAPPING DONE")
    driver.quit()


if __name__ == "__main__":
    run_revenue()