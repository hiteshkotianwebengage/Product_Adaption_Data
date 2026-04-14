import argparse
import os
import time
import random
import json
from utils.logger import logger
import requests
from requests.exceptions import ReadTimeout, ConnectionError, ChunkedEncodingError

from auth.login import init_driver
from auth.cookies import get_session_cookies
from access.request_access import request_access

from config.headers import DASHBOARD_HEADER

from data.fetch_dashboard import fetch_dashboard
from data.parser_dashboard import parse_dashboard
from data.load_lc import load_licence_codes
from data.sheet_writer import (
    get_gsheet_client,
    SPREADSHEET_ID_D_M_F,
    get_or_create_worksheet,
    push_rows
)

from config.settings import (
    PRE_BASE_URLS,
    BASE_URLS,
    ROLE_IDS
)

PROGRESS_FILE = "progress_dashboard.json"


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

def run_dashboard():

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
    # LOGIN FLOW
    # ----------------------

    driver = init_driver("selenium_profile")

    driver.get(f"{PRE_BASE_URLS[REGION]}/admin")
    input("👉 Login and press ENTER...")

    # 🔥 CRITICAL (same as overview/channel)
    publisher_list_url = f"{PRE_BASE_URLS[REGION]}/admin/publisher.html?action=list"
    driver.get(publisher_list_url)
    time.sleep(5)

    cookies = get_session_cookies(driver)

    # ----------------------
    # GOOGLE SHEETS
    # ----------------------

    client = get_gsheet_client()
    spreadsheet = client.open_by_key(SPREADSHEET_ID_D_M_F)

    worksheet = get_or_create_worksheet(
        spreadsheet,
        f"Dashboard {REGION}",
        DASHBOARD_HEADER
    )

    progress = load_progress()

    # ----------------------
    # MAIN LOOP
    # ----------------------

    for i, lc in enumerate(license_codes):

        if lc in progress.get(REGION, []):
            logger.info(f"⏭️ Skipping {lc}")
            continue

        logger.info(f"🔍 [{i+1}/{len(license_codes)}] {lc}")

        # ----------------------
        # STEP 1: REQUEST ACCESS FIRST
        # ----------------------

        request_access(lc, REGION, cookies, BASE_URLS, ROLE_IDS)
        time.sleep(8)

        # ----------------------
        # STEP 2: FETCH & Self Healing Block
        # ----------------------

        res = None
        try:
            res = fetch_dashboard(lc, REGION, cookies, BASE_URLS)
            
            # If request worked but session is dead
            if not res or res.status_code in [401, 403]:
                raise requests.exceptions.RequestException("Auth/Session Expired")

        except (ReadTimeout, ConnectionError, ChunkedEncodingError, requests.exceptions.RequestException) as e:
            logger.warning(f"🔄 Recovery triggered for {lc}: {e}")
            
            # Refresh Session
            driver.get(f"{PRE_BASE_URLS[REGION]}/admin")
            time.sleep(5)
            driver.get(f"{BASE_URLS[REGION]}/accounts/{lc}/engagement/overview/all")
            time.sleep(6)
            
            cookies = get_session_cookies(driver)
            request_access(lc, REGION, cookies, BASE_URLS, ROLE_IDS)
            time.sleep(8)

            # Final Attempt
            try:
                res = fetch_dashboard(lc, REGION, cookies, BASE_URLS)
            except Exception as final_err:
                logger.error(f"❌ Failed after recovery: {final_err}")
                res = None

        # ----------------------
        # FINAL PROCESS
        # ----------------------

        if res and res.status_code == 200:
            rows = parse_dashboard(res.json(), lc)
            
            sheet_push_done = True # Track if GSheet actually worked

            if rows:
                try:
                    push_rows(worksheet, rows)
                    # ✅ Added your requested detailed logger
                    logger.info(f"✅ Data pushed ({len(rows)} rows) → {lc}")
                except Exception as e:
                    logger.error(f"❌ GSheet Push Failed for {lc}: {e}")
                    sheet_push_done = False # 🚩 Mark failure
            else:
                logger.info(f"📭 No data → {lc}")

            # ✅ MARK DONE ONLY IF GSHEET WORKED (or no rows were found)
            if sheet_push_done:
                mark_done(progress, REGION, lc)
                save_progress(progress)
            else:
                logger.warning(f"⚠️ Progress NOT saved for {lc} due to Sheet error.")

        else:
            logger.error(f"❌ Failed → {lc}")

        # ----------------------
        # COOLING
        # ----------------------

        time.sleep(random.uniform(2, 3))

        if (i + 1) % 15 == 0:
            logger.info("😴 Cooling down...")
            time.sleep(20)

    logger.info("✨ DONE!")
    driver.quit()


if __name__ == "__main__":
    run_dashboard()