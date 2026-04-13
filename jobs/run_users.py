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

from config.headers import USERS_HEADER

from data.fetch_users import fetch_users
from data.parser_users import parse_users
from data.load_lc import load_licence_codes
from data.sheet_writer import (
    get_gsheet_client,
    SPREADSHEET_ID_D_M,
    get_or_create_worksheet,
    push_rows
)

from config.settings import (
    PRE_BASE_URLS,
    BASE_URLS,
    ROLE_IDS,
    get_month_info
)

PROGRESS_FILE = "progress_users.json"


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

def run_users():

    parser = argparse.ArgumentParser()
    parser.add_argument("--region", required=True)
    args = parser.parse_args()

    REGION = args.region.upper()

    start_date, end_date, _, _ = get_month_info()

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

    # 🔥 CRITICAL
    driver.get(f"{BASE_URLS[REGION]}/accounts/{license_codes[0]}/engagement/overview/all")
    time.sleep(5)

    cookies = get_session_cookies(driver)

    # ----------------------
    # GOOGLE SHEETS
    # ----------------------

    client = get_gsheet_client()
    spreadsheet = client.open_by_key(SPREADSHEET_ID_D_M)

    sheets = {
        "MAU": get_or_create_worksheet(spreadsheet, f"MAU {REGION}", USERS_HEADER),
        "DAU": get_or_create_worksheet(spreadsheet, f"DAU {REGION}", USERS_HEADER),
        "WAU": get_or_create_worksheet(spreadsheet, f"WAU {REGION}", USERS_HEADER),
    }

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
        # STEP 2: FETCH + SELF HEAL
        # ----------------------

        res = None

        try:
            res = fetch_users(lc, REGION, cookies, BASE_URLS, start_date, end_date)

            if not res or res.status_code in [401, 403]:
                raise requests.exceptions.RequestException("Auth issue")

        except (ReadTimeout, ConnectionError, ChunkedEncodingError, requests.exceptions.RequestException) as e:

            logger.warning(f"🔄 Recovery triggered → {lc}: {e}")

            # 🔁 Refresh session
            driver.get(f"{PRE_BASE_URLS[REGION]}/admin")
            time.sleep(5)

            driver.get(f"{BASE_URLS[REGION]}/accounts/{lc}/engagement/overview/all")
            time.sleep(6)

            cookies = get_session_cookies(driver)

            request_access(lc, REGION, cookies, BASE_URLS, ROLE_IDS)
            time.sleep(8)

            try:
                res = fetch_users(lc, REGION, cookies, BASE_URLS, start_date, end_date)
            except:
                res = None

        # ----------------------
        # FINAL PROCESS
        # ----------------------

        if res and res.status_code == 200:

            try:
                parsed = parse_users(res.json(), lc)
            except Exception as e:
                logger.error(f"❌ Parsing failed → {lc}: {e}")
                continue

            for metric in ["MAU", "DAU", "WAU"]:
                if parsed[metric]:
                    push_rows(sheets[metric], parsed[metric])

            mark_done(progress, REGION, lc)
            save_progress(progress)

            logger.info(f"✅ Data pushed → {lc}")

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
    run_users()