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

from data.fetch_funnel import fetch_funnel
from data.parser_funnel import parse_funnel
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
    ROLE_IDS
)

PROGRESS_FILE = "progress_funnel.json"


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

def run_funnel():

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

    driver.get(f"{BASE_URLS[REGION]}/accounts/{license_codes[0]}/engagement/overview/all")
    time.sleep(5)

    cookies = get_session_cookies(driver)

    # ----------------------
    # SHEET
    # ----------------------

    client = get_gsheet_client()
    spreadsheet = client.open_by_key(SPREADSHEET_ID_D_M)

    header = [
        "License","Funnel ID","Funnel Name","Status",
        "Steps Count","Steps","Completion Time",
        "Created By","Created At","Last Modified At"
    ]

    worksheet = get_or_create_worksheet(
        spreadsheet,
        f"Funnel {REGION}",
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

        request_access(lc, REGION, cookies, BASE_URLS, ROLE_IDS)
        time.sleep(6)

        page = 1
        all_rows = []

        while True:

            res = None

            try:
                res = fetch_funnel(lc, REGION, cookies, BASE_URLS, page)

                if not res or res.status_code in [401, 403]:
                    raise requests.exceptions.RequestException("Auth issue")

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
                    res = fetch_funnel(lc, REGION, cookies, BASE_URLS, page)
                except:
                    res = None

            if not res or res.status_code != 200:
                logger.error(f"❌ Failed → {lc}")
                break

            data = res.json()

            rows = parse_funnel(data, lc)
            all_rows.extend(rows)

            total_pages = data.get("response", {}).get("data", {}).get("numberOfPages", 1)

            if page >= total_pages:
                break

            page += 1
            time.sleep(1)

        if all_rows:
            push_rows(worksheet, all_rows)
            mark_done(progress, REGION, lc)
            save_progress(progress)
            logger.info(f"✅ Done → {lc}")
        else:
            logger.info(f"📭 No funnels → {lc}")

        time.sleep(random.uniform(2,3))

    logger.info("✨ FUNNEL DONE")
    driver.quit()


if __name__ == "__main__":
    run_funnel()