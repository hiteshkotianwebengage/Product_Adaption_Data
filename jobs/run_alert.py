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

from data.fetch_alert import fetch_alert
from data.parser_alert import parse_alert
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

PROGRESS_FILE = os.path.join("Progress_File", "progress_alert.json")


# ----------------------
# PROGRESS
# ----------------------

def load_progress():
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, "r") as f:
            return json.load(f)
    return {}

def save_progress(progress):
    os.makedirs(os.path.dirname(PROGRESS_FILE), exist_ok=True)
    with open(PROGRESS_FILE, "w") as f:
        json.dump(progress, f, indent=4)

def mark_done(progress, region, lc):
    progress.setdefault(region, [])
    if lc not in progress[region]:
        progress[region].append(lc)


# ----------------------
# MAIN
# ----------------------

def run_alert():

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
        "License","Alert ID","Alert Name","Metric","Description",
        "Frequency","Threshold","Operator","Change Type",
        "Status","Created By","Subscribers",
        "Created At","Updated At","Last Evaluated"
    ]

    worksheet = get_or_create_worksheet(
        spreadsheet,
        f"Alert {REGION}",
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

        logger.info(f"🔍 [{i+1}/{len(license_codes)}] Processing {lc}")

        status = request_access(
            lc,
            REGION,
            cookies,
            BASE_URLS,
            ROLE_IDS
        )

        page = 1
        all_rows = []
        fetch_success = False
        is_subscribed = True
        
        # --- STEP 1: INITIAL FETCH & ACCESS CHECK ---
        res = None
        try:
            res = fetch_alert(lc, REGION, cookies, BASE_URLS, page)
            
            # 🚩 PRIORITY 1: Check for Unsubscribed (403)
            if res is not None and res.status_code == 403:
                logger.warning(f"🚫 Module not subscribed for {lc}.")
                is_subscribed = False
                fetch_success = True 

            # 🚩 PRIORITY 2: Check for Session Expiry (401)
            elif res is None or res.status_code == 401:
                logger.warning(f"🔑 Session expired for {lc}. Re-authenticating...")
                driver.get(f"{PRE_BASE_URLS[REGION]}/admin")
                time.sleep(8)
                driver.get(f"{PRE_BASE_URLS[REGION]}/admin/publisher.html?action=list")
                time.sleep(6)
                cookies = get_session_cookies(driver)
                
                request_access(lc, REGION, cookies, BASE_URLS, ROLE_IDS)
                time.sleep(10)
                res = fetch_alert(lc, REGION, cookies, BASE_URLS, page)
                
                if res and res.status_code == 403:
                    is_subscribed = False
                    fetch_success = True

        except Exception as e:
            logger.error(f"🌐 Network error for {lc}: {e}")
            res = None

        # --- STEP 2: HANDLE UNSUBSCRIBED ACCOUNTS ---
        if not is_subscribed:
            # 🚩 Case 1: The module is NOT paid for (Status 403)
            unsub_row = [lc] + (["UNSUBSCRIBED"] * (len(header) - 1))
            try:
                push_rows(worksheet, [unsub_row])
                logger.info(f"📝 Marked as UNSUBSCRIBED in Sheet → {lc}")
                mark_done(progress, REGION, lc)
                save_progress(progress)
            except Exception as e:
                logger.error(f"❌ Failed to push unsubscribed row for {lc}: {e}")
            continue

        if not res or res.status_code != 200:
            logger.error(f"❌ Permanent failure for {lc}. Moving to next.")
            continue

        # --- STEP 3: PAGINATION LOOP ---
        while True:
            try:
                data = res.json()
                page_data = data.get("response", {}).get("data", {})
                contents = page_data.get("contents", [])
                
                # Check if this is Page 1 and it is totally empty
                if page == 1 and not contents:
                    # 🚩 Case 2: Subscribed but no alerts created (Status 200 + Empty)
                    no_data_row = [lc] + (["NO DATA"] * (len(header) - 1))
                    all_rows.append(no_data_row)
                    fetch_success = True
                    break

                # 🚩 Case 3: Subscribed and has actual data
                rows = parse_alert(data, lc)
                all_rows.extend(rows)

                total_pages = page_data.get("numberOfPages", 1)
                logger.info(f"📄 {lc} | Page {page}/{total_pages} fetched")

                if page >= total_pages:
                    fetch_success = True
                    break

                page += 1
                time.sleep(1.5)
                res = fetch_alert(lc, REGION, cookies, BASE_URLS, page)

            except Exception as e:
                logger.error(f"❌ Error parsing {lc} at page {page}: {e}")
                break

        # --- STEP 4: FINAL PUSH ---
        if fetch_success:
            if all_rows:
                push_rows(worksheet, all_rows)
                # Differentiate log message for clarity
                if "NO DATA" in all_rows[0]:
                    logger.info(f"📭 Marked as NO DATA → {lc}")
                else:
                    logger.info(f"✅ Pushed {len(all_rows)} data rows → {lc}")
            
            mark_done(progress, REGION, lc)
            save_progress(progress)
        else:
            logger.error(f"⚠️ Process incomplete for {lc}")

        time.sleep(random.uniform(2, 4))

    logger.info("✨ ALERT DONE")
    driver.quit()

if __name__ == "__main__":
    run_alert()