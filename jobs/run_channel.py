import time
import random
import json
import os
import argparse
import requests
from requests.exceptions import ReadTimeout, ConnectionError, ChunkedEncodingError

from auth.login import init_driver
from auth.cookies import get_session_cookies
from access.request_access import request_access

from data.fetch_channels import fetch_channel
from data.parser_channels import parse_channel_data

from config.settings import PRE_BASE_URLS, BASE_URLS, ROLE_IDS, get_month_info
from config.channel_config import CHANNELS
from config.headers import CHANNEL_HEADER

from data.load_lc import load_licence_codes
from data.sheet_writer import (
    get_gsheet_client,
    get_or_create_worksheet,
    push_rows,
    SPREADSHEET_ID_O_C
)

from utils.logger import logger
from utils.date_filter import parse_iso_date

PROGRESS_FILE = "progress_channels.json"


# ----------------------
# PROGRESS HELPERS
# ----------------------

def load_progress():
    if os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, "r") as f:
                return json.load(f) or {}
        except:
            return {}
    return {}


def save_progress(progress):
    with open(PROGRESS_FILE, "w") as f:
        json.dump(progress, f, indent=4)


def mark_done(progress, region, channel, lc):
    progress.setdefault(region, {}).setdefault(channel, [])
    if lc not in progress[region][channel]:
        progress[region][channel].append(lc)


# ----------------------
# MAIN FUNCTION
# ----------------------

def run_channel():

    parser = argparse.ArgumentParser()
    parser.add_argument("--region", required=True)
    args = parser.parse_args()

    REGION = args.region.upper()
    logger.info(f"🚀 Channel job started for {REGION}")

    # ----------------------
    # DATE
    # ----------------------

    start_date, end_date, start_label, end_label = get_month_info()

    from datetime import datetime
    start_dt = datetime.strptime(start_date[:10], "%Y-%m-%d")
    end_dt = datetime.strptime(end_date[:10], "%Y-%m-%d")

    month_name = start_label[:3] + "'" + start_label[-2:]

    # ----------------------
    # LOAD LCs
    # ----------------------

    lc_region_list = load_licence_codes()
    license_codes = [lc for lc, r in lc_region_list if r == REGION]

    if not license_codes:
        logger.error(f"❌ No license codes for {REGION}")
        return

    # ----------------------
    # LOGIN FLOW (CORRECT)
    # ----------------------

    driver = init_driver("selenium_profile")

    # Login
    driver.get(f"{PRE_BASE_URLS[REGION]}/admin")
    input("👉 Login and press ENTER...")

    # 🔥 CRITICAL → open publisher page ONCE
    publisher_list_url = f"{PRE_BASE_URLS[REGION]}/admin/publisher.html?action=list"
    driver.get(publisher_list_url)
    time.sleep(5)

    cookies = get_session_cookies(driver)

    # ----------------------
    # GOOGLE SHEETS
    # ----------------------

    client = get_gsheet_client()
    spreadsheet = client.open_by_key(SPREADSHEET_ID_O_C)

    tab_name = f"Channel {REGION}"
    worksheet = get_or_create_worksheet(spreadsheet, tab_name, CHANNEL_HEADER)

    progress = load_progress()

    # ----------------------
    # MAIN LOOP
    # ----------------------

    for channel in CHANNELS:

        channel_name = channel["name"]
        logger.info(f"📡 Channel → {channel_name}")

        for i, lc in enumerate(license_codes):

            if lc in progress.get(REGION, {}).get(channel_name, []):
                logger.info(f"⏭️ Skipping {lc}")
                continue

            logger.info(f"🔍 [{i+1}/{len(license_codes)}] {lc}")

            # ----------------------
            # STEP 1: REQUEST ACCESS
            # ----------------------

            request_access(lc, REGION, cookies, BASE_URLS, ROLE_IDS)
            time.sleep(8)   # 🔥 IMPORTANT

            all_campaigns = []
            fetch_success = False
            page = 1

            while True:

                try:
                    res = fetch_channel(lc, channel, REGION, cookies, BASE_URLS, page)
                        # ----------------------
                    # SESSION ISSUE
                    # ----------------------

                    if not res or res.status_code in [401, 403]:
                        raise requests.exceptions.RequestException("Auth/Session Expired")

                    
                except (ReadTimeout, ConnectionError, ChunkedEncodingError) as e:
                    logger.warning(f"🔌 Network hiccup for {lc} on page {page}. Retrying in 5s...")
                    
                    driver.get(f"{PRE_BASE_URLS[REGION]}/admin")
                    time.sleep(4)

                    publisher_list_url = f"{PRE_BASE_URLS[REGION]}/admin/publisher.html?action=list"
                    driver.get(publisher_list_url)
                    time.sleep(5)

                    cookies = get_session_cookies(driver)

                    # 🔁 Re-request access after refresh
                    request_access(lc, REGION, cookies, BASE_URLS, ROLE_IDS)
                    time.sleep(8)

                    # 4. 🔥 THE RETRY: Try the request again now that we refreshed
                    try:
                        res = fetch_channel(lc, channel, REGION, cookies, BASE_URLS, page)
                    except Exception as retry_err:
                        logger.error(f"❌ Retry also failed: {retry_err}")
                        res = None

                if not res or res.status_code != 200:
                    logger.error(f"❌ Failed → {lc}")
                    fetch_success = False
                    break

                data = res.json()
                resp_data = data.get("response", {}).get("data", {})
                contents = resp_data.get("contents", [])
                total_pages = resp_data.get("numberOfPages", 1)

                if contents:
                    all_campaigns.extend(contents)

                if not contents:
                    fetch_success = True
                    break

                # ----------------------
                # PARSE
                # ----------------------

                rows = parse_channel_data(
                    contents,
                    lc,
                    channel_name,
                    start_dt,
                    end_dt,
                    month_name
                )

                # ----------------------
                # EARLY STOP
                # ----------------------

                oldest_in_page = None
                for item in contents:
                    dt = parse_iso_date(item.get("createdOn"))
                    if dt:
                        dt = dt.replace(tzinfo=None)
                        if not oldest_in_page or dt < oldest_in_page:
                            oldest_in_page = dt

                if oldest_in_page and oldest_in_page < start_dt:
                    fetch_success = True
                    break

                # Case 3: Reached the very last page
                if page >= total_pages:
                    fetch_success = True # 🔥 CRITICAL: You missed this in your draft
                    break

                page += 1
                time.sleep(random.uniform(0.7, 1.1))

            # ----------------------
            # MARK DONE
            # ----------------------

            if fetch_success:
                rows = parse_channel_data(all_campaigns, lc, channel_name, start_dt, end_dt, month_name)
                
                sheet_push_done = True # Assume true if no data
                if rows:
                    try:
                        push_rows(worksheet, rows)
                        logger.info(f"✅ Data pushed ({len(rows)} rows) → {lc}")
                    except Exception as e:
                        logger.error(f"❌ GSheet Push Failed for {lc}: {e}")
                        sheet_push_done = False # 🚩 Set to false on error

                # 🔥 ONLY save progress if EVERYTHING worked
                if sheet_push_done:
                    mark_done(progress, REGION, channel_name, lc)
                    save_progress(progress)
                    logger.info(f"💾 Completed and Saved → {lc}")
                else:
                    logger.warning(f"⚠️ Failed at Sheet step. {lc} will be retried next time.")
            else:
                logger.error(f"⚠️ Failed at Fetch step. {lc} will be retried next time.")

            # ----------------------
            # COOLING
            # ----------------------

            if (i + 1) % 10 == 0:
                logger.info("😴 Cooling down...")
                time.sleep(15)

    logger.info("🎯 CHANNEL SCRAPPER COMPLETED")
    driver.quit()


# ----------------------
# ENTRY POINT
# ----------------------

if __name__ == "__main__":
    run_channel()