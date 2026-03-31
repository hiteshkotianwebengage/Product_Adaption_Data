import time
import random
import json
import os
import argparse

from auth.login import init_driver
from auth.cookies import get_session_cookies
from access.request_access import request_access

from data.fetch_channel_campaign import fetch_channel
from data.parser_channels import parse_channel_data

from config.settings import PRE_BASE_URLS, BASE_URLS, ROLE_IDS, get_backfill_months
from config.channel_config import CHANNELS, CHANNEL_HEADER

from data.load_lc import load_licence_codes
from data.sheet_writer import (
    get_gsheet_client, get_or_create_worksheet, push_rows)

from utils.logger import logger
from utils.date_filter import parse_iso_date

PROGRESS_FILE = "progress_channels.json"


# ----------------------
# RESUME HELPERS
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
# MAIN SCRAPER
# ----------------------

def run_scrapper(region):

    logger.info(f"🚀 Starting SCRAPPER for {region}")

    # Load data
    lc_region_list = load_licence_codes()
    license_codes = [lc for lc, r in lc_region_list if r == region]
    backfill_months = get_backfill_months()

    if not license_codes:
        logger.error(f"❌ No license codes for {region}")
        return

    # Login
    driver = init_driver("selenium_profile")
    driver.get(f"{PRE_BASE_URLS[region]}/admin")
    input("👉 Login and press ENTER...")

    cookies = get_session_cookies(driver)

    # IF session expires 
    def refresh_session(driver, region):

        logger.info("🔄 Attempting silent session refresh...")

        driver.get(f"{PRE_BASE_URLS[region]}/admin")
        time.sleep(5)

        cookies = get_session_cookies(driver)

        # 🔍 Validate session by checking a quick API call
        test_url = f"{BASE_URLS[region]}/api/v1/accounts"

        try:
            import requests
            res = requests.get(test_url, cookies=cookies)

            if res.status_code == 200:
                logger.info("✅ Session restored silently")
                return cookies

        except:
            pass

        # ❗ FALLBACK → manual login
        logger.warning("⚠️ Silent refresh failed → manual login needed")
        input("👉 Please login manually and press ENTER...")

        driver.get(f"{PRE_BASE_URLS[region]}/admin")
        time.sleep(5)

        cookies = get_session_cookies(driver)

        return cookies

    # Google Sheets
    client = get_gsheet_client()
    SPREADSHEET_ID = "16OVEV-QXpPNKTlTMvnxcIy52ULE9ILUNL14ox_VNOhI"
    spreadsheet = client.open_by_key(SPREADSHEET_ID)

    # Create month sheets
    worksheets = {}
    for m in backfill_months:
        tab_name = f"{m['month_name']} {region}"
        worksheets[m["month_name"]] = get_or_create_worksheet(
            spreadsheet,
            tab_name,
            CHANNEL_HEADER
        )

    progress = load_progress()
    session_refreshed = False

    # ----------------------
    # MAIN LOOP
    # ----------------------

    for channel in CHANNELS:

        channel_name = channel["name"]
        logger.info(f"📡 Processing Channel: {channel_name}")

        for i, lc in enumerate(license_codes):

            if lc in progress.get(region, {}).get(channel_name, []):
                logger.info(f"⏭️ Skipping {lc}")
                continue

            logger.info(f"🔍 [{i+1}/{len(license_codes)}] {lc}")

            all_campaigns = []
            page = 1
            success = False

            while True:

                logger.info(f"➡️ Calling page {page}")

                res = fetch_channel(lc, channel, region, cookies, BASE_URLS, page)

                # -------- ACCESS HANDLING --------
                # -------- ACCESS HANDLING --------
                if res.status_code == 403:
                    logger.info(f"🔒 Requesting access → {lc}")
                    status = request_access(lc, region, cookies, BASE_URLS, ROLE_IDS)

                    if status == 200:
                        logger.info(f"✅ Access granted → retrying {lc}")
                        time.sleep(random.uniform(3, 6))
                        continue   # Retry the API call

                    # If status is NOT 200, it's likely a session issue
                    logger.warning(f"⚠️ Access failed for {lc}. Session might be expired. Refreshing...")
                    
                    # 🔄 REFRESH SESSION
                    new_cookies = refresh_session(driver, region)
                    
                    if new_cookies:
                        cookies = new_cookies # Update the main cookie variable
                        logger.info("✅ Session refreshed. Retrying access request...")
                        
                        # Retry access with NEW cookies
                        status = request_access(lc, region, cookies, BASE_URLS, ROLE_IDS)
                        
                        if status == 200:
                            logger.info(f"✅ Access granted after refresh → retrying {lc}")
                            time.sleep(random.uniform(3, 6))
                            continue # Successfully recovered
                    
                    # If we reach here, both original and refreshed attempts failed
                    logger.error(f"❌ Access failed permanently for {lc} even after refresh attempt.")
                    break

                if res.status_code != 200:
                    logger.error(f"❌ API Failed {lc} ({res.status_code})")
                    logger.error(res.text[:200])
                    break

                data = res.json()

                # -------- API ERROR --------
                if data.get("response", {}).get("status") == "error":

                    logger.info(f"🔒 API access issue → {lc}")

                    status = request_access(lc, region, cookies, BASE_URLS, ROLE_IDS)

                    if status != 200:
                        logger.warning(f"⚠️ Cannot access → {lc}")
                        break

                    time.sleep(random.uniform(4, 7))

                    res = fetch_channel(lc, channel, region, cookies, BASE_URLS, page)
                    data = res.json()

                resp_data = data.get("response", {}).get("data", {})
                contents = resp_data.get("contents", [])
                total_pages = resp_data.get("numberOfPages", 1)

                logger.info(f"📄 Page {page} | {len(contents)} items")

                success = True

                if not contents:
                    break

                # -------- EARLY STOP --------
                oldest_date = None

                for item in contents:
                    dt = parse_iso_date(item.get("createdOn"))
                    if dt:
                        dt = dt.replace(tzinfo=None)
                        if not oldest_date or dt < oldest_date:
                            oldest_date = dt

                if oldest_date and oldest_date < backfill_months[0]["start_dt"]:
                    logger.info(f"🛑 Early stop → {lc}")
                    break

                all_campaigns.extend(contents)

                if page >= total_pages:
                    break

                page += 1

                time.sleep(random.uniform(0.5, 1.2))

            # -------- PROCESS DATA --------
            if success:

                # ✅ ALWAYS mark done (even if empty)
                mark_done(progress, region, channel_name, lc)
                save_progress(progress)

                if all_campaigns:
                    for month in backfill_months:
                        rows = parse_channel_data(
                            all_campaigns,
                            lc,
                            channel_name,
                            month
                        )
                        if rows:
                            push_rows(worksheets[month["month_name"]], rows)

                    logger.info(f"✅ Completed {lc} with data")

                else:
                    logger.info(f"📭 No campaigns → marked done {lc}")

            else:
                logger.warning(f"⚠️ Skipping {lc} (no success)")

            # -------- COOLING --------
            time.sleep(random.uniform(1.5, 3))

            if (i + 1) % 10 == 0:
                logger.info("😴 Cooling down...")
                time.sleep(random.uniform(20, 40))

    logger.info("🎯 SCRAPPER COMPLETED")
    driver.quit()


# ----------------------
# ENTRY POINT
# ----------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--region", required=True)
    args = parser.parse_args()

    run_scrapper(args.region.upper())