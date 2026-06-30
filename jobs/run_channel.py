import time
import json
import os
import argparse
import random
from datetime import datetime
from requests.exceptions import ReadTimeout, ConnectionError, ChunkedEncodingError, RequestException

from threading import Lock
from concurrent.futures import ThreadPoolExecutor, as_completed

from auth.login import init_driver
from auth.cookies import get_session_cookies
from access.request_access import request_access

from data.fetch_channel import (fetch_campaigns, fetch_campaign_detail, fetch_campaign_aggregates, fetch_journey,)
from data.parser_channel import (parse_campaign_row,)

from config.settings import PRE_BASE_URLS, BASE_URLS, ROLE_IDS, get_month_info
from config.headers import CHANNEL_HEADER
from config.channel_config import CHANNELS

from data.load_lc import load_licence_codes
from data.sheet_writer import (
    get_gsheet_client,
    get_or_create_worksheet,
    push_rows,
    SPREADSHEET_ID_O_C
)

from utils.logger import logger

PROGRESS_FILE = os.path.join("Progress_File", "progress_channels.json")

session_lock = Lock()
journey_cache_lock = Lock()

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
    os.makedirs(os.path.dirname(PROGRESS_FILE), exist_ok=True)
    with open(PROGRESS_FILE, "w") as f:
        json.dump(progress, f, indent=4)


def mark_done(progress, region, lc):
    progress.setdefault(region, [])
    if lc not in progress[region]:
        progress[region].append(lc)

# ----------------------
# Browser Session Recovery
# ----------------------

def refresh_browser_session(driver, region):
    logger.info("🔄 Refreshing administrative cookies and access permissions...")
    driver.get(f"{PRE_BASE_URLS[region]}/admin")
    time.sleep(4)
    
    publisher_list_url = f"{PRE_BASE_URLS[region]}/admin/publisher.html?action=list"
    driver.get(publisher_list_url)
    time.sleep(5)
    
    return get_session_cookies(driver)

# ----------------------
# Helper metrics 
# ----------------------

def has_valid_metrics(response_obj):
    if response_obj.get("message") == "No data":
        return False
    
    data_list = response_obj.get("data", [])
    if not data_list or not isinstance(data_list, list):
        return False
        
    dimensions_list = data_list[0].get("dimensions", [])
    if not dimensions_list or not isinstance(dimensions_list, list):
        return False

    metrics_list = dimensions_list[0].get("metrics", [])
    if not metrics_list or not isinstance(metrics_list, list):
        return False
        
    return True

# ----------------------
# Core Campaign Pipeline Worker
# ----------------------

def process_campaign(campaign, lc, channel, metrics_from, metrics_to, month_name, REGION, session_context, BASE_URLS, journey_cache):
    status = (campaign.get("status", "").upper())
    sent_status = (campaign.get("sentStatus", {}).get("status", "").upper())

    if status == "DRAFT" or sent_status == "UPCOMING": 
        return None

    campaign_id = campaign.get("id")
    if not campaign_id:
        return None

    # GATE 1: THE AGGREGATE CHECK
    aggregate_res = fetch_campaign_aggregates(lc=lc, channel=channel, campaign_id=campaign_id, metrics_from=metrics_from, metrics_to=metrics_to, region=REGION, cookies=session_context["cookies"], base_urls=BASE_URLS)
    if aggregate_res.status_code in [401, 403]:
        raise RequestException("⚠️ Unauthorised or Expired Session detected during data extraction")
    if aggregate_res.status_code != 200:
        return None

    aggregates = aggregate_res.json()
    response = aggregates.get("response", {})
    
    if not has_valid_metrics(response):
        return None  # Early gate exit: Saves Detail and Journey calls!

    # HEAVY ASSET EXTRACTION
    detail_res = fetch_campaign_detail(lc=lc, channel=channel, campaign_id=campaign_id, region=REGION, cookies=session_context["cookies"], base_urls=BASE_URLS)
    if detail_res.status_code in [401, 403]:
        raise RequestException("⚠️ Session Expired during Detail fetch")
    if detail_res.status_code != 200:
        return None

    campaign_detail = detail_res.json()
    journey_id = campaign_detail.get("response", {}).get("data", {}).get("experiment", {}).get("journeyId")

    journey = {}
    if journey_id:
        with journey_cache_lock:
            cached_journey = journey_cache.get(journey_id)
            
        if cached_journey:
            journey = cached_journey
        else:
            try:
                journey_res = fetch_journey(lc=lc, journey_id=journey_id, region=REGION, cookies=session_context["cookies"], base_urls=BASE_URLS)
                if journey_res.status_code in (401, 403):
                    raise RequestException("⚠️ Session Expired during Journey fetch")
                if journey_res.status_code == 200:
                    journey = journey_res.json()
                    with journey_cache_lock:
                        journey_cache[journey_id] = journey
            except RequestException:
                raise
            except Exception as e:
                logger.debug(f"Journey parse failed: {e}")

    return parse_campaign_row(
        campaign=campaign, 
        campaign_detail=campaign_detail, 
        journey=journey, 
        aggregates=aggregates, 
        lc=lc, 
        month_name=month_name
    )

# ----------------------
# Session Refresh
# ----------------------

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

# ----------------------
# MAIN RUN ENGINE
# ----------------------

def run_channel():
    parser = argparse.ArgumentParser()
    parser.add_argument("--region", required=True)
    args = parser.parse_args()

    REGION = args.region.upper()
    logger.info(f"🚀 Channel job started for {REGION}")

    start_date, end_date, start_label, end_label = get_month_info()
    start_dt = datetime.strptime(start_date[:10], "%Y-%m-%d")
    end_dt = datetime.strptime(end_date[:10], "%Y-%m-%d")
    month_name = start_label[:3] + "'" + start_label[-2:]

    metrics_from = f"{start_dt.strftime('%Y-%m-%d')}T00:00:00.000+05:30"
    metrics_to = f"{end_dt.strftime('%Y-%m-%d')}T23:59:59.999+05:30"

    lc_region_list = load_licence_codes()
    license_codes = [lc for lc, r in lc_region_list if r == REGION]

    if not license_codes:
        logger.error(f"❌ No license codes for {REGION}")
        return

    driver = init_driver("selenium_profile")
    driver.get(f"{PRE_BASE_URLS[REGION]}/admin")
    input("👉 Login and press ENTER...")

    publisher_list_url = f"{PRE_BASE_URLS[REGION]}/admin/publisher.html?action=list"
    driver.get(publisher_list_url)
    time.sleep(5)

    session_context = {
        "cookies": get_session_cookies(driver),
        "updated_at": time.time()
    }

    client = get_gsheet_client()
    spreadsheet = client.open_by_key(SPREADSHEET_ID_O_C)

    tab_name = f"Channel {REGION}"
    worksheet = get_or_create_worksheet(spreadsheet, tab_name, CHANNEL_HEADER)

    progress = load_progress()
    journey_cache = {}

    for idx, lc in enumerate(license_codes):
        if lc in progress.get(REGION, []):
            continue

        logger.info(f"💼 Extracting LC: {lc} [{idx + 1}/{len(license_codes)}]")
        
        # Fire access request ONCE and capture its return status
        status = request_access(lc, REGION, session_context["cookies"], BASE_URLS, ROLE_IDS)

        # If unauthorized/expired, execute recovery logic immediately before proceeding
        if status in [401, 403] or status != 200:
            logger.warning(f"⚠️ Access failed (Status {status}) for {lc}. Session might be expired. Refreshing...")
            
            with session_lock:
                new_cookies = refresh_session(driver, REGION)
                if new_cookies:
                    session_context["cookies"] = new_cookies
                    session_context["updated_at"] = time.time()
                    
                    # Retry once with recovered session
                    status = request_access(lc, REGION, session_context["cookies"], BASE_URLS, ROLE_IDS)
            
            if status != 200:
                logger.error(f"❌ Access failed permanently for {lc} even after refresh attempt. Skipping to next account.")
                continue # Safely skip this account since authentication failed

        logger.info(f"✅ Access established for LC: {lc}. Launching worker pipeline...")
        time.sleep(4)

        account_rows = []
        account_pipeline_success = True
        total_targets_checked = 0
        
        with ThreadPoolExecutor(max_workers=8) as executor:
            for channel in CHANNELS:
                if not account_pipeline_success:
                    break
                    
                page_no = 1
                channel_total_pages = None  # Tracked dynamically from the first page response
                
                while True:
                    try:
                        campaign_res = fetch_campaigns(lc, channel, region=REGION, cookies=session_context["cookies"], base_urls=BASE_URLS, page_no=page_no)
                        if campaign_res.status_code in [401, 403]:
                            raise RequestException("⚠️ Unauthorised or Expired Session")
                        if campaign_res.status_code != 200:
                            account_pipeline_success = False
                            break

                        campaign_data = campaign_res.json()
                        resp_data = campaign_data.get("response", {}).get("data", {})
                        campaigns = resp_data.get("contents", [])

                        if not campaigns:
                            break
                        
                        # 📢 Log 1: Print exactly once when starting the channel to state total pages found
                        if page_no == 1:
                            channel_total_pages = resp_data.get("numberOfPages", 1)
                            logger.info(f" 📥 [{channel['name']}] Discovered {channel_total_pages} total pages to process.")

                        future_map = {}
                        for campaign in campaigns:
                            campaign_id = campaign.get("id", "UNKNOWN_ID")
                            f = executor.submit(
                                process_campaign,
                                campaign=campaign,
                                lc=lc,
                                channel=channel,
                                metrics_from=metrics_from,
                                metrics_to=metrics_to,
                                month_name=month_name,
                                REGION=REGION,
                                session_context=session_context,
                                BASE_URLS=BASE_URLS,
                                journey_cache=journey_cache
                            )
                            future_map[f] = campaign_id

                        total_submitted_this_page = len(future_map)
                        total_targets_checked += total_submitted_this_page

                        # Process this page's threads completely silently (no logs inside this loop)
                        for future in as_completed(future_map):
                            current_camp_id = future_map[future]
                            try:
                                row = future.result()
                                if row:
                                    account_rows.append(row)
                            except (ReadTimeout, ConnectionError, ChunkedEncodingError, RequestException):
                                account_pipeline_success = False
                                with session_lock:
                                    if (time.time() - session_context["updated_at"]) > 10:
                                        logger.info(f"🔄 Thread token expired alert via campaign {current_camp_id}. Refreshing credentials...")
                                        new_cookies = refresh_session(driver, REGION)
                                        session_context["cookies"] = new_cookies
                                        session_context["updated_at"] = time.time()
                                        request_access(lc, REGION, new_cookies, BASE_URLS, ROLE_IDS)
                                        time.sleep(6)
                            except Exception as e:
                                logger.debug(f"Campaign processing error on ID {current_camp_id}: {e}")

                        if not account_pipeline_success:
                            break

                        if page_no >= channel_total_pages:
                            break
                        page_no += 1
                        
                    except (ReadTimeout, ConnectionError, ChunkedEncodingError, RequestException) as list_err:
                        logger.warning(f"🔌 Connection exception inside list loop: {list_err}. Processing healing...")
                        with session_lock:
                            if (time.time() - session_context["updated_at"]) > 10:
                                session_context["cookies"] = refresh_browser_session(driver, REGION)
                                session_context["updated_at"] = time.time()
                                request_access(lc, REGION, session_context["cookies"], BASE_URLS, ROLE_IDS)
                                time.sleep(6)
                        continue
                    except Exception as e:
                        logger.error(f"Unexpected listing error for LC {lc} on page {page_no}: {e}")
                        account_pipeline_success = False
                        break

                # 📢 Log 2: Print exactly once when the entire channel finishes processing
                if account_pipeline_success and channel_total_pages:
                    logger.info(f" ✨ [{channel['name']}] Successfully processed all {channel_total_pages} pages.")

        # Log total stats aggregated cleanly per license code
        if total_targets_checked > 0:
            logger.info(f"📊 Tracking summary for {lc}: Formatted {len(account_rows)} valid metric rows out of {total_targets_checked} targets checked.")

        # Commit collected data to Google Sheet
        if account_pipeline_success:
            sheet_push_done = True
            if account_rows:
                try:
                    push_rows(worksheet, account_rows)
                    logger.info(f"✅ Sheet payload committed successfully ({len(account_rows)} records added) → {lc}")
                except Exception as e:
                    logger.error(f"❌ Sheet push execution failed for account {lc}: {e}")
                    sheet_push_done = False

            if sheet_push_done:
                mark_done(progress, REGION, lc)
                save_progress(progress)
            else:
                logger.warning(f"⚠️ Tracking exception caught. {lc} flagged for automated retry.")
        else:
            logger.error(f"⚠️ Core extraction pipeline failure caught on target mining loop for account {lc}.")

        if (idx + 1) % 15 == 0:
            logger.info("😴 Strategy cool-down period active. Sleeping for 10s...")
            time.sleep(10)

    logger.info("🎯 CHANNEL METRICS WRITER SCRIPT PIPELINE COMPLETED")
    driver.quit()

if __name__ == "__main__":
    run_channel()