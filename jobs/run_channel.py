import time
import json
import os
import argparse
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

# Dynamic Utilities Imported From External Modules
from utils.logger import logger
from utils.progress_manager import (progress_lock, load_progress, save_progress_raw, save_page_checkpoint)
from utils.helpers import has_valid_metrics

# ----------------------
# GLOBAL THREADING LOCKS
# ----------------------
session_lock = Lock()
journey_cache_lock = Lock()

# ----------------------
# MODULAR REFRESH SESSION
# ----------------------

def refresh_session(driver, region):
    logger.warning("🔄 Session expired or access denied → executing recovery refresh...")
    
    # 1. Re-anchor the administrative window
    driver.get(f"{PRE_BASE_URLS[region]}/admin")
    time.sleep(8)

    # 2. Open the critical publisher list page to generate tokens
    publisher_list_url = f"{PRE_BASE_URLS[region]}/admin/publisher.html?action=list"
    driver.get(publisher_list_url)
    time.sleep(6)

    # 3. Capture fresh cookies and hand them back
    return get_session_cookies(driver)

# ----------------------
# CORE CAMPAIGN WORKER THREAD
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

    # STARTUP LOGISTIC ANCHORS
    driver = init_driver("selenium_profile")
    driver.get(f"{PRE_BASE_URLS[REGION]}/admin")
    input("👉 Login and press ENTER...")

    publisher_list_url = f"{PRE_BASE_URLS[REGION]}/admin/publisher.html?action=list"
    driver.get(publisher_list_url)
    time.sleep(6)

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
        account_state = progress.get(REGION, {}).get(lc, {})
        if account_state.get("status") == "SUCCESS":
            continue

        logger.info(f"💼 Extracting LC: {lc} [{idx + 1}/{len(license_codes)}]")
        
        status = request_access(lc, REGION, session_context["cookies"], BASE_URLS, ROLE_IDS)

        if status in [401, 403] or status != 200:
            logger.warning(f"⚠️ Access failed (Status {status}) for {lc}. Session might be expired. Refreshing...")
            with session_lock:
                if (time.time() - session_context["updated_at"]) > 10:
                    session_context["cookies"] = refresh_session(driver, REGION)
                    session_context["updated_at"] = time.time()
                    status = request_access(lc, REGION, session_context["cookies"], BASE_URLS, ROLE_IDS)
            
            if status != 200:
                logger.error(f"❌ Access failed permanently for {lc} even after refresh attempt. Skipping account.")
                continue

        logger.info(f"✅ Access established for LC: {lc}. Launching worker pipeline...")
        time.sleep(4)

        channel_states = account_state.get("channels", {})
        account_rows = []
        account_pipeline_success = True
        total_targets_checked = 0
        
        # Upgraded thread execution footprint to the sweet spot (12 workers)
        with ThreadPoolExecutor(max_workers=12) as executor:
            for channel in CHANNELS:
                ch_name = channel['name']

                if not account_pipeline_success:
                    break

                ch_checkpoint = channel_states.get(ch_name, {"last_successful_page": 0, "completed": False})
                if ch_checkpoint["completed"]:
                    logger.info(f"⏭️ Skipping {ch_name} for {lc} (Already marked completed in progress file)")
                    continue
                
                page_no = ch_checkpoint["last_successful_page"] + 1
                channel_total_pages = None

                logger.info(f"🚀 Starting/Resuming [{ch_name}] for {lc} at Page {page_no}")
                
                while True:
                    try:
                        campaign_res = fetch_campaigns(lc, channel, region=REGION, cookies=session_context["cookies"], base_urls=BASE_URLS, page_no=page_no)
                        
                        # 🔄 AUTHENTICATION SELF-HEALING ENGINE
                        if campaign_res.status_code in [401, 403]:
                            with session_lock:
                                if (time.time() - session_context["updated_at"]) > 10:
                                    session_context["cookies"] = refresh_session(driver, REGION)
                                    session_context["updated_at"] = time.time()
                                    status = request_access(lc, REGION, session_context["cookies"], BASE_URLS, ROLE_IDS)
                                    if status == 200:
                                        time.sleep(10)
                            continue

                        if campaign_res.status_code != 200:
                            account_pipeline_success = False
                            break

                        campaign_data = campaign_res.json()
                        resp_data = campaign_data.get("response", {}).get("data", {})
                        campaigns = resp_data.get("contents", [])

                        if not campaigns:
                            with progress_lock:
                                save_page_checkpoint(REGION, lc, ch_name, page_no - 1, completed=True)
                            break
                        
                        if page_no == 1 or channel_total_pages is None:
                            channel_total_pages = resp_data.get("numberOfPages", 1)
                            logger.info(f" 📥 [{ch_name}] Total volume footprint: {channel_total_pages} pages.")

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

                        for future in as_completed(future_map):
                            try:
                                row = future.result()
                                if row:
                                    account_rows.append(row)
                            except Exception as e:
                                logger.debug(f"Campaign exception skipped: {e}")

                        # 📈 MEMORY SHIELD: Flush to GSheets every 50 pages for safety
                        if page_no % 50 == 0 and account_rows:
                            logger.info(f"💾 Milestone Hit! Flushing {len(account_rows)} rows to sheet for safety...")
                            push_rows(worksheet, account_rows)
                            account_rows.clear()
                            with progress_lock:
                                save_page_checkpoint(REGION, lc, ch_name, page_no, completed=(page_no >= channel_total_pages))

                        # 💾 SAVE STATE CHECKPOINT (Optimized to write every 10 pages or terminal page)
                        elif page_no % 10 == 0 or page_no >= channel_total_pages:
                            with progress_lock:
                                save_page_checkpoint(REGION, lc, ch_name, page_no, completed=(page_no >= channel_total_pages))
                            
                        page_no += 1
                        
                    except (ReadTimeout, ConnectionError, ChunkedEncodingError, RequestException) as list_err:
                        logger.warning(f"🔌 Connection drop caught inside list tracking: {list_err}. Activating healing...")
                        with session_lock:
                            if (time.time() - session_context["updated_at"]) > 10:
                                session_context["cookies"] = refresh_session(driver, REGION)
                                session_context["updated_at"] = time.time()
                                request_access(lc, REGION, session_context["cookies"], BASE_URLS, ROLE_IDS)
                                time.sleep(6)
                        continue
                    except Exception as e:
                        logger.error(f"Unexpected listing error for LC {lc} on page {page_no}: {e}")
                        account_pipeline_success = False
                        break

                if account_pipeline_success and channel_total_pages:
                    logger.info(f" ✨ [{ch_name}] Successfully processed all {channel_total_pages} pages.")

        if total_targets_checked > 0:
            logger.info(f"📊 Tracking summary for {lc}: Formatted {len(account_rows)} residual valid metric rows out of {total_targets_checked} targets checked.")

        if account_pipeline_success:
            sheet_push_done = True
            if account_rows:
                try:
                    push_rows(worksheet, account_rows)
                    logger.info(f"✅ Final sheet payload committed successfully ({len(account_rows)} records added) → {lc}")
                    account_rows.clear()
                except Exception as e:
                    logger.error(f"❌ Sheet push execution failed for account {lc}: {e}")
                    sheet_push_done = False

            if sheet_push_done:
                with progress_lock:
                    try:
                        current_data = load_progress()
                        if REGION not in current_data:
                            current_data[REGION] = {}
                        if lc not in current_data[REGION]:
                            current_data[REGION][lc] = {}
                        
                        current_data[REGION][lc]["status"] = "SUCCESS"
                        current_data[REGION][lc]["updated_at"] = datetime.now().isoformat()
                        save_progress_raw(current_data)
                    except Exception as e:
                        logger.error(f"❌ Failed to mark account success: {e}")
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