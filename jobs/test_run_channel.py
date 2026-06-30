import time
import argparse
from datetime import datetime
from threading import Lock
from requests.exceptions import RequestException

from auth.login import init_driver
from auth.cookies import get_session_cookies
from access.request_access import request_access

from data.fetch_channel import fetch_campaigns
from config.settings import PRE_BASE_URLS, BASE_URLS, ROLE_IDS
from config.channel_config import CHANNELS
from data.load_lc import load_licence_codes
from utils.logger import logger

# IMPORT THE ACTUAL PROCESS CAMPAIGN FUNCTION TO TEST IT DIRECTLY
from jobs.run_channel import process_campaign 

def run_targeted_test():
    parser = argparse.ArgumentParser()
    parser.add_argument("--region", required=True)
    args = parser.parse_args()
    REGION = args.region.upper()

    logger.info(f"🧪 Testing Actual process_campaign Function Logic for {REGION}")

    # Test configurations
    test_months = [
        {
            "label": "Feb'26",
            "from": "2026-02-01T00:00:00.000+05:30",
            "to": "2026-02-28T23:59:59.999+05:30"
        },
        {
            "label": "Jun'26",
            "from": "2026-06-01T00:00:00.000+05:30",
            "to": "2026-06-30T23:59:59.999+05:30"
        }
    ]

    # Target campaign IDs to audit
    TARGET_CAMPAIGN_IDS = {"2nesoe7", "311lse1"}

    # Driver Setup
    driver = init_driver("selenium_profile")
    driver.get(f"{PRE_BASE_URLS[REGION]}/admin")
    input("👉 Log in via browser and press ENTER here to begin...")

    publisher_list_url = f"{PRE_BASE_URLS[REGION]}/admin/publisher.html?action=list"
    driver.get(publisher_list_url)
    time.sleep(5)

    session_context = {"cookies": get_session_cookies(driver)}
    journey_cache = {}

    lc_region_list = load_licence_codes()
    license_codes = [lc for lc, r in lc_region_list if r == REGION]

    for lc in license_codes:
        request_access(lc, REGION, session_context["cookies"], BASE_URLS, ROLE_IDS)
        time.sleep(4)

        for channel in CHANNELS:
            page_no = 1
            while True:
                try:
                    res = fetch_campaigns(lc, channel, region=REGION, cookies=session_context["cookies"], base_urls=BASE_URLS, page_no=page_no)
                    if res.status_code != 200:
                        break
                    
                    data = res.json().get("response", {}).get("data", {})
                    campaigns = data.get("contents", [])
                    if not campaigns:
                        break

                    for campaign in campaigns:
                        c_id = campaign.get("id")
                        if c_id in TARGET_CAMPAIGN_IDS:
                            c_name = campaign.get("title", "Unknown")
                            
                            logger.info(f"\n🎯 Target Found: {c_name} ({c_id})")
                            
                            for m_info in test_months:
                                logger.info(f"⏳ Passing to process_campaign for Window: {m_info['label']}")
                                
                                try:
                                    # Executing the real processing logic
                                    row = process_campaign(
                                        campaign=campaign,
                                        lc=lc,
                                        channel=channel,
                                        metrics_from=m_info["from"],
                                        metrics_to=m_info["to"],
                                        month_name=m_info["label"],
                                        REGION=REGION,
                                        cookies=session_context["cookies"],
                                        BASE_URLS=BASE_URLS,
                                        journey_cache=journey_cache
                                    )
                                    
                                    if row:
                                        logger.info(f"📊 [Evaluation result for {m_info['label']}]: ✅ PASSED (Returned valid row data dict)")
                                    else:
                                        logger.info(f"📊 [Evaluation result for {m_info['label']}]: 🚫 DISCARDED (Returned None)")
                                        
                                except Exception as ex:
                                    logger.error(f"💥 Exception caught inside process_campaign pipeline: {ex}", exc_info=True)

                    if page_no >= data.get("numberOfPages", 1):
                        break
                    page_no += 1

                except Exception as loop_err:
                    logger.error(f"Error checking master list loop: {loop_err}")
                    break

    logger.info("🎯 INTEGRATION PROCESS_CAMPAIGN TEST RUN COMPLETE.")
    driver.quit()

if __name__ == "__main__":
    run_targeted_test()