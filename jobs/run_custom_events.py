import os
import json
import time
import random
import argparse
import traceback
import requests

from auth.login import init_driver
from auth.cookies import get_session_cookies

from access.request_access import request_access

from data.fetch_custom_events import (
    fetch_event_list,
    fetch_event_metadata
)

from data.parser_custom_events import (
    parse_custom_event_metadata
)

from data.load_lc import load_licence_codes

from data.sheet_writer import (
    get_gsheet_client,
    get_or_create_worksheet,
    push_rows,
    SPREADSHEET_ID_C_R_A
)

from config.settings import (
    PRE_BASE_URLS,
    BASE_URLS,
    ROLE_IDS
)

from config.headers import (
    CUSTOM_EVENTS_HEADER
)

from utils.logger import logger

from requests.exceptions import (
    ReadTimeout,
    ConnectionError,
    ChunkedEncodingError
)

PROGRESS_FILE = os.path.join("Progress_File", "progress_custom_events.json")


# ---------------------------------------------------------
# PROGRESS
# ---------------------------------------------------------

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

# ---------------------------------------------------------
# RUNNER
# ---------------------------------------------------------

def run_custom_events():

    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--region",
        required=True
    )

    args = parser.parse_args()

    REGION = args.region.upper()

    logger.info(
        f"🚀 Custom Events Started | {REGION}"
    )

    # ---------------------------------------------------------
    # LOAD LCS
    # ---------------------------------------------------------

    lc_region_list = load_licence_codes()

    license_codes = [

        lc
        for lc, r in lc_region_list
        if r == REGION
    ]

    if not license_codes:

        logger.error(
            f"❌ No LC Found | {REGION}"
        )

        return

    # ---------------------------------------------------------
    # LOGIN
    # ---------------------------------------------------------

    driver = init_driver("selenium_profile")

    driver.get(
        f"{PRE_BASE_URLS[REGION]}/admin"
    )

    input("👉 Login and press ENTER...")

    events_url = (
        f"{PRE_BASE_URLS[REGION]}"
        f"/accounts/{license_codes[0]}"
        f"/data-management/events/attributes"
    )

    driver.get(events_url)

    time.sleep(8)

    cookies = get_session_cookies(driver)

    # ---------------------------------------------------------
    # SHEETS
    # ---------------------------------------------------------

    client = get_gsheet_client()

    spreadsheet = client.open_by_key(
        SPREADSHEET_ID_C_R_A
    )

    worksheet = get_or_create_worksheet(
        spreadsheet,
        f"Custom Events {REGION}",
        CUSTOM_EVENTS_HEADER
    )

    # ---------------------------------------------------------
    # PROGRESS
    # ---------------------------------------------------------

    progress = load_progress()

    # ---------------------------------------------------------
    # MAIN LOOP
    # ---------------------------------------------------------

    for i, lc in enumerate(license_codes):

        if lc in progress.get(REGION, []):

            logger.info(f"⏭️ Skipping | {lc}")

            continue

        logger.info(
            f"🔍 [{i+1}/{len(license_codes)}] {lc}"
        )

        final_rows = []

        # ---------------------------------------------------------
        # ACCESS
        # ---------------------------------------------------------

        try:

            request_access(
                lc,
                REGION,
                cookies,
                BASE_URLS,
                ROLE_IDS
            )

            logger.info(
                f"✅ Access Granted | {lc}"
            )

            time.sleep(5)

        except Exception as e:

            logger.warning(
                f"⚠️ Access Failed | {lc} | {e}"
            )

        # ---------------------------------------------------------
        # FETCH EVENTS
        # ---------------------------------------------------------

        all_events = []

        page = 1

        while True:

            response = fetch_event_list(
                lc=lc,
                region=REGION,
                cookies=cookies,
                base_urls=BASE_URLS,
                page=page
            )

            if (
                not response
                or response.status_code != 200
            ):
                break

            data = response.json()

            resp_data = (
                data.get("response", {})
                    .get("data", {})
            )

            contents = (
                resp_data.get("contents")
                or []
            )

            total_pages = (
                resp_data.get("totalPages")
                or resp_data.get("numberOfPages")
                or 1
            )

            if not contents:
                break

            all_events.extend(contents)

            logger.info(
                f"{lc} | "
                f"Page={page}/{total_pages} | "
                f"Events={len(contents)}"
            )

            if page >= total_pages:
                break

            page += 1

            time.sleep(0.5)

        logger.info(
            f"📦 Total Events | "
            f"{lc} | {len(all_events)}"
        )

        # ---------------------------------------------------------
        # LOOP EVENTS
        # ---------------------------------------------------------

        for idx, event in enumerate(all_events, start=1):

            event_name = str(
                event.get("name", "")
            ).strip()

            if not event_name:
                continue

            logger.info(
                f"➡️ {idx}/{len(all_events)} | "
                f"{event_name}"
            )

            # ---------------------------------------------------------
            # EVENT LEVEL
            # ---------------------------------------------------------

            event_statuses = {

                s.get("sdk"): s.get("value")

                for s in event.get("status", [])
                if isinstance(s, dict)
            }

            final_rows.append([

                str(lc),
                event_name,
                "--EVENT_LEVEL--",
                "EVENT",
                "NO_DATA",
                str(event.get("ignored")),

                event_statuses.get(
                    "web",
                    "NO_DATA"
                ),

                event_statuses.get(
                    "android",
                    "NO_DATA"
                ),

                event_statuses.get(
                    "ios",
                    "NO_DATA"
                ),

                event_statuses.get(
                    "none",
                    "NO_DATA"
                )
            ])

            # ---------------------------------------------------------
            # ATTRIBUTE LEVEL
            # ---------------------------------------------------------

            try:

                metadata_response = fetch_event_metadata(
                    lc=lc,
                    event_name=event_name,
                    region=REGION,
                    cookies=cookies,
                    base_urls=BASE_URLS
                )

                if (
                    not metadata_response
                    or metadata_response.status_code != 200
                ):

                    continue

                metadata_data = (
                    metadata_response.json()
                )

                parsed_rows = (
                    parse_custom_event_metadata(
                        metadata_data,
                        lc,
                        event_name
                    )
                )

                if parsed_rows:
                    final_rows.extend(parsed_rows)

                time.sleep(
                    random.uniform(0.05, 0.15)
                )

            except Exception as e:

                logger.error(
                    f"❌ Event Failed | "
                    f"{event_name} | {e}"
                )

        # ---------------------------------------------------------
        # PUSH
        # ---------------------------------------------------------

        logger.info(
            f"📊 Final Rows | "
            f"{lc} | {len(final_rows)}"
        )

        try:

            chunk_size = 1000

            for start in range(
                0,
                len(final_rows),
                chunk_size
            ):

                chunk = final_rows[
                    start:start + chunk_size
                ]

                push_rows(
                    worksheet,
                    chunk
                )

                logger.info(
                    f"✅ Chunk "
                    f"{start + len(chunk)}"
                    f"/{len(final_rows)}"
                )

                time.sleep(2)

            mark_done(
                progress,
                REGION,
                lc
            )

            save_progress(progress)

            logger.info(
                f"💾 Saved | {lc}"
            )

        except Exception as e:

            logger.error(
                f"❌ Push Failed | "
                f"{lc} | {repr(e)}"
            )

            logger.error(
                traceback.format_exc()
            )

    logger.info(
        f"🎯 Custom Events Completed | "
        f"{REGION}"
    )

    driver.quit()


if __name__ == "__main__":

    run_custom_events()