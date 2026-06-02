import time
import requests

from utils.logger import logger

session = requests.Session()


# ---------------------------------------------------------
# EVENT LIST
# ---------------------------------------------------------

def fetch_event_list(
    lc,
    region,
    cookies,
    base_urls,
    page=1
):

    try:

        base = base_urls.get(region).rstrip('/')

        url = (
            f"{base}"
            f"/api/v2/accounts/{lc}/event/list"
        )

        params = {
            "actionEntity": "event",
            "category": "application",
            "pageSize": 100,
            "page": page,
            "eventName": "",
            "query": ""
        }

        headers = {
            "Content-Type": "application/json;charset=UTF-8",
            "Accept": "application/json, text/plain, */*",
            "User-Agent": "Mozilla/5.0"
        }

        for attempt in range(2):

            try:

                res = session.get(
                    url,
                    headers=headers,
                    params=params,
                    cookies=cookies,
                    timeout=10
                )

                if res.status_code == 200:
                    return res

                logger.warning(
                    f"⚠️ Event List Retry "
                    f"| {lc} "
                    f"| Status={res.status_code}"
                )

            except requests.exceptions.RequestException as e:

                logger.warning(
                    f"⚠️ Event List Network Retry "
                    f"| {lc} | {e}"
                )

            time.sleep(1)

    except Exception as e:

        logger.error(
            f"❌ Event List Failed "
            f"| {lc} | {e}"
        )

    return None


# ---------------------------------------------------------
# EVENT METADATA
# ---------------------------------------------------------

def fetch_event_metadata(
    lc,
    event_name,
    region,
    cookies,
    base_urls
):

    try:

        base = base_urls.get(region).rstrip('/')

        url = (
            f"{base}"
            f"/api/v3/accounts/{lc}/event/metadata"
        )

        params = {
            "category": "application"
        }

        payload = {
            "eventName": event_name
        }

        headers = {
            "Content-Type": "application/json;charset=UTF-8",
            "Accept": "application/json, text/plain, */*",
            "User-Agent": "Mozilla/5.0"
        }

        for attempt in range(2):

            try:

                res = session.post(
                    url,
                    headers=headers,
                    params=params,
                    json=payload,
                    cookies=cookies,
                    timeout=5
                )

                if res.status_code == 200:
                    return res

                logger.warning(
                    f"⚠️ Metadata Retry "
                    f"| {lc} "
                    f"| {event_name} "
                    f"| Status={res.status_code}"
                )

            except requests.exceptions.RequestException as e:

                logger.warning(
                    f"⚠️ Metadata Network Retry "
                    f"| {lc} "
                    f"| {event_name} "
                    f"| {e}"
                )

            time.sleep(1)

    except Exception as e:

        logger.error(
            f"❌ Metadata Failed "
            f"| {lc} "
            f"| {event_name} "
            f"| {e}"
        )

    return None