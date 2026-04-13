from time import time

import requests
import traceback
import urllib3
from utils.logger import logger

session = requests.Session()

def fetch_overview(lc, region, cookies, base_urls, payload):
    try:
        base = base_urls.get(region).rstrip('/')
        url = f"{base}/api/v2/accounts/{lc}/stats"

        headers = {
            "Content-Type": "application/json;charset=UTF-8",
            "Accept": "application/json, text/plain, */*",
            "X-Requested-With": "XMLHttpRequest",
            "x-accounts-id": lc, 
            "Origin": base,
            "Referer": f"{base}/accounts/{lc}/engagement/overview/all",
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
        }

        # Adding verify=False and a shorter timeout to force a quick failure/success
        for _ in range(2):
            try:
                res = session.post(
                    url,
                    headers=headers,
                    json=payload,
                    cookies=cookies,
                    timeout=10
                )
                return res
            except:
                time.sleep(1)

    except Exception as e:
        # Keep one generic catch-all for logging
        logger.error(f"❌ API Failed | {lc} | {url} | {e}")
        return None