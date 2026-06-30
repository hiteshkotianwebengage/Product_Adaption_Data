# For now we have not found it but we can keep it ready to change only few thins

import requests

session = requests.Session()

def fetch_conversion(lc, conversion_id, region, cookies, base_urls):

    url = {
        f"{base_urls[region]}"
        f"/api/v1/accounts/{lc}"
        f"/conversions/{conversion_id}"
    }

    headers = {
        "Accept": "application/json, text/plain, */*",
        "x-accounts-id": lc,
        "User-Agent": "Mozilla/5.0"
    }

    return session.get(
        url, headers=headers, cookies=cookies, timeout=60
    )