import requests

session = requests.Session()

def fetch_dashboard(lc, region, cookies, base_urls):

    url = f"{base_urls[region]}/api/v1/accounts/{lc}/dashboard"

    headers = {
        "Accept": "application/json, text/plain, */*",
        "User-Agent": "Mozilla/5.0",
        "Referer": f"{base_urls[region]}/accounts/{lc}/custom-dashboard/list"
    }

    try:
        res = session.get(
            url,
            headers=headers,
            cookies=cookies,
            timeout=60
        )
        return res

    except Exception:
        return None