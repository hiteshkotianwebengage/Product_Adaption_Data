import requests

session = requests.Session()

def fetch_revenue_mapping(lc, region, cookies, base_urls):

    url = f"{base_urls[region]}/api/v2/accounts/{lc}/revenue/mapping"

    headers = {
        "Accept": "application/json",
        "User-Agent": "Mozilla/5.0",
        "Referer": f"{base_urls[region]}/accounts/{lc}/data-management/events/revenue"
    }

    try:
        res = session.get(
            url,
            headers=headers,
            cookies=cookies,
            timeout=10
        )
        return res
    except:
        return None