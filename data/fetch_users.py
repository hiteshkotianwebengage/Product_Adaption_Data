import requests

session = requests.Session()

def fetch_users(lc, region, cookies, base_urls, start_date, end_date):

    url = f"{base_urls[region]}/api/v1/accounts/{lc}/publisher-metrics"

    params = {
        "from": start_date,
        "to": end_date,
        "metrics": "DAU,MAU,WAU"
    }

    headers = {
        "Accept": "application/json",
        "User-Agent": "Mozilla/5.0",
        "Referer": f"{base_urls[region]}/accounts/{lc}/users/overview"
    }

    try:
        res = session.get(
            url,
            headers=headers,
            params=params,
            cookies=cookies,
            timeout=60
        )
        return res
    except:
        return None