import requests

session = requests.Session()

def fetch_funnel(lc, region, cookies, base_urls, page=1):

    url = f"{base_urls[region]}/api/v2/accounts/{lc}/funnels"

    params = {
        "pageNo": page,
        "query": ""
    }

    headers = {
        "Accept": "application/json",
        "User-Agent": "Mozilla/5.0",
        "Referer": f"{base_urls[region]}/accounts/{lc}/analytics"
    }

    try:
        res = session.get(
            url,
            headers=headers,
            params=params,
            cookies=cookies,
            timeout=10
        )
        return res
    except:
        return None