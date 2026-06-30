import requests

session = requests.Session()

def fetch_channel(lc, channel, region, cookies, base_urls, page_no=1):
    endpoint = channel['endpoint']
    version = channel['version']

    url = f"{base_urls[region]}/api/{version}/accounts/{lc}/{endpoint}"

    params = {
        "pageNo": page_no,
        "pageSize": 10
    }

    params.update(channel.get("params", {}))

    headers = {
        "Accept": "application/json, text/plain, */*",
        "Connection": "keep-alive", # 👈 Add this
        "x-accounts-id": lc,
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36..."
    }

    res = session.get(
        url,
        headers=headers,
        params=params,
        cookies=cookies,
        timeout=60
    )

    return res