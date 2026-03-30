import requests

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
        "Accept": "application/json",
        "User-Agent": "Mozilla/5.0"
    }

    res = requests.get(url, headers=headers, params=params, cookies=cookies)

    return res