import requests

def fetch_overview(lc, region, cookies, base_urls, payload):

    url = f"{base_urls[region]}/api/v2/accounts/{lc}/stats"

    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json, text/plain, */*",
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36",
        "Referer": f"{base_urls[region]}/accounts/{lc}/engagement/overview/all"
    }

    res = requests.post(
        url,
        headers= headers,
        json=payload,
        cookies=cookies
    )

    return res