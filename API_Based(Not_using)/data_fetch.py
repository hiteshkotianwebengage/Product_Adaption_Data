# We are not using this now as we distibuted into differnt sheets this was the starting file
import requests
import json
import pandas as pd
import time

df = pd.read_csv("/Users/admin/Desktop/Code Directory/Product_Adaption_Data/config/license_code.csv")
license_codes = df["Licence Code"].dropna().tolist()

BASE_URLS = {
    "INDIA": "https://dashboard.in.webengage.com",
    "GLOBAL": "https://dashboard.webengage.com",
    "KSA": "https://dashboard.ksa.webengage.com"
}


start_date = "2026-02-01T00:00:00.000+05:30"
end_date = "2026-02-28T23:59:59.999+05:30"

headers = {
    "accept": "application/json",
    "content-type": "application/json",
    "authorization": "Basic cDFvODJrWTprb3czakpzOQ==",
    "cookie": "WebKlipperAuth=ecGonltzW8DsnpX36WSJ; _we_us=1774581233525",
}

region = "GLOBAL"

def clean(val):
    return val if val not in [None, "-", ""] else 0

def request_access(license_code, region, headers):

    base_url = BASE_URLS[region]
    url = f"{base_url}/admin/request-access.html?action=save"

    payload = {
        "action": "save",
        "licenseCode": license_code,
        "roleEId": "abd40ke",
        "duration": "1",
        "comment": "auto access request"
    }

    headers_access = {
        "content-type": "application/x-www-form-urlencoded",
        "cookie": headers["cookie"],
    }

    res = requests.post(url, headers=headers_access, data=payload)

    print(f"\n🔐 Request access for {license_code}")
    print("STATUS:", res.status_code)
    print("RESPONSE:", res.text[:300])  # 🔥 IMPORTANT

    return res.status_code == 200

def fetch_overview(license_code, region, start_date, end_date, headers):

    base_url = BASE_URLS[region]
    url = f"{base_url}/api/v2/accounts/{license_code}/stats"

    payload = {
        "campaignStatType": "CHANNEL_STATS_OVERVIEW",
        "splitBy": "CAMPAIGN_CHANNEL",
        "startTime": start_date,
        "endTime": end_date,
        "channels": [
            "OVERALL", "PUSH_NOTIFICATION", "IN_APP_NOTIFICATION",
            "SMS", "ON_SITE_NOTIFICATION", "WEB_PUSH",
            "EMAIL", "WHATSAPP", "WEB_PERSONALIZATION"
        ],
        "containerTypes": ["ALL"],
        "isFunnelView": False,
        "tags": []
    }

# ----------- API CALL -----------
    res = requests.post(url, headers=headers, json=payload)

    if res.status_code != 200:
        print(f"❌ API Failed for {license_code}: {res.status_code}")
        print(res.text[:200])  # preview error
        return None, res.status_code

    return res.json(), 200

def parse_overview(data, license_code):

    rows = []

    try:
        dimension = data['response']['data'][0]['dimensions']
    except:
        print(f"⚠️ Bad response for {license_code}")
        return []

    for item in dimension:

        channel = item['value']

        if channel == "OVERALL":
            continue

        metrics = {m['name'] : m['value'] for m in item['metrics']}

        users = clean(metrics.get("users"))
        campaigns = clean(metrics.get("campaigns"))
        deliveries = clean(metrics.get("deliveries"))
        clicks = clean(metrics.get("clicks"))
        revenue = clean(metrics.get("revenue"))
        conversions = clean(metrics.get("goal"))

        ctr = (clicks / deliveries) if deliveries else 0
        cvr = (conversions / deliveries) if deliveries else 0

        rows.append([
            license_code,
            channel,
            users,
            campaigns,
            deliveries,
            round(ctr * 100, 2),
            round(cvr * 100, 2),
            revenue
        ])

    return rows
    
all_rows = []

for lc in license_codes:

    data, status = fetch_overview(lc, region, start_date, end_date, headers)

    if status == 403:
        print(f"🔒 No access → requesting: {lc}")

        success = request_access(lc, region, headers)

        if success:
            print("⏳ Waiting for access...")
            time.sleep(3)

            data, status = fetch_overview(lc, region, start_date, end_date, headers)

        else:
            print(f"❌ Access request failed: {lc}")
            continue

    if data:
        rows = parse_overview(data, lc)
        all_rows.extend(rows)

for row in all_rows:
     print(row)