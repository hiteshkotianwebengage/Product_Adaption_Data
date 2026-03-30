import requests

def request_access(lc, region, cookies, base_urls, role_ids):

    url = f"{base_urls[region]}/admin/internal-role/request-access.html"

    payload = {
        "licenseCode": lc,
        "roleEId": role_ids[region],
        "duration": "1",
        "comment": "auto access"
    }

    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }

    params = {
        "action": "save"
    }

    res = requests.post(
        url, headers=headers, data=payload, params=params, cookies=cookies
    )

    return res.status_code