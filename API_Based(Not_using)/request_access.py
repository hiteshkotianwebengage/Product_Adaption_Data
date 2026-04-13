# We are not using this file too 
import requests

BASE_URLS = {
    "INDIA": "https://dashboard.in.webengage.com",
    "GLOBAL": "https://dashboard.webengage.com",
    "KSA": "https://dashboard.ksa.webengage.com"
}

region = "GLOBAL"

license_code = "~c2ab2ba2"

with open("cookie.txt", "r") as f:
    cookie_string = f.read().strip()

def request_access(license_code, region, cookie):
    base_url = BASE_URLS[region]

    clean_lc = license_code.replace("~", "")

    url = f"{BASE_URLS[region]}/admin/request-access.html"

    payload = {
        "licenseCode": clean_lc,
        "roleEId": "abd40ke",
        "duration": "1",
        "comment": "auto access request"
    }

    headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
        'Authorization': 'Basic cDFvODJrWTprb3czakpzOQ==',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Origin': base_url,
        'Referer': f'{base_url}/admin/internal-role/request-access.html?action=list&licenseCode={clean_lc}',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36',
        'Cookie': cookie_string  # Pass the full string from cookie.txt
    }

    params = {"action": "save"}

    res = requests.post(url, headers=headers, data=payload, params=params)

    print("\n🔐 REQUEST ACCESS DEBUG")
    print("STATUS:", res.status_code)
    print("RESPONSE:", res.text[:500])

    if res.status_code == 200:
        print("✅ Success: Request saved.")
    else:
        print(f"❌ Failed: {res.status_code}")

    return res

if __name__ == "__main__":
    request_access(license_code, region, cookie_string)