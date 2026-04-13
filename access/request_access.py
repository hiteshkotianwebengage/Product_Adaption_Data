import requests
from utils.logger import logger

def request_access(lc, region, cookies, base_urls, role_ids):

    base = base_urls[region].rstrip("/")

    # 🔥 STEP 1: HIT LIST ENDPOINT (VERY IMPORTANT)
    list_url = f"{base}/admin/internal-role/request-access.html?action=list&licenseCode={lc}"

    try:
        requests.get(
            list_url,
            cookies=cookies,
            headers={"User-Agent": "Mozilla/5.0"},
            timeout=10
        )
    except Exception:
        pass  # Not critical, but helps session

    # 🔥 STEP 2: ACTUAL ACCESS REQUEST
    save_url = f"{base}/admin/internal-role/request-access.html?action=save"

    payload = {
        "action": "save",               # ✅ ADD THIS (missing earlier)
        "licenseCode": lc,
        "roleEId": role_ids[region],
        "duration": "2",
        "comment": "Automated Access"
    }

    headers = {
        "User-Agent": "Mozilla/5.0",
        "Content-Type": "application/x-www-form-urlencoded",
        "Referer": f"{base}/admin/internal-role/request-access.html?action=list&licenseCode={lc}"
    }

    try:
        res = requests.post(
            save_url,
            headers=headers,
            data=payload,
            cookies=cookies,
            timeout=15
        )

        # 🔥 DEBUG LOG (VERY IMPORTANT)
        if res is None:
            logger.error(f"❌ No response → {lc}")
            return None

        if "login" in res.text.lower():
            logger.error(f"❌ Session expired → {lc}")
            return None

        if "Role Approved Successfully" in res.text:
            logger.info(f"✅ Access granted → {lc}")
            return 200

        # fallback
        logger.warning(f"⚠️ Access unclear → {lc} | Status: {res.status_code}")
        return res.status_code

    except Exception as e:
        logger.error(f"❌ Request Access Error → {lc}: {e}")
        return None