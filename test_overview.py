import time
import requests
from auth.login import init_driver
from auth.cookies import get_session_cookies
from access.request_access import request_access
from data.fetch_overview import fetch_overview
from config.settings import PRE_BASE_URLS, BASE_URLS, ROLE_IDS, get_month_info

def test_full_chain():
    REGION = "GLOBAL"
    lc = "aa133088" 
    
    # --- STEP 1: LOGIN ---
    print("🚀 Step 1: Login & Initialize Session")
    driver = init_driver("selenium_profile")
    driver.get(f"{PRE_BASE_URLS[REGION]}/admin")
    input("👉 Log in, navigate to ANY dashboard page, then press ENTER...")

    # --- STEP 2: REQUEST ACCESS (DO THIS FIRST) ---
    # We capture initial cookies just to perform the 'Request Access' action
    initial_cookies = get_session_cookies(driver)
    print(f"🔍 Requesting access for {lc}...")
    
    status = request_access(lc, REGION, initial_cookies, BASE_URLS, ROLE_IDS)
    
    if status == 200:
        print("🔓 Step 2 Success: Access Requested/Granted!")
        print("⏳ Waiting 15s for backend permissions to propagate...")
        time.sleep(15) 
    else:
        print(f"❌ Failed to request access. Status: {status}")
        driver.quit()
        return

    # --- STEP 3: SYNC & RE-CAPTURE COOKIES ---
    # Now that we have access, we navigate to the SPECIFIC account 
    # This forces the server to issue the correct session cookies for this LC
    print(f"🌐 Syncing session for account {lc}...")
    driver.get(f"https://dashboard.webengage.com/accounts/{lc}/engagement/overview/all")
    time.sleep(5) 
    
    final_cookies = get_session_cookies(driver)
    
    # --- STEP 4: DATA FETCH ---
    start_date, end_date, _, _ = get_month_info()
    payload = {
        "campaignStatType": "CHANNEL_STATS_OVERVIEW",
        "splitBy": "CAMPAIGN_CHANNEL",
        "startTime": start_date,
        "endTime": end_date,
        "channels": ["PUSH_NOTIFICATION", "EMAIL", "SMS", "WHATSAPP", "WEB_PUSH"],
        "containerTypes": ["ALL", "ONETIME", "TRIGGERED", "RECURRING", "TRANSACTION", "JOURNEY", "RELAY"],
        "isFunnelView": False,
        "tags": []
    }

    print(f"📊 Step 4: Fetching Data for {lc}...")
    res = fetch_overview(lc, REGION, final_cookies, BASE_URLS, payload)
    
    if res and res.status_code == 200:
        print("💎 SUCCESS! Data received.")
        print(f"📄 Data Preview: {res.text[:200]}...")
    else:
        status_code = res.status_code if res else "No Response"
        print(f"❌ Fetch failed. Status: {status_code}")

    driver.quit()

if __name__ == "__main__":
    test_full_chain()