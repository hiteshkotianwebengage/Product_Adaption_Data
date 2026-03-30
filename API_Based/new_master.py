import os
import time
import pandas as pd
import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# --- 1. SETUP & PATHS ---
REGION = "GLOBAL"
PRE_BASE_URL = "https://p1o82kY:kow3jJs9@dashboard.webengage.com"
BASE_URL = "https://dashboard.webengage.com"
csv_path = "/Users/admin/Desktop/Code Directory/Product_Adaption_Data/config/license_code.csv"

# Load Licenses
df = pd.read_csv(csv_path)
license_codes = df["Licence Code"].dropna().tolist()

# --- 2. BROWSER LOGIN ---
script_dir = os.path.dirname(os.path.abspath(__file__))
user_data_dir = os.path.join(script_dir, "selenium_profile")

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument(f"--user-data-dir={user_data_dir}")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 60)

print(f"🚀 Opening Dashboard for {REGION}...")
driver.get(f"{PRE_BASE_URL}/admin")

# Wait for manual SSO if needed
if "login" in driver.current_url or "sso" in driver.current_url:
    print("🔐 Please login via SSO in the browser...")
    input("👉 Press ENTER here AFTER you see the dashboard load...")

# --- 3. THE REQUEST LOOP ---
def request_access_api(lc, session_cookies):
    # Clean tilde if present
    # clean_lc = lc.replace("~", "")
    
    # Use the EXACT path from your successful cURL
    url = f"{BASE_URL}/admin/internal-role/request-access.html"
    
    headers = {
        'Authorization': 'Basic cDFvODJrWTprb3czakpzOQ==', # Your Basic Auth
        'Content-Type': 'application/x-www-form-urlencoded',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36',
        'Referer': f'{BASE_URL}/admin/internal-role/request-access.html?action=list&licenseCode={license_codes}',
        'Origin': BASE_URL
    }

    payload = {
        'licenseCode': license_codes,
        'roleEId': 'abd40ke',
        'duration': '1',
        'comment': 'Auto-access script'
    }

    # We use requests.Session to keep the connection "warm"
    response = requests.post(url, headers=headers, data=payload, params={'action': 'save'}, cookies=session_cookies)
    return response.status_code

# --- 4. EXECUTION ---
print(f"🔄 Processing {len(license_codes)} licenses...")

# Capture cookies from the ACTIVE browser session
selenium_cookies = driver.get_cookies()
# Convert Selenium cookies to a format 'requests' understands
session_cookies = {c['name']: c['value'] for c in selenium_cookies}

for lc in license_codes:
    status = request_access_api(lc, session_cookies)
    if status == 200:
        print(f"✅ Access requested for {lc}")
    else:
        print(f"❌ Failed for {lc} (Status: {status})")
    
    time.sleep(1) # Small delay to avoid spamming the server

print("\n🎯 All requests completed!")
driver.quit()