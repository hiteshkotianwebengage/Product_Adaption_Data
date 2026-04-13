# Not using this file too
import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

# --------------------------------
# Region config
# --------------------------------

REGION = "GLOBAL"

REGION_CONFIG = {
    "INDIA": {
        "base_url": "https://p1o82kY:kow3jJs9@dashboard.in.webengage.com",
        "publisher_url": "https://dashboard.in.webengage.com/admin/publisher.html?action=list"
    },
    "GLOBAL": {
        "base_url": "https://p1o82kY:kow3jJs9@dashboard.webengage.com",
        "publisher_url": "https://dashboard.webengage.com/admin/publisher.html?action=list"
    },
    "KSA": {
        "base_url": "https://p1o82kY:kow3jJs9@dashboard.ksa.webengage.com",
        "publisher_url": "https://dashboard.ksa.webengage.com/admin/publisher.html?action=list"
    }
}

# --------------------------------
# Driver setup
# --------------------------------

script_dir = os.path.dirname(os.path.abspath(__file__))
user_data_dir = os.path.join(script_dir, "selenium_profile_new")

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument(f"--user-data-dir={user_data_dir}")

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()), 
    options=options
)

wait = WebDriverWait(driver, 120)

# --------------------------------
# STEP 1 - OPEN DASHBOARD
# --------------------------------

BASE_URL = REGION_CONFIG[REGION]['base_url']

driver.get(f"{BASE_URL}/admin")

print("⏳ Waiting for page to settle ......")

time.sleep(3)

# --------------------------------
# Step 2 - This is for manually sso login and publisher page navigation
# --------------------------------

if "login" in driver.current_url or "sso" in driver.current_url:
    print("🔐 Please login via SSO...")
    input("👉 Press ENTER after login is completed...")

print(f"🔍 Current Url: {driver.current_url}")

# --------------------------------
# Step 3 - Navigate to publisher
# --------------------------------

driver.get(REGION_CONFIG[REGION]["publisher_url"])

try:
    wait.until(EC.presence_of_element_located((By.NAME, "licenseCode")))
    print("🎯 Publisher loaded")
except:
    print("⚠️ Retry loading publisher...")
    driver.get(REGION_CONFIG[REGION]["publisher_url"])
    wait.until(EC.presence_of_element_located((By.NAME, "licenseCode")))

# --------------------------------
# Step 3 - Extract Cookies
# --------------------------------
def get_cookie_string(driver):
    cookies = driver.get_cookies()

    cookie_string = "; ".join(
        [f"{c['name']}={c['value']}" for c in cookies]
    )

    return cookie_string

cookie_string = get_cookie_string(driver)

print("\n 🍪Cookie Captured: \n ")
print(cookie_string)

with open("cookie.txt", "w") as f:
    f.write(cookie_string)

print("\n✅ Cookie saved to cookie.txt")

# Keep browser open if needed
# input("\nPress ENTER to close browser...")
# driver.quit()