'''
    1) Change REGION
    2) It will check if the LC exist in the region or not also if it is in service or not
    3) Check if the json file of credential file path is right - This is for me to remember if in future i change the system
'''

import time
import os
import gspread
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from google.oauth2.service_account import Credentials

# ===============================
# CONFIG
# ===============================

SHEET_ID = "1o5QRUGQYptkwe1NdsZcfgD44fQCjjkfY2D_DSftSYa4"
SHEET_NAME = "LC_Check"


# ===============================
# Region
# ===============================

REGION = "INDIA"

REGION_CONFIG = {
    "INDIA": {
        "base_url": "https://p1o82kY:kow3jJs9@dashboard.in.webengage.com",
        "publisher_url": "https://p1o82kY:kow3jJs9@dashboard.in.webengage.com/admin/publisher.html?action=list",
        "column": 2
    },
    "GLOBAL": {
        "base_url": "https://p1o82kY:kow3jJs9@dashboard.webengage.com",
        "publisher_url": "https://p1o82kY:kow3jJs9@dashboard.webengage.com/admin/publisher.html?action=list",
        "column": 3
    },
    "KSA": {
        "base_url": "https://p1o82kY:kow3jJs9@dashboard.ksa.webengage.com",
        "publisher_url": "https://p1o82kY:kow3jJs9@dashboard.ksa.webengage.com/admin/publisher.html?action=list",
        "column": 4
    }
}


# ===============================
# GOOGLE SHEETS
# ===============================

def init_sheet():

    scopes = ["https://www.googleapis.com/auth/spreadsheets"]

    creds = Credentials.from_service_account_file(
        "/Users/admin/Desktop/Code Directory/Product_Adaption_Data/Credential File/mycred-googlesheet.json",
        scopes=scopes
    )

    client = gspread.authorize(creds)

    sheet = client.open_by_key(SHEET_ID).worksheet(SHEET_NAME)

    rows = sheet.get_all_values()

    license_codes = [row[0] for row in rows[1:] if row[0].strip() != ""]

    print("Total LC found:", len(license_codes))

    return sheet

# ===============================
# SEARCH LICENSE
# ===============================

# ========== We need no change from here till the edit button clicked ========== #

# ---------- STEP 0: SETUP PERSISTENT PROFILE ----------
script_dir = os.path.dirname(os.path.abspath(__file__))
user_data_dir = os.path.join(script_dir, "selenium_profile")

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument(f"--user-data-dir={user_data_dir}")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 120) 

# Step 1: Open admin dashboard
BASE_URL = REGION_CONFIG[REGION]["base_url"]
driver.get(f"{BASE_URL}/admin")

# --- GIVE IT A MOMENT TO REDIRECT ---
print("⏳ Waiting for page to settle...")
time.sleep(3) 

# ---------- STEP 2: SMART PAGE DETECTION ----------
print(f"🔍 Current URL: {driver.current_url}")

# Give it one refresh if we aren't where we expect to be
if "publisher.html" not in driver.current_url:
    print("🔄 Refreshing to ensure state...")
    driver.refresh()
    time.sleep(2)

if "publisher.html" in driver.current_url:
    print("✅ Already on Publishers page via Cookies.")

else:
    try:
        print("🔍 Checking if Publishers link is visible...")
        # Try a more generic XPath that finds the link even if text is weird
        publisher_xpath = "//a[contains(@href, 'publisher.html')]"
        wait_short = WebDriverWait(driver, 15) # Increased to 15
        
        link = wait_short.until(EC.element_to_be_clickable((By.XPATH, publisher_xpath)))
        link.click()
        print("✅ Clicked Publishers from sidebar.")

    except Exception:
        print("📍 Sidebar link not found. Starting Deep Navigation...")

        # FIX: The profile head might be nested. Let's use a simpler selector.
        try:
            profile_xpath = "//div[contains(@class,'pop-over__head')] | //div[contains(@class,'noselect')]"
            dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, profile_xpath)))
            driver.execute_script("arguments[0].click();", dropdown) # JS click is safer here
            print("✅ Opened Profile Dropdown")
            time.sleep(1.5)

            # Click Super Admin - Use a more flexible text match
            super_admin_xpath = "//a[normalize-space()='Super Admin' or contains(text(),'Super Admin')]"
            sa_btn = wait.until(EC.element_to_be_clickable((By.XPATH, super_admin_xpath)))
            sa_btn.click()
            print("✅ Clicked Super Admin")

            # Final check for Publisher link after landing in Super Admin
            final_publisher_xpath = "//a[contains(@href,'publisher.html')]"
            wait.until(EC.element_to_be_clickable((By.XPATH, final_publisher_xpath))).click()
        except Exception as e:
            print(f"❌ Deep Navigation failed: {e}")
            driver.save_screenshot("nav_failure.png")
            # If everything fails, try going to the URL directly as a last resort
            print("🚀 Attempting direct URL navigation...")
            PUBLISHER_URL = REGION_CONFIG[REGION]["publisher_url"]
            driver.get(PUBLISHER_URL)

print("🎯 SUCCESS: You are now on the Publishers page.")

def search_by_license(driver, wait, license_code):
    license_input = wait.until(
        EC.presence_of_element_located((By.NAME, "licenseCode"))
    )
    license_input.clear()
    license_input.send_keys(license_code)

    search_btn = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//button[normalize-space()='Search']")
        )
    )
    search_btn.click()
    print("✅ Succesfully entered the LC and clicked search button")

def detect_license_status(driver, license_code):

    # LC exists
    if driver.find_elements(
        By.XPATH,
        f"//tr[contains(., '{license_code}')]"
    ):
        return "Exist"

    # churned customer
    if driver.find_elements(
        By.XPATH,
        "//*[contains(normalize-space(.), 'Request a Demo')]"
    ):
        return "Service stopped"

    # region mismatch
    if driver.find_elements(
        By.XPATH,
        "//h2[contains(text(),'Oops')]"
    ):
        return "Region mismatch"

    return "Unknown"

# ===============================
# MAIN
# ===============================

sheet = init_sheet()

print("✅ Connected to Google Sheets")

license_codes = [c for c in sheet.col_values(1)[1:] if c.strip()]

column_index = REGION_CONFIG[REGION]["column"]

for row_index, code in enumerate(license_codes, start=2):

    print(f"\n▶ Processing {code}")

    try:

        search_by_license(driver, wait, code)

        time.sleep(2)

        result = detect_license_status(driver, code)

        print("Result:", result)

        sheet.update_cell(row_index, column_index, result)

    except Exception as e:

        print("Error:", e)

        sheet.update_cell(row_index, column_index, "ERROR")

    finally:
        # 🔥 RESET PAGE FOR NEXT LC
        driver.get(REGION_CONFIG[REGION]["publisher_url"])
        time.sleep(2)