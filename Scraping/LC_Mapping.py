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
import argparse

parser = argparse.ArgumentParser()

parser.add_argument(
    "--region",
    required=True,
    choices=["INDIA", "GLOBAL", "KSA"]
)

args = parser.parse_args()

REGION = args.region.upper()

# ===============================
# CONFIG
# ===============================

SHEET_ID = "1o5QRUGQYptkwe1NdsZcfgD44fQCjjkfY2D_DSftSYa4"
SHEET_NAME = "LC_Check"

# ===============================
# Region
# ===============================

REGION_CONFIG = {
    "INDIA": {
        "base_url": "https://p1o82kY:kow3jJs9@dashboard.in.webengage.com",
        "publisher_url": "https://p1o82kY:kow3jJs9@dashboard.in.webengage.com/admin/publisher.html?action=list",
        "column": 3
    },
    "GLOBAL": {
        "base_url": "https://p1o82kY:kow3jJs9@dashboard.webengage.com",
        "publisher_url": "https://p1o82kY:kow3jJs9@dashboard.webengage.com/admin/publisher.html?action=list",
        "column": 4
    },
    "KSA": {
        "base_url": "https://p1o82kY:kow3jJs9@dashboard.ksa.webengage.com",
        "publisher_url": "https://p1o82kY:kow3jJs9@dashboard.ksa.webengage.com/admin/publisher.html?action=list",
        "column": 5
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

    # rows = sheet.get_all_values()
    # This below line will help us start from 2nd column so everything remains same
    all_codes = sheet.col_values(2)

    license_codes = [code for code in all_codes[1:] if code.strip() != ""]

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
# print(f"🔍 Current URL: {driver.current_url}")

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
        publisher_xpath = "//a[contains(@href, 'publisher.html')]"
        wait_short = WebDriverWait(driver, 15)
        link = wait_short.until(EC.element_to_be_clickable((By.XPATH, publisher_xpath)))
        link.click()
        print("✅ Clicked Publishers from sidebar.")
    except Exception:
        print("📍 Sidebar link not found. Executing fixed Dropdown Navigation...")
        try:
            # FIX: Targeting the profile dropdown header accurately using the class from your HTML
            dropdown_xpath = "//div[contains(@class, 'pop-over__head')]"
            dropdown = wait.until(EC.presence_of_element_located((By.XPATH, dropdown_xpath)))
            
            # Use JavaScript click to reliably trigger the dropdown
            driver.execute_script("arguments[0].click();", dropdown)
            print("✅ Opened Profile Dropdown via JS")
            time.sleep(1.5)

            # FIX: Targeting the precise Super Admin link element you provided
            super_admin_xpath = "//a[contains(@href, '/admin') and text()='Super Admin']"
            sa_btn = wait.until(EC.presence_of_element_located((By.XPATH, super_admin_xpath)))
            
            # Click Super Admin via JS to bypass overlay blockages
            driver.execute_script("arguments[0].click();", sa_btn)
            print("✅ Clicked Super Admin via JS")
            time.sleep(3)

            # Final check for Publisher link after landing in Super Admin
            final_publisher_xpath = "//a[contains(@href,'publisher.html')]"
            final_pub = wait.until(EC.element_to_be_clickable((By.XPATH, final_publisher_xpath)))
            driver.execute_script("arguments[0].click();", final_pub)
        except Exception as e:
            print(f"❌ Deep Navigation failed: {e}")
            driver.save_screenshot("nav_failure.png")
            print("🚀 Last resort: Direct URL navigation...")
            PUBLISHER_URL = REGION_CONFIG[REGION]["publisher_url"]
            driver.get(PUBLISHER_URL)

print("🎯 SUCCESS: Landing complete on Publishers page.")

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

# 2 is for the 2nd column where the LC are there and we are skipping the header by [1:]
license_codes = [c for c in sheet.col_values(2)[1:] if c.strip()]

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