'''
    Things to change when we switch between India, GLobal, KSA
    1) Below Client inside init_google_sheet function change the sheet name
    2) In the step 0 below the Driver varibale change between driver.get 
    3) In the step 2 while direct url for publisher change between driver.get
    4) In the step E we have urls to directly go to publisher after data is fetching
    5) This is pretty main we have to change the LC based on the global, india, ksa
'''

'''
    For above steps we changed it now things are simple
    1) Change REGION
    2) If required change the license_codes inside the REGION_CONFIG
    3) Check if the json file of credential file path is right - This is for me to remember if in future i change the system
    4) Check if the sheet name is right we associated with the region
'''
import os
import time
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import gspread
from google.oauth2.service_account import Credentials
from selenium.webdriver.common.action_chains import ActionChains

# =======================
# REGION CONFIG
# =======================
REGION = "GLOBAL"   # options: "INDIA", "GLOBAL", "KSA"

REGION_CONFIG = {
    "INDIA": {
        "base_url": "https://dashboard.in.webengage.com",
        "sheet_name": "User India",
        "publisher_url": "https://dashboard.in.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "in~~134106216"
        ]
    },
    "GLOBAL": {
        "base_url": "https://dashboard.webengage.com",
        "sheet_name": "User Global",
        "publisher_url": "https://dashboard.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "~c2ab3108","~c2ab3083","~15ba2063d","~47b65ca0","~15ba20218","aa133168","d3a4a339","~10a5cbb7a","~2024a887","~7167d473","11b565ab8","aa13306d","~7167d4b4"
        ]
    },
    "KSA": {
        "base_url": "https://dashboard.ksa.webengage.com",
        "sheet_name": "User KSA",
        "publisher_url": "https://dashboard.ksa.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "ksa~82617408"
        ]
    }
}

def init_google_sheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(
        "/Users/admin/Desktop/Python Script/agreement_file_pasting/mycred-googlesheet.json",
        scopes=scopes
    )
    client = gspread.authorize(creds)

    sheet_name = REGION_CONFIG[REGION]["sheet_name"]
    sheet = client.open_by_key(
        "1aw7c-PMc0wv5_QOcQdijEQFxcf9rzFN4vAo9HamVlDw"
    ).worksheet(sheet_name)

    return sheet

try:
    sheet = init_google_sheet()
    print("✅ Connected to Google Sheets")
except Exception as e:
    print(f"❌ Failed to connect to Google Sheets: {e}")
    exit()


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
    
def check_if_result_exists(driver, license_code, timeout=5):
    """
    Returns False if LC does not appear in publisher table (wrong region / invalid)
    """
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: (
                d.find_elements(By.XPATH, f"//tr[contains(., '{license_code}')]")
                or "No data" in d.page_source
            )
        )
    except:
        return False

    rows = driver.find_elements(By.XPATH, f"//tr[contains(., '{license_code}')]")
    return len(rows) > 0

def open_actions_dropdown(driver, wait, license_code):
    print(f"⏳ Opening Actions dropdown for {license_code}...")
    
    # Improved XPath: Finds the specific TR containing the LC, then the toggle button inside it
    toggle_xpath = f"//tr[contains(., '{license_code}')]//button[contains(@class, 'dropdown-toggle')]"
    
    try:
        dropdown_btn = wait.until(EC.element_to_be_clickable((By.XPATH, toggle_xpath)))
        # Force scroll and use JS click to bypass any overlapping UI elements
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", dropdown_btn)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", dropdown_btn)
        print("✅ Successfully clicked the specific dropdown toggle")
    except Exception as e:
        print(f"⚠️ Could not click dropdown for {license_code}: {e}")
        # If it fails, we try a refresh as a fallback
        driver.refresh()
        time.sleep(2)

def click_request_access(wait):
    request_access = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(@class,'requestAccess')]")
        )
    )
    request_access.click()
    print("✅ Succesfully opened the access dialog box")

def handle_request_modal(wait, driver):
    print("⏳ Handling request modal...")
    
    # Give the modal a second to fully animate and render
    time.sleep(2)

    # 1. SMART IFRAME DETECTION
    # We only switch if an iframe actually exists AND is visible
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    switched = False
    for frame in iframes:
        if frame.is_displayed():
            driver.switch_to.frame(frame)
            switched = True
            print("↔️ Switched to active iframe")
            break

    try:
        # 2. WAIT FOR DROPDOWN
        # We use a very broad XPath to find the select by its ID or Name
        dropdown_xpath = "//select[@id='roleIdField' or @name='roleEId']"
        role_dropdown = wait.until(
            EC.presence_of_element_located((By.XPATH, dropdown_xpath))
        )
        
        # 3. FORCE SELECTION VIA JAVASCRIPT
        # This is the most reliable way to select 'Viewer' regardless of UI quirks
        print("🚀 # React-safe role selection (required only during first-time access request)")
        Select(role_dropdown).select_by_visible_text("Viewer")
        
        # 4. FILL COMMENT
        comment_box = wait.until(EC.presence_of_element_located((By.ID, "commentText")))
        comment_box.clear()
        comment_box.send_keys("access request")

        # 5. CLICK REQUEST
        request_btn = driver.find_element(By.XPATH, "//button[contains(.,'Request')]")
        driver.execute_script("arguments[0].click();", request_btn)
        print("✅ Successfully submitted access request")

    except Exception as e:
        print(f"❌ Error inside modal: {e}")
        driver.save_screenshot("error_modal_view.png")
    
    finally:
        if switched:
            driver.switch_to.default_content()
            print("↔️ Switched back to main content")

def close_modal_if_exists(driver):
    """Closes the modal only if it is still visible."""
    try:
        # Check for the close button with a very short timeout
        close_btn = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.ID, "cboxClose"))
        )
        close_btn.click()
        print("✅ Manually closed the dialog box")
        time.sleep(1)
    except:
        print("ℹ️ Modal already closed or close button not found (which is fine)")

def click_edit(driver, wait, license_code):
    print(f"⏳ Clicking Edit for {license_code}...")
    
    # Find the Edit link inside the row that matches the License Code
    edit_xpath = f"//tr[contains(., '{license_code}')]//a[contains(@href, '/publisher/edit')]"
    
    edit_btn = wait.until(
        EC.element_to_be_clickable((By.XPATH, edit_xpath))
    )
    
    edit_btn.click()
    print("✅ Clicked the specific edit button")


def click_users(wait):
    print("⏳ Clicking Users...")

    users_xpath = (
        "//a[contains(@href,'/users/overview') and .//span[text()='Users']]"
    )

    wait.until(
        EC.element_to_be_clickable((By.XPATH, users_xpath))
    ).click()

    print("✅ Users opened")

def switch_users_delta(wait, driver, mode):
    """ mode = 'WoW' or 'MoM' """
    print(f"🔁 Switching Users delta to {mode}")

    dropdown_head = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//th//div[contains(@class,'pop-over__head')]")
        )
    )

    driver.execute_script("arguments[0].click();", dropdown_head)

    option = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, f"//span[normalize-space()='{mode}']")
        )
    )

    driver.execute_script("arguments[0].click();", option)

    # 🔹 CRITICAL: wait for DOM refresh
    time.sleep(1.2)
    wait_for_users_table_or_empty(driver)

def extract_users_table(driver, license_code, delta_type):
    rows_data = []

    has_data = wait_for_users_table_or_empty(driver)

    if not has_data:
        print(f"⚠️ No Users data for {delta_type}")
        return [[
            license_code,
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            delta_type,
            "NO_DATA"
        ]]

    rows = driver.find_elements(
        By.XPATH,
        "//tbody/tr[contains(@class,'table__row')]"
    )

    for row in rows:
        try:
            cells = row.find_elements(By.XPATH, "./td")
            if len(cells) < 4:
                continue

            channel = cells[0].text.strip() or "UNKNOWN"
            reach_pct = cells[1].text.strip() or "0%"
            reach_count = cells[2].text.strip() or "0"

            # Change %
            change = cells[3].text.strip() or "0.00%"

            rows_data.append([
                license_code,
                channel,
                reach_pct,
                reach_count,
                delta_type,
                change
            ])

        except:
            continue

    return rows_data

def extract_users_reachability(driver, wait, license_code):
    print("📥 Extracting Users Reachability (WoW + MoM)...")

    all_rows = []

    for mode in ["WoW", "MoM"]:
        switch_users_delta(wait, driver, mode)
        rows = extract_users_table(driver, license_code, mode)
        all_rows.extend(rows)

    return all_rows

def wait_for_users_table_or_empty(driver, timeout=6):
    """
    Waits for:
    - users rows
    - footer row (Overall)
    - empty state
    Returns True if data rows exist, False otherwise
    """
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: (
                d.find_elements(By.XPATH, "//tbody/tr[contains(@class,'table__row')]")
                or d.find_elements(By.XPATH, "//tfoot/tr")
                or "No data" in d.page_source
            )
        )
    except:
        return False

    rows = driver.find_elements(By.XPATH, "//tbody/tr[contains(@class,'table__row')]")
    return len(rows) > 0


def append_to_sheet(sheet, rows):
    if rows:
        sheet.append_rows(rows, value_input_option="USER_ENTERED")
        print("📤 Data pushed to Google Sheets")
    else:
        print("⚠️ No data to push")

def log_error_to_sheet(sheet, license_code, stage, error_reason):
    print(f"📝 Logging error for {license_code} at stage: {stage}")

    row = [
        license_code,
        "ERROR",
        stage,
        error_reason[:300],  # keep it readable
        time.strftime("%Y-%m-%d %H:%M:%S")
    ]

    sheet.append_row(row, value_input_option="USER_ENTERED")

LICENSE_CODES = REGION_CONFIG[REGION]["license_codes"]

for code in LICENSE_CODES:
    print(f"\n▶ Processing {code}")
    try:
        # Step A: Search and land on result
        search_by_license(driver, wait, code)

        # 🔥 ERROR HANDLING: LC belongs to another region
        if not check_if_result_exists(driver, code):
            log_error_to_sheet(
                sheet,
                code,
                stage="REGION_MISMATCH",
                error_reason="License code not found in this region"
            )
            continue
        
        # Step B: Try to get access (only if needed)
        try:
            open_actions_dropdown(driver, wait, code) # Added driver and code
            
            request_btn_xpath = f"//tr[contains(., '{code}')]//a[contains(@class,'requestAccess')]"
            if len(driver.find_elements(By.XPATH, request_btn_xpath)) > 0:
                click_request_access(wait)
                handle_request_modal(wait, driver)
                time.sleep(2)
                close_modal_if_exists(driver)
            else:
                print("ℹ️ Access already available. Moving to Edit.")
                # If dropdown is open but we don't need it, refresh or Esc
                ActionChains(driver).send_keys(Keys.ESCAPE).perform() 
        except Exception as e:
            print(f"⚠️ Access step skipped: {e}")

        # --- Updated Step C: Pass 'code' to the edit function ---
        main_window = driver.current_window_handle
        click_edit(driver, wait, code)

        # SWITCH TO NEW TAB
        for window_handle in driver.window_handles:
            if window_handle != main_window:
                driver.switch_to.window(window_handle)
                break
        print(f"↔️ Switched to Edit tab for {code}")

        # Step D: Extract Data
        try:
            click_users(wait)

            users_rows = extract_users_reachability(driver, wait, code)
            append_to_sheet(sheet, users_rows)

            time.sleep(1)

            users_rows = extract_users_reachability(driver, wait, code)

            append_to_sheet(sheet, users_rows)

        except Exception as e:
            log_error_to_sheet(
                sheet,
                code,
                stage="USERS_REACHABILITY",
                error_reason=str(e)
            )
    
    finally:
        # Step E: Cleanup for next iteration
        # Close extra tabs and go back to the list
        if len(driver.window_handles) > 1:
            driver.close() # Closes current (Edit) tab
            driver.switch_to.window(main_window)
        
        PUBLISHER_URL = REGION_CONFIG[REGION]["publisher_url"]
        driver.get(PUBLISHER_URL)

        time.sleep(2)
