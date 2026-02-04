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
        "sheet_name": "Overview India",
        "publisher_url": "https://dashboard.in.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "in~~13410618d","in~~15ba205db","in~aa131896","in~58adcd07","in~11b564357","in~~10a5cba77","in~14507c76b","in~58adcc11","in~~10a5cbac3","in~~134106208","in~~2024c233","in~~2024c207","in~~47b66709","in~82617226","in~14507c71d","in~~991991d0","in~~10a5cbb34","in~76aa298","in~311c46c9","in~~2024c2aa","in~~71680b3c","in~76aa1d7","in~~134106266","in~~15ba20741","in~aa131676","in~aa131652","in~~991992c7","in~~c2ab3781","in~~1341062c1","in~~71680c30","in~~47b6677d","in~11b56417b","in~~71680b90"
        ]
    },
    "GLOBAL": {
        "base_url": "https://dashboard.webengage.com",
        "sheet_name": "Overview Global",
        "publisher_url": "https://dashboard.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "~134105a52","~716800b0","~15ba2020b","~15ba20105","~2024bb2d","~134105a04","76aa78b","~1341056a0","~71680655","~134105965","~9919868d","~99198a29","58add69d","14507cc74","aa132225","311c4c14","311c4c11","76aa762","11b564830","76aab88","~71680627","11b5646ca","~c2ab3083","~15ba2019a","~99198226","82617775","~c2ab3042","311c4bc4","~10a5cb20c","14507cc0a","~15ba1db98","~47b661c8","~2024bb90","82617779","~47b660d4","~134105353","~47b66257","~15ba20080","58add667","~1341056c2","14507d197","~10a5cb283","~15ba1dd60","~9919879c","d3a4b631","~c2ab2d61","~47b65736","~2024b742","14507d028","d3a4a4aa","76abad2","~134104786","76aba82","~15ba1dd60","76ab97b","145080010","d3a4b4a3","~2024a938","82618947.0","~15ba1cdc0","76ab953","~7167d478","~13410487b","76ab940"
        ]
    },
    "KSA": {
        "base_url": "https://dashboard.ksa.webengage.com",
        "sheet_name": "Overview KSA",
        "publisher_url": "https://dashboard.ksa.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "ksa~~15ba20526","ksa~82617412","ksa~~716809b4","ksa~14507c890","ksa~~2024c070","ksa~76aa41c","ksa~~716809bd","ksa~11b564409","ksa~~47b6652a","ksa~aa1318a1","ksa~~47b6652c","ksa~58adcd4c","ksa~11b564406","ksa~~15ba2051c","ksa~aa131897","ksa~~134106071","ksa~82617408","ksa~58adcd55","ksa~~15ba20523","ksa~~10a5cb9bd","ksa~d3a49d49","ksa~11b5643db","ksa~d3a49d4a","ksa~d3a49d44","ksa~58adcd54","ksa~aa13189b","ksa~311c489a","ksa~82617404","ksa~~2024c08a","ksa~~134106076","ksa~14507c891","ksa~~716809c9","ksa~d3a49d46","ksa~~47b66537","ksa~~134106074","ksa~~2024c07d","ksa~~2024c085","ksa~82617402","ksa~~13410607a","ksa~11b564403","ksa~~716809ba","ksa~~10a5cb9c4","ksa~~99199078","ksa~~15ba20518","ksa~~134106069","ksa~311c4892","ksa~~99199083","ksa~aa1318a0","ksa~~13410606b","ksa~11b5643d5","ksa~~134106080","ksa~58adcd47","ksa~58adcd44","ksa~11b5643d3","ksa~826173db","ksa~d3a49d41","ksa~~134106084","ksa~~99199073","ksa~~99199087","ksa~aa131893","ksa~~15ba20531","ksa~76aa3da","ksa~~2024c091","ksa~826173dc","ksa~82617401","ksa~aa131890","ksa~14507c89c","ksa~~10a5cb9d1","ksa~~2024c08d","ksa~~47b66522"
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
        "1sathL7caATX3PnV2urKhp8UpCLOcwdTIxzkJkA4Kar4"
    ).worksheet(sheet_name)

    return sheet

try:
    sheet = init_google_sheet()
    print("✅ Connected to Google Sheets")
except Exception as e:
    print(f"❌ Failed to connect to Google Sheets: {e}")
    exit()

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

# ========== Till here no change things will be same ========== #

def go_to_overview(driver, wait, account_id):
    base_url = REGION_CONFIG[REGION]["base_url"]
    url = f"{base_url}/accounts/{account_id}/engagement/overview/all"

    driver.get(url)

# If no Data available -> check
def is_overview_no_data(driver):
    return "Channel statistics not available" in driver.page_source

def extract_overview_table(driver, license_code):
    print("📥 Extracting Overview table...")

    # ✅ ONLY true no-data case
    if "Channel statistics not available" in driver.page_source:
        print("ℹ️ Channel statistics not available → NO_DATA")
        return [[
            license_code,
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            "NO_DATA"
        ]]

    rows = driver.find_elements(
        By.XPATH,
        "//tbody/tr[contains(@class,'table__row')]"
    )

    # 🚨 SAFETY: page loaded but table missing (rare)
    if not rows:
        print("⚠️ No rows found, treating as NO_DATA")
        return [[
            license_code,
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            "NO_DATA"
        ]]

    rows_data = []

    for row in rows:
        cells = row.find_elements(By.XPATH, "./td")

        # Overview table has 6 metrics + channel
        if len(cells) < 6:
            continue

        rows_data.append([
            license_code,
            cells[0].text.strip() or "-",  # Channel
            cells[1].text.strip() or "-",
            cells[2].text.strip() or "-",
            cells[3].text.strip() or "-",
            cells[4].text.strip() or "-",
            cells[5].text.strip() or "-"
        ])

    print(f"✅ Extracted {len(rows_data)} rows")
    return rows_data

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
            account_id = code

            go_to_overview(driver, wait, account_id)

            overview_rows = extract_overview_table(driver, code)
            append_to_sheet(sheet, overview_rows)

        except Exception as e:
            log_error_to_sheet(
                sheet,
                code,
                stage="OVERVIEW",
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
