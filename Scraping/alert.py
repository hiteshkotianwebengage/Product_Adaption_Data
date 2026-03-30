'''
    Things to change when we switch between India, GLobal, KSA
    1) Below Client inside init_google_sheet function change the sheet name
    2) In the step 0 below the Driver varibale change between driver.get 
    3) In the step 2 while direct url for publisher change between driver.get
    4) Inside the go_to_Alert_events function we switch between the url
    5) In the step 5 we have urls to directly go to publisher after data is fetching
    6) This is pretty main we have to change the LC based on the global, india, ksa
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
REGION = "INDIA"   # options: "INDIA", "GLOBAL", "KSA"

REGION_CONFIG = {
    "INDIA": {
        "base_url": "https://p1o82kY:kow3jJs9@dashboard.in.webengage.com",
        "sheet_name": "Alert Event India",
        "publisher_url": "https://p1o82kY:kow3jJs9@dashboard.in.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "in~~15ba20672","in~~c2ab3721","in~82617205","in~82617246","in~58adcb94","in~82617203","in~~134106115","in~d3a49b80","in~~99199283","in~~71680c0c"
        ]
    },
    "GLOBAL": {
        "base_url": "https://p1o82kY:kow3jJs9@dashboard.webengage.com",
        "sheet_name": "Alert Event Global",
        "publisher_url": "https://p1o82kY:kow3jJs9@dashboard.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "d3a4ab38","76aac96","d3a4a457","~134105a60","58add423","14507cd00","d3a4ac1c","~1341059b6","~10a5cab6a","~10a5cac40","~7167db84","14507cc77","58add283","9,91,98,968","~15ba1da68","58add2d9","~134105a52","76aa813","aa13266b","826174d0","82617c25","~47b6607a","8,26,17,822","~10a5cb63c","~c2ab3108","~134105a8c","~15ba20105","8,26,17,757","aa131c59","311c4b69","~2024bb2d","14507cd4d","~13410604b","~2024b99a","~9919839a","aa131c84","76aa858","13,41,05,251","~99198a20","~1341056a0","~2024b6c6","~15ba20116","76aa833","76aa7a3","14507cba6","14507d14d","7,16,80,655","~2024bb10","~10a5cb557","~c2ab313b","13,41,05,965","~1341059c5","58adc5c7","~7168057d","~47b65848","d3a4a32d","311c4c4b","~47b66045","~15ba201a6","76aa124","~9919868d","76a9c30","~311c4b76","58adca91","76aa76b","82617b34","~2024bada","d3a4a403","8261827a","7,16,80,577","~99198a29","~2024bad5","~134105b84","~c2ab3033","~15ba1ddc2","8,26,18,089","58add346","~1341061bb","d3a4a69c","58add2da","~716805d8","~15ba20214","~10a5cb6b0","~1341056bb","aa132703","~aa1321c5","~c2ab2c0c","~9919871c","~47b66614","~991981d3","14507cc74","8261786b","11b564b69","aa132225","311c4c14","311c4c11","76aa762","311c4bbb","14507cba8","58add7aa","11b564830","76aac69","76aab88","7,16,80,627","~15ba20042","~oldetmoney","11b564720","~c2ab275a","~15ba1d70a","~c2ab2ba2","8261812c","~47b6665c","~1341059cb","~991989d1","~old2024c085","~oldmagma1","~oldmagma2","~oldmagma3","11b56527d","~c2ab2c08","~11b5646b8","~d3a4a286","76aa7c6","76aa85d","~47b65b94","~7168053d","~c2ab3083","58addc40","76aa844","~d3a49c4c-old","76aa868","~15ba2019a","~7168069d","~oldrangde","~10a5cb533","~10a5cb636","~c2ab26b8","9,91,98,226","~99198a14"
        ]
    },
    "KSA": {
        "base_url": "https://p1o82kY:kow3jJs9@dashboard.ksa.webengage.com",
        "sheet_name": "Alert Event KSA",
        "publisher_url": "https://p1o82kY:kow3jJs9@dashboard.ksa.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "ksa~~15ba20526","ksa~82617412","ksa~~716809b4","ksa~14507c890","ksa~~2024c070","ksa~76aa41c","ksa~~716809bd","ksa~11b564409","ksa~~47b6652a","ksa~aa1318a1","ksa~~47b6652c","ksa~58adcd4c","ksa~11b564406","ksa~~15ba2051c","ksa~aa131897","ksa~~134106071","ksa~82617408","ksa~58adcd55","ksa~~15ba20523","ksa~~10a5cb9bd","ksa~d3a49d49","ksa~11b5643db","ksa~d3a49d4a","ksa~d3a49d44","ksa~58adcd54","ksa~aa13189b","ksa~311c489a","ksa~82617404","ksa~~2024c08a","ksa~~134106076","ksa~14507c891","ksa~~716809c9","ksa~d3a49d46","ksa~~47b66537","ksa~~134106074","ksa~~2024c07d","ksa~~2024c085","ksa~82617402","ksa~~13410607a","ksa~11b564403","ksa~~716809ba","ksa~~10a5cb9c4","ksa~~99199078","ksa~~15ba20518","ksa~~134106069","ksa~311c4892","ksa~~99199083","ksa~aa1318a0","ksa~~13410606b","ksa~11b5643d5","ksa~~134106080","ksa~58adcd47","ksa~58adcd44","ksa~11b5643d3","ksa~826173db","ksa~d3a49d41","ksa~~134106084","ksa~~99199073","ksa~~99199087","ksa~aa131893","ksa~~15ba20531","ksa~76aa3da","ksa~~2024c091","ksa~826173dc","ksa~82617401","ksa~aa131890","ksa~14507c89c","ksa~~10a5cb9d1","ksa~~2024c08d","ksa~~47b66522","ksa~~c2ab3537","ksa~311c4898","ksa~~716809d4","ksa~76aa402","ksa~~99199088","ksa~~134106087","ksa~~9919906d","ksa~~2024c083","ksa~~2024c084","ksa~~c2ab3537","ksa~311c4898","ksa~~2024c07c","ksa~~716809d4","ksa~76aa402","ksa~~99199088","ksa~~134106087"
        ]
    }
}

def init_google_sheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(
        "/Users/admin/Desktop/Code Directory/Product_Adaption_Data/Credential File/mycred-googlesheet.json",
        scopes=scopes
    )
    client = gspread.authorize(creds)

    sheet_name = REGION_CONFIG[REGION]["sheet_name"]
    sheet = client.open_by_key(
        "1D0-O3OX3TOmZRMRZmyYoubRxBtekAB05b50lcaIOweM"
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

def should_skip_account(driver):

    # ❌ Case 1: Oops page
    if driver.find_elements(By.XPATH, "//h2[contains(text(),'Oops')]"):
        print("⛔ Oops page → Skipping")
        return True, "Region mismatch"

    # ❌ Case 2: Request Demo (ROBUST FIX)
    if (
        driver.find_elements(By.XPATH, "//*[contains(normalize-space(.), 'Request a Demo')]")
        or "Request a Demo" in driver.page_source
    ):
        print("⛔ Request Demo → Skipping")
        return True, "Service stopped"

    return False, None

def check_if_result_exists(driver, license_code, timeout=5):
    """
Returns False if LC does not appear in publisher table
    (wrong region / invalid LC)
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

def open_data_platform(driver, wait):
    print("⏳ Opening Data Platform sidebar (hybrid React-safe)...")

    li = wait.until(
        EC.presence_of_element_located((By.ID, "nav-data-platform"))
    )

    head = li.find_element(By.XPATH, ".//div[contains(@class,'menu__group__head')]")

    # Always scroll first
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", head)
    time.sleep(0.2)

    # 1️⃣ Try normal JS click (React listener)
    driver.execute_script("arguments[0].click();", head)
    time.sleep(0.5)

    # 2️⃣ Check if React actually opened it
    class_now = li.get_attribute("class")
    if "menu__group--is-active" in class_now:
        print("✅ Data Platform opened via click")
        return

    print("⚠️ Click did not open menu, forcing state...")

    # 3️⃣ Force class (DOM state)
    driver.execute_script("""
        arguments[0].classList.add('menu__group--is-active');
    """, li)

    # 4️⃣ Dispatch synthetic click event (React fallback)
    driver.execute_script("""
        arguments[0].dispatchEvent(
            new MouseEvent('click', { bubbles: true })
        );
    """, head)

    time.sleep(0.5)

    # Final confirmation
    if "menu__group--is-active" in li.get_attribute("class"):
        print("✅ Data Platform opened via hybrid fallback")
    else:
        raise Exception("❌ Failed to open Data Platform sidebar")
    
def go_to_alerts(driver, wait, account_id):
    base_url = REGION_CONFIG[REGION]["base_url"]
    url = f"{base_url}/accounts/{account_id}/alerts"

    driver.get(url)
    time.sleep(2)  # critical for modal detection

def is_alerts_locked(driver):
    # modal visible
    if "Locked feature" in driver.page_source:
        return True

    # redirected silently (very common)
    if not driver.current_url.endswith("/alerts"):
        return True

    return False

def click_data_management(wait):
    print("⏳ Clicking Data Management...")

    data_management_xpath = (
        "//a[contains(@href,'/data-management/system/attributes') and .//span[text()='Data Management']]"
    )

    wait.until(
        EC.element_to_be_clickable((By.XPATH, data_management_xpath))
    ).click()

    print("✅ Clicked Data Management")

def click_Alert_events(wait):
    print("⏳ Clicking Alert Events tab...")

    Alert_events_xpath = (
        "//a[contains(@href,'/data-management/events/attributes') and normalize-space()='Alert Events']"
    )

    wait.until(
        EC.element_to_be_clickable((By.XPATH, Alert_events_xpath))
    ).click()

    print("✅ Alert Events opened")

def wait_for_alerts_ready(driver, timeout=8):
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: (
                "Locked feature" in d.page_source
                or "/alerts" in d.current_url
            )
        )
        return True
    except:
        return False

def extract_Alert_events_page(driver, license_code):
    rows_data = []
    # Wait for at least one row to be present before scraping
    has_data = wait_for_table_or_empty(driver)

    if not has_data:
        return []

    rows = driver.find_elements(By.XPATH, "//tr[contains(@class,'table__row')]")

    for row in rows:
        try:
            cells = row.find_elements(By.XPATH, "./td")
            if len(cells) < 8: continue 

            # 1. Event Name 
            event_name = cells[0].find_element(By.XPATH, ".//span[contains(@class,'text-ellipsis')]").get_attribute("title").strip()

            # 2. Is PII / Data Type / Personalization
            is_pii = cells[1].text.strip() or "NO"
            data_type = cells[2].text.strip() or "NULL"
            personalization = cells[3].text.strip() or "Disabled"

            # 3. Status Columns (Website, Android, iOS, Others)
            def get_status(cell):
                try:
                    return cell.find_element(By.CLASS_NAME, "status-label").text.strip()
                except:
                    return "NULL"

            website = get_status(cells[4])
            android = get_status(cells[5])
            ios = get_status(cells[6])
            others = get_status(cells[7])

            rows_data.append([
                license_code, event_name, is_pii, data_type, 
                personalization, website, android, ios, others
            ])
        except Exception as e:
            print(f"⚠️ Skipping a row due to error: {e}")
            continue

    return rows_data

def extract_alerts(driver, license_code):
    print("📥 Extracting Alerts...")

    # hard wait to allow React to paint content
    time.sleep(2)

    rows = driver.find_elements(
        By.XPATH, "//tbody/tr[contains(@class,'table__row')]"
    )

    if not rows:
        return [[
            license_code,
            "NO_DATA",
            "",
            "",
            "",
            "",
            "",
            ""
        ]]

    rows_data = []

    for row in rows:
        try:
            cells = row.find_elements(By.XPATH, "./td")
            if len(cells) < 6:
                continue

            # 1️⃣ Alert Name (use title – most stable)
            name_container = cells[0].find_element(
                By.XPATH, ".//div[@title]"
            )
            alert_name = name_container.get_attribute("title").strip()

            # Template flag
            template = "Template" if name_container.find_elements(
                By.XPATH, ".//span[contains(@class,'template-banner')]"
            ) else ""

            # 2️⃣ Frequency
            frequency = cells[1].text.strip() or "-"

            # 3️⃣ Status
            status = cells[2].text.strip() or "-"

            # 4️⃣ Subscribers (handles '-', emails, +N)
            subscribers = cells[3].text.strip() or "-"

            # 5️⃣ Last Check
            last_check = cells[4].text.strip() or "-"

            # 6️⃣ Created On
            created_on = cells[5].text.strip() or "-"

            rows_data.append([
                license_code,
                alert_name,
                template,
                frequency,
                status,
                subscribers,
                last_check,
                created_on
            ])

        except Exception as e:
            print(f"⚠️ Skipping alert row due to parse error: {e}")
            continue

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

def wait_for_table_or_empty(driver, timeout=6):
    """
    Waits for:
    - table rows
    - empty state
    Returns True if rows exist, False otherwise
    """
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: (
                d.find_elements(By.XPATH, "//tr[contains(@class,'table__row')]")
                or "No data" in d.page_source
            )
        )
    except:
        return False

    rows = driver.find_elements(By.XPATH, "//tr[contains(@class,'table__row')]")
    return len(rows) > 0

LICENSE_CODES = REGION_CONFIG[REGION]["license_codes"]


for code in LICENSE_CODES:
    print(f"\n▶ Processing {code}")
    try:
        # Step A: Search and land on result
        search_by_license(driver, wait, code)
        time.sleep(2)

        # 🔥 FAST SKIP CHECK
        skip, reason = should_skip_account(driver)

        if skip:
            log_error_to_sheet(
                sheet,
                code,
                stage="SKIPPED",
                error_reason=reason
            )

            driver.get(REGION_CONFIG[REGION]["publisher_url"])
            wait.until(EC.presence_of_element_located((By.NAME, "licenseCode")))
            continue

        # 🚫 REGION MISMATCH GUARD (same as Users / Dashboards)
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

        # Step D: Extract Alerts
        try:
            account_id = code
            
            go_to_alerts(driver, wait, account_id)
            # wait for either alerts OR locked modal
            wait_for_alerts_ready(driver)

            if is_alerts_locked(driver):
                log_error_to_sheet(
                    sheet,
                    code,
                    stage="ALERTS_LOCKED",
                    error_reason="Alerts feature not enabled for this account"
                )
                continue

            alerts_rows = extract_alerts(driver, code)
            append_to_sheet(sheet, alerts_rows)

        except Exception as e:
            log_error_to_sheet(
                sheet,
                code,
                stage="ALERTS",
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
