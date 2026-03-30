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
from datetime import datetime, timedelta

# =======================
# REGION CONFIG
# =======================
REGION = "GLOBAL"   # options: "INDIA", "GLOBAL", "KSA"

REGION_CONFIG = {
    "INDIA": {
        "base_url": "https://p1o82kY:kow3jJs9@dashboard.in.webengage.com",
        "sheet_name": "Overview India",
        "publisher_url": "https://p1o82kY:kow3jJs9@dashboard.in.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "in~~9919912c","in~~99199192","in~~134106156","in~~15ba205db","in~58adcd07","in~~47b66667","in~14507c784","in~76aa392","in~~10a5cba77","in~8261729b","in~14507c76b","in~58adcc4a","in~~47b665d8","in~58adcc70","in~58adcc59","in~~c2ab363b","in~~2024c156","in~~c2ab3517","in~76aa2c9","in~~71680ad0","in~~15ba2065c","in~~1341061ac","in~58adcc11","in~76aa35b","in~~134106132","in~~15ba205d1","in~~1341061b6","in~~10a5cbb1d","in~11b5642a5","in~~47b66670","in~~15ba20633","in~311c4742","in~311c4724","in~~134106208","in~14507c728","in~~15ba20670","in~~15ba206a9","in~~2024c233","in~~47b666d5","in~76aa1d8","in~~99199258","in~~71680b65","in~~10a5cbb42","in~aa1318ab","in~~c2ab3714","in~58adcb8b","in~d3a49b94","in~d3a49bac","in~~10a5cbb66","in~~c2ab36d5","in~~c2ab3671","in~~47b66716","in~82617217","in~d3a49ba1","in~14507c71d","in~76aa298","in~~10a5cba3a","in~~10a5cbba6","in~11b5641db","in~~134106263","in~58adcb79","in~~15ba20672","in~58adcb94","in~~134106273","in~11b5641d0","in~d3a49b43","in~76aa1b3","in~826171c3","in~~134106267","in~~99199277","in~~10a5cbb38","in~~71680bd5","in~~2024c1d7","in~~2024c218","in~~47b6668a","in~~134106213","in~~15ba206a8","in~~10a5cbb61","in~11b564256","in~~9919921b","in~~134106216","in~~71680b69","in~311c46d4","in~311c46d3","in~58adcb85","in~~2024c249","in~311c4665","in~~71680c19","in~~2024c085","in~d3a49b5d","in~~10a5cbc14","in~~134106266","in~11b5641a0","in~~71680bb9","in~~71680c2b","in~58adcb08","in~aa13163a","in~~2024c2b8","in~~1341062bb","in~~10a5cbc2d","in~311c4664","in~14507c641","in~~71680c30","in~~47b6677d","in~~1341062c2","in~~99199081","in~14507c63b","in~11b564246","in~aa131665","in~aa131632","in~311c464b","in~311c4766","in~~c2ab36a7","in~82617256","in~14507c69a","in~aa13177d","in~~71680b7d","in~58adcb46","in~~9919926b","old-in~~99199240","in~aa13168a","in~311c4685","in~~47b66712","in~58adcb45","in~14507c6ad","in~~134106261","in~~134106255","in~~c2ab36cb","in~~99199213"
        ]
    },
    "GLOBAL": {
        "base_url": "https://p1o82kY:kow3jJs9@dashboard.webengage.com",
        "sheet_name": "Overview Global",
        "publisher_url": "https://p1o82kY:kow3jJs9@dashboard.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "~47b6574d","~47b66864","8261829c","58add307","~2024b5d8","d3a4a457","311c5625","311c5642","~15ba20153","~7167db84","9,91,98,968","~15ba1da68","~134105a52","~716800b0","826174d0","~47b65b6c","~47b6607a","~15ba2020b","~15ba1d691","~c2ab3108","76ab0a5","~15ba20105","~2024bb2d","~134105a04","14507cd4d","~13410604b","76aa78b","~99198a20","~1341056a0","~2024bb10","~10a5cb677","~c2ab313b","13,41,05,965","311c4c4b","~47b66045","~9919868d","~311c4b76","58adca91","76aa76b","82617b34","~2024bada","8261827a","~99198a29","~2024bad5","58add69d","~134105b84","d3a4a6dd","~c2ab3033","8,26,18,089","~1341061bb","d3a4a69c","58add2da","aa132703","~aa1321c5","~c2ab2c0c","~9919871c","~47b66614","14507cc74","11b564b69","aa132225","311c4c14","311c4c11","76aa762","311c4bbb","14507cba8","11b564830","76aac69","7,16,80,627","~oldetmoney","~c2ab275a","~15ba1d70a","~c2ab2ba2","~47b6665c","~991989d1","~old2024c085","~oldmagma1","~oldmagma2","~oldmagma3","11b5646ca","~11b5646b8","~d3a4a286","76aa85d","~47b65b94","~c2ab3083","~d3a49c4c-old","~15ba2019a","~oldrangde","~10a5cb636","9,91,98,226","~99198a14","8,26,17,775","~134105aac","311c4bc4","d3a4a420","~10a5cb20c","14507cc0a","~15ba1db98","~14507ccb9","~134105a45","~15ba200d7","~11b564836","~14507ccc0","~7168071b","~2024bb90","8,26,17,869","~oldUPES","14507cba1","~47b65a1c","8,26,17,779","~10a5cb53d","13,41,05,732","~311c4dc3","~2024bb26","~76aab32","old~2024c1a3","826182a0","~10a5cb24b","11b56488a","~15ba20234","~15ba2063d","~47b661ab","~47b65875","9,91,98,624","14507cd97","13,41,05,353","d3a4ab04","~ c2ab260c","~47b66257","~c2ab3091","~c2ab30aa","~15ba20080","d3a4a3b7","~aa131bd4","~134105b82","~134105a36","58add667","aa131752","11b564bc3","~14507d169"
        ]
    },
    "KSA": {
        "base_url": "https://p1o82kY:kow3jJs9@dashboard.ksa.webengage.com",
        "sheet_name": "Overview KSA",
        "publisher_url": "https://p1o82kY:kow3jJs9@dashboard.ksa.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "ksa~~15ba20526"
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

    # ⚡ FAST check (no page_source)
    if driver.find_elements(By.XPATH, "//h2[contains(text(),'Oops')]"):
        return True, "Region mismatch"

    if driver.find_elements(By.XPATH, "//a[.//span[text()='Request a Demo']]"):
        return True, "Service stopped"

    return False, None

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

from datetime import datetime, timedelta

def get_previous_month_range():

    today = datetime.today()

    first_day_this_month = today.replace(day=1)

    last_day_prev_month = first_day_this_month - timedelta(days=1)

    first_day_prev_month = last_day_prev_month.replace(day=1)

    start_label = first_day_prev_month.strftime("%B %-d, %Y")
    end_label = last_day_prev_month.strftime("%B %-d, %Y")

    return start_label, end_label

# MANUAL_START_DAY = 1
# MANUAL_END_DAY = 28

# def get_previous_month_range():

#     start_day = MANUAL_START_DAY
#     end_day = MANUAL_END_DAY

#     return start_day, end_day

def go_to_overview(driver, wait, account_id):
    base_url = REGION_CONFIG[REGION]["base_url"]
    url = f"{base_url}/accounts/{account_id}/engagement/overview/all"

    driver.get(url)

def wait_for_overview_page(driver):
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//span[contains(text(),'Last')]"))
    )

def select_previous_month_date(driver, wait):

    start_label, end_label = get_previous_month_range()

    print(f"📅 Selecting date: {start_label} → {end_label}")

    # open date dropdown
    wait.until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(),'Last')]"))
    ).click()

    # click custom dates
    wait.until(
        EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='Custom dates']"))
    ).click()

    # click previous month arrow
    prev_btn = wait.until(
        EC.element_to_be_clickable((
            By.XPATH,
            "//button[@aria-label='Move backward to switch to the previous month']"
        ))
    )

    prev_btn.click()

    print("⬅️ Moved calendar to previous month")

    time.sleep(1)

    # start date
    start_date = wait.until(
        EC.element_to_be_clickable((
            By.XPATH,
            f"//button[contains(@aria-label,'{start_label}')]"
        ))
    )

    start_date.click()

    # end date
    end_date = wait.until(
        EC.element_to_be_clickable((
            By.XPATH,
            f"//button[contains(@aria-label,'{end_label}')]"
        ))
    )

    end_date.click()

    # apply
    wait.until(
        EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='APPLY']"))
    ).click()

    WebDriverWait(driver, 10).until(
    lambda d: (
        d.find_elements(By.XPATH, "//tbody//tr")
        or d.find_elements(By.XPATH, "//h5[contains(text(),'Channel statistics not available')]")
    )

)

    print("✅ Date range applied")

# If no Data available -> check
def is_overview_no_data(driver):
    return "Channel statistics not available" in driver.page_source

def wait_for_overview_state(driver, timeout=10):
    """
    Waits until either:
    - table rows appear
    - OR no-data message appears
    Returns: "DATA" | "NO_DATA"
    """
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: (
                d.find_elements(By.XPATH, "//tbody//tr")
                or d.find_elements(By.XPATH, "//h5[contains(text(),'Channel statistics not available')]")
            )
        )
    except:
        return "UNKNOWN"

    if driver.find_elements(By.XPATH, "//h5[contains(text(),'Channel statistics not available')]"):
        return "NO_DATA"

    if driver.find_elements(By.XPATH, "//tbody//tr"):
        return "DATA"

    return "UNKNOWN"

def extract_overview_table(driver, license_code, start_label, end_label):
    print("📥 Extracting Overview table...")

    state = wait_for_overview_state(driver)

    if state == "NO_DATA":
        print("ℹ️ No campaigns → NO_DATA")
        return [[
            license_code,
            "NO_DATA","NO_DATA","NO_DATA","NO_DATA",
            "NO_DATA","NO_DATA","NO_DATA",
            start_label,
            end_label
        ]]

    if state != "DATA":
        print("⚠️ Unknown state → treating as NO_DATA")
        return [[
            license_code,
            "NO_DATA","NO_DATA","NO_DATA","NO_DATA",
            "NO_DATA","NO_DATA","NO_DATA",
            start_label,
            end_label
        ]]

    # ✅ DATA exists
    rows = driver.find_elements(
        By.XPATH,
        "//tbody/tr[contains(@class,'table__row')]"
    )

    rows_data = []

    for row in rows:
        cells = row.find_elements(By.XPATH, "./td")

        if len(cells) < 7:
            continue

        rows_data.append([
            license_code,
            cells[0].text.strip() or "-",
            cells[1].text.strip() or "-",
            cells[2].text.strip() or "-",
            cells[3].text.strip() or "-",
            cells[4].text.strip() or "-",
            cells[5].text.strip() or "-",
            cells[6].text.strip() or "-",
            start_label,
            end_label
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
        # Step A: Search
        search_by_license(driver, wait, code)

        # wait for either result / oops / request demo
        WebDriverWait(driver, 5).until(
            lambda d: (
                d.find_elements(By.XPATH, f"//tr[contains(., '{code}')]")
                or d.find_elements(By.XPATH, "//h2[contains(text(),'Oops')]")
                or d.find_elements(By.XPATH, "//a[.//span[text()='Request a Demo']]")
            )
        )

        # 🔥 Step B: Skip check FIRST
        skip, reason = should_skip_account(driver)

        if skip:
            print(f"⏭ Skipping {code} → {reason}")

            log_error_to_sheet(sheet, code, "SKIPPED", reason)

            driver.get(REGION_CONFIG[REGION]["publisher_url"])
            wait.until(EC.presence_of_element_located((By.NAME, "licenseCode")))
            continue

        # 🚫 Step C: LC not found
        if not check_if_result_exists(driver, code):
            log_error_to_sheet(
                sheet,
                code,
                "REGION_MISMATCH",
                "License code not found in this region"
            )

            driver.get(REGION_CONFIG[REGION]["publisher_url"])
            continue

        # Step D: Access handling
        try:
            open_actions_dropdown(driver, wait, code)

            request_btn_xpath = f"//tr[contains(., '{code}')]//a[contains(@class,'requestAccess')]"

            if driver.find_elements(By.XPATH, request_btn_xpath):
                click_request_access(wait)
                handle_request_modal(wait, driver)
                time.sleep(2)
                close_modal_if_exists(driver)
            else:
                print("ℹ️ Access already available")
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()

        except Exception as e:
            print(f"⚠️ Access step skipped: {e}")

        # Step E: Edit
        main_window = driver.current_window_handle
        click_edit(driver, wait, code)

        for window_handle in driver.window_handles:
            if window_handle != main_window:
                driver.switch_to.window(window_handle)
                break

        # Step F: Overview
        go_to_overview(driver, wait, code)

        wait_for_overview_page(driver)  # ✅ NEW

        start_label, end_label = get_previous_month_range()

        select_previous_month_date(driver, wait)

        overview_rows = extract_overview_table(driver, code, start_label, end_label)
        append_to_sheet(sheet, overview_rows)

    except Exception as e:
        log_error_to_sheet(sheet, code, "OVERVIEW", str(e))

    finally:
        if len(driver.window_handles) > 1:
            driver.close()
            driver.switch_to.window(main_window)

        driver.get(REGION_CONFIG[REGION]["publisher_url"])
        wait.until(EC.presence_of_element_located((By.NAME, "licenseCode")))
        time.sleep(1)