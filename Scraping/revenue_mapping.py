'''
    Things to change when we switch between India, GLobal, KSA
    1) Below Client inside init_google_sheet function change the sheet name
    2) In the step 0 below the Driver varibale change between driver.get 
    3) In the step 2 while direct url for publisher change between driver.get
    4) Inside the go_to_revenue_mapping function we switch between the url
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
REGION = "KSA"   # options: "INDIA", "GLOBAL", "KSA"

REGION_CONFIG = {
    "INDIA": {
        "base_url": "https://p1o82kY:kow3jJs9@dashboard.in.webengage.com",
        "sheet_name": "Revenue Mapping India",
        "publisher_url": "https://p1o82kY:kow3jJs9@dashboard.in.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "in~~10a5cbc14","in~aa131655","in~~47b66750","in~76aa241","in~~2024c2a0","in~14507c681","in~aa131667","in~~134106266","in~11b5641a0","in~~15ba20741","in~~2024c2c1","in~~71680b61","in~76aa206","in~~10a5cbc2c","in~aa131676","in~~71680bb9","in~~c2ab3735","in~aa131652","in~14507c67b","in~aa131675","in~14507c65b","in~11b5641a9","in~~2024c27c","in~11b5641aa","in~d3a49b10","in~~15ba20753","in~d3a49b0b","in~~991992c6","in~~15ba2076c","in~~71680c2b","in~82617199","in~58adcb50","in~76aa201","in~~47b66733","in~~10a5cbc25","in~aa131650","in~aa13163a","in~11b56418d","in~11b564191","in~~2024c2b8","in~311c4663","in~76aa1a2","in~~15ba2074d","in~~c2ab3781","in~~1341062bb","in~~991992c4","in~~10a5cbc2d","in~~1341062c1","in~~991992cc","in~311c4664","in~14507c641","in~~71680c30","in~aa13164b","in~~991992a4","in~~15ba20759","in~~15ba205c0","in~~2024c231","in~76aa1ac","in~11b5641b1","in~~47b6677d","in~58adcb36","in~aa13166b","in~~991992d1","in~~1341062c2","in~~99199081","in~14507c63b","in~~99199278","in~14507c666","in~aa131632","in~76aa20d","in~311c464b","in~311c4766","in~11b564177","in~11b564172","in~d3a49ad8","in~~47b66782","in~11b564181","in~~c2ab3789","in~311c4646","in~~c2ab36a7","in~~47b6665b","in~~99199306","in~~c2ab3794","in~~991992cd","in~~2024c262","in~aa13162a","in~~47b66752","in~82617256","in~14507c69a","in~~99199213","in~~71680b7d","in~58adcb46","in~~10a5cbbb6","in~~9919926b","old-in~~99199240","in~aa13168a","in~311c4685","in~58adcb73","in~14507c6ad","in~~134106261","in~~134106255","in~~71680bb8","in~~c2ab36cb","in~aa131694","old-in~~99199240"
        ]
    },
    "GLOBAL": {
        "base_url": "https://p1o82kY:kow3jJs9@dashboard.webengage.com",
        "sheet_name": "Revenue Mapping Global",
        "publisher_url": "https://p1o82kY:kow3jJs9@dashboard.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "~d3a4a667","~11b56421","~11b565abc","~c2ab1c88","~d3a4b624","~old~11b564403","~bit24newtempold","82617855","~2024c07c","~991978c7","~311c60ad","~82618978","~oldin~~c2ab3761","~145080023","~11b565971","~11b565961","~134104919","~old~76ab96b","~2024a939","~82618947","~7167d4a4"
        ]
    },
    "KSA": {
        "base_url": "https://p1o82kY:kow3jJs9@dashboard.ksa.webengage.com",
        "sheet_name": "Revenue Mapping KSA",
        "publisher_url": "https://p1o82kY:kow3jJs9@dashboard.ksa.webengage.com/admin/publisher.html?action=list",
        "license_codes": [
            "ksa~~2024c070","ksa~~47b6652a","ksa~58adcd4c","ksa~11b564406","ksa~82617404","ksa~d3a49d46","ksa~~2024c07d","ksa~~99199078","ksa~~99199083","ksa~~13410606b","ksa~11b5643d5","ksa~~134106080","ksa~58adcd44","ksa~~134106084","ksa~~99199087","ksa~aa131893","ksa~14507c89c","ksa~~47b66522","ksa~~c2ab3537","ksa~~99199088","ksa~~2024c083","ksa~~2024c084"
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
        print("🚀 Forcing selection via JavaScript...")
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


def open_data_platform(driver, wait):
    print("⏳ Ensuring Data Platform menu is open...")

    data_platform_li = wait.until(
        EC.presence_of_element_located((By.ID, "nav-data-platform"))
    )

    class_attr = data_platform_li.get_attribute("class")

    if "menu__group--is-active" not in class_attr:
        print("🔓 Opening Data Platform sidebar via JS")
        driver.execute_script(
            "arguments[0].classList.add('menu__group--is-active');",
            data_platform_li
        )
        time.sleep(0.5)
    else:
        print("✅ Data Platform already open")

def go_to_revenue_mapping(driver, wait, account_id):
    base_url = REGION_CONFIG[REGION]["base_url"]
    url = f"{base_url}/accounts/{account_id}/data-management/events/revenue"

    driver.get(url)

    # wait for rows OR empty state
    wait_for_revenue_or_empty(driver)
    print("✅ Landed on Revenue Mapping (via URL)")


def click_data_management(wait):
    print("⏳ Clicking Data Management...")

    data_management_xpath = (
        "//a[contains(@href,'/data-management/system/attributes') and .//span[text()='Data Management']]"
    )

    wait.until(
        EC.element_to_be_clickable((By.XPATH, data_management_xpath))
    ).click()

    print("✅ Clicked Data Management")

def click_revenue_mapping(wait):
    print("⏳ Waiting for Revenue Mapping tab...")

    revenue_tab_xpath = (
        "//a[contains(@href,'/data-management/events/revenue') and normalize-space()='Revenue Mapping']"
    )

    wait.until(
        EC.element_to_be_clickable((By.XPATH, revenue_tab_xpath))
    ).click()

    print("✅ Revenue Mapping opened")

def extract_revenue_mapping_data(driver, wait, license_code):
    print("📥 Extracting Revenue Mapping data...")
    data_rows = []

    account_name = wait.until(
        EC.presence_of_element_located((By.ID, "we-account-name"))
    ).text.strip()

    try:
        currency = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((
                By.XPATH,
                "//div[@aria-label='Select a currency']//span[contains(@class,'handle-text-overflow')]"
            ))
        ).text.strip()
    except:
        currency = "NULL"

    # ⏳ Wait briefly for rows OR empty state
    has_data = wait_for_revenue_or_empty(driver)

    if not has_data:
        print("⚠️ No Revenue Mapping found — inserting NO_DATA row")
        return [[
            license_code,
            account_name,
            currency,
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            ""
        ]]

    rows = driver.find_elements(
        By.XPATH,
        "//div[contains(@class,'row') and .//i[contains(@class,'fl-delete')]]"
    )

    for row in rows:
        try:
            dropdowns = row.find_elements(By.XPATH, ".//div[contains(@class,'r-ss-trigger')]")
            if len(dropdowns) >= 2:
                event_name = dropdowns[0].text.strip()
                attribute_name = dropdowns[1].text.strip()

                data_rows.append([
                    license_code,
                    account_name,
                    currency,
                    event_name,
                    attribute_name,
                    "SUCCESS",
                    ""
                ])
        except:
            continue

    if not data_rows:
        return [[
            license_code,
            account_name,
            currency,
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            ""
        ]]

    print(f"✅ Extracted {len(data_rows)} revenue mappings")
    return data_rows


def append_to_sheet(sheet, rows):
    if rows:
        sheet.append_rows(rows, value_input_option="USER_ENTERED")
        print("📤 Data pushed to Google Sheets")
    else:
        print("⚠️ No data to push")

def log_error_to_sheet(sheet, license_code, error_reason):
    print(f"📝 Logging error for {license_code}")

    row = [
        license_code,
        "NULL",
        "NULL",
        "NULL",
        "NULL",
        "ERROR",
        error_reason[:300]  # prevent huge stack traces
    ]

    sheet.append_row(row, value_input_option="USER_ENTERED")

def wait_for_revenue_or_empty(driver, timeout=6):
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: (
                d.find_elements(By.XPATH, "//i[contains(@class,'fl-delete')]")
                or "No data" in d.page_source
            )
        )
    except:
        return False

    rows = driver.find_elements(By.XPATH, "//i[contains(@class,'fl-delete')]")
    return len(rows) > 0

# Here we start to loop in the LC in the site 

LICENSE_CODES = REGION_CONFIG[REGION]["license_codes"]

for code in LICENSE_CODES:
    print(f"\n▶ Processing {code}")

    try:
        # Step A: Search
        search_by_license(driver, wait, code)

        # 🔥 Step B: Skip check
        skip, reason = should_skip_account(driver)

        if skip:
            print(f"⏭ Skipping {code} → {reason}")

            log_error_to_sheet(sheet, code, reason)

            driver.get(REGION_CONFIG[REGION]["publisher_url"])
            wait.until(EC.presence_of_element_located((By.NAME, "licenseCode")))
            continue

        # 🚫 Step C: LC not found
        if not driver.find_elements(By.XPATH, f"//tr[contains(., '{code}')]"):
            print("⛔ LC not found in this region")

            log_error_to_sheet(sheet, code, "Region mismatch")

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

        # Step F: Revenue
        go_to_revenue_mapping(driver, wait, code)

        rows = extract_revenue_mapping_data(driver, wait, code)
        append_to_sheet(sheet, rows)

        print(f"✅ Success for {code}")

    except Exception as e:
        print(f"❌ Failed {code}: {e}")
        log_error_to_sheet(sheet, code, str(e))

    finally:
        if len(driver.window_handles) > 1:
            driver.close()
            driver.switch_to.window(main_window)

        driver.get(REGION_CONFIG[REGION]["publisher_url"])
        wait.until(EC.presence_of_element_located((By.NAME, "licenseCode")))
        time.sleep(1)