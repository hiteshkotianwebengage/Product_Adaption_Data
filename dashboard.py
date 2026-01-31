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

def init_google_sheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(
        "/Users/admin/Desktop/Python Script/agreement_file_pasting/mycred-googlesheet.json", scopes=scopes
    )
    client = gspread.authorize(creds)

    sheet = client.open_by_key("1sathL7caATX3PnV2urKhp8UpCLOcwdTIxzkJkA4Kar4").worksheet("Copy of Dashboard Global")
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
driver.get("https://dashboard.in.webengage.com/admin")

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
            driver.get("https://dashboard.in.webengage.com/admin/publisher.html?action=list")

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
        print("🚀 Forcing selection via JavaScript...")
        driver.execute_script(
            "arguments[0].value = '~32537i7'; arguments[0].dispatchEvent(new Event('change'));", 
            role_dropdown
        )
        
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


def click_dashboards(wait):
    print("⏳ Clicking Dashboards...")

    dashboard_xpath = (
        "//a[contains(@href,'/custom-dashboard/list') and .//span[text()='Dashboards']]"
    )

    wait.until(
        EC.element_to_be_clickable((By.XPATH, dashboard_xpath))
    ).click()

    print("✅ Dashboards opened")

def wait_for_dashboard_table_or_empty(driver, timeout=8):
    """
    Waits for either:
    - dashboard rows
    - empty dashboard state
    - pagination (even if empty)
    Returns True if rows exist, False if NO_DATA
    """
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: (
                d.find_elements(By.CLASS_NAME, "table__row")
                or d.find_elements(By.CLASS_NAME, "pagination")
                or "No dashboards" in d.page_source
            )
        )
    except:
        return False

    rows = driver.find_elements(By.CLASS_NAME, "table__row")
    return len(rows) > 0

def extract_dashboard_page(driver, license_code):
    rows_data = []

    rows = driver.find_elements(
        By.XPATH,
        "//tbody/tr[contains(@class,'table__row')]"
    )

    for row in rows:
        try:
            cells = row.find_elements(By.XPATH, "./td")
            if len(cells) < 5:
                continue

            # 1️⃣ Dashboard Name
            dashboard_name = cells[0].find_element(
                By.XPATH, ".//a"
            ).get_attribute("title").strip()

            # 2️⃣ Cards Count
            cards = cells[1].text.strip()

            # 3️⃣ Last Updated
            last_updated = cells[2].find_element(
                By.XPATH, ".//span"
            ).get_attribute("title").strip()

            # 4️⃣ Tags
            try:
                tags = ",".join([
                    t.text.strip()
                    for t in cells[3].find_elements(By.CLASS_NAME, "pill-text")
                ])
            except:
                tags = ""

            rows_data.append([
                license_code,
                dashboard_name,
                cards,
                last_updated,
                tags
            ])

        except Exception as e:
            print(f"⚠️ Skipping dashboard row: {e}")
            continue

    return rows_data


def extract_all_dashboards(driver, license_code):
    print("📥 Extracting Dashboards (all pages)...")

    has_data = wait_for_dashboard_table_or_empty(driver)

    # 🔹 NO DATA FAST EXIT
    if not has_data:
        print("⚠️ No dashboards found — fast skipping")
        return [[
            license_code,
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            "NO_DATA"
        ]]

    all_data = []
    page_number = 1

    while True:
        print(f"📄 Scraping Dashboard Page {page_number}...")
        page_data = extract_dashboard_page(driver, license_code)
        all_data.extend(page_data)

        try:
            next_li = driver.find_element(
                By.XPATH,
                "//li[contains(@class,'pagination__next')]"
            )

            if "is-disabled" in next_li.get_attribute("class"):
                break

            next_link = next_li.find_element(By.TAG_NAME, "a")
            driver.execute_script("arguments[0].click();", next_link)
            time.sleep(1.5)
            page_number += 1

        except:
            break

    if not all_data:
        return [[
            license_code,
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            "NO_DATA"
        ]]

    print(f"✅ Total Dashboards Extracted: {len(all_data)}")
    return all_data


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


LICENSE_CODES = [
    
    "in~~15ba205d1","in~~10a5cbb1d","in~311c4742","in~311c4724","in~~991991c4","in~~134106220","in~311c488b","in~11b564274","in~14507c728","in~~15ba20670","in~~15ba2068a","in~~47b666d5","in~76aa1d8","in~76aa247","in~~99199258","in~~2024c207","in~~c2ab36a2","in~aa1318ab","in~d3a49bac","in~~c2ab3671","in~~47b66716","in~~991991d0","in~~10a5cbb34","in~76aa273","in~~99199217","in~76aa298","in~~1341061c6","in~d3a49b75","in~58adcb79","in~~15ba20672","in~~71680bbb","in~58adcb94","in~11b5641d0","in~d3a49b43","in~~10a5cbb79","in~826171c3","in~58adcb40","in~14507c6a9","in~76aa22a","in~~15ba20749","in~311c467c","in~~47b66689","in~~2024c1d7","in~~2024c218","in~~71680b3c","in~~47b6668a","in~~99199244","in~~c2ab36ad","in~76aa268","in~~9919921b","in~~134106216","in~~71680b92","in~76aa245","in~311c46d4","in~311c46d3","in~58adcb85","in~~2024c249","in~76aa1c0","in~11b564332","in~~71680c19","in~~15ba20752","in~~2024c2a0","in~14507c681","in~~2024c2c1","in~76aa206","in~aa131675","in~14507c65b","in~11b5641a9","in~d3a49b10","in~~71680c2b","in~58adcb50","in~~10a5cbc25","in~aa13163a","in~11b56418d","in~311c4663","in~~c2ab3781","in~~1341062bb","in~~991992c4","in~~10a5cbc2d","in~~1341062c1","in~~991992cc","in~14507c641","in~~71680c30","in~aa13164b","in~~15ba20759","in~~15ba205c0","in~~2024c231","in~~47b6677d","in~58adcb36","in~aa13166b","in~~991992d1","in~~1341062c2","in~~99199081","in~14507c63b","in~aa131665","in~~71680b90","in~14507c666","in~aa131632","in~76aa20d","in~311c464b","in~311c4766","in~~71680c4c","in~11b564172","in~~47b66782","in~11b564181"
]

for code in LICENSE_CODES:
    print(f"\n▶ Processing {code}")
    try:
        # Step A: Search and land on result
        search_by_license(driver, wait, code)
        
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
            click_dashboards(wait)

            # Wait for dashboard table to load
            dashboard_rows = extract_all_dashboards(driver, code)
            append_to_sheet(sheet, dashboard_rows)

            time.sleep(1)

            dashboard_rows = extract_all_dashboards(driver, code)

            append_to_sheet(sheet, dashboard_rows)

        except Exception as e:
            error_msg = str(e)
            print(f"❌ Dashboard extraction failed for {code}: {error_msg}")

            log_error_to_sheet(
                sheet,
                code,
                stage="DASHBOARD_EXTRACTION",
                error_reason=error_msg
            )
    
    finally:
        # Step E: Cleanup for next iteration
        # Close extra tabs and go back to the list
        if len(driver.window_handles) > 1:
            driver.close() # Closes current (Edit) tab
            driver.switch_to.window(main_window)
        
        driver.get("https://dashboard.in.webengage.com/admin/publisher.html?action=list")
        time.sleep(2)
