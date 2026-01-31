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

    sheet = client.open_by_key("1sathL7caATX3PnV2urKhp8UpCLOcwdTIxzkJkA4Kar4").worksheet("User India")
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


LICENSE_CODES = [
    "in~d3a49ca9","in~~9919912c","in~311c47dd","in~14507c7a6","in~~c2ab3610","in~~991990bc","in~~13410618d","in~~99199192","in~~134106156","in~311c4863","in~~15ba205db","in~311c4838","in~11b5642c7","in~~2024c18c","in~76aa307","in~8261735b","in~58adcc15","in~aa131896","in~826173c2","in~14507c7cd","in~58adcd07","in~~10a5cba6d","in~58adcbdb","in~~2024c1a8","in~~47b66667","in~14507c80b","in~76aa34a","in~14507c789","in~14507c784","in~11b564357","in~~c2ab364c","in~~c2ab3662","in~76aa392","in~~10a5cba77","in~8261729b","in~~10a5cbb14","in~14507c76b","in~311c4774","in~~47b66639","in~14507c77b","in~~47b66647","in~58adcc4a","in~~47b665d8","in~~2024c179","in~58adcc70","in~58adcc59","in~~c2ab363b","in~76aa2c5","in~~10a5cba63","in~~9919913a","in~82617341","in~11b5642aa","in~11b56430c","in~~1341061ba","in~~2024c1cd","in~~10a5cbb09","in~~2024c156","in~~134106180","in~aa131782","in~14507c7c3","in~d3a49c7d","in~11b5642a4","in~~c2ab3517","in~76aa2c9","in~~71680ad0","in~~2024c1c6","in~~47b66678","in~~15ba2065c","in~~99199151","in~~99199168","in~~c2ab361a","in~14507c838","in~~1341061ac","in~58adcc11","in~76aa35b","in~~2024c081","in~~134106132","in~~c2ab35d2","in~58adcc57","in~~13410619c","in~~71680aa0","in~d3a49cb0","in~58adcc36","in~aa1317cc","in~~10a5cbb29","in~~15ba205d1","in~~1341061b6","in~~1341061b5","in~~10a5cbb1d","in~d3a49c80","in~11b5642a5","in~~10a5cbac3","in~58adcbd2","in~58adcbda","in~~c2ab3690","in~~47b66670","in~~2024c1bc","in~~c2ab3695","in~~991991c8","in~~15ba20633","in~~47b6663c","in~~1341061cb","in~76aa2a6","in~d3a49c1a","in~~47b66699","in~311c4742","in~~99199207","in~311c472d","in~311c474b","in~~134106200","in~311c4724","in~~134106208","in~~991991c4","in~~134106220","in~~c2ab3675","in~311c488b","in~11b564274","in~~10a5cbb2d","in~~71680b39","in~8261728c","in~14507c728","in~~15ba20670","in~~15ba206a3","in~~15ba206bb","in~~15ba206a9","in~aa1316d4","in~~15ba2068a","in~8261726b","in~~2024c233","in~76aa24d","in~~47b666d5","in~76aa1d8","in~8261722b","in~76aa247","in~~2024c254","in~~99199258","in~311c4708","in~~c2ab35bb","in~~c2ab368c","in~d3a49b8c","in~311c4773","in~~71680b78","in~311c4703","in~~71680b65","in~311c4744","in~11b5642a0","in~~2024c207","in~76aa2b4","in~~10a5cbb42","in~~c2ab36a2","in~~99199206","in~11b564260","in~14507c6ca","in~~71680a90","in~~99199205","in~aa1318ab","in~~47b66709","in~~134106257","in~~c2ab3714","in~d3a49b66","in~aa1316dd","in~58adcb8b","in~~47b666c4","in~d3a49b94","in~d3a49bac","in~aa1316cc","in~~71680b93","in~~10a5cbb66","in~~c2ab36d5","in~aa1316c3","in~~15ba206c2","in~76aa221","in~~c2ab3671","in~~47b66716","in~76aa23c","in~82617226","in~82617217","in~~c2ab36d4","in~8261723d","in~~2024c255","in~d3a49ba1","in~d3a49b86","in~~15ba20658","in~~2024c247","in~14507c71d","in~~991991d0","in~~10a5cbb34","in~76aa273","in~~15ba20652","in~~99199217","in~76aa298","in~~1341061c6","in~~1341061c8","in~aa1316aa","in~~10a5cba3a","in~~10a5cbad8","in~~10a5cbba6","in~~10a5cbbb5","in~76aa21c","in~~c2ab36d9","in~14507c738","in~~15ba206d5","in~14507c6b7","in~d3a49b75","in~11b5641db","in~58adcb61","in~~15ba20690","in~76aa1c5","in~~134106263","in~58adcb79","in~aa1316d1","in~~15ba20672","in~~c2ab3721","in~~71680bbb","in~82617205","in~14507c695","in~11b56420d","in~82617246","in~58adcb94","in~82617203","in~~134106115","in~58adcc83","in~d3a49b80","in~~99199283","in~76aa200","in~~134106273","in~~71680c0c","in~826171d8","in~~1341061b2","in~311c46c9","in~11b5641d0","in~~2024c239","in~~99199233","in~76aa1d3","in~d3a49b43","in~58adcb4b","in~~134106259","in~~10a5cbba9","in~76aa1b3","in~826171d3","in~~10a5cbb79","in~~10a5cbbc9","in~~134106294","in~11b5641ba","in~~47b66722","in~~2024c242","in~d3a49b5a","in~826171c3","in~58adcb40","in~d3a49b3a","in~~2024c276","in~~47b66714","in~~71680ba6","in~~10a5cbbdd","in~~99199068","in~~71680bad","in~~13410624c","in~~10a5cbb76","in~~134106267","in~14507c68c","in~14507c6a9","in~~c2ab3708","in~76aa22a","in~76aa1d0","in~~10a5cbc11","in~~10a5cbc11","in~~10a5cbba4","in~~47b666dc","in~311c4671","in~58adcb30","in~311c467b","in~~47b6675b","in~~991992ab","in~~134106286","in~d3a49c4c","in~~99199277","in~~10a5cbb38","in~~71680c14","in~~13410629c","in~~15ba20749","in~~10a5cbb14","in~~71680bd5","in~311c467c","in~~2024c2aa","in~~71680b12","in~~47b66689","in~~2024c1d7","in~~2024c218","in~311c474a","in~~71680b3c","in~~47b6668a","in~~99199244","in~~2024c246","in~11b564276","in~~134106213","in~~15ba206a8","in~~10a5cbb61","in~11b564256","in~~c2ab36ad","in~76aa268","in~~9919921b","in~~134106216","in~~71680b69","in~~71680b92","in~76aa245","in~311c46d4","in~311c46d3","in~58adcb85","in~~2024c249","in~11b5641d1","in~~47b66730","in~aa131685","in~826172a7","in~76aa1c0","in~311c4665","in~11b564332","in~14507c7d2","in~11b564340","in~aa131666","in~11b5641a6","in~~71680c19","in~~134106253","in~~15ba20752","in~~2024c085","in~d3a49b5d","in~76aa1d7","in~~10a5cbc14","in~aa131655","in~~47b66750","in~76aa241","in~~2024c2a0","in~14507c681","in~aa131667","in~~134106266","in~11b5641a0","in~~15ba20741","in~~991992a7","in~~991992b1","in~~2024c2c1","in~~71680b61","in~76aa206","in~~10a5cbc2c","in~826171b0","in~~71680c29","in~aa131676","in~~71680bb9","in~d3a49b24","in~~c2ab3735","in~aa131652","in~14507c67b","in~aa131675","in~14507c65b","in~11b5641a9","in~~2024c27c","in~11b5641aa","in~d3a49b10","in~aa13163d","in~~15ba20753","in~d3a49b0b","in~d3a49b14","in~~991992c6","in~~15ba2076c","in~~71680c2b","in~~1341062c9","in~14507c647","in~82617199","in~~71680c38","in~58adcb50","in~~991992aa","in~76aa201","in~58adcb08","in~~991992c7","in~~47b66733","in~~10a5cbc25","in~aa131650","in~aa13163a","in~11b56418d","in~11b564191","in~~2024c2b8","in~311c4663","in~76aa1a2","in~~15ba2074d","in~~c2ab3781","in~~1341062bb","in~~991992c4","in~~10a5cbc2d","in~~1341062c1","in~~99199240","in~~71680b76","in~~991992cc","in~~47b6678c","in~311c4664","in~14507c641","in~~71680c30","in~aa13164b","in~~991992a4","in~~15ba20759","in~~15ba205c0","in~~2024c231","in~76aa1ac","in~11b5641b1","in~~47b6677d","in~58adcb36","in~aa13166b","in~~991992d1","in~~1341062c2","in~~99199081","in~14507c63b","in~~c2ab3786","in~11b564246","in~~99199278","in~11b56417b","in~aa131665","in~~71680b90","in~14507c666","in~aa131632","in~76aa20d","in~311c464b","in~311c4766","in~~c2ab3761","in~~71680c4c","in~11b564177","in~11b564172","in~d3a49ad8","in~~47b66782","in~11b56417c","in~11b564181","in~~c2ab3789","in~11b5642d2","in~311c4646"
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
        
        driver.get("https://dashboard.in.webengage.com/admin/publisher.html?action=list")
        time.sleep(2)
