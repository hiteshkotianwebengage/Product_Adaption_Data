def get_session_cookies(driver):
    """
    Captures all cookies from Selenium and returns a dictionary.
    Includes a small delay to ensure all session cookies are set.
    """
    # Wait a split second for the browser to finalize session cookies
    import time
    time.sleep(1) 
    
    selenium_cookies = driver.get_cookies()
    cookie_dict = {c['name']: c['value'] for c in selenium_cookies}
    
    # DEBUG: Check if the main session cookie exists
    # For WebEngage, look for 'webengage.session.id' or 'JSESSIONID' or similar
    if not cookie_dict:
        print("⚠️ WARNING: No cookies captured!")
        
    return cookie_dict