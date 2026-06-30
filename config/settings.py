# This pre is becoz we have username and password asked when clicked
# on this url at very start so we will use this thing for starting urls
PRE_BASE_URLS = {
    "INDIA": "https://p1o82kY:kow3jJs9@dashboard.in.webengage.com",
    "GLOBAL": "https://p1o82kY:kow3jJs9@dashboard.webengage.com",
    "KSA": "https://p1o82kY:kow3jJs9@dashboard.ksa.webengage.com"
}

BASE_URLS = {
    "INDIA": "https://dashboard.in.webengage.com",
    "GLOBAL": "https://dashboard.webengage.com",
    "KSA": "https://dashboard.ksa.webengage.com"
}

ROLE_IDS = {
    "INDIA": "~32537i7",
    "GLOBAL": "abd40ke",
    "KSA": "abd40ke"
}
# The drive URL
ROOT_FOLDER_ID = "1ofGHkTIOYcYvJe_IMuuLo8SSW3P69WUT"

Credential_File = "/Users/admin/Desktop/Code Directory/Product_Adaption_Data/Credential File/mycred-googlesheet.json"

from datetime import datetime, timedelta

# Updated config/settings.py (Optional cleanup)
def get_month_info():
    today = datetime.today()
    first_this = today.replace(day=1)
    last_prev = first_this - timedelta(days=1)
    first_prev = last_prev.replace(day=1)

    # API Formats
    # start_date = first_prev.strftime("%Y-%m-%dT00:00:00.000+05:30")
    # end_date = last_prev.strftime("%Y-%m-%dT23:59:59.999+05:30")

    # # Label Formats
    # start_label = first_prev.strftime("%B %d, %Y")
    # end_label = last_prev.strftime("%B %d, %Y")

    # ----------------------------
    # January
    # ----------------------------
    # # Exact ISO format for WebEngage API - January
    # start_date = "2026-01-01T00:00:00.000+05:30"
    # end_date = "2026-01-31T23:59:59.999+05:30"

    # # Labels for your CSV/Print outputs
    # start_label = "January 01, 2026"
    # end_label = "January 31, 2026"

    # ----------------------------
    # February
    # ----------------------------
    # # Exact ISO format for WebEngage API - February
    # start_date = "2026-02-01T00:00:00.000+05:30"
    # end_date = "2026-02-28T23:59:59.999+05:30"

    # # Labels for your CSV/Print outputs
    # start_label = "February 01, 2026"
    # end_label = "February 28, 2026"

    # ----------------------------
    # March
    # ----------------------------
    # # Exact ISO format for WebEngage API - March
    # start_date = "2026-03-01T00:00:00.000+05:30"
    # end_date = "2026-03-31T23:59:59.999+05:30"

    # # Labels for your CSV/Print outputs
    # start_label = "March 01, 2026"
    # end_label = "March 31, 2026"

    # ----------------------------
    # April
    # ----------------------------
    # Exact ISO format for WebEngage API - April
    start_date = "2026-04-01T00:00:00.000+05:30"
    end_date = "2026-04-30T23:59:59.999+05:30"

    # Labels for your CSV/Print outputs
    start_label = "April 01, 2026"
    end_label = "April 30, 2026"

    # ----------------------------
    # MAY
    # ----------------------------
    # # Exact ISO format for WebEngage API - May
    # start_date = "2026-05-01T00:00:00.000+05:30"
    # end_date = "2026-05-31T23:59:59.999+05:30"

    # # Labels for your CSV/Print outputs
    # start_label = "May 01, 2026"
    # end_label = "May 31, 2026"

    # ----------------------------
    # MAY
    # ----------------------------
    # # Exact ISO format for WebEngage API - June
    # start_date = "2026-06-01T00:00:00.000+05:30"
    # end_date = "2026-06-30T23:59:59.999+05:30"

    # # Labels for your CSV/Print outputs
    # start_label = "June 01, 2026"
    # end_label = "June 30, 2026"

    return start_date, end_date, start_label, end_label

# ---------------------------
# Back fill temporary
from datetime import datetime

def get_backfill_months():
    months = [
        ("2025-10-01", "2025-10-31"),
        ("2025-11-01", "2025-11-30"),
        ("2025-12-01", "2025-12-31"),
        ("2026-01-01", "2026-01-31"),
        ("2026-02-01", "2026-02-28"),
        ("2026-03-01", "2026-03-31"),
    ]

    result = []

    for start, end in months:
        # 1. Create the actual datetime objects
        start_dt = datetime.strptime(start, "%Y-%m-%d")
        # Set end_dt to the very end of the day to catch late-night campaigns
        end_dt = datetime.strptime(end, "%Y-%m-%d").replace(hour=23, minute=59, second=59)

        start_api = start_dt.strftime("%Y-%m-%dT00:00:00.000+05:30")
        end_api = end_dt.strftime("%Y-%m-%dT23:59:59.999+05:30")

        result.append({
            "start_dt": start_dt,       # 🔥 CRITICAL: Needed for is_in_month
            "end_dt": end_dt,           # 🔥 CRITICAL: Needed for is_in_month
            "start_api": start_api,
            "end_api": end_api,
            "start_label": start_dt.strftime("%B %d, %Y"),
            "end_label": end_dt.strftime("%B %d, %Y"),
            "month_name": start_dt.strftime("%b'%y")
        })

    return result