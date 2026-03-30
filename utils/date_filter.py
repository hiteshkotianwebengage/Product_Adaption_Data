from datetime import datetime, timezone

def parse_iso_date(date_str):
    if not date_str:
        return None
    
    # Replace Z with UTC offset
    date_str = date_str.replace("Z", "+00:00")
    
    return datetime.fromisoformat(date_str)

def is_in_month(date_str, start_dt, end_dt):

    dt = parse_iso_date(date_str)
    if not dt:
        return False

    # Convert to naive for comparison (optional but safe)
    dt = dt.replace(tzinfo=None)

    return start_dt <= dt <= end_dt