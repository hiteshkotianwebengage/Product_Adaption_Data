from utils.date_filter import is_in_month

def parse_channel_data(contents, lc, channel_name, start_dt, end_dt, month_name):
    rows = []

    for item in contents:
        created_on = item.get("createdOn")
        if not created_on:
            continue

        if not is_in_month(created_on, start_dt, end_dt):
            continue

        stats = item.get("stats", {})

        # Extract values safely
        # Note: Some endpoints use 'click', some use 'clicks'
        clicks = stats.get("clicks", stats.get("click", 0))
        
        # Safe extraction of revenue and conversion variants
        conv = stats.get("conversion", 0)
        ct_conv = stats.get("clickThroughConversion", 0)
        
        rev = stats.get("revenue", 0)
        ct_rev = stats.get("clickThroughRevenue", 0)

        rows.append([
            lc,
            channel_name,
            item.get("id"),
            item.get("title"),
            item.get("status"),
            item.get("category"),
            created_on,
            item.get("startDate"),
            stats.get("sent", 0),
            stats.get("delivered", 0),
            clicks,
            conv,      # Total Conversions
            ct_conv,   # Click Through Conversions
            rev,       # Total Revenue
            ct_rev,    # Click Through Revenue
            month_name
        ])

    return rows