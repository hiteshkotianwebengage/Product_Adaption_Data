from utils.date_filter import is_in_month

def parse_channel_data(contents, lc, channel_name, month_data):

    rows = []

    for item in contents:
        created_on = item.get("createdOn")

        if not is_in_month(created_on, month_data["start_dt"], month_data["end_dt"]):
            continue

        stats = item.get("stats", {})

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
            stats.get("clicks", stats.get("click", 0)),
            stats.get("clickThroughConversion", 0),
            stats.get("revenue", 0),
            month_data["month_name"]
        ])

    return rows