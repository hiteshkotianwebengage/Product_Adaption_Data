def parse_dashboard(data, lc):

    rows = []

    contents = data.get("response", {}).get("data", {}).get("contents", [])

    for item in contents:

        rows.append([
            lc,
            item.get("dashboardEId"),
            item.get("name"),
            item.get("status"),
            item.get("visibility"),
            item.get("dashboardType"),
            item.get("createdBy"),
            item.get("creationTimestamp"),
            item.get("lastModifiedTimestamp"),
            item.get("cardCount")
        ])

    return rows