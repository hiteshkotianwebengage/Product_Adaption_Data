def parse_revenue_mapping(data, lc):

    rows = []

    mappings = data.get("response", {}).get("data", {}).get("revenueMappings", [])

    for item in mappings:

        rows.append([
            lc,
            item.get("id"),
            item.get("eventName"),
            item.get("attributeName"),
            item.get("active")
        ])

    return rows