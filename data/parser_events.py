def parse_events(data, lc):

    rows = []

    contents = data.get("response", {}).get("data", {}).get("contents", [])

    for item in contents:

        # ----------------------
        # SDK STATUS PARSING
        # ----------------------
        web_status = ""
        android_status = ""
        ios_status = ""
        last_received = ""

        for s in item.get("status", []):
            if s.get("sdk") == "web":
                web_status = s.get("value")
                last_received = s.get("lastReceived")
            elif s.get("sdk") == "android":
                android_status = s.get("value")
            elif s.get("sdk") == "ios":
                ios_status = s.get("value")

        # ----------------------
        # MAPPING USAGE PARSING
        # ----------------------
        usage_map = {
            "STRING": 0,
            "INTEGER": 0,
            "BOOLEAN": 0,
            "DATE": 0
        }

        for m in item.get("mappingUsages", []):
            usage_map[m.get("dataType")] = m.get("usage", 0)

        rows.append([
            lc,
            item.get("name"),
            item.get("displayName"),
            item.get("category"),
            item.get("ignored"),
            item.get("personalizationStatus"),
            web_status,
            android_status,
            ios_status,
            last_received,
            usage_map["STRING"],
            usage_map["INTEGER"],
            usage_map["BOOLEAN"],
            usage_map["DATE"]
        ])

    return rows