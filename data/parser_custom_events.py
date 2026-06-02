def parse_custom_event_metadata(
    data,
    lc,
    event_name
):

    rows = []

    resp_data = (
        data.get("response", {})
            .get("data", {})
    )

    mappings = resp_data.get("mappings", [])

    # ---------------------------------------------------------
    # NO ATTRIBUTE
    # ---------------------------------------------------------

    if not mappings:

        rows.append([

            str(lc),
            str(event_name).strip(),
            "--NO_ATTRIBUTE--",
            "ATTRIBUTE",
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            "NO_DATA",
            "NO_DATA"
        ])

        return rows

    # ---------------------------------------------------------
    # ATTRIBUTE LEVEL
    # ---------------------------------------------------------

    for item in mappings:

        status_list = item.get("status") or []

        statuses = {
            s.get("sdk"): s.get("value")
            for s in status_list
            if isinstance(s, dict)
        }

        row = [

            str(lc),
            str(event_name).strip(),
            str(item.get("name", "NO_DATA")),
            "ATTRIBUTE",
            str(item.get("type", "NO_DATA")),
            str(item.get("ignored", "NO_DATA")),

            statuses.get("web", "NO_DATA"),
            statuses.get("android", "NO_DATA"),
            statuses.get("ios", "NO_DATA"),
            statuses.get("none", "NO_DATA")
        ]

        clean_row = [
            "" if v is None else str(v)
            for v in row
        ]

        rows.append(clean_row)

    return rows