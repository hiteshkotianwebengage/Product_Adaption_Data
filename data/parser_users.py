def parse_users(data, lc):

    rows = {
        "MAU" : [],
        "DAU" : [],
        "WAU" : []
    }

    contents = data.get("response", {}).get("data", {})

    for metric in ["MAU", "DAU", "WAU"]:

        metric_data = contents.get(metric, [])

        for platform_obj in metric_data:

            platform = platform_obj.get("name")

            for stat in platform_obj.get("stats", []):
                for date, value in stat.items():

                    rows[metric].append([
                        lc,
                        platform,
                        date,
                        value
                    ])

    return rows