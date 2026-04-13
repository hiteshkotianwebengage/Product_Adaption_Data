def parse_alert(data, lc):

    rows = []

    contents = data.get("response", {}).get("data", {}).get("contents", [])

    for item in contents:

        metric = item.get("monitoredMetricDetails", {}).get("metricName")

        subscribers = item.get("subscribers", [])
        subscribers_str = ", ".join(subscribers) if subscribers else ""

        threshold = item.get("comparisonThresholdValues", [])
        threshold_val = threshold[0] if threshold else ""

        rows.append([
            lc,
            item.get("id"),
            item.get("name"),
            metric,
            item.get("description"),
            item.get("alertFrequency"),
            threshold_val,
            item.get("comparisonOperator"),
            item.get("comparisonChangeType"),
            item.get("status"),
            item.get("createdBy"),
            subscribers_str,
            item.get("createdAt"),
            item.get("updatedAt"),
            item.get("lastEvaluationOn")
        ])

    return rows