def parse_overview(data, lc, start_label, end_label):
    rows = []

    # Check if data exists before accessing index [0]
    resp_data = data.get('response', {}).get('data', [])
    if not resp_data:
        return rows

    dimension = resp_data[0].get('dimensions', [])

    for item in dimension:
        if item['value'] == 'OVERALL':
            continue

        metrics = {m['name']: m['value'] for m in item['metrics']}

        users = metrics.get("users", 0)
        campaigns = metrics.get("campaigns", 0)
        deliveries = metrics.get("deliveries", 0)
        clicks = metrics.get("clicks", 0)
        revenue = metrics.get("revenue", 0)
        conversions = metrics.get("goal", 0)

        ctr = (clicks / deliveries) if deliveries else 0
        cvr = (conversions / deliveries) if deliveries else 0

        rows.append([
            lc,
            item['value'],
            users,
            campaigns,
            deliveries,
            round(ctr * 100, 2),
            round(cvr * 100, 2),
            revenue,
            start_label,
            end_label
        ])

    return rows