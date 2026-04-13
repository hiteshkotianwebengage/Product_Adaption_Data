from scipy import stats


def parse_overview(data, lc, start_label, end_label):
    rows = []

    # Check if data exists before accessing index [0]
    resp_data = data.get('response', {}).get('data', [])
    if not resp_data:
        return rows

    dimension = resp_data[0].get('dimensions', [])

    for item in dimension:
        if item.get('value') == 'OVERALL':
            continue

        metrics = {m['name']: m['value'] for m in item['metrics']}

        users = metrics.get("users", 0)
        campaigns = metrics.get("campaigns", 0)
        deliveries = metrics.get("deliveries", 0)
        clicks = metrics.get("clicks", 0)

        click_revenue = metrics.get("click_revenue", 0)
        click_conversions = metrics.get("click_conversions", 0)

        revenue = metrics.get("revenue", 0)
        conversions = metrics.get("goal", 0)

        ctr = (clicks / deliveries) if deliveries else 0
        cvr = (click_conversions / clicks) if clicks else 0

        rows.append([
            lc,
            item['value'],
            users,
            campaigns,
            deliveries,
            round(ctr * 100, 2),
            round(cvr * 100, 2),
            click_revenue,
            click_conversions,
            revenue,
            start_label,
            end_label
        ])

    return rows