
# ------------------------------------------------------------------
# Constant
# ------------------------------------------------------------------

DEFAULT_NUMERIC = 0
DEFAULT_TEXT = "NA"

# ------------------------------------------------------------------
# Helper Functions
# ------------------------------------------------------------------

def safe_percent(num, den):
    if not den:
        return 0
    return round((num / den) * 100, 2)

def flatten_metrics(response_json):
    """
    Convert WebEngage metrics list into an easy-to-use dictionary.
    Example -------
    Input:
    [
        {"name": "sent", "groupFunction": "SUM", "value": 100},
        {"name": "sent", "groupFunction": "COUNT", "value": 98},
        {"name": "clicks", "groupFunction": "SUM", "value": 20},
    ]
    Output:
    {
        ("sent", "SUM"): 100,
        ("sent", "COUNT"): 98,
        ("clicks", "SUM"): 20,
    }
    """

    flattened = {}

    # Safely unpack WebEngage's deep payload nesting layout
    data_list = response_json.get("response", {}).get("data", [])
    if not isinstance(data_list, list) or not data_list:
        return flattened
    
    dimensions = data_list[0].get("dimensions", [])
    if not dimensions:
        return flattened
    
    metrics = dimensions[0].get("metrics", [])

    for metric in metrics: 
        name = metric.get("name")
        group_fn = metric.get("groupFunction", "SUM")
        category = metric.get("category")
        value = metric.get("value", DEFAULT_NUMERIC)

        # 1. Create readable string key
        # If it has a category namespace (like conversions_Impression-Through or revenue_Click-Through)
        if category:
            key_name = f"{category}_{name}"     # e.g - "failure_Uninstalled"
        else:
            key_name = name                     # e.g - "sent", "clicks"
        
        # 2. Append suffix only if its COUNT metric, keeping default SUM
        if group_fn == "COUNT":
            key_name = f"{key_name}_COUNT"

        flattened[key_name] = value
    return flattened

def parse_campaign_row(campaign, campaign_detail, journey, aggregates, lc, month_name):
    """
    Builds one report row for a campaign.
    - Parameters
    ----------
    campaign : dict - Campaign object returned from fetch_campaigns()

    campaign_detail : dict - Campaign configuration returned from
        fetch_campaign_detail()

    aggregates : dict - Overall campaign metrics returned from
        fetch_campaign_aggregates()

    aggregates_by_id : dict - Aggregate metrics split by variation/control group

    failure_metrics : dict - Failure-category aggregate metrics

    cg_stats : dict - Control-group statistics

    lc : str - License code

    month_name : str - Reporting month
    """

    # ------------------------------------------------------------------
    # 1. Campaign Metadata
    # ------------------------------------------------------------------

    def parse_campaign_metadata(campaign, campaign_detail):
        """
        Extract all non-metric campaign fields required in the final 59-column report.
        """
        # WebEngage packs the core campaign definition inside 'experiment'
        data_root = campaign_detail.get("response", {}).get("data", {}) if isinstance(campaign_detail, dict) else {}
        detail = data_root.get("experiment", {})
        conversion = data_root.get("conversion", {})
        scheduler = data_root.get("scheduler", {})
        stop = scheduler.get("stop", {})

        row = {}
        
        # 1. Base Campaign Information
        row["campaign_name"] = str(campaign.get("title") or "").strip() or DEFAULT_TEXT
        row["campaign_id"] = campaign.get("id", DEFAULT_TEXT)
        row["channel"] = campaign.get("productType", "UNKNOWN")
        row["campaign_type"] = campaign.get("category", DEFAULT_TEXT)
        
        # Priority on explicit master status
        row["status"] = campaign.get("sentStatus", {}).get("status") or campaign.get("status") or DEFAULT_TEXT

        # Safe join protection for tags
        tags = campaign.get("tags") or []
        row["campaign_tags"] = ", ".join(tags) if tags else DEFAULT_TEXT

        row["campaign_start_date"] = campaign.get("startDate", DEFAULT_TEXT)
        row["campaign_end_date"] = stop.get("startDate") or DEFAULT_TEXT
        row["created_by"] = detail.get("createdByUser") or "System"

        # 2. Journey Mapping (Using Section 2 Container Logic)
        is_journey = detail.get("container") == "JOURNEY"
        row["journey_id"] = detail.get("journeyId", DEFAULT_TEXT) if is_journey else DEFAULT_TEXT
        row["journey_name"] = DEFAULT_TEXT

        # 3. Segment Mapping
        row["segment_name"] = data_root.get("segmentName", DEFAULT_TEXT)
        row["segment_id"] = detail.get("segmentId", DEFAULT_TEXT)

        # 4. Conversion Tracking (Extracted from the root goal element block)
        triggers = conversion.get("triggerSet", {}).get("triggers", [])
        
        row["conversion_event"] = ", ".join(
            t.get("name", "") for t in triggers if t.get("name")
        ) or DEFAULT_TEXT

        # Clean the deadline display string cleanly
        # row["conversion_deadline"] = conversion.get("deadline", DEFAULT_TEXT)

        # 5. Control Group Properties
        cg = detail.get("controlGroup")
        row["control_group"] = f"Enabled ({cg}%)" if cg else "Disabled"

        return row

    # ------------------------------------------------------------------
    # 2. Journey Metadata
    # ------------------------------------------------------------------

    def parse_journey(response_json):
        """ Extract Journey metadata """
        if not isinstance(response_json, dict):
            return {}
        
        data = (
            response_json
            .get("response", {})
            .get("data", {})
            .get("contents", [])
        )

        if not isinstance(data, list) or not data:
            return {}
        
        journey = data[0]
        return {
            "journey_id": journey.get("journeyEId", DEFAULT_TEXT),
            "journey_name": journey.get("name", DEFAULT_TEXT),
        }

    # ------------------------------------------------------------------
    # 3. Aggregate Metrics
    # ------------------------------------------------------------------

    def parse_aggregate_metrics(metrics):
        """ Convert flattened aggregate metrics into report columns """

        row = {}

        # Delivery metrics
        row["sent"] = metrics.get("sent", DEFAULT_NUMERIC)
        row["failed"] = metrics.get("failures", DEFAULT_NUMERIC) or metrics.get("total_failures", DEFAULT_NUMERIC)
        row["delivered"] = metrics.get("delivery", DEFAULT_NUMERIC) or metrics.get("deliveries", DEFAULT_NUMERIC)
        # Impression metrics
        row["total_impressions"] = metrics.get("views", DEFAULT_NUMERIC)
        # Dismiss metrics
        row["unique_dismisses"] = metrics.get("close_COUNT", DEFAULT_NUMERIC) or metrics.get("close", DEFAULT_NUMERIC)
        row["total_dismisses"] = metrics.get("close", DEFAULT_NUMERIC)
        # Click metrics
        row["unique_clicks"] = metrics.get("clicks_COUNT", DEFAULT_NUMERIC) or metrics.get("clicks", DEFAULT_NUMERIC)
        row["total_clicks"] = metrics.get("clicks", DEFAULT_NUMERIC)
        # Conversion metrics
        row["unique_conversions"] = metrics.get("conversions_COUNT", DEFAULT_NUMERIC) or metrics.get("conversions", DEFAULT_NUMERIC)
        row["total_conversions"] = metrics.get("conversions", DEFAULT_NUMERIC)
        # Attribution conversions
        row["total_impression_through_conversions"] = (metrics.get("conversions_Impression-Through", DEFAULT_NUMERIC) or metrics.get("conversions_Impression-Through_COUNT", DEFAULT_NUMERIC))
        row["unique_click_through_conversions"] = (metrics.get("conversions_Click-Through_COUNT", DEFAULT_NUMERIC) or metrics.get("conversions_Click-Through", DEFAULT_NUMERIC))
        row["total_click_through_conversions"] = (metrics.get("conversions_Click-Through", DEFAULT_NUMERIC))
        # Attribution revenue
        row["impression_through_revenue"] = metrics.get("revenue_Impression-Through", DEFAULT_NUMERIC)
        row["click_through_revenue"] = metrics.get("revenue_Click-Through", DEFAULT_NUMERIC)
        # Revenue
        row["revenue"] = metrics.get("revenue", DEFAULT_NUMERIC)

        return row

    # ------------------------------------------------------------------
    # 4. Failure Breakdown
    # ------------------------------------------------------------------

    # For now we keep this out of our data 

    # def parse_failure_metrics(metrics):
    #     """ Extract delivery failure breakdown from the aggregate API (failure=true) """
    #     row = {}

    #     row["failed_uninstalled"] = metrics.get("failures_Uninstalled", DEFAULT_NUMERIC)
    #     row["failed_configuration_issue"] = metrics.get("failures_Configuration Issue", DEFAULT_NUMERIC)
    #     row["failed_dnd_queue_drop"] = metrics.get("failures_DND Queue Drop", DEFAULT_NUMERIC)
    #     row["failed_frequency_capping_queue_drop"] = metrics.get("failures_Frequency Capping Queue Drop", DEFAULT_NUMERIC)
    #     row["failed_personalization_error"] = metrics.get("failures_Personalization Error", DEFAULT_NUMERIC)
    #     row["failed_device_not_registered"] = metrics.get("failures_Device Not Registered", DEFAULT_NUMERIC)
    #     row["failed_channel_not_available"] = metrics.get("failures_Channel Not Available", DEFAULT_NUMERIC)
    #     row["failed_other_failures"] = sum([
    #         metrics.get("failures_Other Failures", 0),
    #         metrics.get("failures_Device Push Opted Out", 0),
    #         metrics.get("failures_Channel Opted Out", 0),
    #     ])

    #     return row
    
    # ------------------------------------------------------------------
    # 5. Control Group
    # ------------------------------------------------------------------

    # def parse_control_group_metrics(metrics):
    #     """ Extract control group statistics """
    #     row = {}

    #     row["total_in_control_group"] = metrics.get("control_group") or metrics.get("control_group_COUNT") or DEFAULT_NUMERIC
    #     # Conversions that happened within the Control Group (used for uplift calculation)        
    #     row["unique_control_group_conversions"] = metrics.get("conversions_Control Group") or metrics.get("conversions_Control Group_COUNT") or DEFAULT_NUMERIC

    #     return row

    # ------------------------------------------------------------------
    # 6. Rate Calculations
    # ------------------------------------------------------------------

    def calculate_rates(row):
        """ Calculate all percentage fields for the final 59-column report """
        rates = {}

        # Delivery & Failures
        rates["failed_rate"] = safe_percent(row.get("failed", DEFAULT_NUMERIC), row.get("sent", DEFAULT_NUMERIC))
        rates["delivered_rate"] = safe_percent(row.get("delivered", DEFAULT_NUMERIC), row.get("sent", DEFAULT_NUMERIC))

        # Impression / Open Rate
        rates["total_impression_rate"] = safe_percent(row.get("total_impressions", DEFAULT_NUMERIC), row.get("delivered", DEFAULT_NUMERIC))
        
        # Click Rates
        rates["unique_click_rate"] = safe_percent(row.get("unique_clicks", DEFAULT_NUMERIC), row.get("delivered", DEFAULT_NUMERIC))
        rates["total_click_rate"] = safe_percent(row.get("total_clicks", DEFAULT_NUMERIC), row.get("delivered", DEFAULT_NUMERIC))
        rates["total_click_to_impression_rate"] = safe_percent(row.get("total_clicks", DEFAULT_NUMERIC), row.get("total_impressions", DEFAULT_NUMERIC))

        # General Conversion Rates
        rates["unique_conversion_rate"] = safe_percent(row.get("unique_conversions", DEFAULT_NUMERIC), row.get("delivered", DEFAULT_NUMERIC))
        rates["total_conversion_rate"] = safe_percent(row.get("total_conversions", DEFAULT_NUMERIC), row.get("delivered", DEFAULT_NUMERIC))
        rates["total_conversion_to_impression_rate"] = safe_percent(row.get("total_conversions", DEFAULT_NUMERIC), row.get("total_impressions", DEFAULT_NUMERIC))
        rates["unique_conversion_to_click_rate"] = safe_percent(row.get("unique_conversions", DEFAULT_NUMERIC), row.get("unique_clicks", DEFAULT_NUMERIC))
        rates["total_conversion_to_click_rate"] = safe_percent(row.get("total_conversions", DEFAULT_NUMERIC), row.get("total_clicks", DEFAULT_NUMERIC))

        # Control Group Baseline Rate
        # --------------
        # comment
        # -------------- 
        # rates["unique_control_group_conversion_rate"] = safe_percent(row.get("unique_control_group_conversions", DEFAULT_NUMERIC), row.get("total_in_control_group", DEFAULT_NUMERIC))

        # Special Multi-Touch Attribution Conversion Rates
        rates["total_impression_through_conversion_rate"] = safe_percent(row.get("total_impression_through_conversions", DEFAULT_NUMERIC), row.get("total_impressions", DEFAULT_NUMERIC))
        rates["unique_click_through_conversion_rate"] = safe_percent(row.get("unique_click_through_conversions", DEFAULT_NUMERIC), row.get("unique_clicks", DEFAULT_NUMERIC))
        rates["total_click_through_conversion_rate"] = safe_percent(row.get("total_click_through_conversions", DEFAULT_NUMERIC), row.get("total_clicks", DEFAULT_NUMERIC))

        return rates

    # ------------------------------------------------------------------
    # 7. Build Final Row (Ordered List Matching CHANNEL_HEADER)
    # ------------------------------------------------------------------

    # Extract base metadata definitions
    meta_data = parse_campaign_metadata(campaign, campaign_detail)
    journey_data = parse_journey(journey)

    # Parse standard aggregate metrics blocks
    flat_aggregates = flatten_metrics(aggregates)
    metrics_data = parse_aggregate_metrics(flat_aggregates)

    # Build a combined internal dict for calculations
    combined = {}
    combined.update(meta_data)
    combined.update(journey_data)
    combined.update(metrics_data)
    
    # Calculate calculated ratios
    rates_data = calculate_rates(combined)
    combined.update(rates_data)

    # Merge Journey Name extracted from the lookup back to meta fields
    journey_name = journey_data.get("journey_name", DEFAULT_TEXT)
    # If parse_campaign_metadata hardcoded journey_name as "NA", overwrite it with verified lookups
    if journey_name != DEFAULT_TEXT:
        combined["journey_name"] = journey_name

    # Transform the dictionary into a sequential list mapping directly to CHANNEL_HEADER columns
    ordered_row = [
        # Account
        lc,                                                       # "License"

        # Campaign
        combined.get("channel", "UNKNOWN"),                        # "Channel"
        combined.get("campaign_id", DEFAULT_TEXT),                 # "Campaign ID"
        combined.get("campaign_name", DEFAULT_TEXT),               # "Campaign Name"
        combined.get("status", DEFAULT_TEXT),                      # "Campaign Status"
        combined.get("campaign_type", DEFAULT_TEXT),               # "Campaign Type"
        combined.get("created_by", "System"),                      # "Campaign Created On" (Mapped from createdByUser)
        combined.get("campaign_start_date", DEFAULT_TEXT),         # "Campaign Start Date"
        combined.get("campaign_tags", DEFAULT_TEXT),               # "Campaign Tags"

        # Journey
        combined.get("journey_id", DEFAULT_TEXT),                  # "Journey ID"
        combined.get("journey_name", DEFAULT_TEXT),                # "Journey Name"
        combined.get("status", DEFAULT_TEXT),                      # "Journey Status" (Reuses master campaign status if isolated is unavailable)
        combined.get("created_by", "System"),                      # "Journey Created By"

        # Conversion
        combined.get("conversion_event", DEFAULT_TEXT),            # "Conversion Event"
        # combined.get("conversion_deadline", DEFAULT_TEXT),         # "Conversion Deadline"
        combined.get("control_group", "Disabled"),                 # "Control Group"

        # Delivery
        combined.get("status", DEFAULT_TEXT),                      # "Delivery Status"
        combined.get("sent", DEFAULT_NUMERIC),                     # "Sent"
        combined.get("delivered", DEFAULT_NUMERIC),                # "Delivered"
        combined.get("total_impressions", DEFAULT_NUMERIC),        # "Views"
        combined.get("unique_clicks", DEFAULT_NUMERIC),            # "Clicks"

        # Performance
        combined.get("unique_conversions", DEFAULT_NUMERIC),       # "Conversions"
        combined.get("total_impression_through_conversions", DEFAULT_NUMERIC), # "View Through Conversions"
        combined.get("unique_click_through_conversions", DEFAULT_NUMERIC),      # "Click Through Conversions"

        combined.get("delivered_rate", DEFAULT_NUMERIC),           # "Delivered Rate"
        combined.get("unique_click_rate", DEFAULT_NUMERIC),        # "Click Rate"
        combined.get("unique_conversion_rate", DEFAULT_NUMERIC),   # "Conversion Rate"

        # Revenue
        combined.get("revenue", DEFAULT_NUMERIC),                  # "Revenue"
        combined.get("impression_through_revenue", DEFAULT_NUMERIC),# "View Through Revenue"
        combined.get("click_through_revenue", DEFAULT_NUMERIC),     # "Click Through Revenue"

        # Month
        month_name                                                 # "Month"
    ]

    return ordered_row