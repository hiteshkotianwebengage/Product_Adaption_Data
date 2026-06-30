# This is to get the journey data for the final channel data 

import requests

session = requests.Session()

def get_headers(lc):
    """
    Returns the common headers required by all WebEngage API requests.
    """
    return {
        "Accept": "application/json, text/plain, */*",
        "x-accounts-id": lc,
        "User-Agent": "Mozilla/5.0"
    }

def fetch_campaigns(lc, channel, region, cookies, base_urls, page_no=1,):
    """
    Fetches one page of campaigns for a specific channel
    Parameters
    ----------
    lc : str -> License Code
    channel : dict -> Entry from CHANNELS config
    region : str -> Region code
    cookies : dict -> Active WebEngage session cookies.
    base_urls : dict -> Regional dashboard URLs
    page_no : int -> Campaign page number.
    """
    endpoint = channel["endpoint"]
    version = channel["version"]

    url = (
        f"{base_urls[region]}/api/{version}/accounts/{lc}/{endpoint}"
    )

    params = {
        "pageNo": page_no,
        "pageSize": 100,
        **channel.get("params", {}),
        "tags": "",
        "propertyNames": "",
        "properties": ""
    }

    headers = get_headers(lc)

    response = session.get(url, headers=headers, params=params, cookies=cookies, timeout=60,)
    return response

def fetch_campaign_detail(lc, channel, campaign_id, region, cookies, base_urls,):
    """ Fetch complete campaign definition
    contains: journeyId, goalId, createdByUser, controlGroup, variations, campaign configuration
    """
    endpoint = channel["endpoint"]
    version = "v2"

    url = (
        f"{base_urls[region]}/api/{version}/accounts/{lc}/{endpoint}/{campaign_id}"
    )

    params = {
        "fetchAllVar": "true"
    }

    headers = get_headers(lc)

    response = session.get(url, headers=headers, params=params, cookies=cookies, timeout=60)
    return response

def fetch_campaign_aggregates(lc, channel, campaign_id, metrics_from, metrics_to, region, cookies, base_urls, dimensions=None, failure=False,):
    """
    Fetch campaign metrics for the requested date window.
    """

    endpoint = channel["endpoint"]

    url = (
        f"{base_urls[region]}/api/v3/accounts/{lc}/{endpoint}/{campaign_id}/aggregates"
    )

    params = {
        "metrics": "all",
        "from": metrics_from,
        "to": metrics_to,
        "isFunnelView": "false",
        "flag": "true"
    }

    if dimensions:
        params["dimensions"] = dimensions

    if failure:
        params["failure"] = "true"

    headers = get_headers(lc)
    
    response = session.get(url, headers=headers, params=params, cookies=cookies, timeout=60)
    return response

def fetch_campaign_cg_stats(lc, campaign_id, channel, metrics_from, metrics_to, region, cookies, base_urls,):
    """
    Fetch control-group statistics for a campaign.
    """

    url = (
        f"{base_urls[region]}/api/v2/accounts/{lc}/cg-stats"
    )

    params = {
        "fromDate": metrics_from,
        "toDate": metrics_to,
        "experimentEncodedId": campaign_id,
        "campaignChannel": channel["cg_channel"],
    }

    headers = get_headers(lc)

    response = session.get(url, headers=headers, params=params, cookies=cookies, timeout=60)
    return response

def fetch_journey(lc, journey_id, region, cookies, base_urls,):
    """ Fetch journey metadata for a single journey
    Used to obtain journey name and other journey level information
    """

    url = (f"{base_urls[region]}/api/v1/accounts/{lc}/journeys")

    params = {
        "journeyEId": journey_id,
        "journeyType": "JOURNEY",
    }

    headers = get_headers(lc)
    
    response = session.get(url, headers=headers, params=params, cookies=cookies, timeout=60)
    return response