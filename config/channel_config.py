CHANNELS = [
    {
        "name": "PUSH",
        "endpoint": "push-notifications",
        "version": "v1",
        "params": {"noView": "true", "sdks": "2,3"}
    },
    {
        "name": "SMS",
        "endpoint": "sms-messages",
        "version": "v1",
        "params": {"noView": "true"}
    },
    {
        "name": "EMAIL",
        "endpoint": "emails",
        "version": "v1",
        "params": {"noView": "true"}
    },
    {
        "name": "WEB_PUSH",
        "endpoint": "web-push",
        "version": "v1",
        "params": {"noView": "true"}
    },{
        "name": "RCS",
        "endpoint": "rcs-messages",
        "version": "v1",
        "params": {"noView": "true"}
    },
    {
        "name": "WHATSAPP",
        "endpoint": "whatsapp-messages",
        "version": "v1",
        "params": {"noView": "true"}
    },
    {
        "name": "FACEBOOK",
        "endpoint": "fb-audiences",
        "version": "v2",
        "params": {}
    },
    {
        "name": "GOOGLE",
        "endpoint": "gAd-audiences",
        "version": "v2",
        "params": {}
    }
]

CHANNEL_HEADER = [
    "License",
    "Channel",
    "Campaign ID",
    "Campaign Name",
    "Status",
    "Category",
    "Created On",
    "Start Date",
    "Sent",
    "Delivered",
    "Clicks",
    "Conversions",
    "Revenue"
]