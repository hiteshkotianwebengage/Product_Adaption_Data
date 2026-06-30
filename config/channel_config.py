CHANNELS = [
    {
        "name": "PUSH",
        "endpoint": "push-notifications",
        "cg_channel": "PUSH_NOTIFICATION",
        "version": "v1",
        "params": {"noView": "true", "sdks": "2,3"}
    },
    {
        "name": "SMS",
        "endpoint": "sms-messages",
        "cg_channel": "SMS",
        "version": "v1",
        "params": {"noView": "true"}
    },
    {
        "name": "EMAIL",
        "endpoint": "emails",
        "cg_channel": "EMAIL",
        "version": "v1",
        "params": {"noView": "true"}
    },
    {
        "name": "WEB_PUSH",
        "endpoint": "web-push",
        "cg_channel": "WEB_PUSH",
        "version": "v1",
        "params": {"noView": "true"}
    },{
        "name": "RCS",
        "endpoint": "rcs-messages",
        "cg_channel": "RCS",
        "version": "v1",
        "params": {"noView": "true"}
    },
    {
        "name": "WHATSAPP",
        "endpoint": "whatsapp-messages",
        "cg_channel": "WHATSAPP",
        "version": "v1",
        "params": {"noView": "true"}
    }
]