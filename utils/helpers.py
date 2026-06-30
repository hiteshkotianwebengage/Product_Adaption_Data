import time
import requests
from auth.cookies import get_session_cookies
from config.settings import PRE_BASE_URLS
from utils.logger import logger

# ----------------------
# Helper metrics 
# ----------------------

def has_valid_metrics(response_obj):
    if response_obj.get("message") == "No data":
        return False
    
    data_list = response_obj.get("data", [])
    if not data_list or not isinstance(data_list, list):
        return False
        
    dimensions_list = data_list[0].get("dimensions", [])
    if not dimensions_list or not isinstance(dimensions_list, list):
        return False

    metrics_list = dimensions_list[0].get("metrics", [])
    if not metrics_list or not isinstance(metrics_list, list):
        return False
        
    return True