import os
import json
from datetime import datetime
from threading import Lock
from utils.logger import logger

PROGRESS_FILE = os.path.join("Progress_File", "progress_channels.json")
progress_lock = Lock()

# ----------------------
# PROGRESS HELPERS
# ----------------------

def load_progress():
    if os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, "r") as f:
                return json.load(f) or {}
        except:
            return {}
    return {}

def save_progress_raw(data):
    os.makedirs(os.path.dirname(PROGRESS_FILE), exist_ok=True)
    with open(PROGRESS_FILE, "w") as f:
        json.dump(data, f, indent=4)

# ----------------------
# Helper to save the page checkpoints
# ----------------------

def save_page_checkpoint(region, lc, channel_name, page_number, completed=False):
    """Safely updates the local progress configuration file with precise page tracking checkpoints"""
    try:
        data = load_progress()

        if region not in data:
            data[region] = {}
        if lc not in data[region]:
            data[region][lc] = {"status": "PROCESSING", "channels": {}}
        
        if "channels" not in data[region][lc]:
            data[region][lc]["channels"] = {}

        data[region][lc]["channels"][channel_name] = {
            "last_successful_page": page_number,
            "completed": completed
        }
        data[region][lc]["updated_at"] = datetime.now().isoformat()

        # Update global account status if all tracks complete
        all_done = all(ch.get("completed", False) for ch in data[region][lc]["channels"].values())
        if all_done and completed:
            data[region][lc]["status"] = "SUCCESS"

        save_progress_raw(data)
            
    except Exception as e:
        logger.error(f"❌ Failed to write state checkpoint to progress file: {e}")