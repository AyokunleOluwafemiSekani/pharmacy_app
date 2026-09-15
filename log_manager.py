import json, os
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(BASE_DIR, "admin_logs.json")

def log_event(event_type, message, ip="system"):
    logs = []

    if os.path.exists(LOG_FILE):
        logs = json.load(open(LOG_FILE))

    logs.append({
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "type": event_type,
        "message": message,
        "ip": ip
    })

    json.dump(logs, open(LOG_FILE, "w"), indent=4)

def get_logs():
    if not os.path.exists(LOG_FILE):
        return []
    return json.load(open(LOG_FILE))
