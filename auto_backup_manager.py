import os
import json
from datetime import datetime
from backup_manager import create_backup

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(BASE_DIR, "auto_backup.json")


def load_settings():
    if not os.path.exists(SETTINGS_FILE):
        return {"enabled": False, "time": "01:00", "last_run": None, "last_status": None}
    return json.load(open(SETTINGS_FILE))


def save_settings(data):
    json.dump(data, open(SETTINGS_FILE, "w"), indent=4)


def run_auto_backup_if_due():
    settings = load_settings()

    if not settings["enabled"]:
        return

    now = datetime.now().strftime("%H:%M")

    if now == settings["time"]:
        try:
            create_backup()
            settings["last_run"] = datetime.now().strftime("%Y-%m-%d %H:%M")
            settings["last_status"] = "Success"
        except Exception as e:
            settings["last_run"] = datetime.now().strftime("%Y-%m-%d %H:%M")
            settings["last_status"] = "Failed"

        save_settings(settings)
