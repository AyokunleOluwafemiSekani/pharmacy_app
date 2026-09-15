import json, os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(BASE_DIR, "kpi_settings.json")

def load_kpi_settings():
    if not os.path.exists(SETTINGS_FILE):
        return {}
    return json.load(open(SETTINGS_FILE))

def save_kpi_settings(data):
    json.dump(data, open(SETTINGS_FILE, "w"), indent=4)
