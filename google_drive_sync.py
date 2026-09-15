import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BACKUP_DIR = os.path.join(BASE_DIR, "backups")

# Your Google Drive folder ID
DRIVE_FOLDER_ID = "1Lm9DcKvF3gliO4i63v0sE5RhJb0CjBXn"

# Path to your service account key
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, "google_drive_key.json")

# Authenticate with Google Drive
creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE,
    scopes=["https://www.googleapis.com/auth/drive"]
)

drive_service = build("drive", "v3", credentials=creds)


def upload_backup_to_drive(filename):
    """Uploads an encrypted backup file to Google Drive."""
    file_path = os.path.join(BACKUP_DIR, filename)

    if not os.path.exists(file_path):
        raise FileNotFoundError("Backup file not found: " + filename)

    file_metadata = {
        "name": filename,
        "parents": [DRIVE_FOLDER_ID]
    }

    media = MediaFileUpload(file_path, resumable=True)

    uploaded = drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id"
    ).execute()

    return uploaded.get("id")
def download_backup_from_drive(file_id, save_as):
    """Download a backup file from Google Drive."""
    request = drive_service.files().get_media(fileId=file_id)
    fh = open(save_as, "wb")

    downloader = MediaIoBaseDownload(fh, request)
    done = False

    while not done:
        status, done = downloader.next_chunk()

    fh.close()
    return save_as
from googleapiclient.errors import HttpError

def list_cloud_backups():
    """List encrypted backups stored in Google Drive."""
    try:
        results = drive_service.files().list(
            q=f"'{DRIVE_FOLDER_ID}' in parents",
            fields="files(id, name, size, modifiedTime)"
        ).execute()

        files = results.get("files", [])
        backups = []

        for f in files:
            if f["name"].endswith(".enc"):
                backups.append({
                    "id": f["id"],
                    "name": f["name"],
                    "size": round(int(f.get("size", 0)) / 1024 / 1024, 2),
                    "date": f["modifiedTime"].replace("T", " ").replace("Z", "")
                })

        return backups

    except HttpError:
        return None
def get_cloud_backups():
    cloud = list_cloud_backups()
    formatted = []

    if cloud:
        for b in cloud:
            formatted.append({
                "source": "cloud",
                "name": b["name"],
                "date": b["date"],
                "size": f"{b['size']} MB",
                "id": b["id"]
            })

    return formatted
def get_cloud_sync_status():
    """Return cloud sync health information."""
    try:
        backups = list_cloud_backups()

        return {
            "connected": True,
            "count": len(backups),
            "last_backup": backups[0]["date"] if backups else "None",
            "backups": backups
        }

    except Exception:
        return {
            "connected": False,
            "count": 0,
            "last_backup": "N/A",
            "backups": []
        }
