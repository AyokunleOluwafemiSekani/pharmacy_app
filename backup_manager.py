import os
import shutil
import hashlib
from datetime import datetime
from cryptography.fernet import Fernet
from google_drive_sync import upload_backup_to_drive

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "pharmacy.db")

# Put this inside your OneDrive/Google Drive synced folder
BACKUP_DIR = os.path.join(BASE_DIR, "cloud_backups")

KEY_FILE = os.path.join(BASE_DIR, "backup_key.key")


def load_or_create_key():
    if os.path.exists(KEY_FILE):
        with open(KEY_FILE, "rb") as f:
            key = f.read()
    else:
        key = Fernet.generate_key()
        with open(KEY_FILE, "wb") as f:
            f.write(key)
    return key


def create_backup():
    os.makedirs(BACKUP_DIR, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    raw_backup_name = f"backup_{timestamp}.db"
    raw_backup_path = os.path.join(BACKUP_DIR, raw_backup_name)

    # 1. Copy DB
    shutil.copy2(DB_PATH, raw_backup_path)

    # 2. Encrypt
    key = load_or_create_key()
    fernet = Fernet(key)

    with open(raw_backup_path, "rb") as f:
        data = f.read()

    encrypted_data = fernet.encrypt(data)

    enc_backup_name = f"backup_{timestamp}.enc"
    enc_backup_path = os.path.join(BACKUP_DIR, enc_backup_name)

    with open(enc_backup_path, "wb") as f:
        f.write(encrypted_data)

    # 3. Integrity (SHA256)
    sha256 = hashlib.sha256(encrypted_data).hexdigest()
    hash_file = os.path.join(BACKUP_DIR, f"backup_{timestamp}.sha256")
    with open(hash_file, "w") as f:
        f.write(sha256)

    # 4. Remove raw .db copy (keep only encrypted)
    os.remove(raw_backup_path)

    print(f"Backup created: {enc_backup_name}")
    print(f"SHA256: {sha256}")

def verify_and_restore(enc_backup_path):
    key = load_or_create_key()
    fernet = Fernet(key)

    # Read encrypted data
    with open(enc_backup_path, "rb") as f:
        encrypted_data = f.read()

    # Verify SHA256
    base = os.path.basename(enc_backup_path).replace(".enc", "")
    hash_file = os.path.join(BACKUP_DIR, f"{base}.sha256")

    if not os.path.exists(hash_file):
        raise ValueError("Hash file missing for this backup.")

    with open(hash_file, "r") as f:
        stored_hash = f.read().strip()

    current_hash = hashlib.sha256(encrypted_data).hexdigest()

    if current_hash != stored_hash:
        raise ValueError("Integrity check failed. Backup file may be corrupted.")

    # Decrypt
    data = fernet.decrypt(encrypted_data)

    # Restore DB
    with open(DB_PATH, "wb") as f:
        f.write(data)

    print("Backup verified and restored successfully.")

def create_backup():
    os.makedirs(BACKUP_DIR, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    raw_backup = os.path.join(BACKUP_DIR, f"backup_{timestamp}.db")

    shutil.copy2(DB_PATH, raw_backup)

    key = load_or_create_key()
    fernet = Fernet(key)

    data = open(raw_backup, "rb").read()
    encrypted = fernet.encrypt(data)

    enc_file = os.path.join(BACKUP_DIR, f"backup_{timestamp}.enc")
    open(enc_file, "wb").write(encrypted)

    sha = hashlib.sha256(encrypted).hexdigest()
    open(os.path.join(BACKUP_DIR, f"backup_{timestamp}.sha256"), "w").write(sha)

    os.remove(raw_backup)

    # ⭐ Upload encrypted backup to Google Drive
    upload_backup_to_drive(f"backup_{timestamp}.enc")
def verify_backup_integrity(enc_file):
    """Verify SHA256 integrity of an encrypted backup."""
    sha_file = enc_file.replace(".enc", ".sha256")

    if not os.path.exists(sha_file):
        return False

    expected_sha = open(sha_file, "r").read().strip()
    actual_sha = hashlib.sha256(open(enc_file, "rb").read()).hexdigest()

    return expected_sha == actual_sha


def decrypt_backup(enc_file):
    """Decrypt an encrypted backup file and return raw DB bytes."""
    key = load_or_create_key()
    fernet = Fernet(key)

    encrypted_data = open(enc_file, "rb").read()
    return fernet.decrypt(encrypted_data)


def restore_backup(enc_file):
    """Restore the database from an encrypted backup."""
    if not verify_backup_integrity(enc_file):
        raise Exception("Backup integrity check failed.")

    raw_data = decrypt_backup(enc_file)

    # Write decrypted DB to a temporary file
    temp_db = os.path.join(BACKUP_DIR, "temp_restore.db")
    open(temp_db, "wb").write(raw_data)

    # Replace main DB
    shutil.copy2(temp_db, DB_PATH)
    os.remove(temp_db)

    return True
def get_local_backup_stats():
    backups = list_backups()
    if not backups:
        return None

    latest = sorted(backups, key=lambda x: x["date"], reverse=True)[0]

    return {
        "count": len(backups),
        "latest": latest["date"],
        "latest_name": latest["name"],
        "latest_size": latest["size"]
    }
def get_local_backups():
    backups = list_backups()
    formatted = []

    for b in backups:
        formatted.append({
            "source": "local",
            "name": b["name"],
            "date": b["date"],
            "size": b["size"],
            "id": None
        })

    return formatted
def scan_backup_integrity():
    """Scan all local backups and return integrity results."""
    backups = list_backups()
    results = []

    for b in backups:
        enc_file = os.path.join(BACKUP_DIR, b["name"])
        sha_file = enc_file.replace(".enc", ".sha256")

        # Check SHA file exists
        sha_exists = os.path.exists(sha_file)

        # Check SHA integrity
        if sha_exists:
            expected_sha = open(sha_file, "r").read().strip()
            actual_sha = hashlib.sha256(open(enc_file, "rb").read()).hexdigest()
            sha_valid = (expected_sha == actual_sha)
        else:
            sha_valid = False

        # Check encryption validity
        try:
            key = load_or_create_key()
            fernet = Fernet(key)
            fernet.decrypt(open(enc_file, "rb").read())
            encrypted_ok = True
        except Exception:
            encrypted_ok = False

        results.append({
            "name": b["name"],
            "date": b["date"],
            "size": b["size"],
            "sha_exists": sha_exists,
            "sha_valid": sha_valid,
            "encrypted_ok": encrypted_ok
        })

    return results
def get_scheduler_status():
    settings = load_settings()

    next_run = settings["time"]
    last_run = settings["last_run"]
    last_status = settings["last_status"]
    enabled = settings["enabled"]

    return {
        "enabled": enabled,
        "next_run": next_run,
        "last_run": last_run if last_run else "Never",
        "last_status": last_status if last_status else "N/A"
    }

if __name__ == "__main__":
    create_backup()
