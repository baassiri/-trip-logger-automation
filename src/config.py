import os
import json
import gdown
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
FILE_PATH = os.path.join(DATA_DIR, "INVOICE_MANAGEMENT_AUTO.xlsm")
SERVICE_ACCOUNT_PATH = os.path.join(BASE_DIR, "service_account.json")
FILE_ID = "1LXsBrrREmdBbZQVRmBv6QBu0ZOFu3oS3"

def authenticate_drive():
    """Authenticate with Google Drive using a service account JSON."""
    gauth = GoogleAuth()
    gauth.LoadCredentialsFile(SERVICE_ACCOUNT_PATH)

    try:
        gauth.ServiceAuth()
        print("✅ Authenticated using Service Account.")
    except Exception as e:
        print(f"❌ Authentication failed: {e}")

    return GoogleDrive(gauth)

def upload_to_drive():
    """Uploads the updated Excel file back to Google Drive."""
    drive = authenticate_drive()

    if not os.path.exists(FILE_PATH):
        print("❌ Upload failed: Local file not found!")
        return

    # Debug: Print file size before upload
    print(f"📏 File size before upload: {os.path.getsize(FILE_PATH)} bytes")

    try:
        print(f"📤 Attempting to overwrite {FILE_PATH} on Google Drive...")
        file = drive.CreateFile({'id': FILE_ID})
        file.SetContentFile(FILE_PATH)
        file.Upload()
        print("✅ File uploaded successfully!")

        # Fetch and verify the update
        uploaded_file = drive.CreateFile({'id': FILE_ID})
        uploaded_file.FetchMetadata()
        print(f"✅ Google Drive metadata after upload: {uploaded_file}")

    except Exception as e:
        print(f"❌ Upload failed: {e}")
