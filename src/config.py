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
    
    # Ensure credentials exist
    if not os.path.exists(SERVICE_ACCOUNT_PATH):
        print("❌ Service account credentials not found.")
        return None

    try:
        gauth.LoadCredentialsFile(SERVICE_ACCOUNT_PATH)
        gauth.ServiceAuth()
        print("✅ Authenticated using Service Account.")
    except Exception as e:
        print(f"❌ Authentication failed: {e}")
        return None

    return GoogleDrive(gauth)

def upload_to_drive():
    """Uploads the updated Excel file back to Google Drive, ensuring it is overwritten."""
    drive = authenticate_drive()
    
    if not drive:
        print("❌ Upload aborted: Authentication failed.")
        return

    if not os.path.exists(FILE_PATH):
        print("❌ Upload failed: Local file not found!")
        return

    # Debug: Print file size before upload
    file_size = os.path.getsize(FILE_PATH)
    print(f"📏 File size before upload: {file_size} bytes")

    try:
        print(f"📤 Overwriting {FILE_PATH} on Google Drive...")

        # Step 1: Find the file on Google Drive
        file_list = drive.ListFile({'q': f"'{FILE_ID}' in parents"}).GetList()
        if file_list:
            print("📌 File exists. Overwriting it...")
            file = drive.CreateFile({'id': FILE_ID})  # Overwrite existing file
        else:
            print("📌 File does not exist. Creating a new one...")
            file = drive.CreateFile({'title': "INVOICE_MANAGEMENT_AUTO.xlsm"})

        # Step 2: Upload the updated file
        file.SetContentFile(FILE_PATH)
        file.Upload()

        # Step 3: Fetch metadata to verify the update
        uploaded_file = drive.CreateFile({'id': FILE_ID})
        uploaded_file.FetchMetadata()

        print(f"✅ File uploaded successfully! Metadata: {uploaded_file}")

    except Exception as e:
        print(f"❌ Upload failed: {e}")
