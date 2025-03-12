import os
import json
import gdown
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
import streamlit as st

# Define paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
FILE_PATH = os.path.join(DATA_DIR, "INVOICE_MANAGEMENT_AUTO.xlsm")
SERVICE_ACCOUNT_PATH = os.path.join(BASE_DIR, "service_account.json")
FILE_ID = "1LXsBrrREmdBbZQVRmBv6QBu0ZOFu3oS3"

# Ensure data directory exists
os.makedirs(DATA_DIR, exist_ok=True)

# Regenerate service_account.json if missing (for Streamlit Cloud)
if not os.path.exists(SERVICE_ACCOUNT_PATH):
    print("🔍 service_account.json not found. Attempting to regenerate...")
    if "service_account_json" in st.secrets:
        try:
            with open(SERVICE_ACCOUNT_PATH, "w") as f:
                json.dump(json.loads(st.secrets["service_account_json"]), f, indent=2)
            print("✅ service_account.json successfully created from Streamlit secrets.")
        except Exception as e:
            print(f"❌ Failed to create service_account.json: {e}")
    else:
        print("❌ 'service_account_json' not found in Streamlit secrets!")

def authenticate_drive():
    """Authenticate with Google Drive using a service account JSON."""
    if not os.path.exists(SERVICE_ACCOUNT_PATH):
        print("❌ Authentication failed: service_account.json not found!")
        return None

    gauth = GoogleAuth()
    
    try:
        gauth.LoadCredentialsFile(SERVICE_ACCOUNT_PATH)
        gauth.ServiceAuth()
        print("✅ Authenticated using Service Account.")
        return GoogleDrive(gauth)
    except Exception as e:
        print(f"❌ Authentication error: {e}")
        return None

def upload_to_drive():
    """Uploads the updated Excel file back to Google Drive."""
    drive = authenticate_drive()
    if drive is None:
        print("❌ Upload aborted: Authentication failed.")
        return

    if not os.path.exists(FILE_PATH):
        print("❌ Upload failed: Local file not found!")
        return

    # Debug: Print file size before upload
    file_size = os.path.getsize(FILE_PATH)
    print(f"📏 File size before upload: {file_size} bytes")

    try:
        print(f"📤 Uploading {FILE_PATH} to Google Drive...")
        
        # Overwrite existing file using FILE_ID
        file = drive.CreateFile({'id': FILE_ID})
        file.SetContentFile(FILE_PATH)
        file.Upload()
        print("✅ File successfully uploaded to Google Drive.")

        # Verify upload
        uploaded_file = drive.CreateFile({'id': FILE_ID})
        uploaded_file.FetchMetadata()
        print(f"✅ Google Drive metadata after upload: {uploaded_file}")

    except Exception as e:
        print(f"❌ Upload failed: {e}")

