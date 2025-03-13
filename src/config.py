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

# Download the file from Google Drive if missing
def download_from_drive():
    if not os.path.exists(FILE_PATH):
        print("⚠️ File missing! Downloading from Google Drive...")
        GDRIVE_URL = f"https://drive.google.com/uc?id={FILE_ID}"
        try:
            gdown.download(GDRIVE_URL, FILE_PATH, quiet=False, fuzzy=True)
            print(f"✅ File downloaded successfully: {FILE_PATH}")
        except Exception as e:
            print(f"❌ Download failed: {e}")

# Ensure service account JSON exists for authentication
def setup_service_account():
    if st and not os.path.exists(SERVICE_ACCOUNT_PATH):
        if "service_account_json" in st.secrets:
            try:
                with open(SERVICE_ACCOUNT_PATH, "w") as f:
                    json.dump(json.loads(st.secrets["service_account_json"]), f, indent=2)
                print("✅ service_account.json created from Streamlit secrets.")
            except Exception as e:
                print(f"❌ Failed to create service_account.json: {e}")
        else:
            print("❌ 'service_account_json' not found in Streamlit secrets!")

def authenticate_drive():
    """Authenticate with Google Drive using a service account JSON."""
    gauth = GoogleAuth()
    
    gauth.settings["client_config_backend"] = "service"
    gauth.settings["service_config"] = {
        "client_json_file_path": SERVICE_ACCOUNT_PATH,
        "client_user_email": "streamlit-service-account@drive-api-453511.iam.gserviceaccount.com"  # Add your service account email
    }
    
    try:
        gauth.ServiceAuth()
        print("✅ Authenticated using Service Account.")
    except Exception as e:
        print(f"❌ Authentication failed: {e}")

    return GoogleDrive(gauth)


def upload_to_drive():
    """ Uploads the updated Excel file to Google Drive """
    
    drive = authenticate_drive()
    if drive is None:
        print("❌ Upload aborted: Authentication failed.")
        return

    if not os.path.exists(FILE_PATH):
        print("❌ Upload failed: Local file not found!")
        return

    file_size = os.path.getsize(FILE_PATH)
    print(f"📏 File size before upload: {file_size} bytes")

    try:
        print(f"📤 Attempting to overwrite {FILE_PATH} on Google Drive...")
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



# Ensure file and service account setup before proceeding
setup_service_account()
download_from_drive()
