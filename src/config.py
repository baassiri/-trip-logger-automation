import os
import json
import gdown
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

# Try to import streamlit for secrets management
try:
    import streamlit as st
except ImportError:
    st = None

# Define paths (assuming this file is in src/)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
LOCAL_FILE_PATH = os.path.join(DATA_DIR, "INVOICE_MANAGEMENT_AUTO.xlsm")
CLIENT_SECRETS_PATH = os.path.join(BASE_DIR, "client_secrets.json")

# On Streamlit Cloud, if client_secrets.json is missing, create it from st.secrets.
if st is not None:
    if not os.path.exists(CLIENT_SECRETS_PATH) and "client_secrets_json" in st.secrets:
        try:
            with open(CLIENT_SECRETS_PATH, "w") as f:
                f.write(st.secrets["client_secrets_json"])
            print("✅ client_secrets.json created from Streamlit secrets.")
        except Exception as e:
            print(f"⚠️ Could not create client_secrets.json from secrets: {e}")
    else:
        if os.path.exists(CLIENT_SECRETS_PATH):
            print("✅ client_secrets.json already exists.")
        else:
            print("⚠️ No client_secrets_json found in st.secrets.")

# Ensure 'data' directory exists
os.makedirs(DATA_DIR, exist_ok=True)

# Download the Excel file from Google Drive
print("📥 Checking for the latest Excel file from Google Drive...")
FILE_ID = "1LXsBrrREmdBbZQVRmBv6QBu0ZOFu3oS3"
GDRIVE_URL = f"https://drive.google.com/uc?id={FILE_ID}"
try:
    gdown.download(GDRIVE_URL, LOCAL_FILE_PATH, quiet=False, fuzzy=True)
    print(f"✅ File downloaded successfully: {LOCAL_FILE_PATH}")
except Exception as e:
    print(f"❌ Download failed: {e}")

# Set file path for other scripts
FILE_PATH = LOCAL_FILE_PATH

def authenticate_drive():
    """Authenticate with Google Drive and reuse credentials to prevent login prompts."""
    gauth = GoogleAuth()
    creds_path = os.path.join(BASE_DIR, "credentials.json")
    if os.path.exists(creds_path):
        gauth.LoadCredentialsFile(creds_path)
    if gauth.credentials is None:
        print("🔑 First-time authentication required. Opening browser...")
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        print("🔄 Refreshing expired token...")
        gauth.Refresh()
    else:
        print("✅ Using existing Google authentication.")
    gauth.SaveCredentialsFile(creds_path)
    return GoogleDrive(gauth)

def upload_to_drive():
    """Uploads the updated Excel file back to Google Drive."""
    drive = authenticate_drive()
    try:
        file_list = drive.ListFile({'q': "title='INVOICE_MANAGEMENT_AUTO.xlsm'"}).GetList()
        if file_list:
            file = file_list[0]
            file.SetContentFile(FILE_PATH)
            file.Upload()
            print("✅ Updated Excel file uploaded to Google Drive!")
        else:
            file = drive.CreateFile({'title': "INVOICE_MANAGEMENT_AUTO.xlsm"})
            file.SetContentFile(FILE_PATH)
            file.Upload()
            print("✅ New Excel file uploaded to Google Drive!")
    except Exception as e:
        print(f"❌ Error during file upload: {e}")
