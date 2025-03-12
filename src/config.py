import os
import json
import gdown
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

try:
    import streamlit as st
except ImportError:
    st = None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
LOCAL_FILE_PATH = os.path.join(DATA_DIR, "INVOICE_MANAGEMENT_AUTO.xlsm")
CLIENT_SECRETS_PATH = os.path.join(BASE_DIR, "client_secrets.json")

# 1️⃣ If on Streamlit Cloud, create client_secrets.json from st.secrets
if st is not None:
    if not os.path.exists(CLIENT_SECRETS_PATH) and "client_secrets_json" in st.secrets:
        try:
            secret_value = st.secrets["client_secrets_json"]
            parsed = json.loads(secret_value)
            with open(CLIENT_SECRETS_PATH, "w") as f:
                json.dump(parsed, f, indent=2)
            print("✅ client_secrets.json created from Streamlit secrets.")
        except Exception as e:
            print(f"❌ Failed to create client_secrets.json from st.secrets: {e}")
    else:
        print("✅ client_secrets.json exists or no secrets provided.")
else:
    print("DEBUG: Running locally. Ensure client_secrets.json is present if needed.")

# 2️⃣ Ensure 'data' directory
os.makedirs(DATA_DIR, exist_ok=True)

# 3️⃣ Download the Excel file from Google Drive
print("📥 Checking for the latest Excel file from Google Drive...")
FILE_ID = "1LXsBrrREmdBbZQVRmBv6QBu0ZOFu3oS3"
GDRIVE_URL = f"https://drive.google.com/uc?id={FILE_ID}"
try:
    gdown.download(GDRIVE_URL, LOCAL_FILE_PATH, quiet=False, fuzzy=True)
    print(f"✅ File downloaded successfully: {LOCAL_FILE_PATH}")
except Exception as e:
    print(f"❌ Download failed: {e}")

FILE_PATH = LOCAL_FILE_PATH

def authenticate_drive():
    """Authenticate with Google Drive using CommandLineAuth on Streamlit Cloud, or LocalWebserverAuth locally."""
    gauth = GoogleAuth()

    # Point PyDrive2 to your client_secrets.json
    gauth.settings["client_config_file"] = CLIENT_SECRETS_PATH

    creds_path = os.path.join(BASE_DIR, "credentials.json")
    if os.path.exists(creds_path):
        gauth.LoadCredentialsFile(creds_path)

    if gauth.credentials is None:
        if st is not None:
            # On Streamlit Cloud → use CommandLineAuth
            print("🔑 Using CommandLineAuth (headless) on Streamlit Cloud.")
            gauth.CommandLineAuth()
        else:
            # Local dev → use LocalWebserverAuth
            print("🔑 First-time local auth. Opening browser on port 8080.")
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
