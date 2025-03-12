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
SERVICE_ACCOUNT_PATH = os.path.join(BASE_DIR, "service_account.json")

print("DEBUG: BASE_DIR =", BASE_DIR)

# If running on Streamlit Cloud, create service_account.json from st.secrets if it doesn't exist.
if st is not None:
    if not os.path.exists(SERVICE_ACCOUNT_PATH):
        print("DEBUG: service_account.json not found locally.")
        if "service_account_json" in st.secrets:
            try:
                sa_json = st.secrets["service_account_json"]
                print("DEBUG: Found st.secrets['service_account_json']: ", sa_json[:100], "...")
                parsed = json.loads(sa_json)
                with open(SERVICE_ACCOUNT_PATH, "w") as f:
                    json.dump(parsed, f, indent=2)
                print("✅ service_account.json created successfully from Streamlit secrets.")
            except Exception as e:
                print(f"❌ Failed to create service_account.json from st.secrets: {e}")
        else:
            print("❌ 'service_account_json' not found in st.secrets.")
    else:
        print("✅ service_account.json already exists.")
else:
    print("DEBUG: Running locally. Ensure service_account.json is present if needed.")

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

FILE_PATH = LOCAL_FILE_PATH

def authenticate_drive():
    """Authenticate with Google Drive using a service account JSON."""
    gauth = GoogleAuth()
    # Configure PyDrive2 for service account authentication and add the required key.
    gauth.settings["client_config_backend"] = "service"
    gauth.settings["service_config"] = {
        "client_json_file_path": SERVICE_ACCOUNT_PATH,
        "client_user_email": ""  # leave empty if not impersonating
    }
    creds_path = os.path.join(BASE_DIR, "credentials.json")
    if os.path.exists(creds_path):
        gauth.LoadCredentialsFile(creds_path)
    try:
        gauth.ServiceAuth()  # Uses the service account file; no interactive prompt.
        print("✅ Authenticated using Service Account.")
    except Exception as e:
        print(f"❌ Service account authentication failed: {e}")
    gauth.SaveCredentialsFile(creds_path)
    return GoogleDrive(gauth)

def upload_to_drive():
    drive = authenticate_drive()
    try:
        print(f"📤 Uploading {FILE_PATH} to Google Drive...")
        file = drive.CreateFile({'id': '1LXsBrrREmdBbZQVRmBv6QBu0ZOFu3oS3'})
        file.SetContentFile(FILE_PATH)
        file.Upload()
        print("✅ Overwrote existing XLSM in personal drive!")
        
        # Confirm if file is correctly uploaded
        uploaded_file = drive.CreateFile({'id': '1LXsBrrREmdBbZQVRmBv6QBu0ZOFu3oS3'})
        uploaded_file.FetchMetadata()
        print(f"🔍 File metadata after upload: {uploaded_file}")

    except Exception as e:
        print(f"❌ Error during file upload: {e}")
