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

# If on Streamlit, create service_account.json from secrets
if st is not None:
    if not os.path.exists(SERVICE_ACCOUNT_PATH) and "service_account_json" in st.secrets:
        try:
            sa_json = st.secrets["service_account_json"]
            parsed = json.loads(sa_json)
            with open(SERVICE_ACCOUNT_PATH, "w") as f:
                json.dump(parsed, f)
            print("✅ service_account.json created from Streamlit secrets.")
        except Exception as e:
            print(f"❌ Failed to create service_account.json: {e}")
    else:
        print("✅ service_account.json already exists or no secrets provided.")

os.makedirs(DATA_DIR, exist_ok=True)

# Download Excel
FILE_ID = "1LXsBrrREmdBbZQVRmBv6QBu0ZOFu3oS3"
# Fixed line: properly closed the f-string
GDRIVE_URL = f"https://drive.google.com/uc?id={FILE_ID}"

# Download the Excel file
print("📥 Checking for the latest Excel file from Google Drive...")
try:
    gdown.download(GDRIVE_URL, LOCAL_FILE_PATH, quiet=False, fuzzy=True)
    print(f"✅ File downloaded successfully: {LOCAL_FILE_PATH}")
except Exception as e:
    print(f"❌ Download failed: {e}")

FILE_PATH = LOCAL_FILE_PATH

def authenticate_drive():
    """Authenticate with Google Drive using a service account JSON."""
    gauth = GoogleAuth()

    # Configure PyDrive2 to use 'service' instead of 'installed'
    gauth.settings["client_config_backend"] = "service"
    gauth.settings["service_config"] = {
        "client_json_file_path": SERVICE_ACCOUNT_PATH
    }

    gauth.ServiceAuth()  # No interactive prompt
    print("✅ Authenticated using Service Account.")
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
