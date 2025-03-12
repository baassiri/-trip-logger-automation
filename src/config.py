import os
import requests
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

# Define paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
LOCAL_FILE_PATH = os.path.join(DATA_DIR, "INVOICE_MANAGEMENT_AUTO.xlsm")

# Google Drive File ID (Extracted from the shared link)
FILE_ID = "1LXsBrrREmdBbZQVRmBv6QBu0ZOFu3oS3"
GDRIVE_URL = f"https://drive.google.com/uc?export=download&id={FILE_ID}"

# Ensure 'data' directory exists
os.makedirs(DATA_DIR, exist_ok=True)

# **Download Excel file using requests (Better for Streamlit Cloud)**
def download_excel():
    print("📥 Checking for the latest Excel file from Google Drive...")
    try:
        response = requests.get(GDRIVE_URL, stream=True)
        if response.status_code == 200:
            with open(LOCAL_FILE_PATH, "wb") as f:
                for chunk in response.iter_content(chunk_size=1024):
                    f.write(chunk)
            print(f"✅ File downloaded successfully: {LOCAL_FILE_PATH}")
        else:
            print(f"❌ Download failed with status code: {response.status_code}")
    except Exception as e:
        print(f"❌ Error downloading file: {e}")

# Check if file exists, else download
if not os.path.exists(LOCAL_FILE_PATH):
    download_excel()
else:
    print("✅ Using existing Excel file.")

# Set file path for other scripts
FILE_PATH = LOCAL_FILE_PATH
