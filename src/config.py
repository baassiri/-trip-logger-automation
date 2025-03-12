import os
import gdown

# Define paths
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
DATA_DIR = os.path.join(BASE_DIR, "data")
LOCAL_FILE_PATH = os.path.join(DATA_DIR, "INVOICE MANAGEMENT AUTO.xlsm")

# Google Drive File ID
FILE_ID = "1LXsBrrREmdBbZQVRmBv6QBu0ZOFu3oS3"
GDRIVE_URL = f"https://drive.google.com/uc?id={FILE_ID}"

# Ensure 'data' folder exists
os.makedirs(DATA_DIR, exist_ok=True)

# Download file only if missing
if not os.path.exists(LOCAL_FILE_PATH):
    print("📥 Downloading the Excel file from Google Drive...")
    try:
        gdown.download(GDRIVE_URL, LOCAL_FILE_PATH, quiet=False)
        print("✅ Download successful.")
    except Exception as e:
        print(f"❌ Download failed: {e}")
else:
    print("✅ Using existing Excel file.")

# Set global file path
FILE_PATH = LOCAL_FILE_PATH
