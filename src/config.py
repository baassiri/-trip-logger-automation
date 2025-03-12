import os
import gdown

# Define paths
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOCAL_FILE_PATH = os.path.join(BASE_DIR, "data", "INVOICE MANAGEMENT AUTO.xlsm")

# Google Drive File ID
FILE_ID = "1LXsBrrREmdBbZQVRmBv6QBu0ZOFu3oS3"

# Google Drive Download URL
GDRIVE_URL = f"https://drive.google.com/uc?id={FILE_ID}"

# Ensure 'data' folder exists
os.makedirs(os.path.join(BASE_DIR, "data"), exist_ok=True)

# Download if file doesn't exist
if not os.path.exists(LOCAL_FILE_PATH):
    print("📥 Downloading the Excel file from Google Drive...")
    gdown.download(GDRIVE_URL, LOCAL_FILE_PATH, quiet=False)
else:
    print("✅ Using existing Excel file.")

# Set the file path for other scripts
FILE_PATH = LOCAL_FILE_PATH
