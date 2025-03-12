import os
import gdown
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

# Define paths
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")
LOCAL_FILE_PATH = os.path.join(DATA_DIR, "INVOICE MANAGEMENT AUTO.xlsm")

# Google Drive File ID (Extracted from shared link)
FILE_ID = "1LXsBrrREmdBbZQVRmBv6QBu0ZOFu3oS3"
GDRIVE_URL = f"https://drive.google.com/uc?id={FILE_ID}"

# Ensure 'data' directory exists
os.makedirs(DATA_DIR, exist_ok=True)

# Download file if it doesn't exist
if not os.path.exists(LOCAL_FILE_PATH):
    print("📥 Downloading Excel file from Google Drive...")
    try:
        gdown.download(GDRIVE_URL, LOCAL_FILE_PATH, quiet=False)
        print("✅ Download successful.")
    except Exception as e:
        print(f"❌ Download failed: {e}")
else:
    print("✅ Using existing Excel file.")

# Set file path for other scripts
FILE_PATH = LOCAL_FILE_PATH

# 🔹 Persistent Google Drive authentication
def authenticate_drive():
    """Authenticate with Google Drive and reuse credentials to prevent login prompts."""
    gauth = GoogleAuth()

    # Use existing credentials if available
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

    gauth.SaveCredentialsFile(creds_path)  # Save credentials
    return GoogleDrive(gauth)

# 🔹 Upload updated file back to Google Drive
def upload_to_drive():
    """Uploads the updated Excel file back to Google Drive."""
    drive = authenticate_drive()
    
    try:
        # Check if file exists on Drive
        file_list = drive.ListFile({'q': "title='INVOICE MANAGEMENT AUTO.xlsm'"}).GetList()
        
        if file_list:
            file = file_list[0]  # Update existing file
            file.SetContentFile(FILE_PATH)
            file.Upload()
            print("✅ Updated Excel file uploaded to Google Drive!")
        else:
            # Upload as new file
            file = drive.CreateFile({'title': "INVOICE MANAGEMENT AUTO.xlsm"})
            file.SetContentFile(FILE_PATH)
            file.Upload()
            print("✅ New Excel file uploaded to Google Drive!")

    except Exception as e:
        print(f"❌ Error during file upload: {e}")
