from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import os

# Initialize GoogleAuth
gauth = GoogleAuth()

# Try to load saved client credentials
if os.path.exists("credentials.json"):
    gauth.LoadCredentialsFile("credentials.json")

if gauth.credentials is None:
    # Authenticate if credentials are not available
    gauth.LocalWebserverAuth()
elif gauth.access_token_expired:
    # Refresh token if expired
    gauth.Refresh()
else:
    # Authorize with existing credentials
    gauth.Authorize()

# Save credentials to a file for future use
gauth.SaveCredentialsFile("credentials.json")

# Initialize Google Drive instance
drive = GoogleDrive(gauth)

print("✅ Authentication successful! Credentials saved.")
