"""
Step 1 — Connection test script
Run: python3 test_connection.py
Tests Drive API connection and lists all folders in 02 Automation
"""
import sys, time

print("Testing Google Drive API connection...")
print("="*50)

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    print("✅ Google API libraries installed")
except ImportError as e:
    print(f"❌ Missing library: {e}")
    print("   Run: pip3 install google-api-python-client google-auth")
    sys.exit(1)

import os
CREDS_FILE = os.path.join(os.path.dirname(__file__), "credentials.json")
ROOT_ID    = "1C63nw8Vr9OPjc9WFNUd4Ziok_x8VqXyX"

# Step 1: credentials file
if not os.path.exists(CREDS_FILE):
    print(f"❌ credentials.json not found at: {CREDS_FILE}")
    sys.exit(1)
print(f"✅ credentials.json found")

# Step 2: authenticate
try:
    creds   = service_account.Credentials.from_service_account_file(
                CREDS_FILE, scopes=["https://www.googleapis.com/auth/drive"])
    service = build("drive", "v3", credentials=creds, cache_discovery=False)
    print(f"✅ Authenticated as: {creds.service_account_email}")
except Exception as e:
    print(f"❌ Authentication failed: {e}")
    sys.exit(1)

# Step 3: list root folder (lightweight call)
t0 = time.time()
try:
    resp = service.files().list(
        q=f"'{ROOT_ID}' in parents and trashed=false",
        fields="files(id, name, mimeType)",
        pageSize=20
    ).execute()
    elapsed = round(time.time() - t0, 2)
    files   = resp.get("files", [])
    print(f"✅ Drive API responding ({elapsed}s)")
    print(f"\nFolders found in 02 Automation ({len(files)}):")
    for f in files:
        icon = "📁" if "folder" in f["mimeType"] else "📄"
        print(f"  {icon} {f['name']}  (id: {f['id']})")
except Exception as e:
    print(f"❌ Drive API call failed: {e}")
    sys.exit(1)

print("\n" + "="*50)
print("✅ Connection OK — ready to run pipeline")
