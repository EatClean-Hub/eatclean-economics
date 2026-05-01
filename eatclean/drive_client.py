"""
Google Drive API client using service account credentials.
Handles: list files, download to temp, upload results.
All list/API operations have retry logic with exponential backoff.
"""

import os
import io
import time
import tempfile
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from google.oauth2 import service_account

SCOPES           = ["https://www.googleapis.com/auth/drive"]
CREDENTIALS_FILE = os.path.join(os.path.dirname(__file__), "credentials.json")

FOLDER_IDS = {
    "root":             "1C63nw8Vr9OPjc9WFNUd4Ziok_x8VqXyX",
    "daily_selections": None,
    "invoices":         None,
    "reference":        None,
    "outputs":          None,
}


def get_service():
    """Authenticate and return Drive API service."""
    creds = service_account.Credentials.from_service_account_file(
        CREDENTIALS_FILE, scopes=SCOPES
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def _execute(request, retries=4):
    """Execute a Drive API request with retry on any error."""
    for attempt in range(retries):
        try:
            return request.execute()
        except Exception as e:
            if attempt < retries - 1:
                wait = 2 ** attempt  # 1s, 2s, 4s
                time.sleep(wait)
            else:
                raise


def list_folder(service, folder_id: str) -> list:
    """List all files (not subfolders) in a Drive folder."""
    results = []
    page_token = None
    while True:
        req = service.files().list(
            q=f"'{folder_id}' in parents and mimeType!='application/vnd.google-apps.folder' and trashed=false",
            fields="nextPageToken, files(id, name)",
            pageToken=page_token,
            pageSize=100,
        )
        resp = _execute(req)
        results.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return results


def list_subfolders(service, folder_id: str) -> list:
    """List all subfolders inside a Drive folder."""
    req = service.files().list(
        q=f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
        fields="files(id, name)",
        pageSize=50,
    )
    resp = _execute(req)
    return resp.get("files", [])


def download_file(service, file_id: str, dest_path: str, retries: int = 4):
    """Download a Drive file to a local path with retry."""
    for attempt in range(retries):
        try:
            request = service.files().get_media(fileId=file_id)
            with open(dest_path, "wb") as f:
                downloader = MediaIoBaseDownload(f, request, chunksize=2*1024*1024)
                done = False
                while not done:
                    _, done = downloader.next_chunk()
            return
        except Exception as e:
            if attempt < retries - 1:
                time.sleep(2 ** attempt)
            else:
                raise


def upload_file(service, local_path: str, filename: str, parent_folder_id: str) -> str:
    """Upload a local file to a Drive folder. Returns file ID."""
    file_metadata = {"name": filename, "parents": [parent_folder_id]}
    media = MediaFileUpload(
        local_path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True,
    )
    file = _execute(service.files().create(
        body=file_metadata, media_body=media, fields="id"
    ))
    return file.get("id")


def discover_folders(service) -> dict:
    """Discover folder IDs for 01-04 subfolders inside root."""
    subfolders = list_subfolders(service, FOLDER_IDS["root"])
    mapping = {}
    for sf in subfolders:
        name = sf["name"]
        fid  = sf["id"]
        if name.startswith("01"):
            mapping["daily_selections"] = fid
        elif name.startswith("02"):
            mapping["invoices"] = fid
        elif name.startswith("03"):
            mapping["reference"] = fid
        elif name.startswith("04"):
            mapping["outputs"] = fid
    return mapping


def download_folder_to_temp(service, folder_id: str,
                             tmp_dir: str, subfolder_name: str = "") -> str:
    """
    Download all Excel files from a Drive folder (and subfolders) to local temp dir.
    Returns local folder path.
    """
    local_folder = os.path.join(tmp_dir, subfolder_name) if subfolder_name else tmp_dir
    os.makedirs(local_folder, exist_ok=True)

    # Download files in this folder
    try:
        files = list_folder(service, folder_id)
    except Exception as e:
        print(f"  [WARN] Could not list folder: {e}")
        return local_folder

    for f in files:
        if f["name"].endswith(".xlsx"):
            local_path = os.path.join(local_folder, f["name"])
            try:
                download_file(service, f["id"], local_path)
            except Exception as e:
                print(f"  [WARN] Failed to download {f['name']}: {e}")

    # Recurse into subfolders
    try:
        subfolders = list_subfolders(service, folder_id)
        for sf in subfolders:
            download_folder_to_temp(service, sf["id"], local_folder, sf["name"])
    except Exception as e:
        print(f"  [WARN] Could not list subfolders: {e}")

    return local_folder
