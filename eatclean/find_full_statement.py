"""
Look for the full Wio statement (01/01/2026 - 24/04/2026) in the March 2026 folder.
Lists every file in March 2026 and related bank/statement folders with size + id.
"""

import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from drive_client import get_service, _execute

MARCH_FOLDER_ID = "1ZDHBmIQzrK8y9W_Ddz7-cSb8cKL1kgtN"


def walk(service, folder_id, path):
    req = service.files().list(
        q=f"'{folder_id}' in parents and trashed=false",
        fields="files(id, name, mimeType, size)",
        pageSize=200,
    )
    resp = _execute(req)
    items = resp.get("files", [])
    for c in items:
        is_dir = c["mimeType"] == "application/vnd.google-apps.folder"
        tag = "[DIR]" if is_dir else "[FILE]"
        size = c.get("size", "")
        full = f"{path}/{c['name']}"
        print(f"{tag}  {full}  size={size}  id={c['id']}")
        if is_dir:
            walk(service, c["id"], full)


def main():
    svc = get_service()
    print("March 2026 tree:")
    walk(svc, MARCH_FOLDER_ID, "/March 2026")


if __name__ == "__main__":
    main()
