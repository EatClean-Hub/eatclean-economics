"""
Upload Cash_Flow_2026.xlsx to Drive in a new 'Cash Flow 2026' subfolder
inside the Accounting folder.

Creates the subfolder if it doesn't exist. Uploads the Excel file.
Prints the resulting folder ID and file ID.
"""

import sys, os
sys.path.insert(0, os.path.dirname(__file__))

from drive_client import get_service, _execute, upload_file, list_subfolders

ACCOUNTING_FOLDER_ID = "1XkmMgLanlHSwCmtAOEAn_NGI69e1VYGr"
SUBFOLDER_NAME = "Cash Flow 2026"
LOCAL_XLSX = "/tmp/eatclean_cashflow_2026/Cash_Flow_2026.xlsx"


def ensure_subfolder(service) -> str:
    """Return ID of 'Cash Flow 2026' subfolder; create if missing."""
    existing = list_subfolders(service, ACCOUNTING_FOLDER_ID)
    for sf in existing:
        if sf["name"] == SUBFOLDER_NAME:
            return sf["id"]
    # Create it
    req = service.files().create(
        body={
            "name": SUBFOLDER_NAME,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [ACCOUNTING_FOLDER_ID],
        },
        fields="id",
    )
    resp = _execute(req)
    return resp["id"]


def main():
    if not os.path.exists(LOCAL_XLSX):
        print(f"ERROR: {LOCAL_XLSX} not found. Run cashflow_build.py first.",
              file=sys.stderr)
        sys.exit(1)
    svc = get_service()
    folder_id = ensure_subfolder(svc)
    print(f"'{SUBFOLDER_NAME}' folder id: {folder_id}")
    print(f"Uploading {os.path.basename(LOCAL_XLSX)}...")
    file_id = upload_file(svc, LOCAL_XLSX, "Cash_Flow_2026.xlsx", folder_id)
    print(f"Uploaded. File id: {file_id}")
    print(f"URL: https://drive.google.com/file/d/{file_id}/view")


if __name__ == "__main__":
    main()
