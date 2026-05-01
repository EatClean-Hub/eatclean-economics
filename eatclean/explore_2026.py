"""
Explore the 2026 folder structure in the Accounting Drive folder.
Lists all folders, subfolders, and files relevant to 2026 only.
Read-only discovery. Does not download.
"""

import sys, os, json
sys.path.insert(0, os.path.dirname(__file__))

from drive_client import get_service, _execute

ACCOUNTING_FOLDER_ID = "1XkmMgLanlHSwCmtAOEAn_NGI69e1VYGr"

MONTHS_2026 = [
    "jan 2026", "january 2026", "jan-2026", "january-2026", "01 2026", "01-2026",
    "feb 2026", "february 2026", "feb-2026", "february-2026", "02 2026", "02-2026",
    "mar 2026", "march 2026", "mar-2026", "march-2026", "03 2026", "03-2026",
    "apr 2026", "april 2026", "apr-2026", "april-2026", "04 2026", "04-2026",
    "2026",
]


def list_children(service, folder_id):
    req = service.files().list(
        q=f"'{folder_id}' in parents and trashed=false",
        fields="files(id, name, mimeType, size)",
        pageSize=200,
    )
    resp = _execute(req)
    return resp.get("files", [])


def walk(service, folder_id, path, depth, max_depth, out):
    if depth > max_depth:
        return
    try:
        children = list_children(service, folder_id)
    except Exception as e:
        out.append(f"[ERROR] {path}: {e}")
        return
    for c in children:
        is_folder = c["mimeType"] == "application/vnd.google-apps.folder"
        full = f"{path}/{c['name']}"
        tag = "[DIR]" if is_folder else "[FILE]"
        size = c.get("size", "")
        out.append(f"{tag} {full}  id={c['id']}  size={size}  mime={c['mimeType']}")
        if is_folder:
            walk(service, c["id"], full, depth + 1, max_depth, out)


def main():
    print("Connecting...", file=sys.stderr)
    service = get_service()
    print(f"Walking Accounting folder (max depth 4)...", file=sys.stderr)
    out = []
    walk(service, ACCOUNTING_FOLDER_ID, "", 0, 4, out)

    # Filter to 2026-relevant paths only
    filtered = []
    for line in out:
        low = line.lower()
        if "2026" in low:
            filtered.append(line)

    print(f"Total items: {len(out)}", file=sys.stderr)
    print(f"2026-relevant: {len(filtered)}", file=sys.stderr)
    print("\n".join(filtered))


if __name__ == "__main__":
    main()
