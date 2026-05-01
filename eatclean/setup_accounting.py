"""
EatClean Accounting Pipeline — Setup
Run once to discover the Accounting folder structure and save IDs.

Run from: ~/Documents/eatclean-economics/eatclean/
Command:  python3 setup_accounting.py

Output: accounting_config.json (saved to same folder)
"""

import sys, os, json, time
sys.path.insert(0, os.path.dirname(__file__))

from drive_client import get_service, _execute


def list_folder_recursive(service, folder_id, indent=0, path=""):
    """Recursively list all files and folders. Returns structured dict."""
    results = []
    req = service.files().list(
        q=f"'{folder_id}' in parents and trashed=false",
        fields="files(id, name, mimeType, size, modifiedTime)",
        orderBy="name",
        pageSize=100,
    )
    resp = _execute(req)
    
    for f in resp.get("files", []):
        is_folder = "folder" in f["mimeType"]
        size_kb   = int(f.get("size", 0)) // 1024 if not is_folder else 0
        full_path = f"{path}/{f['name']}" if path else f['name']
        
        icon = "📁" if is_folder else "📄"
        size_str = f"({size_kb}KB)" if size_kb > 0 else ""
        print("  " * indent + f"{icon} {f['name']} {size_str}")
        
        entry = {
            "id":       f["id"],
            "name":     f["name"],
            "path":     full_path,
            "type":     "folder" if is_folder else "file",
            "mimeType": f["mimeType"],
            "size_kb":  size_kb,
            "modified": f.get("modifiedTime", ""),
        }
        
        if is_folder:
            entry["children"] = list_folder_recursive(
                service, f["id"], indent + 1, full_path
            )
        
        results.append(entry)
    
    return results


def find_accounting_folder(service):
    """Search for Accounting folder accessible by service account."""
    req = service.files().list(
        q="name='Accounting' and mimeType='application/vnd.google-apps.folder' and trashed=false",
        fields="files(id, name, parents)",
    )
    resp = _execute(req)
    folders = resp.get("files", [])
    
    if not folders:
        print("❌ Accounting folder not found.")
        print("   Make sure you shared it with:")
        print("   eatclean-pipeline@eatclean-pipeline.iam.gserviceaccount.com")
        return None
    
    if len(folders) > 1:
        print(f"Found {len(folders)} folders named 'Accounting':")
        for i, f in enumerate(folders):
            print(f"  {i+1}. ID: {f['id']}")
        return folders[0]  # Take first
    
    return folders[0]


def main():
    print("=" * 60)
    print("EatClean Accounting — Setup")
    print("=" * 60)
    
    print("\nConnecting to Drive...")
    service = get_service()
    print("✅ Connected")
    
    print("\nSearching for Accounting folder...")
    folder = find_accounting_folder(service)
    if not folder:
        return
    
    print(f"✅ Found: '{folder['name']}' (ID: {folder['id']})\n")
    print("Mapping folder structure...")
    print("-" * 60)
    
    structure = list_folder_recursive(service, folder["id"])
    
    print("-" * 60)
    
    # Count files
    def count_files(items):
        total = 0
        for item in items:
            if item["type"] == "file":
                total += 1
            elif "children" in item:
                total += count_files(item["children"])
        return total
    
    total = count_files(structure)
    print(f"\n✅ Found {total} files across {len(structure)} top-level folders")
    
    # Save config
    config = {
        "accounting_folder_id": folder["id"],
        "accounting_folder_name": folder["name"],
        "structure": structure,
        "last_scanned": time.strftime("%Y-%m-%d %H:%M:%S"),
    }
    
    config_path = os.path.join(os.path.dirname(__file__), "accounting_config.json")
    with open(config_path, "w") as f:
        json.dump(config, f, indent=2)
    
    print(f"\n✅ Config saved to: accounting_config.json")
    print("\nNext step: run python3 accounting_pipeline.py")
    print("=" * 60)


if __name__ == "__main__":
    main()
