"""
Find zero-price dishes — date and client for each occurrence.
Run from: ~/Documents/eatclean-economics/eatclean/
Command:  python3 find_zero_price.py
"""

import sys, os, re, glob, tempfile
import pandas as pd

sys.path.insert(0, os.path.dirname(__file__))
from drive_client import get_service, discover_folders, download_folder_to_temp

TARGETS = [
    'Beef Burrito LOWCAL',
    'Beef Burrito LEAN',
    'Cajun Spiced Fish LEAN',
    'Black Pepper Prawns LEAN',
    'Caponata Pasta LOWCAL',
    'Chicken Cilantro Bowl LEAN',
]

print("Connecting to Drive...")
service = get_service()
folders = discover_folders(service)
print(f"Connected. Downloading daily selections...")

tmp = tempfile.mkdtemp()
download_folder_to_temp(service, folders['daily_selections'], tmp)

files = sorted(glob.glob(os.path.join(tmp, '**', '*.xlsx'), recursive=True))
print(f"Downloaded {len(files)} files. Scanning...\n")

found = {t: [] for t in TARGETS}

for f in files:
    try:
        df = pd.read_excel(f, header=0)
    except Exception:
        continue

    # Standard format
    if 'Person' in df.columns:
        for _, row in df.iterrows():
            person = str(row.get('Person', '')).strip()
            date   = str(row.get('Meal plan for the day', ''))[:10]
            for line in str(row.get('Meals', '')).split('\n'):
                line = line.strip()
                if ':' not in line:
                    continue
                mt, dish_raw = line.split(':', 1)
                dish = re.sub(r'\s*-\s*\d+\s*$', '', dish_raw.strip().replace('(x)', '')).strip()
                dish = re.sub(r'\s+', ' ', dish)
                if dish in TARGETS:
                    found[dish].append(f"{date}  |  {mt.strip():<14}  |  {person}")

print("=" * 70)
print("ZERO-PRICE DISHES — ALL OCCURRENCES")
print("=" * 70)

any_found = False
for dish in TARGETS:
    hits = found[dish]
    if hits:
        any_found = True
        print(f"\n❌ {dish}  ({len(hits)} occurrence{'s' if len(hits)>1 else ''})")
        for h in hits:
            print(f"   {h}")
    else:
        print(f"\n✅ {dish}  — not found in any file")

if not any_found:
    print("\nNo zero-price dishes found. All clear.")

print("\n" + "=" * 70)
