"""
Download the three 2026 Wio bank statement CSVs (Jan, Feb, Mar) to /tmp.
Prints each file's first 30 lines so we can see the CSV schema.
Read-only. No modifications to Drive.
"""

import sys, os, tempfile
sys.path.insert(0, os.path.dirname(__file__))

from drive_client import get_service, download_file

# IDs discovered by explore_2026.py
FILES = [
    ("jan_2026.csv", "19u9QuOs8PmXoEL0ho0vSE9GwrKGdFNem"),  # statement(Jan 1, 2026 - Jan 31, 2027.csv
    ("feb_2026.csv", "1L1aLC87Mlu4QgAfvGqOqONt-IZ6oFGJ1"),  # statement(Feb 1, 2026 - Feb 28, 2026).csv
    ("mar_2026.csv", "1hJmiMFRtnZ6lshpZhHFfS2-E1E8wT5A2"),  # statement(Mar 1, 2026 - Mar 31, 2026).csv
]

OUT_DIR = "/tmp/eatclean_bank_2026"
os.makedirs(OUT_DIR, exist_ok=True)


def main():
    svc = get_service()
    for fname, fid in FILES:
        path = os.path.join(OUT_DIR, fname)
        print(f"Downloading {fname} ({fid})", file=sys.stderr)
        download_file(svc, fid, path)
        print(f"\n=== {fname} (first 30 lines) ===")
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            for i, line in enumerate(f):
                if i >= 30:
                    break
                print(line.rstrip())


if __name__ == "__main__":
    main()
