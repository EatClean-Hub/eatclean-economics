# EatClean Economics Pipeline

Automated unit economics analysis for EatClean UAE (Food Solutions Tech FZCO).

## What it does
- Pulls daily meal selections and Basiligo invoices from Google Drive
- Calculates per-customer kitchen cost, delivery cost, margin
- Flags zero-price dishes and price violations
- Outputs weekly assessment Excel to Google Drive

## Stack
- Python 3
- Google Drive API (service account)
- openpyxl, pandas

## Run
cd eatclean && python3 run.py
