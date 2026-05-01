"""
Parse a Wio bank account statement PDF into structured transactions.

Handles the Wio AED statement format:
  DD/MM/YYYY REFNUMBER DESCRIPTION AMOUNT BALANCE

Multi-line descriptions are uncommon here because each tx sits on one PDF row.
Notes/Original-ref columns are NOT in the PDF (they're only in CSV exports),
so we approximate the "notes" field as empty.
"""

import re
from datetime import datetime

import pdfplumber

# Example matches:
#   02/01/2026 P550745816 From NETWORK INTERNATIONAL LLC 7,171.8 84,276.43
#   05/01/2026 P550661642 Google Workspace_eatcle (rate: 3.7) -621.6 86,778.5
TX_RE = re.compile(
    r"""^(?P<date>\d{2}/\d{2}/\d{4})\s+
        (?P<ref>P\d{6,})\s+
        (?P<desc>.+?)\s+
        (?P<amount>-?[\d,]+(?:\.\d+)?)\s+
        (?P<balance>-?[\d,]+(?:\.\d+)?)\s*$""",
    re.VERBOSE,
)


def _parse_num(s: str) -> float:
    return float(s.replace(",", ""))


def _normalize_date(s: str) -> str:
    """Convert DD/MM/YYYY -> YYYY-MM-DD."""
    return datetime.strptime(s, "%d/%m/%Y").strftime("%Y-%m-%d")


def _classify_tx_type(desc: str) -> str:
    """Best-effort transaction type matching the CSV's column."""
    d = desc.lower()
    if "foreign exchange transaction fee" in d or "subscription fee for" in d:
        return "Fees"
    if d.startswith("from ") or d.startswith("to "):
        return "Transfers"
    return "Card"


def parse_pdf(pdf_path: str) -> list[dict]:
    """
    Parse the Wio statement PDF into a list of transaction dicts shaped like
    the CSV parser output (see cashflow_build.parse_csv).
    """
    txns = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.split("\n"):
                line = line.strip()
                if not line:
                    continue
                m = TX_RE.match(line)
                if not m:
                    continue
                date_iso = _normalize_date(m.group("date"))
                amount = _parse_num(m.group("amount"))
                balance = _parse_num(m.group("balance"))
                desc = m.group("desc").strip()
                # Strip the "(rate: 3.7)" FX-rate annotation for cleaner matching
                desc_clean = re.sub(r"\s*\(rate:\s*[\d.]+\)\s*", " ", desc).strip()
                month_key = date_iso[:7]  # YYYY-MM
                txns.append({
                    "month": month_key,
                    "date": date_iso,
                    "tx_type": _classify_tx_type(desc),
                    "description": desc_clean,
                    "notes": "",
                    "amount": amount,
                    "balance": balance,
                    "ref": m.group("ref"),
                })
    return txns


if __name__ == "__main__":
    import sys, json
    if len(sys.argv) < 2:
        print("Usage: python3 wio_pdf_parser.py <pdf_path>", file=sys.stderr)
        sys.exit(1)
    txns = parse_pdf(sys.argv[1])
    print(f"Parsed {len(txns)} transactions", file=sys.stderr)
    # Quick sanity: by month
    from collections import Counter
    by_month = Counter(t["month"] for t in txns)
    for m in sorted(by_month):
        print(f"  {m}: {by_month[m]} tx", file=sys.stderr)
    if txns:
        print("First tx:", json.dumps(txns[0], default=str))
        print("Last tx:", json.dumps(txns[-1], default=str))
