#!/usr/bin/env python3
"""
sync_cloud.py — Automated Cloud Sync for DCR Ledger
==================================================
Connects to SimpleFIN Bridge API to fetch live transactions for:
  - Capital One Venture X
  - Banco Popular Checking
  - Merrill Lynch

Writes directly to data.json for immediate ingestion by the PWA on GitHub Pages.
"""

import os, sys, json, base64, urllib.request, urllib.parse, re, time
from pathlib import Path
from datetime import datetime

PR_US_MERCHANT_MAP = {
    # Groceries
    "costco": "Groceries", "pueblo": "Groceries", "supermax": "Groceries",
    "econo": "Groceries", "freshmart": "Groceries", "amigo": "Groceries",
    "walmart": "Groceries", "sams": "Groceries", "sams club": "Groceries",
    "trader joe": "Groceries", "whole foods": "Groceries", "target": "Groceries",
    # Dining
    "starbucks": "Dining", "meson": "Dining", "bocaditos": "Dining",
    "la placita": "Dining", "hacienda san pedro": "Dining", "bistro": "Dining",
    "bakery": "Dining", "panaderia": "Dining", "mcdonald": "Dining",
    "wendy": "Dining", "subway": "Dining", "chipotle": "Dining",
    "cava": "Dining", "sweetgreen": "Dining", "uber eats": "Dining", "doordash": "Dining",
    # Transport
    "uber": "Transport", "lyft": "Transport", "total": "Transport",
    "puma": "Transport", "mobil": "Transport", "shell": "Transport",
    "texaco": "Transport", "gulf": "Transport", "autoexpreso": "Transport",
    # Health & Wellness
    "walgreens": "Fitness", "cvs": "Fitness", "farmacia": "Fitness",
    "liv fitness": "Fitness", "planet fitness": "Fitness", "equinox": "Fitness",
    # Subscriptions / Utilities
    "luma": "Utilities", "aaa": "Utilities", "liberty": "Coffee",
    "claro": "Coffee", "t-mobile": "Coffee", "at&t": "Coffee",
    "netflix": "Coffee", "spotify": "Coffee", "apple.com": "Coffee",
    # Shopping
    "amazon": "Shopping", "home depot": "Shopping", "ikea": "Shopping",
    "best buy": "Shopping", "apple store": "Shopping", "zara": "Shopping",
}

def clean_desc(desc: str) -> str:
    cleaned = desc.split("  ")[0].strip()
    cleaned = re.sub(r" (SAN JUAN|CAROLINA|GUAYNABO|LONDON|PRI|USA|GBR) .*$", "", cleaned, flags=re.I).strip()
    return cleaned.title() if cleaned else "Unknown Merchant"

def categorize_tx(clean_name: str) -> str:
    name_lower = clean_name.lower()
    for pattern, cat in PR_US_MERCHANT_MAP.items():
        if pattern in name_lower:
            return cat
    return "Other"

def parse_simplefin_url(access_url: str):
    trimmed = access_url.strip().rstrip("/")
    parsed = urllib.parse.urlparse(trimmed)
    if parsed.username and parsed.password:
        auth_pair = f"{parsed.username}:{parsed.password}".encode("utf-8")
        auth_header = "Basic " + base64.b64encode(auth_pair).decode("ascii")
        host = parsed.netloc.split("@")[-1]
        clean_url = f"{parsed.scheme}://{host}{parsed.path}"
        return clean_url, {"Authorization": auth_header}
    return trimmed, {}

def fetch_simplefin_data(access_url: str, start_days: int = 60):
    clean_url, headers = parse_simplefin_url(access_url)
    if not clean_url.endswith("/accounts"):
        clean_url += "/accounts"
    safe_start = int(time.time()) - (start_days * 24 * 60 * 60)
    fetch_url = f"{clean_url}?start-date={safe_start}"
    req = urllib.request.Request(fetch_url, headers=headers)
    req.add_header("User-Agent", "DCR-Ledger-Cloud-Sync/3.0")
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception:
        try:
            plain_req = urllib.request.Request(clean_url, headers=headers)
            plain_req.add_header("User-Agent", "DCR-Ledger-Cloud-Sync/3.0")
            with urllib.request.urlopen(plain_req, timeout=20) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except Exception as e:
            print(f"Error fetching from SimpleFIN: {e}")
            return None

def sync_to_data_json(access_url: str, output_path: str = "data.json"):
    data = fetch_simplefin_data(access_url)
    if not data or "accounts" not in data:
        print("Failed to get data from SimpleFIN")
        return False

    accounts = data.get("accounts", [])
    all_txs = []
    popular_bal = 0.0
    merrill_bal = 0.00
    card_debt = 0.00

    for acc in accounts:
        name = acc.get("name", "")
        org = acc.get("org", {}).get("name", "") if isinstance(acc.get("org"), dict) else str(acc.get("org", ""))
        full_str = f"{name} {org}".lower()
        bal = float(acc.get("balance", 0.0) or 0.0)

        # Classify account
        account_label = "Banco Popular"
        if "capital one" in full_str or "venture" in full_str:
            account_label = "Capital One Venture X"
            card_debt += abs(bal)
        elif "merrill" in full_str:
            account_label = "Merrill Lynch"
            merrill_bal = bal
        elif "popular" in full_str:
            account_label = "Banco Popular"
            popular_bal = bal

        raw_txs = acc.get("transactions", [])
        for t in raw_txs:
            amt = float(t.get("amount", 0.0) or 0.0)
            if amt < 0: # expense
                posted_ts = t.get("posted", 0)
                date_str = datetime.fromtimestamp(posted_ts).strftime("%Y-%m-%d") if posted_ts else datetime.now().strftime("%Y-%m-%d")
                desc = clean_desc(t.get("description") or t.get("payee") or "Expense")
                cat = categorize_tx(desc)
                all_txs.append({
                    "tx_id": str(t.get("id") or f"{date_str}_{desc}_{abs(amt)}"),
                    "date": date_str,
                    "merchant": desc,
                    "amount": round(abs(amt), 2),
                    "category": cat,
                    "account": account_label
                })

    all_txs.sort(key=lambda x: x["date"], reverse=True)

    payload = {
        "updated_at": datetime.utcnow().isoformat() + "Z",
        "config": {
            "monthly_lifestyle_cap": 3000.00,
            "baseline_utilities": 750.00,
            "mortgage_monthly": 0,
            "hoa_monthly": 375.00
        },
        "net_worth": {
            "total_net_worth": round(popular_bal + merrill_bal + 0 - 0 - 0 - card_debt, 2),
            "liquid_net_worth": round(popular_bal + merrill_bal - card_debt, 2),
            "home_equity": 0,
            "total_assets": round(popular_bal + merrill_bal + 0, 2),
            "total_liabilities": round(0 + 0 + card_debt, 2),
            "merrill_balance": round(merrill_bal, 2),
            "banco_popular": round(popular_bal, 2),
            "credit_card_debt": round(card_debt, 2),
            "house_value": 0,
            "mortgage_balance": 0,
            "merrill_loan": 0
        },
        "transactions": all_txs
    }

    with open(output_path, "w") as f:
        json.dump(payload, f, indent=2)
    print(f"Successfully wrote {len(all_txs)} transactions to {output_path}")
    return True

if __name__ == "__main__":
    url = os.environ.get("SIMPLEFIN_ACCESS_URL")
    if len(sys.argv) > 1 and sys.argv[1].startswith("http"):
        url = sys.argv[1]
    if not url:
        print("Usage: python3 sync_cloud.py <SIMPLEFIN_ACCESS_URL>")
        sys.exit(1)
    sync_to_data_json(url)
