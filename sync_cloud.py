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

# ---------------------------------------------------------------------------
# Encryption-at-rest (AES-256-GCM, PBKDF2-SHA256).
# The published data.json contains ONLY ciphertext. The passphrase never
# leaves the device / CI secret store.
# ---------------------------------------------------------------------------
# Baseline (non-transaction) figures are NOT stored in this public repo.
# Supply them via the LEDGER_BASELINE env var / CI secret as JSON, e.g.
#   {"banco_popular":0,"merrill_balance":0,"house_value":0,
#    "mortgage_balance":0,"merrill_loan":0,"home_equity":0,
#    "monthly_lifestyle_cap":0,"baseline_utilities":0,
#    "mortgage_monthly":0,"hoa_monthly":0}
BASELINE_DEFAULTS = {
    "banco_popular": 0.0, "merrill_balance": 0.0, "house_value": 0.0,
    "mortgage_balance": 0.0, "merrill_loan": 0.0, "home_equity": 0.0,
    "monthly_lifestyle_cap": 0.0, "baseline_utilities": 0.0,
    "mortgage_monthly": 0.0, "hoa_monthly": 0.0,
}

def load_baseline() -> dict:
    b = dict(BASELINE_DEFAULTS)
    raw = os.environ.get("LEDGER_BASELINE")
    if raw:
        try:
            b.update({k: float(v) for k, v in json.loads(raw).items() if k in b})
        except Exception as e:
            print(f"Warning: could not parse LEDGER_BASELINE ({e}); using zeros.")
    return b


ENC_FORMAT = "dcr-ledger-enc-v1"
PBKDF2_ITERATIONS = 250000

def encrypt_payload(payload: dict, passphrase: str) -> dict:
    import hashlib, secrets
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM

    salt = secrets.token_bytes(16)
    iv = secrets.token_bytes(12)
    key = hashlib.pbkdf2_hmac("sha256", passphrase.encode("utf-8"), salt, PBKDF2_ITERATIONS, dklen=32)
    plaintext = json.dumps(payload, separators=(",", ":")).encode("utf-8")
    ciphertext = AESGCM(key).encrypt(iv, plaintext, None)

    b64 = lambda b: base64.b64encode(b).decode("ascii")
    return {
        "format": ENC_FORMAT,
        "encrypted": True,
        "updated_at": payload.get("updated_at"),
        "kdf": {"name": "PBKDF2", "hash": "SHA-256", "iterations": PBKDF2_ITERATIONS},
        "cipher": "AES-GCM",
        "salt": b64(salt),
        "iv": b64(iv),
        "ciphertext": b64(ciphertext),
        "note": "Encrypted ledger payload. No plaintext financial data is stored in this repository.",
    }


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

def sync_to_data_json(access_url: str, output_path: str = "data.json", passphrase: str = None):
    data = fetch_simplefin_data(access_url)
    if not data or "accounts" not in data:
        print("Failed to get data from SimpleFIN")
        return False

    base = load_baseline()
    accounts = data.get("accounts", [])
    all_txs = []
    popular_bal = base["banco_popular"]
    merrill_bal = base["merrill_balance"]
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
            "monthly_lifestyle_cap": base["monthly_lifestyle_cap"],
            "baseline_utilities": base["baseline_utilities"],
            "mortgage_monthly": base["mortgage_monthly"],
            "hoa_monthly": base["hoa_monthly"]
        },
        "net_worth": {
            "total_net_worth": round(popular_bal + merrill_bal + base["house_value"]
                                     - base["mortgage_balance"] - base["merrill_loan"] - card_debt, 2),
            "liquid_net_worth": round(popular_bal + merrill_bal - card_debt, 2),
            "home_equity": round(base["home_equity"], 2),
            "total_assets": round(popular_bal + merrill_bal + base["house_value"], 2),
            "total_liabilities": round(base["mortgage_balance"] + base["merrill_loan"] + card_debt, 2),
            "merrill_balance": round(merrill_bal, 2),
            "banco_popular": round(popular_bal, 2),
            "credit_card_debt": round(card_debt, 2),
            "house_value": round(base["house_value"], 2),
            "mortgage_balance": round(base["mortgage_balance"], 2),
            "merrill_loan": round(base["merrill_loan"], 2)
        },
        "transactions": all_txs
    }

    if not passphrase:
        print("ERROR: no passphrase supplied. Refusing to write plaintext financial data to a public repo.")
        print("Set LEDGER_PASSPHRASE in the environment (or repo secrets) and re-run.")
        return False

    envelope = encrypt_payload(payload, passphrase)
    with open(output_path, "w") as f:
        json.dump(envelope, f, indent=2)
    print(f"Successfully encrypted {len(all_txs)} transactions into {output_path}")
    return True

if __name__ == "__main__":
    url = os.environ.get("SIMPLEFIN_ACCESS_URL")
    if len(sys.argv) > 1 and sys.argv[1].startswith("http"):
        url = sys.argv[1]
    if not url:
        print("Usage: python3 sync_cloud.py <SIMPLEFIN_ACCESS_URL>")
        print("       (requires LEDGER_PASSPHRASE in the environment)")
        sys.exit(1)
    passphrase = os.environ.get("LEDGER_PASSPHRASE")
    if not passphrase:
        print("ERROR: LEDGER_PASSPHRASE is not set. Aborting rather than publishing plaintext.")
        sys.exit(2)
    ok = sync_to_data_json(url, passphrase=passphrase)
    sys.exit(0 if ok else 1)
