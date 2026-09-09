#!/usr/bin/env python3
"""
sync_simplefin.py — DCR SimpleFIN Bridge Sync
==============================================
Connects to SimpleFIN Bridge API to sync live balances for:
  1. Merrill Lynch Investment Account -> Net Worth & Config!C24
  2. Banco Popular Operating Checking -> Net Worth & Config!C29
  3. Credit Cards (Capital One, Amex) -> Net Worth & Config!C30

USAGE:
  python3 sync_simplefin.py                       # Sync using stored URL in Config!C27
  python3 sync_simplefin.py --url <TOKEN_OR_URL>  # Set/claim Access URL & sync
  python3 sync_simplefin.py --dry-run             # Fetch and display without writing to Excel
"""

import os, sys, argparse, json, base64, urllib.request, urllib.parse, re, time
from pathlib import Path
from datetime import datetime

# Import openpyxl from _vendor
sys.path.insert(0, str(Path(__file__).parent.parent / "_vendor"))
sys.path.insert(0, str(Path(__file__).parent / "_vendor"))
import openpyxl

try:
    import sync_budget
    resolve_xlsx_path = getattr(sync_budget, "resolve_xlsx_path", lambda: Path(__file__).parent.parent / "DCR Ledger.xlsx")
except ImportError:
    resolve_xlsx_path = lambda: Path(__file__).parent.parent / "DCR Ledger.xlsx"


def resolve_access_url(input_token_or_url: str) -> str:
    """
    Handles SimpleFIN Setup Tokens (base64) or claim URLs.
    If input is a claim URL or base64 setup token, claims it via POST
    to retrieve the permanent Basic Auth Access URL.
    """
    cleaned = input_token_or_url.strip()
    if not cleaned:
        return ""

    # 1. Check if base64 encoded setup token
    if not cleaned.startswith("http"):
        try:
            decoded = base64.b64decode(cleaned).decode("utf-8").strip()
            if decoded.startswith("http"):
                cleaned = decoded
        except Exception:
            pass

    # 2. If it is a claim URL, POST to claim it
    if "/claim/" in cleaned:
        print(f"Connecting to SimpleFIN Bridge claim service...")
        req = urllib.request.Request(cleaned, data=b"", method="POST")
        req.add_header("User-Agent", "DCR-Ledger-Sync/2.5")
        try:
            with urllib.request.urlopen(req, timeout=15) as resp:
                code = resp.getcode()
                body = resp.read().decode("utf-8").strip()
                if code == 200 and body.startswith("http"):
                    print("✅ Successfully claimed permanent SimpleFIN Access URL!")
                    return body
                elif code == 403:
                    print("❌ Setup token already claimed or expired. Please generate a fresh token on bridge.simplefin.org.")
                    sys.exit(1)
                else:
                    print(f"❌ SimpleFIN claim returned HTTP {code}: {body}")
                    sys.exit(1)
        except urllib.error.HTTPError as he:
            err_body = he.read().decode("utf-8", errors="ignore")
            print(f"❌ Claim error (HTTP {he.code}): {err_body}")
            sys.exit(1)
        except Exception as e:
            print(f"❌ Could not claim SimpleFIN URL: {e}")
            sys.exit(1)

    return cleaned


def parse_simplefin_url(access_url: str):
    """Parses SimpleFIN access URL into base URL and Basic Auth headers."""
    trimmed = access_url.strip().rstrip("/")
    parsed = urllib.parse.urlparse(trimmed)
    if parsed.username and parsed.password:
        auth_pair = f"{parsed.username}:{parsed.password}".encode("utf-8")
        auth_header = "Basic " + base64.b64encode(auth_pair).decode("ascii")
        host = parsed.netloc.split("@")[-1]
        clean_url = f"{parsed.scheme}://{host}{parsed.path}"
        return clean_url, {"Authorization": auth_header}
    return trimmed, {}


def fetch_simplefin_data(access_url: str, start_days: int = 90):
    """Calls SimpleFIN Bridge /accounts endpoint with transactions."""
    clean_url, headers = parse_simplefin_url(access_url)
    if not clean_url.endswith("/accounts"):
        clean_url += "/accounts"
    
    # Try with start-date parameter first
    safe_start = int(time.time()) - (start_days * 24 * 60 * 60)
    fetch_url = f"{clean_url}?start-date={safe_start}"
    req = urllib.request.Request(fetch_url, headers=headers)
    req.add_header("User-Agent", "DCR-Ledger-Sync/2.5")
    
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            return data
    except urllib.error.HTTPError as he:
        # If server rejects start-date query, fallback to plain /accounts
        try:
            plain_req = urllib.request.Request(clean_url, headers=headers)
            plain_req.add_header("User-Agent", "DCR-Ledger-Sync/2.5")
            with urllib.request.urlopen(plain_req, timeout=20) as resp:
                data = json.loads(resp.read().decode("utf-8"))
                return data
        except Exception as e:
            print(f"❌ Error connecting to SimpleFIN Bridge: {e}")
            return None
    except Exception as e:
        print(f"❌ Error connecting to SimpleFIN Bridge: {e}")
        return None


def get_stored_url(xlsx_path: Path) -> str:
    """Reads SimpleFIN Access URL stored in Config!C27 or environment."""
    env_url = os.environ.get("SIMPLEFIN_ACCESS_URL")
    if env_url:
        return env_url.strip()
    if not xlsx_path.exists():
        return ""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    if "Config" in wb.sheetnames:
        val = wb["Config"]["C27"].value
        return str(val).strip() if val else ""
    return ""


PR_US_MERCHANT_MAP = {
    # Groceries
    "costco": "Groceries", "pueblo": "Groceries", "supermax": "Groceries",
    "econo": "Groceries", "freshmart": "Groceries", "amigo": "Groceries",
    "walmart": "Groceries", "sam's": "Groceries", "sams club": "Groceries",
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
    "walgreens": "Health", "cvs": "Health", "farmacia": "Health",
    "liv fitness": "Health", "planet fitness": "Health", "equinox": "Health",
    # Subscriptions / Utilities
    "luma": "Subscriptions", "aaa": "Subscriptions", "liberty": "Subscriptions",
    "claro": "Subscriptions", "t-mobile": "Subscriptions", "at&t": "Subscriptions",
    "netflix": "Subscriptions", "spotify": "Subscriptions", "apple.com": "Subscriptions",
    # Shopping
    "amazon": "Shopping", "home depot": "Shopping", "ikea": "Shopping",
    "best buy": "Shopping", "apple store": "Shopping", "zara": "Shopping",
}

SKIP_TRANSFER_PATTERNS = [
    "payment received", "autopay payment", "credit card payment",
    "payment thank you", "internet payment", "transfer to", "transfer from",
    "online payment"
]


def clean_desc(desc: str) -> str:
    cleaned = desc.split("  ")[0].strip()
    cleaned = re.sub(r"(SAN JUAN|CAROLINA|GUAYNABO|LONDON|PRI|USA|GBR).*$", "", cleaned, flags=re.I).strip()
    return cleaned.title() if cleaned else "Unknown Merchant"


def categorize_tx(clean_name: str) -> str:
    name_lower = clean_name.lower()
    for pattern, cat in PR_US_MERCHANT_MAP.items():
        if pattern in name_lower:
            return cat
    try:
        mapped = sync_budget.map_category(clean_name, "")
        if mapped:
            return mapped
    except Exception:
        pass
    return "Misc"


def classify_account(acc: dict):
    """
    Accurately classifies an account into:
      - 'merrill_investment' (Asset: Investment Brokerage)
      - 'popular_checking'   (Asset: Operating Checking)
      - 'credit_card'        (Liability: Revolving Debt)
      - 'mortgage_loan'      (Liability: Long-term Debt)
      - 'other'
    """
    name = (acc.get("name") or "").strip()
    org_obj = acc.get("org") or {}
    org_name = (org_obj.get("name") or "").strip()
    
    full_str = f"{name} {org_name}".lower()
    acc_lower = name.lower()
    org_lower = org_name.lower()

    # 1. Merrill Lynch Investment
    is_merrill_loan = any(k in acc_lower for k in ["loan", "lma", "facility", "line of credit"])
    if (any(k in org_lower for k in ["merrill", "lynch"]) or 
        any(k in acc_lower for k in ["merrill", "investment advantage", "cma", "brokerage"])) and not is_merrill_loan:
        return "merrill_investment", "Merrill Lynch Investment Account"

    # 2. Banco Popular Operating Checking
    is_popular = any(k in org_lower or k in acc_lower for k in ["banco popular", "popular"])
    if is_popular:
        if any(k in acc_lower for k in ["mortgage", "hipoteca", "fha"]):
            return "mortgage_loan", "Banco Popular Mortgage"
        if any(k in acc_lower for k in ["credit card", "tarjeta", "visa", "mastercard"]):
            return "credit_card", "Banco Popular Credit Card"
        return "popular_checking", "Banco Popular Checking"

    # 3. Capital One Venture X / Cards
    if "capital one" in full_str or "venture" in full_str:
        if any(k in full_str for k in ["venture", "card", "credit"]):
            return "credit_card", "Capital One Venture X"
        return "other_checking", "Capital One Checking"

    # 4. Bank of America
    if "bank of america" in full_str or "bofa" in full_str:
        if any(k in full_str for k in ["card", "credit", "cash rewards", "travel rewards", "customized"]):
            return "credit_card", "Bank of America Credit Card"
        return "other_checking", "Bank of America Checking"

    # 5. Monzo
    if "monzo" in full_str:
        if "flex" in full_str:
            return "credit_card", "Monzo Flex"
        return "other_checking", "Monzo Current"

    # 6. American Express
    if any(k in full_str for k in ["american express", "amex"]):
        return "credit_card", name or "Amex BA Black"

    # 7. Other Credit Cards
    if any(k in acc_lower for k in ["credit card", "card", "visa", "mastercard"]):
        return "credit_card", name

    # 8. Other checking / savings
    if any(k in acc_lower for k in ["checking", "cheque", "cuenta de cheques", "savings", "ahorros"]):
        return "other_checking", name

    return "other", name


def main():
    parser = argparse.ArgumentParser(description="Sync balances and transactions from SimpleFIN Bridge to DCR Ledger.xlsx")
    parser.add_argument("--url", help="SimpleFIN Access URL or Setup Token")
    parser.add_argument("--dry-run", action="store_true", help="Display fetched accounts without modifying Excel")
    args = parser.parse_args()

    xlsx_path = resolve_xlsx_path()
    raw_url = args.url or get_stored_url(xlsx_path)

    if not raw_url:
        print("=" * 65)
        print(" SimpleFIN Bridge Setup Required")
        print("=" * 65)
        print("No SimpleFIN Access URL or Setup Token found.")
        print("1. Log in to https://bridge.simplefin.org")
        print("2. Connect your Merrill Lynch, Banco Popular, and Credit Cards")
        print("3. Click 'Generate Setup Token' or 'Access URL'")
        print("4. Run: python3 Gravity/sync_simplefin.py --url \"<TOKEN_OR_URL>\"")
        print("   OR paste it directly into cell C27 on the Config sheet in DCR Ledger.xlsx")
        print("=" * 65)
        sys.exit(1)

    # Claim or normalize Access URL
    access_url = resolve_access_url(raw_url)

    print(f"Connecting to SimpleFIN Bridge...")
    data = fetch_simplefin_data(access_url)
    if not data or "accounts" not in data:
        print("❌ Could not retrieve account data from SimpleFIN.")
        sys.exit(1)

    accounts = data.get("accounts", [])
    print(f"✅ Successfully fetched {len(accounts)} accounts from SimpleFIN.\n")

    merrill_balance = None
    popular_balance = None
    card_balance = 0.0
    mortgage_balance = None
    matched_ids = {}

    wb = openpyxl.load_workbook(xlsx_path, data_only=False)
    cfg = wb["Config"]
    ws_tx = wb["Transactions"]
    ws_nw = wb["Net Worth"]

    # Read existing transaction IDs
    existing_ids = set()
    for row in ws_tx.iter_rows(min_row=3, values_only=True):
        if row and len(row) >= 8 and row[7]:
            existing_ids.add(str(row[7]).strip())

    new_tx_list = []

    print("--- SIMPLEFIN ACCOUNT AUDIT & CLASSIFICATION ---")
    for acc in accounts:
        acc_id = acc.get("id", "")
        name = acc.get("name", "Unknown Account")
        org = acc.get("org", {}).get("name", "")
        currency = acc.get("currency", "USD")
        raw_bal = acc.get("balance", "0.00")
        try:
            bal = float(raw_bal)
        except ValueError:
            bal = 0.0

        aclass, display_name = classify_account(acc)
        full_desc = f"{name} ({org})" if org else name

        print(f"  • {full_desc:38s}: {currency} {bal:12,.2f}  [{aclass.upper()}]")

        if aclass == "merrill_investment":
            merrill_balance = bal
            matched_ids["merrill"] = acc_id
            tx_count = len(acc.get("transactions", []))
            print(f"    ↳ 💎 Merrill Lynch Portfolio: ${bal:,.2f} (Excluded {tx_count} investment trades from spend ledger)")
            continue

        elif aclass == "popular_checking":
            popular_balance = bal
            matched_ids["popular"] = acc_id
            print(f"    ↳ 🏦 Banco Popular Operating Checking: ${bal:,.2f}")

        elif aclass == "credit_card":
            debt = abs(bal)
            card_balance += debt
            matched_ids.setdefault("cards", []).append(acc_id)
            print(f"    ↳ 💳 Credit Card Liability: ${debt:,.2f} ({display_name})")

        elif aclass == "mortgage_loan":
            mortgage_balance = abs(bal)
            matched_ids["mortgage"] = acc_id
            print(f"    ↳ 🏠 Real Estate Mortgage Liability: ${mortgage_balance:,.2f}")

        # Process spend transactions for banking / credit accounts
        tx_list = acc.get("transactions", [])
        added_count = 0

        for t in tx_list:
            t_id = (t.get("id") or "").strip()
            if not t_id or t_id in existing_ids:
                continue

            raw_amt = float(t.get("amount") or 0)
            if raw_amt == 0:
                continue

            # Expenses in banking feeds are negative; positive in our spend ledger
            amt_spend = -raw_amt

            raw_desc = t.get("description") or t.get("payee") or t.get("memo") or "Unknown"
            clean_merchant = clean_desc(raw_desc)

            if any(p in raw_desc.lower() for p in SKIP_TRANSFER_PATTERNS):
                continue

            posted = t.get("posted") or t.get("transacted_at")
            tx_date = datetime.fromtimestamp(posted) if posted else datetime.now()
            category = categorize_tx(clean_merchant)

            new_tx_list.append({
                "date": tx_date,
                "merchant": clean_merchant,
                "amount": amt_spend,
                "category": category,
                "account": display_name,
                "month": tx_date.month,
                "year": tx_date.year,
                "tx_id": t_id
            })
            existing_ids.add(t_id)
            added_count += 1

        if tx_list:
            print(f"    ↳ Synced {added_count} new transactions (out of {len(tx_list)} available)")

    print("-" * 65)
    print("NET WORTH BALANCE SHEET FEED SUMMARY:")
    if merrill_balance is not None:
        print(f"  • Merrill Lynch Portfolio (Asset)    : ${merrill_balance:,.2f} -> Config!C24 & Net Worth!E10")
    else:
        print("  • Merrill Lynch Portfolio            : Not detected (retaining current baseline)")

    if popular_balance is not None:
        print(f"  • Banco Popular Checking (Asset)     : ${popular_balance:,.2f} -> Config!C29 & Net Worth!E9")
        print(f"    (Operating Waterline preserved     : ${cfg['C7'].value:,.2f} in Config!C7)")
    else:
        print("  • Banco Popular Checking             : Not detected (retaining current waterline)")

    print(f"  • Revolving Credit Cards (Liability) : ${card_balance:,.2f} -> Config!C30 & Net Worth!E18")
    if mortgage_balance is not None:
        print(f"  • Mortgage Principal (Liability)     : ${mortgage_balance:,.2f}")
    print(f"  • New Spend Transactions             : {len(new_tx_list)}")

    if args.dry_run:
        print("\n[DRY RUN] Excel updates skipped.")
        return

    # Update Access URL in Config!C27
    if access_url:
        cfg["C27"] = access_url

    # Update SimpleFIN account mapping in Config!C28
    if matched_ids:
        cfg["C28"] = json.dumps(matched_ids)

    # 1. Update Merrill Lynch Portfolio
    if merrill_balance is not None:
        cfg["C24"] = round(merrill_balance, 2)

    # 2. Update Banco Popular Live Checking in dedicated C29
    if popular_balance is not None:
        cfg["C29"] = round(popular_balance, 2)
        # Ensure Net Worth E9 formula points to C29
        ws_nw["E9"] = "=IF(ISBLANK(Config!C29), Config!C7, Config!C29)"

    # 3. Update Revolving Credit Cards in dedicated C30 & Net Worth E18
    cfg["C30"] = round(card_balance, 2)
    ws_nw["E18"] = "=IF(ISBLANK(Config!C30), 0, Config!C30)"

    # Append new transactions to Transactions sheet
    if new_tx_list:
        new_tx_list.sort(key=lambda x: x["date"])
        
        C_WHITE = "FFFFFF"
        C_CARD_BG = "F8FAFC"
        C_OBSIDIAN = "080F1A"
        C_SLATE_LIGHT = "64748B"
        C_BORDER_SUBTLE = "E2E8F0"
        thin_hairline = openpyxl.styles.Side(border_style="thin", color=C_BORDER_SUBTLE)
        box_border = openpyxl.styles.Border(left=thin_hairline, right=thin_hairline, top=thin_hairline, bottom=thin_hairline)

        start_row = ws_tx.max_row + 1
        for i, t in enumerate(new_tx_list):
            curr_row = start_row + i
            bg_fill = openpyxl.styles.PatternFill("solid", fgColor=C_CARD_BG if curr_row % 2 == 0 else C_WHITE)

            ws_tx.cell(row=curr_row, column=1, value=t["date"]).number_format = "yyyy-mm-dd"
            ws_tx.cell(row=curr_row, column=2, value=t["merchant"]).font = openpyxl.styles.Font(name="Arial", size=9, color=C_OBSIDIAN)

            c3 = ws_tx.cell(row=curr_row, column=3, value=t["amount"])
            c3.number_format = "$#,##0.00"
            c3.font = openpyxl.styles.Font(name="Arial", size=9, color=C_OBSIDIAN)

            ws_tx.cell(row=curr_row, column=4, value=t["category"]).font = openpyxl.styles.Font(name="Arial", size=9, color=C_SLATE_LIGHT)
            ws_tx.cell(row=curr_row, column=5, value=t["account"]).font = openpyxl.styles.Font(name="Arial", size=9, color=C_SLATE_LIGHT)
            ws_tx.cell(row=curr_row, column=6, value=t["month"]).font = openpyxl.styles.Font(name="Arial", size=9, color=C_SLATE_LIGHT)
            ws_tx.cell(row=curr_row, column=7, value=t["year"]).font = openpyxl.styles.Font(name="Arial", size=9, color=C_SLATE_LIGHT)
            ws_tx.cell(row=curr_row, column=8, value=t["tx_id"]).font = openpyxl.styles.Font(name="Arial", size=8, color=C_SLATE_LIGHT)

            ws_tx.cell(row=curr_row, column=9, value=f'=IFERROR(INDEX(\'Monthly Budget\'!$B$6:$B$16, MATCH(D{curr_row}, \'Monthly Budget\'!$C$6:$C$16, 0)), "Discretionary")')
            ws_tx.cell(row=curr_row, column=10, value=f'=IFERROR(INDEX(\'Monthly Budget\'!$D$6:$D$16, MATCH(D{curr_row}, \'Monthly Budget\'!$C$6:$C$16, 0)), 0)').number_format = "$#,##0.00"
            ws_tx.cell(row=curr_row, column=11, value=f'=IF(J{curr_row}>0, C{curr_row}/J{curr_row}, 0)').number_format = "0.0%"
            ws_tx.cell(row=curr_row, column=12, value=f'=IF(C{curr_row}>250, "🔴 Large Tx", IF(C{curr_row}>100, "🟡 Medium Tx", "🟢 Regular"))')

            ws_tx.row_dimensions[curr_row].height = 20
            for col in range(1, 13):
                c = ws_tx.cell(row=curr_row, column=col)
                c.border = box_border
                c.fill = bg_fill

    # Ensure native AutoFilter covers all transactions
    ws_tx.auto_filter.ref = f"A2:L{ws_tx.max_row}"

    wb.save(xlsx_path)
    print(f"\n✅ Successfully updated {xlsx_path.name} with live balances!")
    print(f"  • Merrill Lynch balance updated in Config!C24 & Net Worth sheet.")
    print(f"  • Banco Popular live checking updated in Config!C29 & Net Worth sheet.")
    print(f"  • Revolving credit card debt updated in Config!C30 & Net Worth sheet.")
    print(f"  • Permanent Access URL safely stored in Config!C27.")


if __name__ == "__main__":
    main()
