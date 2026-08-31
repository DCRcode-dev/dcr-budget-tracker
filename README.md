# Victoria Court Budget Tracker — Operations & Setup Guide

Welcome to the **Victoria Court Budget Tracker** system. This system was custom-engineered for **Daniel Cruz Rosso (DCR)** in San Juan, Puerto Rico to provide:
1. **The Ask**: Exact dollar calculation of your Renovation Funding Gap to request from your parents.
2. **Living Budget Gauge**: A single unambiguous over/under dial against your **$2,000/month** living spend ceiling.
3. **Mortgage & Runway Ring-Fence**: Full monitoring of your **$3,305.55/month** November mortgage and 3-month cash runway reserve.

---

## ⏱️ Quick Setup Guide (Zero to First Sync in < 20 Minutes)

### Step 1: Set Up SimpleFIN Bridge ($15 / year)
1. Go to [https://bridge.simplefin.org](https://bridge.simplefin.org) and create an account.
2. Follow the on-screen prompts to connect your two primary US accounts:
   - **Banco Popular de Puerto Rico** (Checking)
   - **Capital One** (Venture X Credit Card)
3. Once connected, click **Generate Access URL** and copy the resulting URL token to your clipboard.

---

### Step 2: Create the Google Sheet & Paste the Code
1. Open Google Drive ([drive.google.com](https://drive.google.com)).
2. Click **+ New** → **Google Sheets**, and name the spreadsheet:  
   `Victoria Court Budget Tracker`
3. In the top menu, go to **Extensions** → **Apps Script**.
4. Delete any default code in the editor, open [Code.gs](file:///Users/dcr/Documents/LLMs/Projects/Budgeting/Code.gs), copy all of its contents, and paste it into the editor.
5. Click the **Save** icon (diskette).

---

### Step 3: Run the Setup Routine
1. In the Apps Script toolbar dropdown, select the function **`setupTracker`** and click **Run**.
2. When Google asks for authorization:
   - Click **Review permissions**
   - Select your Google account
   - Click **Advanced** → **Go to Victoria Court Budget Tracker (unsafe)**
   - Click **Allow**
3. Once the script finishes, switch back to your Google Sheet. You will see the entire formatted workbook built with all cards, Named Ranges, conditional formatting, and seed data.

---

### Step 4: Connect SimpleFIN & Sync
1. Refresh your Google Sheet in your browser. You will now see a custom menu at the top: **💰 Budget Tracker**.
2. Click **💰 Budget Tracker** → **Setup: store SimpleFIN credentials**.
3. Paste the SimpleFIN Access URL you copied in Step 1 and click **OK**.
4. Click **💰 Budget Tracker** → **Sync banks now**.
   - Your account balances will populate automatically and 90 days of transactions will be imported and categorized.
5. Click **💰 Budget Tracker** → **Setup: enable daily sync**.
   - This sets an automatic background trigger to refresh your data every morning at 06:00 AM AST.

---

### Step 5: Fill in Your Editable Blue Cells
Go to the **Config** tab. Cells with **BLUE TEXT** are user inputs:
- `UTILITIES_EST`: Enter your estimated monthly utilities (Electricity + Water + Internet + Phone).
- `HOA_AMOUNT`: Enter the Victoria Court HOA monthly fee (note: home insurance is included inside the HOA).
- `MORTGAGE_FIRST_PAYMENT_DATE`: Update once Banco Popular confirms the exact November date.

---

## 🛠️ How to Re-Sync Renovation Figures from Mac

The renovation database mirrors your local file `Victoria Court Tracker.xlsx` at `/Users/dcr/Documents/LLMs/Projects/Miramar/02_Project_Management/`.

Whenever you update your Excel file:
1. Open `Victoria Court Tracker.xlsx` in Excel.
2. Select the items in the Renovation table, right-click, and choose **Copy**.
3. In your Google Sheet, click **💰 Budget Tracker** → **Re-sync renovation figures**.
4. Paste the text into the prompt window and click **OK**.
5. Your dashboard figures, contractor disbursements, and The Ask funding gap will update immediately.

---

## 📥 How to Import Monzo (GBP) & Bank of America Statements

For your winding-down accounts:
1. Export a CSV statement from Monzo or Bank of America.
2. Open the CSV in a text editor, select all lines, and copy them.
3. In Google Sheets, click **💰 Budget Tracker** → **Import CSV (Monzo / BofA)**.
4. Enter the account name (e.g. `Monzo Current` or `Bank of America`).
5. Paste the CSV lines and click **OK**.
   - Monzo amounts in GBP will automatically convert to USD using the `GBP_USD_RATE` (1.27) and deduplicate against existing rows using SHA-256 signatures.

---

## ⚠️ KNOWN_OPEN_ITEMS

* Utilities monthly estimate not yet supplied.
* Victoria Court HOA monthly amount not yet supplied.
* Exact November mortgage first-payment date unconfirmed.
* The $550,000 loan sits $8,713 above the 2026 FHA national floor of $541,287; San Juan is very likely a floor county. Confirm with Peter Grell Rodríguez at Banco Popular whether the loan amount holds, and whether Condominio Victoria Court has FHA project approval.
* The remaining $550,000 of the purchase price is unsettled between DCR and his father.
* 50 renovation items still carry no price; the funding gap understates reality until priced.
