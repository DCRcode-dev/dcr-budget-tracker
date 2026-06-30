#!/bin/bash
# DCR Budget Sync — double-click this file to run the sync
# It will read your Amex/ and Monzo/ CSV files and update DCR_Budget_Tracker.xlsx

# Move to the script's directory (works wherever this file is saved)
cd "$(dirname "$0")"

echo "================================"
echo " DCR Budget Sync"
echo "================================"
echo ""

# Run the sync
python3 sync_budget.py

echo ""
echo "Press any key to close..."
read -n 1
