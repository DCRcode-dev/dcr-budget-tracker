#!/usr/bin/env zsh
# DCR Budget Tracker Local Server — double-click to start
# Serves the mobile-friendly web app on port 8080 and opens it in the browser.

# Move to the script's directory (works wherever this file is saved)
cd "$(dirname "$0")"

echo "=================================================="
echo " Starting DCR Spend Tracker Local Dev Server..."
echo "=================================================="
echo ""

# Open the browser in the background
open "http://localhost:8080"

# Start the python backend server
python3 server.py

echo ""
echo "Press any key to close..."
read -n 1
