#!/usr/bin/env zsh
# DCR Budget Tracker — GitHub Pages Deploy Script
set -e

echo "=========================================================="
echo "      🚀 DCR Budget Tracker GitHub Uploader"
echo "=========================================================="
echo ""
echo "This script initializes a git repository, creates a new public"
echo "GitHub repository using your credentials, and pushes your tracker PWA."
echo ""

# Move to the script's directory
cd "$(dirname "$0")"

# Ask for credentials
read "USERNAME?Enter your GitHub username: "
if [ -z "$USERNAME" ]; then
    echo "Error: Username cannot be empty."
    exit 1
fi

# Disable echo for token prompt
stty -echo
read "TOKEN?Enter your GitHub Personal Access Token (PAT): "
stty echo
echo ""
if [ -z "$TOKEN" ]; then
    echo "Error: Token cannot be empty. (Create one at: https://github.com/settings/tokens with 'repo' scope)"
    exit 1
fi

read "REPO_NAME?Enter repository name [default: dcr-budget-tracker]: "
REPO_NAME=${REPO_NAME:-dcr-budget-tracker}

echo ""
echo "Step 1: Creating GitHub repository 'https://github.com/$USERNAME/$REPO_NAME'..."

# Create repo via GitHub API
RESPONSE=$(curl -s -o /dev/null -w "%{http_code}" \
  -H "Authorization: token $TOKEN" \
  -H "Accept: application/vnd.github.v3+json" \
  https://api.github.com/user/repos \
  -d '{"name":"'"$REPO_NAME"'", "private": false, "description": "DCR Spend Tracker & Observation Mobile Web App"}')

if [ "$RESPONSE" -eq 201 ]; then
    echo "✅ Repository created successfully!"
elif [ "$RESPONSE" -eq 422 ]; then
    echo "⚠️ Repository already exists or invalid parameters. Proceeding to push anyway..."
else
    echo "❌ Failed to create repository (HTTP status code: $RESPONSE). Please check your token and try again."
    exit 1
fi

echo ""
echo "Step 2: Initializing local Git repository and committing files..."

# Setup Git locally
if [ ! -d ".git" ]; then
  git init
fi
git config user.name "Daniel Cruz-Rosso"
git config user.email "daniel.cruz.rosso@gmail.com"

# Add tracking files (private spreadsheets and statements are ignored in .gitignore)
git add index.html manifest.json sw.js icon-192.png icon-512.png server.py Run\ Server.command sync_budget.py Run\ Sync.command BudgetSync.gs deploy.sh .gitignore

read "COMMIT_MSG?Enter commit message [default: deploy budget PWA]: "
COMMIT_MSG=${COMMIT_MSG:-deploy budget PWA}

git commit -m "$COMMIT_MSG" || echo "Nothing new to commit."
git branch -M main

# Add remote (update if exists)
git remote remove origin 2>/dev/null || true
git remote add origin "https://$USERNAME:$TOKEN@github.com/$USERNAME/$REPO_NAME.git"

echo ""
echo "Step 3: Pushing to GitHub main branch..."
git push -u origin main

echo ""
echo "=========================================================="
echo "🎉 SUCCESS! Your Budget Tracker has been pushed to GitHub."
echo "=========================================================="
echo "Playbook Repo Link: https://github.com/$USERNAME/$REPO_NAME"
echo "PWA App Link:        https://localhost:8080 (Local)"
echo "                     https://$USERNAME.github.io/$REPO_NAME (GitHub Pages)"
echo "=========================================================="
