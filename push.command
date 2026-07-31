#!/bin/bash
export NVM_DIR="$HOME/.nvm"
export PATH="$NVM_DIR/versions/node/v20.20.2/bin:$PATH"

cd "/Users/pratikpalkar/ganesh chaturthi 2026"

# Remove stale lock file if present
rm -f .git/index.lock

echo "📤 Pushing to GitHub..."
git add index.html
git commit -m "fix: use new Apps Script deployment with Drive scope"
git push origin HEAD:main
echo "✅ Done!"
read -p "Press Enter to close..."
