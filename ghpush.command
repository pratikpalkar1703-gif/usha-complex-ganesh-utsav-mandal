#!/bin/bash
export NVM_DIR="$HOME/.nvm"
export PATH="$NVM_DIR/versions/node/v20.20.2/bin:$PATH"
export PATH="/usr/local/bin:/opt/homebrew/bin:$PATH"

cd "/Users/pratikpalkar/ganesh chaturthi 2026"
rm -f .git/index.lock

# Try gh CLI first
if command -v gh &>/dev/null; then
  echo "✅ gh CLI found, using it to push..."
  gh auth status
  git push origin HEAD:main
else
  echo "❌ gh not found, trying GCM..."
  # Try using git credential manager
  git config --global credential.helper manager
  git push origin HEAD:main
fi

echo ""
echo "Done!"
read -p "Press Enter to close..."
