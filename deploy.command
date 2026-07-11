#!/bin/bash
# Load nvm node 20
export NVM_DIR="$HOME/.nvm"
export PATH="$NVM_DIR/versions/node/v20.20.2/bin:$PATH"

cd ~/Downloads/ganpati

echo "📦 Pushing Apps Script..."
clasp push --force
clasp deploy -i "AKfycbxSBb-UfpneLJ0oNU7qr0Ow4SEcmPXxxbHBN48o6ajtpQjYlZ1CZ6yiWphktOc3rVA_jw" -d "update $(date '+%d %b %Y %H:%M')"
echo "✅ Apps Script deployed!"

echo ""
echo "🌐 Pushing to GitHub..."
git add .
git commit -m "update $(date '+%d %b %Y %H:%M')"
git push --force origin HEAD:main
echo "✅ Site deployed!"

echo ""
echo "🎉 All done! Site updates in ~1 minute."
echo "   https://pratikpalkar1703-gif.github.io/usha-complex-ganesh-utsav-mandal/"
read -p "Press Enter to close..."
