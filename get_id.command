#!/bin/bash
export NVM_DIR="$HOME/.nvm"
export PATH="$NVM_DIR/versions/node/v20.20.2/bin:$PATH"

cd "/Users/pratikpalkar/ganesh chaturthi 2026"

# Get the new deployment ID from clasp list
echo "Listing deployments..."
clasp deployments 2>&1 | tee "/Users/pratikpalkar/ganesh chaturthi 2026/deployments.txt"
echo ""
echo "Saved to deployments.txt"
read -p "Press Enter to close..."
