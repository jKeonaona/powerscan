#!/usr/bin/env bash
set -e

echo "=== PowerScan Deploy ==="

cd /var/www/powerscan

git fetch origin

OLD_SHA=$(git rev-parse --short HEAD)
NEW_SHA=$(git rev-parse --short origin/main)
echo "Updating: ${OLD_SHA} → ${NEW_SHA}"

git reset --hard origin/main

source venv/bin/activate

echo "Installing pip dependencies..."
pip install -q -r requirements.txt

pm2 restart powerscan

sleep 2

echo ""
pm2 list | grep powerscan

echo ""
echo "=== Deploy complete ==="
