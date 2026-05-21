#!/bin/bash
set -e
cd /Users/kim_yeong/my-project/pubgm-jp-product
set -a; source .env; set +a

echo "$(date): Generating dashboard..."
python3 generate.py

git diff --quiet || (git add -u && git commit -m "Update cache $(date +'%Y-%m-%d')" && git push)

echo "$(date): Deploying to Cloudflare..."
npx wrangler deploy
echo "$(date): Done."
