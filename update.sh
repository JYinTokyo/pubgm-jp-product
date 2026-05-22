#!/bin/bash
set -e
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"
cd /Users/kim_yeong/my-project/pubgm-jp-product
set -a; source .env; set +a

echo "$(date): Generating dashboard..."
python3 generate.py

echo "$(date): Deploying to Cloudflare..."
npx wrangler deploy
echo "$(date): Done."

git diff --quiet || (git add -u && git commit -m "Update cache $(date +'%Y-%m-%d')" && git push) || echo "$(date): git push skipped (no remote auth)"
