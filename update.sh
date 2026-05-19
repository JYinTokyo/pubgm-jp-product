#!/bin/bash
set -e
cd /Users/kim_yeong/my-project/pubgm-jp-product
set -a; source .env; set +a

echo "$(date): Generating dashboard..."
python3 generate.py

git add dist/index.html product_cache.json 2>/dev/null || true
git diff --staged --quiet || (git commit -m "Update dashboard $(date +'%Y-%m-%d')" && git push)

echo "$(date): Deploying to Cloudflare..."
npx wrangler deploy
echo "$(date): Done."
