#!/bin/bash
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"
cd /Users/kim_yeong/my-project/pubgm-jp-product
set -a; source .env; set +a

TODAY=$(date +'%Y-%m-%d')
SUCCESS_FLAG="/tmp/pubgm-jp-product-done-${TODAY}"

# 오늘 이미 성공했으면 스킵
if [ -f "$SUCCESS_FLAG" ]; then
  echo "$(date): Already updated today, skipping."
  exit 0
fi

# VPN/Databricks 연결 확인
if ! curl -sf --max-time 5 "https://krafton-hq.cloud.databricks.com" -o /dev/null 2>/dev/null; then
  echo "$(date): Databricks unreachable (VPN off?), skipping."
  exit 0
fi

set -e
echo "$(date): Generating dashboard..."
python3 generate.py

echo "$(date): Deploying to Cloudflare..."
npx wrangler deploy
echo "$(date): Done."

touch "$SUCCESS_FLAG"

git diff --quiet || (git add -u && git commit -m "Update cache ${TODAY}" && git push) || echo "$(date): git push skipped (no remote auth)"
