#!/usr/bin/env bash
# 一鍵更新：拉最新碼 → 推到你自己的 Apps Script → 重新部署同一個網頁 App。
# 不會動到你的試算表/資料，也不用重跑 setup()。需先裝好 clasp 並 `clasp login`。
set -e
cd "$(dirname "$0")"

echo "▸ 拉取最新程式碼…"
git pull --ff-only

echo "▸ 推到你的 Apps Script…"
clasp push -f

# 自動找出網頁 App 的版本化部署 ID（排除會自動跟最新碼的 @HEAD），重新部署成新版本，
# 這樣 /exec 網址不變、iPhone 上的 App 直接更新。
DEP="$(clasp deployments | grep -v '@HEAD' | grep -oE 'AKfyc[A-Za-z0-9_-]+' | head -1 || true)"
if [ -n "$DEP" ]; then
  echo "▸ 重新部署：$DEP"
  clasp deploy -i "$DEP"
  echo "✓ 完成。iPhone 上的 App 重新整理即是新版。"
else
  echo "⚠ 找不到版本化部署。請先在 Apps Script 部署一次網頁 App（部署 → 新增部署作業 → 網頁應用程式）。"
  exit 1
fi
