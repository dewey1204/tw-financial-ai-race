#!/bin/bash
# 每日定時執行台灣金融業 AI 動態追蹤（由 macOS launchd 每天上午 09:00 自動觸發）
set -euo pipefail

PROJECT_DIR="/Users/dewey/git/tw-financial-ai-race"
LOG_DIR="$PROJECT_DIR/logs"
AGY_BIN="/Users/dewey/.local/bin/agy"
TS="$(date +"%Y-%m-%d_%H%M%S")"

mkdir -p "$LOG_DIR"
cd "$PROJECT_DIR"

# 設定 PATH 確保包含常用執行路徑（包含 Python 與 git）
export PATH="/Users/dewey/.local/bin:/usr/local/bin:/Library/Frameworks/Python.framework/Versions/3.13/bin:/usr/bin:/bin:/usr/sbin:/sbin:$PATH"

LOG_FILE="$LOG_DIR/tracker_${TS}.log"

echo "=== [${TS}] 開始執行每日金融業 AI 動態追蹤 ===" | tee -a "$LOG_FILE"

"$AGY_BIN" -p "請執行 tw-fin-ai-tracker 技能（參考 .agents/skills/tw-fin-ai-tracker/SKILL.md）：搜尋過去 24 小時台灣各大金控與旗下金融機構之最新 AI 與 Agent 應用進展，產出 daily_reports/YYYY-MM-DD.md，透過 python3 .agents/skills/tw-fin-ai-tracker/scripts/update_excel.py 寫入 金融業AI場景應用.xlsx，並 git commit 與 push 同步儀表板。" \
  --dangerously-skip-permissions \
  --print-timeout 15m0s \
  >> "$LOG_FILE" 2>&1

echo "=== [$(date +"%Y-%m-%d_%H%M%S")] 執行完成 ===" | tee -a "$LOG_FILE"
