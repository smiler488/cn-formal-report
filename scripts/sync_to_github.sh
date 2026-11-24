#!/usr/bin/env bash
# 提交并同步 cn-formal-report 到 GitHub。
# 用法：scripts/sync_to_github.sh "fix: 简述本次规则变更"
set -euo pipefail
REPO_DIR="$(cd "$(dirname "$0")/.." && pwd)"
cd "$REPO_DIR"
MSG="${1:-chore: update cn-formal-report rules}"

if ! git rev-parse --is-inside-work-tree >/dev/null 2>&1; then
  echo "尚未初始化 git 仓库，请先执行：git init && git remote add origin https://github.com/smiler488/cn-formal-report.git" >&2
  exit 1
fi

git add -A
if git diff --cached --quiet; then
  echo "无变更，跳过提交"
  exit 0
fi

git commit -m "$MSG"
git push origin HEAD
echo "已同步：$MSG"
