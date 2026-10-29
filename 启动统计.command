#!/bin/bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON="$ROOT_DIR/.venv/bin/python"

if [[ ! -x "$PYTHON" ]]; then
  echo "尚未安装依赖，请先双击“安装依赖.command”。"
  exit 1
fi

cd "$ROOT_DIR"
"$PYTHON" "$ROOT_DIR/发票统计工具/gui.py"
