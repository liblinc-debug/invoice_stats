#!/bin/bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
REQUESTED_TARGET="${1:-auto}"
SYSTEM_NAME="$(uname -s)"

case "$SYSTEM_NAME" in
  Darwin)
    PLATFORM="macos"
    PYTHON_CMD=(python3)
    VENV_DIR="${TMPDIR:-/tmp}/invoice-tool-build-venv-$UID"
    VENV_PYTHON="$VENV_DIR/bin/python"
    BUNDLE_MODE=(--onedir)
    ;;
  MINGW*|MSYS*|CYGWIN*)
    PLATFORM="windows"
    if command -v py >/dev/null 2>&1; then
      PYTHON_CMD=(py -3)
    elif command -v python >/dev/null 2>&1; then
      PYTHON_CMD=(python)
    else
      echo "未找到 Python，请先安装 Python 3.9 或更高版本。" >&2
      exit 1
    fi
    VENV_DIR="$ROOT_DIR/.build-venv"
    VENV_PYTHON="$VENV_DIR/Scripts/python.exe"
    BUNDLE_MODE=(--onefile)
    ;;
  *)
    echo "暂不支持在 $SYSTEM_NAME 上构建 Windows 或 macOS 程序。" >&2
    exit 1
    ;;
esac

if [[ "$REQUESTED_TARGET" != "auto" && "$REQUESTED_TARGET" != "$PLATFORM" ]]; then
  echo "PyInstaller 不支持交叉编译：当前系统只能构建 $PLATFORM 版本。" >&2
  echo "请在目标系统上运行同一脚本。" >&2
  exit 1
fi

echo "正在准备 $PLATFORM 构建环境…"
"${PYTHON_CMD[@]}" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 9) else 1)' || {
  echo "Python 版本过低，请安装 Python 3.9 或更高版本。" >&2
  exit 1
}
"${PYTHON_CMD[@]}" -m venv "$VENV_DIR"
"$VENV_PYTHON" -m pip install --upgrade pip
"$VENV_PYTHON" -m pip install -r "$ROOT_DIR/发票统计工具/requirements-build.txt"

DIST_DIR="$ROOT_DIR/bin/$PLATFORM"
mkdir -p "$DIST_DIR" "$ROOT_DIR/build"

echo "正在构建 $PLATFORM 可执行程序…"
"$VENV_PYTHON" -m PyInstaller \
  --noconfirm \
  --clean \
  "${BUNDLE_MODE[@]}" \
  --windowed \
  --name "发票整理工具" \
  --distpath "$DIST_DIR" \
  --workpath "$ROOT_DIR/build/pyinstaller-$PLATFORM" \
  --specpath "$ROOT_DIR/build" \
  "$ROOT_DIR/发票统计工具/gui.py"

if [[ "$PLATFORM" == "macos" ]]; then
  OUTPUT="$DIST_DIR/发票整理工具.app"
else
  OUTPUT="$DIST_DIR/发票整理工具.exe"
fi

if [[ ! -e "$OUTPUT" ]]; then
  echo "构建失败：未找到 $OUTPUT" >&2
  exit 1
fi

echo
echo "构建完成：$OUTPUT"
