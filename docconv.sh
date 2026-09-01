#!/usr/bin/env bash
# 문서 변환기 실행 (macOS / Linux)
set -euo pipefail
export PYTHONIOENCODING=utf-8
cd "$(dirname "$0")"

if [ "${1:-}" = "--gui" ] || [ $# -eq 0 ]; then
    exec python3 -m docconv --gui
fi
exec python3 -m docconv "$@"
