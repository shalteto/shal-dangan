#!/usr/bin/env bash
set -euo pipefail

# ripgrepでTODOコメントを収集しファイルへ保存する
if ! command -v rg >/dev/null 2>&1; then
  echo "rg is required (ripgrep not found)" >&2
  exit 1
fi

out="${1:-tmp/todo-scan.txt}"
mkdir -p "$(dirname "$out")"

rg --no-heading --line-number --glob '!.venv/**' --glob '!skills/**' --glob '!node_modules/**' 'TODO' . >"$out"

echo "saved: $out"
