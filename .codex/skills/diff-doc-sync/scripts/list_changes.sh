#!/usr/bin/env bash
set -euo pipefail

# git差分を一覧表示するユーティリティ
# オプション:
#   --cached    : ステージ済みのみ
#   --working   : 未ステージのみ
#   --base <rev>: 比較元（デフォルトHEAD）

mode=""
base="HEAD"

while [[ $# -gt 0 ]]; do
  case "$1" in
    --cached) mode="--cached"; shift ;;
    --working) mode=""; base=""; shift ;; # working tree vs index
    --base) base="$2"; shift 2 ;;
    *) break ;;
  esac
done

if ! command -v git >/dev/null 2>&1; then
  echo "git not found" >&2
  exit 1
fi

if [[ -z "$base" && -z "$mode" ]]; then
  # working tree diff vs index
  git diff --name-status
else
  git diff ${base:+${base}} ${mode} --name-status
fi
