#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT_DIR"

echo "[FERA build] Iniciando builds Linux..."
"$ROOT_DIR/fera_linux_partial.sh"
"$ROOT_DIR/fera_linux_full.sh"
echo "[FERA build] Builds concluidos em ./dist/FERA_LINUX e ./dist/FERA_FULL_LINUX"
