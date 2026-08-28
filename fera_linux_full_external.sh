#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT_DIR"

DIST_DIR="./dist"
SOURCE_DIR="$DIST_DIR/fera_full_external"
TARGET_DIR="$DIST_DIR/FERA_FULL_LINUX_EXTERNAL"

mkdir -p "$DIST_DIR"

echo "[FERA build] Buildando FERA_FULL_LINUX_EXTERNAL..."
pyinstaller --clean --noconfirm feralinux_full_noconsole_external.spec

rm -rf "$TARGET_DIR"
mv "$SOURCE_DIR" "$TARGET_DIR"

if [ -d "$DIST_DIR/printer_interface_linux" ]; then
    rm -rf "$TARGET_DIR/printer_interface_linux"
    cp -R "$DIST_DIR/printer_interface_linux" "$TARGET_DIR/printer_interface_linux"
else
    echo "[FERA build] Aviso: $DIST_DIR/printer_interface_linux nao encontrado."
fi

if [ -f "./FERA.pdf" ]; then
    cp -f "./FERA.pdf" "$TARGET_DIR/FERA.pdf"
elif [ -f "$DIST_DIR/FERA.pdf" ]; then
    cp -f "$DIST_DIR/FERA.pdf" "$TARGET_DIR/FERA.pdf"
else
    echo "[FERA build] Aviso: FERA.pdf nao encontrado."
fi

echo "[FERA build] Concluido: $TARGET_DIR"
