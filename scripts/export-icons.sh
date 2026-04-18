#!/usr/bin/env bash
# Régénère docs/assets/icon-{16,32,64,80}.png depuis design/palo_connect.svg
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"
command -v inkscape >/dev/null || { echo "Installez inkscape." >&2; exit 1; }
for s in 16 32 64 80; do
  inkscape "design/palo_connect.svg" -o "docs/assets/icon-${s}.png" -w "$s" -h "$s" --export-type=png
  echo "docs/assets/icon-${s}.png"
done
