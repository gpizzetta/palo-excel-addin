#!/bin/sh
# Genere functions-bundle.js (palo-api + functions) pour le fallback Excel Desktop V1.0.
# Sans npm : a lancer avant chaque deploiement si functions.js ou palo-api.js changent.
set -e
cd "$(dirname "$0")"
OUT="functions-bundle.js"
{
  echo "/* Palo OLAP — bundle genere (ne pas editer). Voir build-bundle.sh */"
  cat ./assets/palo-api.js
  # Conserver PALO_CDN_BASE / PALO_ASSET_VERSION ; sauter seulement le preload importScripts.
  awk 'NR<=3 { print; next } /^\(function paloPreloadBundleForJsOnlyRuntime/,/^\}\)\(\);$/ { next } { print }' ./functions.js
} > "$OUT"
echo "OK: $OUT ($(wc -c < "$OUT" | tr -d ' ') octets)"
