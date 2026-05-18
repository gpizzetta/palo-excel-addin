#!/bin/sh
# Genere functions-bundle.js (palo-api + functions) pour le fallback Excel Desktop V1.0.
# Sans npm : a lancer avant chaque deploiement si functions.js ou palo-api.js changent.
set -e
cd "$(dirname "$0")"
OUT="functions-bundle.js"
{
  echo "/* Palo OLAP — bundle genere (ne pas editer). Voir build-bundle.sh */"
  cat ./assets/palo-api.js
  # Sauter le bloc importScripts en tete de functions.js (deja inclus via palo-api).
  awk 'BEGIN{skip=0} /^\/\* global CustomFunctions/{skip=1} skip && /^\(function paloFunctionsBootstrap/{skip=0} !skip {print}' ./functions.js
} > "$OUT"
echo "OK: $OUT ($(wc -c < "$OUT" | tr -d ' ') octets)"
