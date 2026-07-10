#!/bin/sh
# Genere functions.js ET functions-bundle.js (palo-api + functions-core.js).
# Sur Excel Desktop, le runtime CF charge souvent functions.js seul (sans importScripts).
# Il faut donc que functions.js deploye contienne tout le code Palo.
# Editer functions-core.js (source), puis : ./build-bundle.sh
set -e
cd "$(dirname "$0")"
SRC="./functions-core.js"
for OUT in functions.js functions-bundle.js; do
  {
    echo "/* Palo OLAP — genere depuis functions-core.js + palo-api.js. Ne pas editer. */"
    cat ./assets/palo-cf-polyfills.js
    cat ./assets/palo-api.js
    awk 'NR<=3 { print; next } /^\(function paloPreloadBundleForJsOnlyRuntime/,/^\}\)\(\);$/ { next } { print }' "$SRC"
  } > "$OUT"
  echo "OK: $OUT ($(wc -c < "$OUT" | tr -d ' ') octets)"
done
