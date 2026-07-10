#!/bin/sh
# Met a jour le numero de version partout (source de verite : version.json).
# Usage : ./bump-version.sh 1.0.2.3
set -e
cd "$(dirname "$0")"
NEW="${1:-}"
if [ -z "$NEW" ]; then
  echo "Usage: $0 <version>   ex. $0 1.0.2.3" >&2
  exit 1
fi
OLD=$(grep -o '"version"[[:space:]]*:[[:space:]]*"[^"]*"' version.json | sed 's/.*"\([^"]*\)"$/\1/')
if [ -z "$OLD" ]; then
  echo "Impossible de lire version.json" >&2
  exit 1
fi
echo "Version $OLD -> $NEW"
DATE=$(date +%Y-%m-%d)
sed -i "s/\"version\"[[:space:]]*:[[:space:]]*\"$OLD\"/\"version\": \"$NEW\"/" version.json
sed -i "s/\"built\"[[:space:]]*:[[:space:]]*\"[^\"]*\"/\"built\": \"$DATE\"/" version.json
for f in manifest.xml shared-runtime.html commands.html functions.html taskpane.html \
  functions-core.js assets/taskpane.js commands.js; do
  if [ -f "$f" ]; then
    sed -i "s/$OLD/$NEW/g" "$f"
  fi
done
sed -i "s/1\.0\.2\.0/$NEW/g" taskpane.html 2>/dev/null || true
./build-bundle.sh
echo "OK version $NEW (ancienne $OLD)"
