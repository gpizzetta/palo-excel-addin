#!/bin/sh
# Promouvoir le BETA valide vers la PROD (racine docs/).
# A lancer seulement apres tests sur staging/manifest.xml
set -e
cd "$(dirname "$0")"

PROD_ID="9e35717b-2d76-4b84-b9f2-c7c1df86d901"
BETA_ID="b8c4e29a-3f61-4a5e-9d2b-7a1e0c4f8b32"
PROD_BASE="https://gpizzetta.github.io/palo-excel-addin/"
STAGING_BASE="https://gpizzetta.github.io/palo-excel-addin/staging/"

if [ ! -f staging/manifest.xml ]; then
  echo "Pas de docs/staging/. Lancez ./publish-staging.sh d'abord." >&2
  exit 1
fi

echo "Promotion staging -> prod (racine docs/)"

for item in staging/*; do
  name=$(basename "$item")
  if [ "$name" = "manifest.xml" ]; then
    continue
  fi
  rm -rf "$name"
  cp -a "$item" .
done

sed \
  -e "s|$STAGING_BASE|$PROD_BASE|g" \
  -e "s|<Id>$BETA_ID</Id>|<Id>$PROD_ID</Id>|" \
  -e 's|DefaultValue="Palo OLAP Add-in (BETA)"|DefaultValue="Palo OLAP Add-in"|g' \
  -e 's|<ProviderName>Palo OLAP Add-in (BETA)</ProviderName>|<ProviderName>Palo OLAP Add-in</ProviderName>|' \
  -e 's|id="Palo.Group.Actions.Label" DefaultValue="Palo BETA"|id="Palo.Group.Actions.Label" DefaultValue="Palo"|g' \
  -e 's|Ouvre la gestion des connexions Palo (canal BETA).|Ouvre la gestion des connexions Palo.|g' \
  staging/manifest.xml > manifest.xml

find . -maxdepth 2 -type f \( -name '*.js' -o -name '*.html' \) ! -path './staging/*' -exec sed -i \
  -e "s|$STAGING_BASE|$PROD_BASE|g" \
  -e 's|https://gpizzetta.github.io/palo-excel-addin/staging"|https://gpizzetta.github.io/palo-excel-addin"|' \
  {} +

sed -i \
  -e 's|<title>Palo OLAP Add-in (BETA)</title>|<title>Palo OLAP Add-in</title>|' \
  -e 's|<h2>Connexion Palo (BETA)</h2>|<h2>Connexion Palo</h2>|' \
  shared-runtime.html

VERSION=$(grep -o '"version"[[:space:]]*:[[:space:]]*"[^"]*"' version.json | sed 's/.*"\([^"]*\)"$/\1/')
echo ""
echo "OK prod v$VERSION"
echo "  Manifeste utilisateurs : ${PROD_BASE}manifest.xml"
echo "  Commit conseille         : git add -A docs/ && git commit -m \"release: v$VERSION\""
