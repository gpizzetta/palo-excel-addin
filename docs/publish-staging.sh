#!/bin/sh
# Publie une copie BETA dans docs/staging/ sans toucher la prod (racine docs/).
# Les testeurs chargent : https://gpizzetta.github.io/palo-excel-addin/staging/manifest.xml
set -e
cd "$(dirname "$0")"

PROD_ID="9e35717b-2d76-4b84-b9f2-c7c1df86d901"
BETA_ID="b8c4e29a-3f61-4a5e-9d2b-7a1e0c4f8b32"
PROD_BASE="https://gpizzetta.github.io/palo-excel-addin/"
STAGING_BASE="https://gpizzetta.github.io/palo-excel-addin/staging/"

echo "Publication BETA -> docs/staging/"

rm -rf staging
mkdir staging

for item in *; do
  case "$item" in
    staging|publish-staging.sh|promote-to-production.sh|manifest.local.xml|manifest.staging.xml)
      continue
      ;;
    *)
      cp -a "$item" staging/
      ;;
  esac
done

# Manifest BETA : autre Id, URLs /staging/, libellés ruban et volet
sed \
  -e "s|$PROD_BASE|$STAGING_BASE|g" \
  -e "s|<Id>$PROD_ID</Id>|<Id>$BETA_ID</Id>|" \
  -e 's|DefaultValue="Palo OLAP Add-in"|DefaultValue="Palo OLAP Add-in (BETA)"|g' \
  -e 's|<ProviderName>Palo OLAP Add-in</ProviderName>|<ProviderName>Palo OLAP Add-in (BETA)</ProviderName>|' \
  -e 's|id="Palo.Group.Actions.Label" DefaultValue="Palo"|id="Palo.Group.Actions.Label" DefaultValue="Palo BETA"|g' \
  -e 's|Ouvre la gestion des connexions Palo.|Ouvre la gestion des connexions Palo (canal BETA).|g' \
  manifest.xml > staging/manifest.xml

# Assets staging : URLs CDN + libellés volet
find staging -type f \( -name '*.js' -o -name '*.html' \) -exec sed -i \
  -e "s|$PROD_BASE|$STAGING_BASE|g" \
  -e 's|https://gpizzetta.github.io/palo-excel-addin"|https://gpizzetta.github.io/palo-excel-addin/staging"|' \
  {} +

sed -i \
  -e 's|<title>Palo OLAP Add-in</title>|<title>Palo OLAP Add-in (BETA)</title>|' \
  -e 's|<h2>Connexion Palo</h2>|<h2>Connexion Palo (BETA)</h2>|' \
  staging/shared-runtime.html

VERSION=$(grep -o '"version"[[:space:]]*:[[:space:]]*"[^"]*"' version.json | sed 's/.*"\([^"]*\)"$/\1/')
echo ""
echo "OK staging v$VERSION"
echo "  Manifeste testeurs : ${STAGING_BASE}manifest.xml"
echo "  Ruban              : groupe « Palo BETA »"
echo "  Commit conseille    : git add staging/ && git commit -m \"staging: v$VERSION\""
echo "  La prod (racine) n'est pas modifiee tant que vous n'appelez pas ./promote-to-production.sh"
