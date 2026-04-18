# Icônes sources (SVG)

Fichiers **SVG** conservés dans le dépôt pour l’édition (Inkscape, etc.).  
Le manifeste Office utilise uniquement les **PNG** dans `docs/assets/` (pas les SVG).

## Fichiers

| Fichier | Rôle |
|---------|------|
| `palo_connect.svg` | Source principale pour les icônes PNG du complément |
| `palo.svg` | Variante / ressource graphique |

## Régénérer les PNG

```bash
./scripts/export-icons.sh
```

Ou manuellement à partir de `palo_connect.svg` :

```bash
cd "$(git rev-parse --show-toplevel)"
for s in 16 32 64 80; do
  inkscape "design/palo_connect.svg" -o "docs/assets/icon-${s}.png" -w "$s" -h "$s" --export-type=png
done
```

Puis incrémente `<Version>` dans `docs/manifest.xml` et pousse sur `main`.
