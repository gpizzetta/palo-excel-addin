<<<<<<< HEAD
# palo-excel-addin
=======
# palo-excel-addin

Complément Excel minimal, **100 % statique** (HTML + JS dans `docs/`), sans Node ni build.

- Fonction **`=PALO.HELLO()`** → retourne la chaîne `hello world`.
- Manifeste : **`https://gpizzetta.github.io/palo-excel-addin/manifest.xml`** (après activation de GitHub Pages sur la branche `main`, dossier `/docs`).

## Fichiers

| Fichier | Rôle |
|---------|------|
| `docs/manifest.xml` | Manifeste Office |
| `docs/functions.html` | Page hôte Office.js |
| `docs/functions.js` | `PALO.HELLO` |
| `docs/functions.json` | Métadonnées de la fonction |
| `docs/assets/` | Icônes PNG |

Validation du manifeste (optionnel, avec outil Node installé ailleurs) :  
`npx office-addin-manifest validate docs/manifest.xml`

## Licence

MIT — voir `LICENSE`.
>>>>>>> d3d2746 (Minimal static add-in: PALO.HELLO(), remove Node tooling)
