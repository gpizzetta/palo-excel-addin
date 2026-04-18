# palo-excel-addin

Complément Excel minimal, **100 % statique** (HTML + JS dans `docs/`), sans Node ni build.

- Fonction **`=PALO.HELLO()`** → retourne la chaîne `hello world`.
- Manifeste : **`https://gpizzetta.github.io/palo-excel-addin/manifest.xml`** (GitHub Pages : branche `main`, dossier `/docs`).
- Contrôle du déploiement : **`https://gpizzetta.github.io/palo-excel-addin/`** — page `index.html` qui affiche la `<Version>` lue dans le manifeste publié (après `git push`, attendre ~1–2 min puis actualiser).

## Mises à jour et cache Excel

Office met souvent en **cache** le manifeste et les scripts. Pour forcer une nouvelle version :

1. **Incrémenter** `<Version>` dans `docs/manifest.xml` à chaque publication (ex. `1.0.1.0` → `1.0.2.0`). C’est la méthode recommandée par Microsoft pour signaler une mise à jour.
2. **Retirer** le complément puis **le recharger** (nouveau fichier manifest ou même URL Pages après `git push`).
3. Si rien ne change : fermer Excel, navigation privée, ou vider le cache Office (selon la plateforme).

Ne change `<Id>` **que** si tu veux un **nouveau** complément aux yeux d’Excel (sinon garde le même GUID).

Optionnel : après un déploiement, attendre 1–2 minutes (CDN GitHub Pages) avant de retester.

## Fichiers

| Fichier | Rôle |
|---------|------|
| `docs/manifest.xml` | Manifeste Office (`<Version>` à bump à chaque release) |
| `docs/functions.html` | Page hôte Office.js |
| `docs/functions.js` | `PALO.HELLO` |
| `docs/functions.json` | Métadonnées de la fonction |
| `docs/assets/` | Icônes **PNG** servies par Pages (`icon-16` … `icon-80`) — utilisées par le manifeste |
| `design/*.svg` | Sources vectorielles (**non** référencées par le manifeste ; voir `design/README.md`) |

Validation du manifeste (optionnel) :  
`npx office-addin-manifest validate docs/manifest.xml`

## Licence

MIT — voir `LICENSE`.
