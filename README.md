# palo-excel-addin

Complément Excel minimal, **100 % statique** (HTML + JS dans `docs/`), sans Node ni build.

- **`=PALO.HELLO()`** → `hello world` (test minimal).
- **Ruban → onglet « Palo » → groupe « Serveur » → « Connexion »** : volet **URL**, **utilisateur**, **mot de passe**, **Enregistrer** (paramètres du classeur). Sur **Excel pour Microsoft 365 (Windows / Mac)**, l’onglet **Palo** apparaît à côté des onglets du ruban. **Excel dans le navigateur** : les commandes ruban des compléments **téléversés** sont souvent **invisibles ou incomplètes** — ouvrir le complément via **Insertion → Compléments → Palo** (volet Connexion), ou utiliser **Excel bureau**.
- **`=PALO.VERSION()`** → numéro de version du **bundle JS** chargé (à comparer à `<Version>` du manifeste). Utile si **`#NOM?`** sur d’autres fonctions : cache Excel / ancien script — retirer le complément, repousser, retélécharger le manifeste.
- **`=PALO.INFO("https://hôte:port/chemin")`** → `GET` en **CORS** vers l’URL ; statut HTTP ou erreur (CORS, réseau, TLS). Palo en **HTTP** seul peut échouer depuis Excel Online ; **HTTPS** + CORS côté serveur souvent nécessaires.

À chaque release, aligner **`ADDIN_VERSION`** dans `docs/functions.js`, **`?v=…`** dans `docs/manifest.xml`, `docs/functions.html`, `docs/taskpane-connection.html`, et `<Version>` du manifeste (même numéro, ex. `1.0.6.0`).
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
| `docs/functions.json` | Métadonnées des fonctions |
| `docs/commands.html` | Commandes du ruban |
| `docs/taskpane-connection.html` + `taskpane-connection.js` | Volet Connexion (URL / user / mot de passe) |
| `docs/assets/` | Icônes **PNG** (`icon-16` … `icon-80`) |
| `design/*.svg` | Sources vectorielles (**non** référencées par le manifeste ; voir `design/README.md`) |

Validation du manifeste (optionnel) :  
`npx office-addin-manifest validate docs/manifest.xml`

## Licence

MIT — voir `LICENSE`.
