# LIKE / COPY (héritage Jedox) et action « Palo » sur `PALO.DATAC`

Ce document décrit le **comportement historique** du client tableur (bibliothèque **PaloSpreadsheetFuncs**, Jedox ~5.1, Excel 2010 / COM) pour les chaînes **`LIKE`** et **`COPY`**, et la **reprise prévue** dans ce complément Office.js : depuis le popup **Action** sur une cellule en `PALO.DATAC`, l’utilisateur saisit une commande du même style que sous l’ancien plugin.

## 1. Où est le code d’origine (C++) ?

**Référence canonique** (code de l’ancien plugin Excel 2010, COM / Jedox ~5.1) — URL à conserver dans la documentation du projet :

[https://github.com/gpizzetta/jedox-mirror/tree/master/molap/client_libraries/5.1/PaloSpreadsheetFuncs](https://github.com/gpizzetta/jedox-mirror/tree/master/molap/client_libraries/5.1/PaloSpreadsheetFuncs)

(Dépôt **jedox-mirror**, dossier **`PaloSpreadsheetFuncs`**.)

- **`SpreadsheetFuncsBase::FPaloSetdata`** — lorsque la **valeur** écrite est une **chaîne** (et non un nombre seul), le client parse des **jetons** (séparés par des espaces) et applique des modes spéciaux dont **`LIKE`** et **`COPY`** (voir `parseCopyParams`, `parsePath`, `CellCopyWrapper` dans `SpreadsheetFuncsBase.cpp`).
- **`PALO.DATAC`** côté historique correspond surtout à **`FPaloGetdataC`** (lecture par noms d’éléments) ; la sémantique **`LIKE` / `COPY`** n’est **pas** dans `FPaloGetdataC`, mais dans la **voie d’écriture** `FPaloSetdata` / `FPaloSetdataA`.

En résumé : sous l’ancien Excel, l’utilisateur voyait souvent une cellule **liée au cube** (affichage type données) et la **saisie** d’une chaîne interprétée comme commande d’écriture ; dans le moteur client, c’est la branche **chaîne** de **`FPaloSetdata`** qui implémente **`LIKE`** et **`COPY`**.

## 2. Chaîne saisie : jetons et mot-clé

Le C++ découpe la valeur texte en **jetons** (espaces ; certains cas fusionnent `;` entre jetons). Les mots-clés reconnus incluent notamment :

| Jeton | Rôle |
|--------|------|
| **`LIKE`** | Active le mode « copier / ajuster une valeur **depuis** une autre coordonnée du cube » (voir §3). |
| **`COPY`** | Variante « copie » sans la logique `LIKE` numérique (voir `parseCopyParams` : `COPY` vs `LIKE`). |
| **`WITHRULES`** | Indique d’utiliser les règles lors de l’opération de copie côté API cube. |
| **`PREDICT` / `PREDICTLR`** | Autres branches (régression linéaire), hors périmètre du premier portage. |

Le **premier** jeton numérique ou spécial (`#…`, `##…`, `!…`, etc.) joue le rôle de **valeur** ou de modificateur selon le cas ; le détail exact est dans `parseCopyParams` et la suite de `FPaloSetdata` (splashing `#`, `!`, `!!`, pourcentages, etc.).

## 3. Chemin « modèle » après `LIKE` ou `COPY` : `dim1;dim2` et `dimension:élément`

Fonction **`parsePath`** (`SpreadsheetFuncsBase.cpp`) : la partie **chemin source** (après les mots-clés dans les jetons) est découpée par **`;`**.

Chaque segment peut être :

1. **Un seul nom** `Element` — le client cherche **dans quelle dimension** cet élément existe (unicité requise) et remplace la coordonnée correspondante dans une copie du **chemin cible** (celui de la cellule éditée).
2. **`NomDimension:Valeur`** — position explicite : la dimension nommée reçoit la valeur indiquée pour construire le chemin **source** utilisé par la copie.

Le résultat est un tableau `coords` (coordonnées complètes du cube) passé à **`CellCopyWrapper`** / **`CellCopy`** : copie depuis `coords` vers le **path** de la cellule cible, avec options (`LIKE` + pourcentage, `LIKE` + addition, etc.).

Exemples de **saisie** (conceptuellement, comme sous l’ancien plugin) :

- `100 LIKE Réalisé;2025` — si `Réalisé` et `2025` sont résolus sans ambiguïté sur les dimensions du cube, copie / ajustement avec valeur **100** et chemin résolu.
- `COPY DimensionComptable:Actif;Année:2025` — copie depuis le chemin explicite vers la cellule courante.

Les séparateurs et guillemets suivent les mêmes règles que le tokenizer C++ (`StringTokenizer`, `unQuote`).

## 4. Limites côté API HTTP

- **`PALO.DATAC`** dans `docs/functions.js` ne fait que la **lecture** HTTP (`/cell/value`).
- **`PALO.SETDATA`** envoie une **valeur** via **`/cell/replace`** (splash numérique). Le paramètre **`value`** de cette route est **numérique** : envoyer la chaîne brute **`300 like 2025`** provoque une erreur du type *conversion failed* (comportement normal).
- Les commandes **`LIKE`** / **`COPY`** du plugin 2010 sont interprétées **côté client** (C++), puis traduites en appels cube (dont **`CellCopy`** avec valeur cible pour LIKE). Voir `cell_copy.api` dans [jedox-mirror](https://github.com/gpizzetta/jedox-mirror/blob/master/molap/server/5.1/Api/cell_copy.api) : **`GET /cell/copy`** accepte **`name_path`** (source), **`name_path_to`** (cible) et un **`value`** optionnel (*« The numeric value of the target cube cell »*) — c’est ce qui matérialise le **LIKE** avec nouvelle valeur agrégée.

## 5. Comportement UI retenu (ce dépôt)

Lorsque l’utilisateur ouvre **Action** sur une formule **`PALO.DATAC`** :

1. Parcours **guidé** : **définir** / **additionner** (`/cell/replace` + splash si consolidé) et **copier** depuis un chemin explicite (`/cell/copy` sans `value`).
2. **Mode avancé** (détails LIKE/COPY) : le popup reconnaît une forme **`nombre like chemin_partiel`** ou **`copy chemin_partiel`** (insensible à la casse pour `like` / `copy`). Voir **§ 5.2**.
3. **Contrainte** : base, cube et chaque coordonnée du **`PALO.DATAC`** doivent être des **littéraux** `"…"` pour les appels HTTP depuis le dialogue (sauf résolution `datac_r` côté commandes). Sinon un message d’erreur invite à simplifier la formule.
4. La **formule** `=PALO.DATAC(...)` **n’est pas modifiée** après une écriture depuis le popup.

### 5.1. Modes splash HTTP (`/cell/replace`, paramètre `splash` 0–5)

Sur un chemin **consolidé**, le popup propose une liste alignée sur **`normalizeSplashMode`** dans `docs/functions.js` (entiers attendus par l’API HTTP Jedox). Ce ne sont **pas** les préfixes `#` / `!` / `!!` saisis au clavier dans le tableur, mais la **sémantique** est proche — voir [Splashing overview](https://knowledgebase.jedox.com/jedox/planning/splashing-overview.htm).

| `splash` | Alias texte (`PALO.SETDATA` / `normalizeSplashMode`) | Rôle (écriture sur consolidé) |
|----------|------------------------------------------------------|-------------------------------|
| **0** | — | Pas de splash ; pas de décomposition automatique ; souvent inadapté à un consolidé « pur » sans autre option. |
| **1** | `default` | **Default** : comportement splash **par défaut du serveur** — la valeur est répartie sur les cellules **de base** sous le consolidé selon la **logique Jedox** (poids de consolidation, règles internes). Ce n’est **pas** strictement une division **égale** sur toutes les bases (`#` seul en saisie manuelle). En pratique la répartition peut **ressembler** à une répartition **pondérée** ou liée à l’**existant** selon version ; le détail exact = **doc serveur / version**. |
| **2** | `add`, `add_base` | Comme **`!!`** ([doc Jedox](https://knowledgebase.jedox.com/jedox/planning/splashing-overview.htm)) : **la même valeur est ajoutée à chaque** cellule de base liée au consolidé — **pas** une répartition du total entre enfants (pour « splitter » une valeur sur les bases, viser **`#`** / mode **1**). |
| **3** | `set`, `set_base` | Comme **`!`** : **la même valeur est écrite sur chaque** cellule de base (toutes les bases liées reçoivent la valeur saisie). |
| **4** | `set_populated` | Comme **`!#`** ([variations #](https://knowledgebase.jedox.com/jedox/planning/splash-parameter-hashtag.htm)) : n’écrit que sur les cellules de base **déjà peuplées** pour ce consolidé et cette tranche du cube. |
| **5** | `add_populated` | Comme **`!!#`** : n’ajoute que sur les bases **déjà peuplées**. |

**Modes 4 et 5 — erreur sur consolidé** : ce n’est **pas** un manque de **path de référence** (à la `LIKE`) ; l’URL reste le chemin **`PALO.DATAC`** courant. Si **aucune** intersection de base sous le consolidé n’a encore de valeur (tout à zéro / vide), Jedox n’a **aucune cible** pour `!#` / `!!#` et le serveur peut renvoyer une erreur — comportement attendu. Il faut d’abord **peupler** au moins une base (ou utiliser **1–3** selon le besoin). En second lieu, vérifier que votre **version** OLAP accepte bien `splash=4` et `splash=5` sur `/cell/replace`.

**Ordre HTTP vs C++** : pour **2** et **3**, l’API HTTP utilise **2 = add base** et **3 = set base**, alors que certains commentaires C++ du miroir Palo inversent les libellés `MODE_SPLASH_SET` / `MODE_SPLASH_ADD` — le complément suit **HTTP** pour l’URL `splash=`.

### 5.2. LIKE / COPY en mode avancé (chemins partiels → `/cell/copy`)

Aligné sur la [doc Jedox LIKE](https://knowledgebase.jedox.com/jedox/planning/splashing-command-like.htm) et sur **`parsePath`** côté **`PaloSpreadsheetFuncs`** (§3) :

- La **cible** est toujours le **chemin complet** de la cellule courante (arguments `PALO.DATAC` après base et cube).
- Le texte après **`like`** ou **`copy`** est une liste de segments séparés par **`;`**. Chaque segment est soit **`NomDimension:élément`**, soit un **seul nom d’élément** (le client cherche dans **quelle dimension** ce nom existe ; **unicité** requise, sinon message invitant à utiliser `Dimension:élément`).
- Les dimensions **non mentionnées** dans ces segments reprennent les **mêmes éléments** que sur le chemin **cible** (ex. `300 like 2025` sur `2024,demographie,titi` → source `2025,demographie,titi` si `2025` n’existe que sur la dimension temps).

**LIKE** : après construction du chemin **source** et **cible**, le popup appelle **`GET /cell/copy`** avec **`name_path`** = source (virgules), **`name_path_to`** = cible, **`function=0`**, et **`value`** = le nombre saisi (ex. **300**). C’est l’équivalent fonctionnel du « LIKE avec nouvelle valeur » côté tableur.

**COPY** : même appel **sans** paramètre **`value`** (copie des valeurs source vers la cible). La case **use_rules** du bloc « Copier depuis un autre chemin » s’applique aussi au mode avancé.

**Négatif** (Jedox) : saisie du type **`'-300 like …`** (apostrophe devant le nombre) est reconnue pour le signe.

**Non couvert pour l’instant** : ordre libre des jetons (`like 300 2025`), `! like` / `!! like`, `WITHRULES`, `PREDICT`, pourcentages et autres branches de **`FPaloSetdata`** — à étendre en s’appuyant sur le C++ si besoin.

## 6. Références externes

- [Jedox — Splashing command Like](https://knowledgebase.jedox.com/jedox/planning/splashing-command-like.htm) — chemins partiels, ambiguïté `Dimension:élément`.
- [Jedox — Data functions](https://knowledgebase.jedox.com/jedox/planning/jedox-data-functions.htm) — rôles de `PALO.DATA`, `PALO.DATAC`, `PALO.SETDATA`, collecte.
- [Jedox — Splashing overview](https://knowledgebase.jedox.com/jedox/planning/splashing-overview.htm) — `#`, `!`, `!!`, `!#`, `!!#`, LIKE.
- [Jedox — Splashing troubleshooting](https://knowledgebase.jedox.com/jedox/planning/splashing-troubleshooting.htm) — écriture consolidée / erreurs.
- Section **G** du [cahier des charges](./CAHIER_DES_CHARGES.md) — lien Excel 2010 / miroir C++.
