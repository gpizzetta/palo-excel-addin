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

## 4. Limites de ce complément Office.js (état actuel)

- **`PALO.DATAC`** dans `docs/functions.js` ne fait que la **lecture** HTTP (`/cell/value`).
- **`PALO.SETDATA`** envoie la valeur au serveur via **`/cell/replace`** avec un mode **splash** numérique ; il **ne réimplémente pas** encore toute la machine à états C++ (`LIKE` / `COPY` / `##` / `!` / etc.) côté JavaScript.
- Pour avancer sans dupliquer tout le C++ tout de suite, le flux **Action → popup** prépare une formule **`PALO.SETDATA(...)`** avec la **chaîne utilisateur en premier argument** et le **même chemin** que `PALO.DATAC`, afin que l’écriture passe par le même pipeline serveur ; l’enrichissement complet du parser LIKE/COPY pourra suivre (ou s’appuyer sur une évolution serveur si disponible).

## 5. Comportement UI retenu (ce dépôt)

Lorsque l’utilisateur ouvre **Action** sur une formule **`PALO.DATAC`** :

1. Le popup affiche un **champ texte** pour saisir par exemple  
   `100 LIKE dimension1:élémA;dimension2:élémB`  
   ou  
   `COPY dimension1:élémA;dimension2:élémB`  
   (et variantes documentées côté Jedox / §2–3).
2. Un bouton envoie la chaîne au serveur via **`GET /cell/replace`** (même contrat que `PALO.SETDATA` dans `functions.js`) avec **splash `1`** (*default* HTTP), en utilisant le **chemin** dérivé des arguments de **`PALO.DATAC`**. La **formule de la cellule reste `=PALO.DATAC(...)`** ; elle n’est pas remplacée par `PALO.SETDATA`.
3. **Contrainte actuelle du popup** : base, cube et chaque coordonnée doivent être des **littéraux** `"…"` dans la formule (pas de références de cellules résolues côté dialogue). Sinon un message d’erreur invite à simplifier la formule ou à utiliser une autre voie.
4. Tant que le parser complet LIKE/COPY n’est pas porté en JS, le **comportement sur le cube** dépend surtout du **serveur** et de la chaîne envoyée telle quelle ; la doc §2–3 reste la **référence fonctionnelle** pour un portage ultérieur côté client.

## 6. Références externes

- [Jedox — Data functions](https://knowledgebase.jedox.com/jedox/planning/jedox-data-functions.htm) — rôles de `PALO.DATA`, `PALO.DATAC`, `PALO.SETDATA`, collecte.
- [Jedox — Splashing troubleshooting](https://knowledgebase.jedox.com/jedox/planning/splashing-troubleshooting.htm) — écriture consolidée / erreurs.
- Section **G** du [cahier des charges](./CAHIER_DES_CHARGES.md) — lien Excel 2010 / miroir C++.
