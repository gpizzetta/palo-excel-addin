# Cahier des charges — écriture sur une cellule PALO.DATAC

Référence de développement. Le contrat est le comportement de Jedox 5.1, pas la saisie texte.

Source auditée : `gpizzetta/jedox-mirror`, `molap/client_libraries/5.1/PaloSpreadsheetFuncs/SpreadsheetFuncsBase.cpp`, fonction `FPaloSetdata` (ligne 886) et `parseCopyParams` (ligne 1058, mot `LIKE` ligne 1087).

`PALO.DATAC` ne parse pas ces chaînes. C’est une lecture. La chaîne est la valeur d’écriture passée à `FPaloSetdata` (double-clic historique, `PALO.SETDATA`, `PALO.SETDATAIF`, `PALO.SETDATAA`). `FPaloSetdataBulk` (ligne 2121) n’interprète aucun préfixe : il envoie les valeurs telles quelles avec un seul splash.

L’interface du plugin peut être des listes et des boutons. Chaque ligne ci-dessous doit produire le même appel serveur.

## 1. Décision avant tout préfixe

La valeur Excel est soit un nombre, soit une chaîne.

| Valeur Excel | Cellule Palo actuelle | Action |
|---|---|---|
| Nombre | quelconque | `/cell/replace` de ce nombre, splash choisi par l’appelant |
| Chaîne | type string | `/cell/replace` de la chaîne brute, splash `NONE`. Aucun `like`, `!!` ou `#` n’est lu |
| Chaîne vide | type numérique | `/cell/replace` de `0`, splash choisi par l’appelant |
| Chaîne non vide | type numérique | grammaire des sections 2 à 5 |
| autre type | | erreur |

La cellule courante est lue avant d’interpréter la chaîne (`CellExport`). Si cette lecture échoue, l’écriture est refusée.

## 2. Découpage de la chaîne

Uniquement quand la cellule Palo est numérique et la saisie non vide.

1. Découper sur les espaces.
2. Retirer les guillemets du premier jeton.
3. Jeter les jetons vides.
4. Si un jeton se termine par `;` et qu’un jeton suit, les coller. `2025; demo` devient `2025;demo`.
5. S’il ne reste qu’un jeton, appliquer la section 3, dans cet ordre.
6. Sinon, appliquer la section 4 (`LIKE`, `COPY`, `PREDICT`).
7. Si la section 4 ne reconnaît pas la commande, relire la chaîne entière comme un nombre (section 6) et faire `/cell/replace` avec le splash de l’appelant. Un `%` sur cette voie est une erreur.

## 3. Un seul jeton : préfixes

Le nombre suit la section 6. `cur` est la valeur numérique déjà consolidée sur la cellule.

| Saisie | Calcul côté client | Appel | Splash |
|---|---|---|---|
| `##10` | `cur + 10` | `/cell/replace` | `DEFAULT` |
| `##10%` | `cur * (100 + 10) / 100` | `/cell/replace` | `DEFAULT` |
| `#10` | `10` | `/cell/replace` | `DEFAULT` |
| `#10%` | `cur * 10 / 100` | `/cell/replace` | `DEFAULT` |
| `\|10` | `10` | `/cell/replace` | `DEFAULT` |
| `\|10%` | | erreur | |
| `!10` | `10` | `/cell/replace` | `SET` |
| `!10%` | | erreur | |
| `!!10` | `10` | `/cell/replace` | `ADD` |
| `!!10%` | | erreur | |
| `?10` | | `/cell/goalseek`, type `GOALSEEK_COMPLETE`, zone vide | aucun |
| `?10%` | | erreur | |

L’ordre de test est obligatoire : `##` avant `#`, `!!` avant `!`.

`#10` et `|10` ne calculent pas une part égale dans le client. Ils envoient le nombre en splash `DEFAULT`. La répartition (poids si le parent vaut 0, proportions si les bases sont remplies, vidage si la valeur est 0) est celle du serveur.

`##` et les formes `%` calculent d’abord le nouveau total, puis le passent en splash `DEFAULT`.

`!` écrit le montant entier sur chaque base (`SET`). `!!` ajoute le montant entier à chaque base (`ADD`).

Noms de splash du code 5.1, et entiers déjà utilisés par le plugin :

| Nom 5.1 | Entier plugin | Effet sur un regroupement |
|---|---|---|
| `NONE` | 0 | écriture refusée |
| `DEFAULT` | 1 | répartit le total |
| `ADD` | 2 | ajoute le montant à chaque base |
| `SET` | 3 | remplace chaque base par le montant |

Les modes 4 et 5 du serveur (`SET` / `ADD` sur les bases peuplées seulement, syntaxes de vue `!#` et `!!#`) ne sont pas dans ce parseur 5.1.

## 4. Plusieurs jetons : LIKE, COPY, PREDICT

Mots reconnus sans tenir compte de la casse : `LIKE`, `COPY`, `PREDICT`, `PREDICTLR`, `WITHRULES`.

Un seul mot de commande. Un seul jeton de chemin. `WITHRULES` peut être placé n’importe où et passe `use_rules=1`. Sinon `use_rules=0`.

Le premier jeton, s’il n’est pas un mot de commande, est le nombre. Il peut commencer par `#` ou `##`.

### LIKE — `/cell/copy` avec une nouvelle valeur

Le splash n’est pas utilisé.

Chemin source = chemin de la cellule, dont on remplace les coordonnées du jeton de chemin (section 5). Chemin cible = chemin de la cellule. Valeur = le nombre, après les ajustements suivants.

| Saisie | Valeur envoyée à `/cell/copy` |
|---|---|
| `10000 like 2025;demo` | `10000` |
| `##1000 like 2025` | `1000` + valeur numérique de la source |
| `#50% like 2025` | `50%` de la valeur numérique de la source |
| `##10% like 2025` | `110%` de la valeur numérique de la source |
| `10000 like 2025 withrules` | `10000`, `use_rules=1` |
| `10% like 2025` | erreur : un `%` exige `#` ou `##` |
| `like 2025` | forme hors contrat : le 5.1 n’a pas de nombre défini |

`##` devant le nombre, avec `LIKE`, signifie « ajouter à la source », pas le splash `ADD`.

### COPY — `/cell/copy` sans nouvelle valeur

Les valeurs de la source sont copiées telles quelles. Un nombre écrit devant `COPY` est parsé puis ignoré. `10000 copy 2025` copie, il ne pose pas 10000.

### PREDICT — régression linéaire

`PREDICT` et `PREDICTLR` appellent `CellPredictLinearRegression` avec la zone décrite par le chemin (section 5) et la cellule cible. Le nombre parsé devant `PREDICT` n’est pas transmis au serveur dans ce code 5.1.

## 5. Chemin source

Séparateur entre coordonnées : `;`.

| Forme | Sens |
|---|---|
| `2025` | élément dont le nom n’existe que dans une dimension du cube. S’il existe dans deux dimensions : erreur. S’il n’existe pas : erreur |
| `2025;demo` | deux remplacements, chacun par nom unique |
| `Annee:2025` | dimension `Annee`, élément `2025`. Dimension inconnue : erreur |
| `2025:demo` | dimension nommée `2025`, élément `demo`. Ce n’est pas « l’élément 2025 et l’élément demo » |

Les guillemets autour d’un nom sont retirés. Les autres coordonnées restent celles de la cellule cible.

Pour `PREDICT`, la même grammaire décrit une zone. Plusieurs éléments d’une dimension remplacent cette dimension. Forme avec parenthèses : `Annee:(2024;2025)`.

## 6. Nombre

Fonction `_string2Double` (ligne 1967).

- La première virgule devient un point. `10,5` vaut 10,5.
- Suffixe `K` ou `k` : multiplier par 1 000. `M` ou `m` : multiplier par 1 000 000.
- Suffixe `%` : le nombre est un pourcentage, pas une valeur absolue. Le `%` n’est accepté que là où les sections 3 et 4 le disent.
- Tout autre caractère après le nombre est une erreur.
- Un espace dans le nombre ne survit pas au découpage de la section 2. `10 000` n’est pas le nombre 10000 sur la voie préfixe.

## 7. Hors de cette chaîne

- Zone verrouillée (`LockingArea`) : argument de `PALO.SETDATAA`, pas un mot de la chaîne.
- Splash de l’appelant : utilisé seulement pour un nombre nu, une chaîne vide, ou une chaîne qui n’est ni un préfixe ni `LIKE` / `COPY` / `PREDICT`.
- Écriture groupée : pas de parseur. Un splash unique pour toutes les valeurs.

## 8. Exemples d’acceptation

Cellule cible `Paris, 2026, demo`, valeur consolidée actuelle `100`, cellule Palo numérique.

| Intention | Ancienne chaîne | Résultat attendu |
|---|---|---|
| Poser 50 avec le splash du menu | `50` | `/cell/replace` 50, splash de l’appelant |
| Augmenter le total de 10 | `##10` | `/cell/replace` 110, splash `DEFAULT` |
| Passer le total à 110 % | `##10%` | `/cell/replace` 110, splash `DEFAULT` |
| Passer le total à 10 % | `#10%` | `/cell/replace` 10, splash `DEFAULT` |
| Écrire 10 sur chaque base | `!10` | `/cell/replace` 10, splash `SET` |
| Ajouter 10 à chaque base | `!!10` | `/cell/replace` 10, splash `ADD` |
| Reprendre la répartition de 2025 et demo, total 10000 | `10000 like 2025;demo` | `/cell/copy` depuis ce chemin vers `Paris, 2026, demo`, valeur 10000, `use_rules=0` |
| Même modèle, en ajoutant 1000 au total de la source | `##1000 like 2025;demo` | `/cell/copy`, valeur = 1000 + total source |
| Copier les chiffres de 2025 sans les changer d’échelle | `copy 2025` | `/cell/copy` sans paramètre `value` |
| Atteindre 80 par goal seek | `?80` | `/cell/goalseek` type complet |
| Cellule Palo déjà en texte | `##10` | la chaîne `##10` est stockée, splash `NONE` |

## 9. Périmètre retenu pour le plugin

Décision du 8 novembre 2026, tableau d’intentions annoté. Le menu affiche l’intention. Le numéro de splash reste interne. `!` sur une base, Reprendre comme sans regroupement, Atteindre (`?`) et le splash 0 sur un regroupement ne sont pas à implémenter.

Ordre du popup. L’action est le premier champ, parce qu’elle décide des saisies suivantes.

1. Action. La liste dépend du regroupement, déjà détecté avant l’ouverture.
2. Valeur. Absente pour Copier. Pour une action en pourcentage, le libellé est un pourcentage.
3. Chemin source, une liste par dimension. Présent pour Copier, et, seulement avec un regroupement, pour Reprendre une répartition, Reprendre et ajouter, Pourcentage du modèle, Ajouter en % au modèle.
4. Appliquer et Annuler, en bas à droite.

Sans regroupement :

| Menu | À faire |
|---|---|
| Remplacer — `/cell/replace`, splash 0 | oui |
| Ajouter — splash 2 | oui |
| Ajouter en % — écart calculé, puis splash 2 | oui |
| Pourcentage de la cellule — pourcentage calculé, puis splash 0 | oui |
| Copier — `/cell/copy` sans valeur | oui |
| `!` — même effet que Remplacer | non |
| Reprendre comme — jugé incohérent sur une base | non |
| Atteindre — `/cell/goalseek` | non |

Avec regroupement :

| Menu | À faire |
|---|---|
| Repartir — splash 1 | oui |
| Ajouter au total — total recalculé, splash 1 | oui |
| Ajouter un % au total — total recalculé, splash 1 | oui |
| Passer le total à % — total recalculé, splash 1 | oui |
| Écrire sur chaque base — splash 3 | oui |
| Ajouter à chaque base — splash 2 | oui |
| Reprendre une répartition, Reprendre et ajouter, Pourcentage du modèle, Ajouter en % au modèle — `/cell/copy` | oui |
| Copier — `/cell/copy` sans valeur | oui |
| Atteindre — `/cell/goalseek` | non |
| Splash 0 | non, Palo refuse |

Hors de ce périmètre : `WITHRULES`, `PREDICT`, suffixes `K` et `M`, splash 4 et 5.
