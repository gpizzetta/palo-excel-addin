# Handoff — PALO.DATAB (lecture groupée)

État au 2026-10-08. Version locale du complément : **1.0.3.3**. Pas encore en ligne tant que `docs/` n’est pas poussé sur GitHub Pages et que le manifeste n’est pas republié dans Microsoft 365.

## Ruban en double (1.0.3.3)

Excel Desktop affichait deux groupes « Palo » parce que le ruban était déclaré dans `VersionOverrides` 1.0 (`commands.html`) et dans le 1.1 (`shared-runtime.html`). Le 1.0 n’a plus de `DesktopFormFactor` : plus de ruban, plus de menu contextuel, plus de `FunctionFile` vers `commands.html`. Le seul ruban restant est celui du bloc 1.1 : volet, actions et formules passent par `shared-runtime.html`. Le bloc 1.0 garde uniquement les fonctions personnalisées, pour un Excel qui ne lirait pas le 1.1.

## Où on en est

`PALO.DATAB` est ajoutée. `PALO.DATAC` n’a pas changé de chemin.

| Formule | Chemin |
| --- | --- |
| `PALO.DATAC` | Inchangé. Dans les formules Excel : un `/cell/value` par cellule (`name_path`). |
| `PALO.DATAB` | Même signature (`servdb`, cube, coord1…coord20). Les appels du même cube partis dans une fenêtre de 24 ms partent ensemble en `/cell/values` avec des identifiants (`paths=id,id:id,id`). |

Une seule cellule `DATAB`, ou `PALO_DATAB_BATCH_MS = 0`, ou `PALO_DISABLE_BATCH` vrai : lecture unitaire `/cell/value` en noms, comme `DATAC`.

Si la conversion noms → identifiants échoue, chaque formule du paquet non encore résolue retombe sur ce `/cell/value`. Un paquet `/cell/values` qui échoue, ou qui revient entièrement vide, retombe sur `/cell/value` par identifiants pour les cellules de ce paquet seulement.

Les noms qui contiennent une virgule passent par les identifiants dans le chemin groupé. Ils ne sont plus coupés par la virgule du `name_path`.

## Fichiers touchés

- `docs/assets/palo-api.js` — file séparée `_databBatchQueues` / `requestDatabCellValue` / `_flushDatabBatch`. `cellBatchDelayMs` et `_flushCellValueBatch` (ceux de `DATAC`) ne sont pas modifiés.
- `docs/functions-core.js` — fonction `DATAB` et `CustomFunctions.associate("DATAB", DATAB)`.
- `docs/functions.json` — métadonnée `DATAB`, 22 paramètres comme `DATAC`.
- `docs/functions.js` et `docs/functions-bundle.js` — régénérés par `docs/bump-version.sh`.
- Manifeste et `?v=1.0.3.2` — mis à jour par le même script.

`docs/staging/` n’a pas été republie. Il est encore en 1.0.3.0 / 1.0.3.1 selon les fichiers, sans `DATAB`.

## Ce qu’il ne faut pas faire

- Ne pas enlever le `if (paloInCustomFunctionsRuntime()) return 0` de `cellBatchDelayMs`.
- Ne pas enlever le bloc « une par une » au début de `_flushCellValueBatch`. C’est le verrou de `DATAC`.
- Ne pas faire appeler `requestDatabCellValue` par `DATAC`.

`DATAB` utilise `setTimeout` dans le runtime des formules. C’est le mécanisme qui avait planté Excel Online pour `DATAC`, et il est volontairement limité à cette nouvelle formule. Un plantage de ce runtime arrête aussi les autres formules Palo, parce qu’elles partagent la page. Premier essai sur Excel Desktop.

## Comment tester

1. Pousser `docs/` et republier le manifeste `1.0.3.2`. Le pied du volet Connexion doit afficher **1.0.3.2**.
2. Copier une feuille. Remplacer `PALO.DATAC` par `PALO.DATAB` sur la copie seulement.
3. Comparer les valeurs, les vides, un nom avec une virgule, deux cubes, et le temps de calcul.
4. Une formule isolée `DATAB` doit donner la même valeur qu’une `DATAC` à côté (les deux passent par `/cell/value`).
5. Un bloc de `DATAB` du même cube doit montrer des appels `/cell/values` (trace `datab-cell-values-batch` si `PALO_BULK_TRACE` ou `PALO_DEBUG` est actif).

Pour couper le groupé sans retirer la formule : dans la console du runtime, `PALO_DATAB_BATCH_MS = 0` ou `PALO_DISABLE_BATCH = true`, puis recalculer. `DATAB` retombe alors sur `/cell/value`.

## Suite possible

Si les valeurs de `DATAB` collent à `DATAC` et que le temps baisse sur un gros classeur, on pourra faire emprunter ce chemin à `DATAC`. Pas avant cette comparaison. Le chargement complet en mémoire (`/cell/export` + bouton Actualiser) reste une option à part, non commencée.
