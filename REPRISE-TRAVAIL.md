# Reprise de travail — Palo Excel Add-in

Document de passation pour continuer le développement sur une autre machine / une nouvelle session.
Dernière mise à jour : 2026-05-20.

---

## Contexte projet

- **Repo** : `palo-excel-addin` — complément Excel M365, assets dans `docs/`
- **Déploiement** : GitHub Pages → `https://gpizzetta.github.io/palo-excel-addin/`
- **Manifeste** : `docs/manifest.xml` (déploiement admin M365 + sideload)
- **Version actuelle** : **1.0.2.6** (vérifier `docs/version.json`)
- **Architecture** :
  - Shared runtime : `docs/shared-runtime.html`
  - Formules : `docs/functions-core.js` → `./build-bundle.sh` génère `docs/functions.js`
  - Ruban : `docs/commands.js`
  - API Palo : `docs/assets/palo-api.js`

### Fichiers importants

| Fichier | Rôle |
|---------|------|
| `docs/functions-core.js` | Source des formules Excel (éditer ici) |
| `docs/functions.json` | Métadonnées des custom functions |
| `docs/assets/palo-api.js` | Client HTTP Palo, connexions, `/cell/replace`, etc. |
| `docs/commands.js` | Ruban, Snapshot, insertion formules |
| `docs/manifest.xml` | Ruban, CF, shared runtime, raccourcis |
| `docs/build-bundle.sh` | Regénère `functions.js` / `functions-bundle.js` |
| `docs/bump-version.sh` | Propagation de version partout |
| `TODO.md` | Backlog (mentionne `PaloSpreadsheetfuncs` 5.1 — **absent du repo**) |

---

## Besoin métier : partage de documents budget

### Workflow visé

1. **Finance** (avec plugin + connexion) prépare des classeurs avec formules Palo → valeurs affichées = **références** pour le budget.
2. **Utilisateurs budget** (sans accès Palo) remplissent le document en s’appuyant sur ces références.
3. **Finance** rouvre le fichier avec le plugin et utilise les formules d’**insertion** pour pousser les saisies vers le cube.

### Contrainte

Les **valeurs de référence doivent rester visibles** même si l’utilisateur n’a pas le plugin ou pas de connexion.

### Options évaluées

| Approche | Verdict |
|----------|---------|
| Fichier avec formules `PALO.*` seulement, sans plugin | **Non fiable** — `#NAME?` au recalcul |
| Plugin installé partout, sans connexion | **Insuffisant seul** — recalcul → erreurs `#PALO!` ou cellules vides |
| **Snapshot** (ruban) sur une copie avant partage | **Fiable** — formules → valeurs en dur |
| **`PALO.DATAS`** (safe) — garder dernière valeur sans connexion | **Piste technique** — voir section dédiée |

### Workflow recommandé (opérationnel aujourd’hui)

1. Finance : master avec formules Palo live.
2. **Fichier > Enregistrer sous** → copie budget (`Budget_2026_saisie.xlsx`).
3. Sur la copie : bouton **Snapshot** du ruban (convertit toutes les formules en valeurs affichées).
4. **Enregistrer** la copie et la partager.
5. Utilisateurs remplissent les zones de saisie (sans Palo).
6. Finance récupère le fichier et injecte via `PALO_SETDATA` / futures formules d’écriture.

**Fichiers à conserver :**

| Fichier | Rôle |
|---------|------|
| `Budget_2026_master.xlsx` | Formules live, évolution du modèle |
| `Budget_2026_saisie.xlsx` | Copie snapshot, partagée |

### Évolution possible : bouton « Préparer copie budget »

Automatiser : snapshot **uniquement des cellules `@PALO.*`** + calcul manuel — **pas encore implémenté**.

---

## Piste `PALO.DATAS` (S = safe)

### Idée

Sans connexion, au lieu de renvoyer `""` ou `#PALO!`, retourner la **dernière valeur connue**.

### Pourquoi ce n’est pas trivial

Une custom function **remplace toujours** le contenu de la cellule par son retour. « Laisser intact » = **retourner explicitement** la valeur à afficher.

### Comportement actuel de `PALO.DATAC` (sans connexion)

Dans `functions-core.js`, en cas d’erreur :
- debug / timeout / HTTP → `#PALO! …`
- sinon → `""` (cellule vide)

→ Les références disparaissent au recalcul, elles ne sont pas figées.

### Proposition produit

| Fonction | Usage | Sans connexion |
|----------|--------|----------------|
| **`PALO.DATAC`** | Finance, modèles live | Erreur explicite (comportement actuel) |
| **`PALO.DATAS`** | Colonnes de référence partagées | Retourner la **dernière valeur en cache** |

### Implémentation envisagée (non faite)

1. Cache en mémoire (shared runtime) : `adresse cellule → dernière valeur OK`.
2. `OfficeRuntime.storage` : clé `servdb + cube + coords` (ne voyage **pas** avec le `.xlsx`).
3. Lecture via `invocation.address` + `Excel.run` (shared runtime) — fragile.
4. Cache embarqué dans le classeur (feuille cachée) — le plus robuste pour le partage.

**Limite** : sans plugin du tout → `#NAME?` ; `DATAS` n’aide que si le complément est installé.

**Statut** : discussion seulement, **pas implémenté**.

---

## À faire : `PALO.SETDATAIF`

### Besoin

Implémenter `PALO.SETDATAIF` comme l’ancien plugin Excel 2010 (COM), pour l’écriture conditionnelle vers le cube (workflow budget finance).

### Ce qui existe déjà

- **`PALO_SETDATA`** dans `docs/functions-core.js` :
  - Signature : `(value, splash, servdb, cubeName, ...coordinates)`
  - Appelle `cellReplaceByIds` → API `/cell/replace`
  - Retourne `1` si OK, `0` si blocage amont, `#PALO!` en erreur
- **Ruban** : insertion `PALO_SETDATA` dans `docs/commands.js`
- **Métadonnées** : entrée dans `docs/functions.json`

### Ce qui manque dans le repo

- Aucune implémentation de `SETDATAIF`
- Source de référence `exemple_addin_palo/client_libraries/5.1/PaloSpreadsheetfuncs` citée dans `TODO.md` mais **absente du dépôt**
- Pas de doc interne détaillée sur le retour exact Excel 2010

### Doc Jedox (publique)

#### Signature — deux variantes selon version

**Ancienne / on-prem (proche Excel 2010 COM)**  
[Vedox OLAP Functions (on-prem)](https://knowledgebase-onprem.jedox.com/jedox/planning/palo-olap-functions.htm)

```excel
=PALO.SETDATAIF(Condition; Valeur; Splash; "Serveur/Base"; "Cube"; Coord1; Coord2; …)
```

**Récente**  
[Jedox Data Functions](https://knowledgebase.jedox.com/jedox/planning/jedox-data-functions.htm)

```excel
=PALO.SETDATAIF(Condition; Valeur; Splash; Check Splash Thr; "Serveur/Base"; "Cube"; Coord1; …)
```

#### Comportement documenté

- **Condition = VRAI** → identique à `PALO.SETDATA` (écriture cube).
- **Condition = FAUX** → **aucune écriture**.
- **`Check Splash Thr`** : filet de sécurité splashing (paramètre API `check_threshold` sur `/cell/replace`) — **non implémenté** chez nous.

### Écarts actuels vs plugin historique

| Point | Plugin 2010 / Jedox | Notre add-in |
|-------|---------------------|--------------|
| Nom lecture | `PALO.SETDATA` | `PALO.PALO_SETDATA` (id `PALO_SETDATA` dans `functions.json`) |
| Nom écriture conditionnelle | `PALO.SETDATAIF` | **Absent** |
| `Check Splash Thr` | Présent (versions récentes) | **Absent** |
| Retour si condition fausse | Doc : « rien ne se passe » — valeur exacte **à confirmer** | N/A |
| Retour si succès | Probablement `1` / `TRUE` | `1` |
| `servdb` dynamique `/DWH` | Connexion active du volet | **Implémenté** dans `palo-api.js` |

### Implémentation proposée (simple)

Wrapper autour de `PALO_SETDATA` :

```javascript
async function SETDATAIF(condition, value, splash, servdb, cubeName) {
  var coordinates = sanitizePaloCoordinates(Array.prototype.slice.call(arguments, 4));
  if (!paloConditionIsTrue(condition)) {
    return 0; // À valider sur fichier Excel 2010 réel
  }
  return PALO_SETDATA(value, splash, servdb, cubeName, ...coordinates);
}
```

**Enregistrement** : `CustomFunctions.associate("SETDATAIF", SETDATAIF)` + entrée dans `functions.json`  
→ Excel : `=PALO.SETDATAIF(...)` (namespace `PALO`).

Puis : `./build-bundle.sh` et bump version si déploiement.

### Points à valider demain (bloquants pour « à l’identique »)

1. **Exemple réel** : copier une formule `PALO.SETDATAIF` depuis un classeur Excel 2010.
2. **4 ou 5 paramètres fixes** avant les coordonnées ? (`Check Splash Thr` présent ou non ?)
3. **Valeur de retour** quand condition = FAUX (`0`, `""`, autre ?).
4. **Évaluation de la condition** : `TRUE`/`FALSE`, `1`/`0`, chaînes ?
5. Fournir **`PaloSpreadsheetfuncs` 5.1** ou dossier `exemple_addin_palo` si disponible ailleurs.
6. Faut-il aussi aligner **`PALO.SETDATA`** (nom + `Check Splash Thr`) pour compatibilité documents existants ?

### API `/cell/replace` — `check_threshold`

Paramètre serveur pour le filet de sécurité splashing.  
Notre `cellReplace` / `cellReplaceByIds` dans `palo-api.js` ne le passe **pas encore**.

---

## Autres éléments déjà en place (session précédente)

- **Snapshot** : ruban `paloSnapshotWorkbookValues` — convertit tout le classeur en valeurs (avec confirmation `palo-snapshot-confirm.html`).
- **Raccourci** : `Ctrl+Alt+A` → `paloRibbonAction` (`docs/shortcuts.json` + `ExtendedOverrides` manifeste).
- **`servdb` dynamique** : `"/DWH"` utilise la connexion active du volet.
- **Compatibilité formules** : objectif dans `TODO.md` d’aligner sur `PaloSpreadsheetfuncs` — pas fait.

---

## Commandes utiles

```bash
# Regénérer le bundle fonctions
cd docs && ./build-bundle.sh

# Monter de version (exemple)
cd docs && ./bump-version.sh 1.0.2.7

# Sideload local
npx office-addin-debugging start docs/manifest.xml desktop
```

---

## Plan de reprise suggéré (demain)

### Priorité 1 — `PALO.SETDATAIF`

1. Obtenir 1–2 formules réelles depuis un classeur Excel 2010.
2. Implémenter dans `functions-core.js` + `functions.json`.
3. Tester : condition VRAI → écriture ; FAUX → pas d’appel API.
4. Décider `Check Splash Thr` (maintenant ou plus tard).
5. `build-bundle.sh`, bump version, republier manifeste admin si besoin.

### Priorité 2 — Workflow budget (si besoin)

- Snapshot ciblé (Palo seulement) ou `PALO.DATAS` selon choix produit.

### Priorité 3 — Compatibilité noms

- Envisager alias `SETDATA` / `SETDATAIF` (sans préfixe `PALO_`) pour ouverture de vieux classeurs.

---

## Liens

- Add-in déployé : https://gpizzetta.github.io/palo-excel-addin/
- Manifeste : https://gpizzetta.github.io/palo-excel-addin/manifest.xml
- Support : https://gpizzetta.github.io/palo-excel-addin/support.html
- Doc Jedox SETDATAIF : https://knowledgebase.jedox.com/jedox/planning/jedox-data-functions.htm
- Doc Jedox on-prem : https://knowledgebase-onprem.jedox.com/jedox/planning/palo-olap-functions.htm

---

## Message pour la prochaine instance IA

> Nous développons un add-in Excel M365 pour Palo/Jedox. Prochaine tâche : **implémenter `PALO.SETDATAIF`** en s’appuyant sur `PALO_SETDATA` existant. Lire **`REPRISE-TRAVAIL.md`** à la racine du repo. Il manque des formules Excel 2010 réelles et la lib `PaloSpreadsheetfuncs` pour garantir une compatibilité 100 % — commencer par lire `functions-core.js` (PALO_SETDATA) et `functions.json`, puis implémenter le wrapper conditionnel. Ne pas oublier `build-bundle.sh` après modification de `functions-core.js`.
