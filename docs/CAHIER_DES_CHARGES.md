# Cahier des charges — fonctions PALO dans l’add-in Office 365

## Contexte

- **Cible** : réimplémenter, en **JavaScript** (fonctions personnalisées Excel / Office.js), l’équivalent des fonctions exposées historiquement par le **plugin Excel 2010** (code C++ du type **`PaloSpreadsheetFuncs`** dans la branche client *molap*, ex. `molap/client_libraries/5.1/PaloSpreadsheetFuncs`).
- **Référence code disponible localement** : dépôt **`palo-server`**, fichier  
  `Library/Parser/PaloFunctionNodeFactory.cpp` — enregistrement des fonctions **`palo.*`** utilisées dans le **langage de règles** du serveur (noms en **minuscules**, sémantique OLAP).
- **Convention Excel** : l’add-in classique expose souvent les mêmes capacités sous des noms **`PALO.*`** (majuscules). Le mapping proposé ci‑dessous est **`PALO.NOM` ↔ `palo.nom`** (à valider ligne à ligne avec le source C++ *PaloSpreadsheetFuncs* lorsque vous l’aurez dans le dépôt).

> **À faire** : ajouter dans ce dépôt (ou en sous-module) le répertoire `PaloSpreadsheetFuncs` pour compléter signatures exactes, paramètres optionnels et fonctions **non** présentes dans le moteur de règles (connexion, sous-ensembles, etc.).

---

## A. Fonctions déjà identifiées dans `palo-server` (moteur de règles)

Liste issue de `PaloFunctionNodeFactory::registerFunctions()` — **21 fonctions**.

| # | Nom moteur (`palo.*`) | Nom cible add-in (`PALO.*`) | Rôle (résumé) |
|---|------------------------|-----------------------------|----------------|
| 1 | `palo.cubedimension` | `PALO.CUBEDIMENSION` | Métadonnée de dimension d’un cube (nom / indice de dimension). |
| 2 | `palo.data` | `PALO.DATA` | Lecture (conceptuellement) d’une valeur de cellule cube selon le chemin d’éléments. |
| 3 | `palo.echild` | `PALO.ECHILD` | Enfant d’un élément dans la dimension. |
| 4 | `palo.echildcount` | `PALO.ECHILDCOUNT` | Nombre d’enfants. |
| 5 | `palo.ecount` | `PALO.ECOUNT` | Nombre d’éléments (contexte dimension). |
| 6 | `palo.efirst` | `PALO.EFIRST` | Premier élément (souvent sous un parent). |
| 7 | `palo.eindent` | `PALO.EINDENT` | Indentation / profondeur hiérarchique. |
| 8 | `palo.eindex` | `PALO.EINDEX` | Index d’un élément. |
| 9 | `palo.eischild` | `PALO.EISCHILD` | Test parent-enfant. |
| 10 | `palo.elevel` | `PALO.ELEVEL` | Niveau hiérarchique. |
| 11 | `palo.ename` | `PALO.ENAME` | Nom d’un élément. |
| 12 | `palo.enext` | `PALO.ENEXT` | Élément suivant. |
| 13 | `palo.eparent` | `PALO.EPARENT` | Parent d’un élément. |
| 14 | `palo.eparentcount` | `PALO.EPARENTCOUNT` | Nombre de parents (selon modèle Palo). |
| 15 | `palo.eprev` | `PALO.EPREV` | Élément précédent. |
| 16 | `palo.esibling` | `PALO.ESIBLING` | Frère / voisin dans la dimension. |
| 17 | `palo.etoplevel` | `PALO.ETOPLEVEL` | Élément de plus haut niveau / racine utile. |
| 18 | `palo.etype` | `PALO.ETYPE` | Type d’élément (consolidation, base, etc.). |
| 19 | `palo.eweight` | `PALO.EWEIGHT` | Poids de consolidation. |
| 20 | `palo.marker` | `PALO.MARKER` | Marqueur / contexte de calcul (règles). |
| 21 | `palo.eoffset` | `PALO.EOFFSET` | Décalage d’élément (offset dans la dimension). |

**Implémentations de référence côté serveur** (pour analyse des paramètres) : répertoire  
`palo-server/Library/Parser/` — fichiers `FunctionNodePalo*.h` / `.cpp` correspondants.

---

## B. Fonctions typiques du complément Excel (hors liste A — à confirmer dans *PaloSpreadsheetFuncs*)

Ces familles apparaissent dans la documentation produit **Jedox / Palo** pour Excel ; elles ne sont **pas** dans `PaloFunctionNodeFactory.cpp` ci‑dessus. Elles relèvent souvent de la **connexion HTTP**, des **vues** ou de l’**écriture** dans les cellules.

| Famille | Exemples | Priorité proposée |
|---------|----------|-------------------|
| Connexion / serveur | `PALO.SERVER_INFO`, `PALO.LOGIN` (ou équivalent selon ancien add-in) | P0 — base pour tout appel API |
| Données cube | `PALO.DATA`, `PALO.DATAC`, `PALO.SETDATA`, variantes | P0–P1 |
| Dimensions / listes | `PALO.SUBSET`, filtres, tris | P1 |
| Métadonnées | bases, cubes, dimensions, attributs | P1 |
| Règles / script | souvent absent ou limité côté feuille | P2 |

*À compléter* une fois le code **`PaloSpreadsheetFuncs`** disponible (liste exhaustive + signatures).

---

## C. Phases de livraison (proposition)

| Phase | Contenu | Dépendances |
|-------|---------|-------------|
| **P0** | Connexion (URL, utilisateur, mot de passe) stockée ; appels API authentifiés (MD5 / token selon API Palo du serveur) ; `PALO.VERSION` / test réseau (`PALO.INFO`). | HTTPS + CORS sur le serveur Palo |
| **P1** | `PALO.DATA` (lecture cellule) + sous-ensemble des `PALO.E*` les plus utilisés (`ENAME`, `EPARENT`, `ECHILD`, `ELEVEL`, …). | P0 |
| **P2** | Reste des fonctions section A + écriture si exposée par l’API (`PALO.SETDATA` ou équivalent HTTP). | P1 |
| **P3** | Fonctions section B (subsets, vues, batch) + optimisation (requêtes groupées). | P2 |

---

## D. Contraintes Office 365 / JavaScript

- Les **fonctions personnalisées** sont **asynchrones** (Promises) ; pas d’accès direct au ruban depuis le même thread que certaines API.
- **Volatilité** : limiter le nombre d’appels réseau par recalcul ; envisager **batch** côté serveur ou file d’appels.
- **Sécurité** : ne pas stocker les mots de passe en clair dans le classeur sans chiffrement ; préférer jetons à durée limitée si l’API le permet.
- **Excel Online** : respect **HTTPS**, **CORS**, parfois pas de ruban pour compléments téléversés — prévoir ouverture du volet depuis **Insertion → Compléments**.

---

## E. Suivi

- [ ] Importer / référencer le code **`PaloSpreadsheetFuncs`** (5.1) pour validation de la liste exhaustive.
- [ ] Valider le mapping **`PALO.*`** vs routes HTTP réelles du serveur (`PaloHttpServer`, jobs, etc.).
- [ ] Tracer une colonne « implémenté / partiel / reporté » dans une table de suivi (feuille projet ou issues GitHub).

---

*Document généré pour le dépôt **palo-excel-addin** ; à mettre à jour lorsque les sources C++ du plugin Excel 2010 seront disponibles dans l’arborescence du projet.*
