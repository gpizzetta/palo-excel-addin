# Cahier des charges — fonctions PALO dans l’add-in Office 365

## Contexte

- **Cible** : réimplémenter, en **JavaScript** (fonctions personnalisées Excel / Office.js), l’équivalent des fonctions exposées historiquement par le **plugin Excel** (bibliothèque client **`PaloSpreadsheetFuncs`**, branche *molap* type Jedox 5.1).
- **Référence client (code C++)** — miroir public (**plugin Excel 2010 / COM**, Jedox ~5.1) — **URL canonique** (à conserver telle quelle dans la doc) :  
  [https://github.com/gpizzetta/jedox-mirror/tree/master/molap/client_libraries/5.1/PaloSpreadsheetFuncs](https://github.com/gpizzetta/jedox-mirror/tree/master/molap/client_libraries/5.1/PaloSpreadsheetFuncs)  
  Dossier équivalent dans le dépôt : `molap/client_libraries/5.1/PaloSpreadsheetFuncs/`. Fichiers principaux : `include/PaloSpreadsheetFuncs/SpreadsheetFuncs.h` (déclarations des entrées `FPalo*`), `SpreadsheetFuncs.cpp` (implémentations).
- **Référence moteur de règles** (autre dépôt) : **`palo-server`**, `Library/Parser/PaloFunctionNodeFactory.cpp` — enregistrement des fonctions **`palo.*`** en **minuscules** dans le langage de règles OLAP (recouvre partiellement la sémantique « éléments » du client tableur).
- **Convention Excel** : l’add-in classique expose les capacités sous des noms **`PALO.*`** (majuscules). Le tableau section **B** propose un **`PALO.*` mécanique** dérivé du suffixe CamelCase après `FPalo` (voir ci‑dessous). Les intitulés **réels** de l’add-in Excel 2010 peuvent différer (alias historiques du type `PALO.DATA` vs `PALO.GETDATA`) : à valider contre la couche d’intégration Excel si vous y avez accès.

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

**Implémentations de référence côté serveur** : `palo-server/Library/Parser/` — fichiers `FunctionNodePalo*.h` / `.cpp` correspondants.

---

## B. Bibliothèque client `PaloSpreadsheetFuncs` (inventaire C++)

### B.1 Périmètre

Les **entrées exposées au tableur** dans Jedox 5.1 sont les méthodes `void FPalo…(GenericCell&, …)` déclarées dans `SpreadsheetFuncs.h` (hors helpers listés en **B.3**). **Dénombrement : 150** entrées.

### B.2 Convention `PALO.*` proposée (mécanique)

Pour chaque `FPaloFooBar`, suffixe `FooBar` → séparation CamelCase en mots → **majuscules avec underscores** → préfixe `PALO.`  
Exemples : `FPaloElementName` → `PALO.ELEMENT_NAME` ; `FPaloGetdataAC` → `PALO.GETDATA_AC`.  
Cela sert de **base stable** pour le cahier des charges ; les noms « marketing » historiques Excel peuvent être documentés en alias dans le code de l’add-in Office 365.

### B.3 Helpers internes (non comptés dans les 150)

Implémentations dans `SpreadsheetFuncs.cpp` / `SpreadsheetFuncs.h`, utilisées en interne ou pour le dispatch :

| Méthode | Rôle |
|---------|------|
| `FPaloGetdataAggregation` | Dispatch des agrégations (appelée par les variantes Sum / Avg / Count / Max / Min). |
| `FPaloGetdataACTIntern` | Variante interne pour chemins de données avec options (collapse / texte). |
| `FPaloParseSubsetParams` | Construction des paramètres de sous-ensemble à partir des arguments de la feuille. |

### B.4 Inventaire complet (150 entrées)

| # | Catégorie | Méthode C++ | Nom cible proposé |
|---|-----------|-------------|-------------------|
| 1 | Admin bases | `FPaloDatabaseAddCube` | `PALO.DATABASE_ADD_CUBE` |
| 2 | Admin bases | `FPaloDatabaseAddDimension` | `PALO.DATABASE_ADD_DIMENSION` |
| 3 | Admin bases | `FPaloDatabaseDeleteCube` | `PALO.DATABASE_DELETE_CUBE` |
| 4 | Admin bases | `FPaloDatabaseDeleteDimension` | `PALO.DATABASE_DELETE_DIMENSION` |
| 5 | Admin bases | `FPaloDatabaseLoadCube` | `PALO.DATABASE_LOAD_CUBE` |
| 6 | Admin bases | `FPaloDatabaseRenameDimension` | `PALO.DATABASE_RENAME_DIMENSION` |
| 7 | Admin bases | `FPaloDatabaseUnloadCube` | `PALO.DATABASE_UNLOAD_CUBE` |
| 8 | Admin bases | `FPaloRootAddDatabase` | `PALO.ROOT_ADD_DATABASE` |
| 9 | Admin bases | `FPaloRootDeleteDatabase` | `PALO.ROOT_DELETE_DATABASE` |
| 10 | Admin bases | `FPaloRootListDatabases` | `PALO.ROOT_LIST_DATABASES` |
| 11 | Admin bases | `FPaloRootListDatabasesExt` | `PALO.ROOT_LIST_DATABASES_EXT` |
| 12 | Admin bases | `FPaloRootSaveDatabase` | `PALO.ROOT_SAVE_DATABASE` |
| 13 | Admin bases | `FPaloRootUnloadDatabase` | `PALO.ROOT_UNLOAD_DATABASE` |
| 14 | Admin cubes | `FPaloCubeClear` | `PALO.CUBE_CLEAR` |
| 15 | Admin cubes | `FPaloCubeRename` | `PALO.CUBE_RENAME` |
| 16 | Admin dimensions | `FPaloDimensionClear` | `PALO.DIMENSION_CLEAR` |
| 17 | Cache, SVS, drill | `FPaloCellDrillTrough` | `PALO.CELL_DRILL_TROUGH` |
| 18 | Cache, SVS, drill | `FPaloEndCacheCollect` | `PALO.END_CACHE_COLLECT` |
| 19 | Cache, SVS, drill | `FPaloSVSInfo` | `PALO.SVS_INFO` |
| 20 | Cache, SVS, drill | `FPaloSVSRestart` | `PALO.SVS_RESTART` |
| 21 | Cache, SVS, drill | `FPaloStartCacheCollect` | `PALO.START_CACHE_COLLECT` |
| 22 | Connexion / serveur / licence | `FPaloActivateLicense` | `PALO.ACTIVATE_LICENSE` |
| 23 | Connexion / serveur / licence | `FPaloChangePassword` | `PALO.CHANGE_PASSWORD` |
| 24 | Connexion / serveur / licence | `FPaloChangeUserPassword` | `PALO.CHANGE_USER_PASSWORD` |
| 25 | Connexion / serveur / licence | `FPaloConnectionUser` | `PALO.CONNECTION_USER` |
| 26 | Connexion / serveur / licence | `FPaloDisconnect` | `PALO.DISCONNECT` |
| 27 | Connexion / serveur / licence | `FPaloGetGroups` | `PALO.GET_GROUPS` |
| 28 | Connexion / serveur / licence | `FPaloGetGroupsForSID` | `PALO.GET_GROUPS_FOR_SID` |
| 29 | Connexion / serveur / licence | `FPaloGetUserForSID` | `PALO.GET_USER_FOR_SID` |
| 30 | Connexion / serveur / licence | `FPaloInit` | `PALO.INIT` |
| 31 | Connexion / serveur / licence | `FPaloLicenseInfo` | `PALO.LICENSE_INFO` |
| 32 | Connexion / serveur / licence | `FPaloPing` | `PALO.PING` |
| 33 | Connexion / serveur / licence | `FPaloRegisterServer` | `PALO.REGISTER_SERVER` |
| 34 | Connexion / serveur / licence | `FPaloRemoveConnection` | `PALO.REMOVE_CONNECTION` |
| 35 | Connexion / serveur / licence | `FPaloServerInfo` | `PALO.SERVER_INFO` |
| 36 | Connexion / serveur / licence | `FPaloSetClientDescription` | `PALO.SET_CLIENT_DESCRIPTION` |
| 37 | Connexion / serveur / licence | `FPaloSetSvs` | `PALO.SET_SVS` |
| 38 | Données cube & cellule | `FPaloCellCopy` | `PALO.CELL_COPY` |
| 39 | Données cube & cellule | `FPaloGetdata` | `PALO.GETDATA` |
| 40 | Données cube & cellule | `FPaloGetdataA` | `PALO.GETDATA_A` |
| 41 | Données cube & cellule | `FPaloGetdataAC` | `PALO.GETDATA_AC` |
| 42 | Données cube & cellule | `FPaloGetdataAT` | `PALO.GETDATA_AT` |
| 43 | Données cube & cellule | `FPaloGetdataATC` | `PALO.GETDATA_ATC` |
| 44 | Données cube & cellule | `FPaloGetdataAV` | `PALO.GETDATA_AV` |
| 45 | Données cube & cellule | `FPaloGetdataAggregationAvg` | `PALO.GETDATA_AGGREGATION_AVG` |
| 46 | Données cube & cellule | `FPaloGetdataAggregationCount` | `PALO.GETDATA_AGGREGATION_COUNT` |
| 47 | Données cube & cellule | `FPaloGetdataAggregationMax` | `PALO.GETDATA_AGGREGATION_MAX` |
| 48 | Données cube & cellule | `FPaloGetdataAggregationMin` | `PALO.GETDATA_AGGREGATION_MIN` |
| 49 | Données cube & cellule | `FPaloGetdataAggregationSum` | `PALO.GETDATA_AGGREGATION_SUM` |
| 50 | Données cube & cellule | `FPaloGetdataC` | `PALO.GETDATA_C` |
| 51 | Données cube & cellule | `FPaloGetdataExport` | `PALO.GETDATA_EXPORT` |
| 52 | Données cube & cellule | `FPaloGetdataT` | `PALO.GETDATA_T` |
| 53 | Données cube & cellule | `FPaloGetdataTC` | `PALO.GETDATA_TC` |
| 54 | Données cube & cellule | `FPaloGetdataV` | `PALO.GETDATA_V` |
| 55 | Données cube & cellule | `FPaloGoalSeek` | `PALO.GOAL_SEEK` |
| 56 | Données cube & cellule | `FPaloSetdata` | `PALO.SETDATA` |
| 57 | Données cube & cellule | `FPaloSetdataA` | `PALO.SETDATA_A` |
| 58 | Données cube & cellule | `FPaloSetdataAIf` | `PALO.SETDATA_A_IF` |
| 59 | Données cube & cellule | `FPaloSetdataBulk` | `PALO.SETDATA_BULK` |
| 60 | Données cube & cellule | `FPaloSetdataIf` | `PALO.SETDATA_IF` |
| 61 | IDs & noms | `FPaloGetCubeId` | `PALO.GET_CUBE_ID` |
| 62 | IDs & noms | `FPaloGetCubeName` | `PALO.GET_CUBE_NAME` |
| 63 | IDs & noms | `FPaloGetDimensionId` | `PALO.GET_DIMENSION_ID` |
| 64 | IDs & noms | `FPaloGetDimensionName` | `PALO.GET_DIMENSION_NAME` |
| 65 | IDs & noms | `FPaloGetElementId` | `PALO.GET_ELEMENT_ID` |
| 66 | IDs & noms | `FPaloGetElementName` | `PALO.GET_ELEMENT_NAME` |
| 67 | Listes & métadonnées | `FPaloCubeInfo` | `PALO.CUBE_INFO` |
| 68 | Listes & métadonnées | `FPaloCubeListDimensions` | `PALO.CUBE_LIST_DIMENSIONS` |
| 69 | Listes & métadonnées | `FPaloDatabaseInfo` | `PALO.DATABASE_INFO` |
| 70 | Listes & métadonnées | `FPaloDatabaseListCubes` | `PALO.DATABASE_LIST_CUBES` |
| 71 | Listes & métadonnées | `FPaloDatabaseListDimensions` | `PALO.DATABASE_LIST_DIMENSIONS` |
| 72 | Listes & métadonnées | `FPaloDatabaseListDimensionsExt` | `PALO.DATABASE_LIST_DIMENSIONS_EXT` |
| 73 | Listes & métadonnées | `FPaloDimensionInfo` | `PALO.DIMENSION_INFO` |
| 74 | Listes & métadonnées | `FPaloDimensionListCubes` | `PALO.DIMENSION_LIST_CUBES` |
| 75 | Listes & métadonnées | `FPaloDimensionListElements` | `PALO.DIMENSION_LIST_ELEMENTS` |
| 76 | Listes & métadonnées | `FPaloDimensionListElements2` | `PALO.DIMENSION_LIST_ELEMENTS2` |
| 77 | Listes & métadonnées | `FPaloDimensionMaxLevel` | `PALO.DIMENSION_MAX_LEVEL` |
| 78 | Listes & métadonnées | `FPaloDimensionReducedChildrenListElements` | `PALO.DIMENSION_REDUCED_CHILDREN_LIST_ELEMENTS` |
| 79 | Listes & métadonnées | `FPaloDimensionReducedFlatListElements` | `PALO.DIMENSION_REDUCED_FLAT_LIST_ELEMENTS` |
| 80 | Listes & métadonnées | `FPaloDimensionReducedTopListElements` | `PALO.DIMENSION_REDUCED_TOP_LIST_ELEMENTS` |
| 81 | Listes & métadonnées | `FPaloDimensionSimpleChildrenListElements` | `PALO.DIMENSION_SIMPLE_CHILDREN_LIST_ELEMENTS` |
| 82 | Listes & métadonnées | `FPaloDimensionSimpleFlatListElements` | `PALO.DIMENSION_SIMPLE_FLAT_LIST_ELEMENTS` |
| 83 | Listes & métadonnées | `FPaloDimensionSimpleTopListElements` | `PALO.DIMENSION_SIMPLE_TOP_LIST_ELEMENTS` |
| 84 | Listes & métadonnées | `FPaloDimensionTopElementsCount` | `PALO.DIMENSION_TOP_ELEMENTS_COUNT` |
| 85 | Listes & métadonnées | `FPaloElementListAncestors` | `PALO.ELEMENT_LIST_ANCESTORS` |
| 86 | Listes & métadonnées | `FPaloElementListConsolidationElements` | `PALO.ELEMENT_LIST_CONSOLIDATION_ELEMENTS` |
| 87 | Listes & métadonnées | `FPaloElementListDescendants` | `PALO.ELEMENT_LIST_DESCENDANTS` |
| 88 | Listes & métadonnées | `FPaloElementListParents` | `PALO.ELEMENT_LIST_PARENTS` |
| 89 | Listes & métadonnées | `FPaloElementListSiblings` | `PALO.ELEMENT_LIST_SIBLINGS` |
| 90 | Règles & conversion | `FPaloCubeConvert` | `PALO.CUBE_CONVERT` |
| 91 | Règles & conversion | `FPaloCubeRuleCreate` | `PALO.CUBE_RULE_CREATE` |
| 92 | Règles & conversion | `FPaloCubeRuleDelete` | `PALO.CUBE_RULE_DELETE` |
| 93 | Règles & conversion | `FPaloCubeRuleModify` | `PALO.CUBE_RULE_MODIFY` |
| 94 | Règles & conversion | `FPaloCubeRuleParse` | `PALO.CUBE_RULE_PARSE` |
| 95 | Règles & conversion | `FPaloCubeRules` | `PALO.CUBE_RULES` |
| 96 | Règles & conversion | `FPaloCubeRulesDelete` | `PALO.CUBE_RULES_DELETE` |
| 97 | Règles & conversion | `FPaloCubeRulesMove` | `PALO.CUBE_RULES_MOVE` |
| 98 | Subsets & sous-cube | `FPaloCoordinatesToArray` | `PALO.COORDINATES_TO_ARRAY` |
| 99 | Subsets & sous-cube | `FPaloExpandTypeChildren` | `PALO.EXPAND_TYPE_CHILDREN` |
| 100 | Subsets & sous-cube | `FPaloExpandTypeLeaves` | `PALO.EXPAND_TYPE_LEAVES` |
| 101 | Subsets & sous-cube | `FPaloExpandTypeSelf` | `PALO.EXPAND_TYPE_SELF` |
| 102 | Subsets & sous-cube | `FPaloExpandTypesToArray` | `PALO.EXPAND_TYPES_TO_ARRAY` |
| 103 | Subsets & sous-cube | `FPaloSubcube` | `PALO.SUBCUBE` |
| 104 | Subsets & sous-cube | `FPaloSubset` | `PALO.SUBSET` |
| 105 | Subsets & sous-cube | `FPaloSubsetAliasFilter` | `PALO.SUBSET_ALIAS_FILTER` |
| 106 | Subsets & sous-cube | `FPaloSubsetBasicFilter` | `PALO.SUBSET_BASIC_FILTER` |
| 107 | Subsets & sous-cube | `FPaloSubsetDataFilter` | `PALO.SUBSET_DATA_FILTER` |
| 108 | Subsets & sous-cube | `FPaloSubsetSize` | `PALO.SUBSET_SIZE` |
| 109 | Subsets & sous-cube | `FPaloSubsetSortingFilter` | `PALO.SUBSET_SORTING_FILTER` |
| 110 | Subsets & sous-cube | `FPaloSubsetStructuralFilter` | `PALO.SUBSET_STRUCTURAL_FILTER` |
| 111 | Subsets & sous-cube | `FPaloSubsetTextFilter` | `PALO.SUBSET_TEXT_FILTER` |
| 112 | Verrous & transactions | `FPaloCubeCommit` | `PALO.CUBE_COMMIT` |
| 113 | Verrous & transactions | `FPaloCubeLock` | `PALO.CUBE_LOCK` |
| 114 | Verrous & transactions | `FPaloCubeLocks` | `PALO.CUBE_LOCKS` |
| 115 | Verrous & transactions | `FPaloCubeRollback` | `PALO.CUBE_ROLLBACK` |
| 116 | Verrous & transactions | `FPaloEventLockBegin` | `PALO.EVENT_LOCK_BEGIN` |
| 117 | Verrous & transactions | `FPaloEventLockEnd` | `PALO.EVENT_LOCK_END` |
| 118 | Vues | `FPaloViewAreaDefinition` | `PALO.VIEW_AREA_DEFINITION` |
| 119 | Vues | `FPaloViewAreaGet` | `PALO.VIEW_AREA_GET` |
| 120 | Vues | `FPaloViewAxisDefinition` | `PALO.VIEW_AXIS_DEFINITION` |
| 121 | Vues | `FPaloViewAxisGet` | `PALO.VIEW_AXIS_GET` |
| 122 | Vues | `FPaloViewAxisGetIndex` | `PALO.VIEW_AXIS_GET_INDEX` |
| 123 | Vues | `FPaloViewAxisGetSize` | `PALO.VIEW_AXIS_GET_SIZE` |
| 124 | Vues | `FPaloViewDimension` | `PALO.VIEW_DIMENSION` |
| 125 | Vues | `FPaloViewSubsetDefinition` | `PALO.VIEW_SUBSET_DEFINITION` |
| 126 | Éléments (CRUD & navigation) | `FPaloElementAdd` | `PALO.ELEMENT_ADD` |
| 127 | Éléments (CRUD & navigation) | `FPaloElementAlias` | `PALO.ELEMENT_ALIAS` |
| 128 | Éléments (CRUD & navigation) | `FPaloElementChildcount` | `PALO.ELEMENT_CHILDCOUNT` |
| 129 | Éléments (CRUD & navigation) | `FPaloElementChildname` | `PALO.ELEMENT_CHILDNAME` |
| 130 | Éléments (CRUD & navigation) | `FPaloElementCount` | `PALO.ELEMENT_COUNT` |
| 131 | Éléments (CRUD & navigation) | `FPaloElementCreateBulk` | `PALO.ELEMENT_CREATE_BULK` |
| 132 | Éléments (CRUD & navigation) | `FPaloElementDelete` | `PALO.ELEMENT_DELETE` |
| 133 | Éléments (CRUD & navigation) | `FPaloElementDeleteBulk` | `PALO.ELEMENT_DELETE_BULK` |
| 134 | Éléments (CRUD & navigation) | `FPaloElementFirst` | `PALO.ELEMENT_FIRST` |
| 135 | Éléments (CRUD & navigation) | `FPaloElementIndent` | `PALO.ELEMENT_INDENT` |
| 136 | Éléments (CRUD & navigation) | `FPaloElementIndex` | `PALO.ELEMENT_INDEX` |
| 137 | Éléments (CRUD & navigation) | `FPaloElementIsChild` | `PALO.ELEMENT_IS_CHILD` |
| 138 | Éléments (CRUD & navigation) | `FPaloElementLevel` | `PALO.ELEMENT_LEVEL` |
| 139 | Éléments (CRUD & navigation) | `FPaloElementMove` | `PALO.ELEMENT_MOVE` |
| 140 | Éléments (CRUD & navigation) | `FPaloElementMoveBulk` | `PALO.ELEMENT_MOVE_BULK` |
| 141 | Éléments (CRUD & navigation) | `FPaloElementName` | `PALO.ELEMENT_NAME` |
| 142 | Éléments (CRUD & navigation) | `FPaloElementNext` | `PALO.ELEMENT_NEXT` |
| 143 | Éléments (CRUD & navigation) | `FPaloElementParentcount` | `PALO.ELEMENT_PARENTCOUNT` |
| 144 | Éléments (CRUD & navigation) | `FPaloElementParentname` | `PALO.ELEMENT_PARENTNAME` |
| 145 | Éléments (CRUD & navigation) | `FPaloElementPrev` | `PALO.ELEMENT_PREV` |
| 146 | Éléments (CRUD & navigation) | `FPaloElementRename` | `PALO.ELEMENT_RENAME` |
| 147 | Éléments (CRUD & navigation) | `FPaloElementSibling` | `PALO.ELEMENT_SIBLING` |
| 148 | Éléments (CRUD & navigation) | `FPaloElementType` | `PALO.ELEMENT_TYPE` |
| 149 | Éléments (CRUD & navigation) | `FPaloElementUpdate` | `PALO.ELEMENT_UPDATE` |
| 150 | Éléments (CRUD & navigation) | `FPaloElementWeight` | `PALO.ELEMENT_WEIGHT` |

---

## C. Lien entre section A (règles) et section B (client)

- Les **21** fonctions `palo.*` du moteur décrivent surtout la **navigation dimension / lecture de données** dans les règles. Elles ont des **équivalents** dans la section B sous d’autres noms — par exemple `palo.ename` ↔ `FPaloElementName` (`PALO.ELEMENT_NAME`), `palo.data` ↔ lecture via `FPaloGetdata` / variantes (`PALO.GETDATA`, `PALO.GETDATA_C`, …).
- La section B ajoute tout ce qui est **propre au tableur** : connexion (`PALO.INIT`, `PALO.REGISTER_SERVER`, …), **subsets**, **vues**, **export**, **verrous**, **administration** des objets, etc. — absent du registre `PaloFunctionNodeFactory`.
- Les fonctions **`palo.marker`**, **`palo.cubedimension`**, **`palo.eoffset`** n’ont pas d’homonyme évident en tant que `FPalo*` dans le tableau B : elles restent **spécifiques au langage de règles** (à mapper côté Office 365 selon besoin métier ou via API serveur).

---

## D. Phases de livraison (proposition)

| Phase | Contenu | Dépendances |
|-------|---------|-------------|
| **P0** | Connexion (URL, utilisateur, mot de passe) stockée ; appels API authentifiés (MD5 / token selon API Palo du serveur) ; équivalents de `PALO.INIT` / `PALO.REGISTER_SERVER` / test réseau (`PALO.PING`, `PALO.SERVER_INFO`). | HTTPS + CORS sur le serveur Palo |
| **P1** | Lecture cellule (`PALO.GETDATA` ou alias `PALO.DATA`) + sous-ensemble des fonctions **éléments** les plus utilisées alignées sur section A / B. | P0 |
| **P2** | Couverture étendue section B (écriture, listes, métadonnées) selon priorités produit. | P1 |
| **P3** | Subsets, vues, batch, verrous, règles — optimisation (requêtes groupées). | P2 |

---

## E. Contraintes Office 365 / JavaScript

- Les **fonctions personnalisées** sont **asynchrones** (Promises) ; pas d’accès direct au ruban depuis le même thread que certaines API.
- **Volatilité** : limiter le nombre d’appels réseau par recalcul ; envisager **batch** côté serveur ou file d’appels.
- **Sécurité** : ne pas stocker les mots de passe en clair dans le classeur sans chiffrement ; préférer jetons à durée limitée si l’API le permet.
- **Excel Online** : respect **HTTPS**, **CORS**, parfois pas de ruban pour compléments téléversés — prévoir ouverture du volet depuis **Insertion → Compléments**.

---

## G. Héritage Excel 2010 — `PALO.DATAC`, écriture cube et syntaxes « commande »

### G.1 Où est le « code historique » ?

- **Ce dépôt (`palo-excel-addin`)** ne contient **pas** l’add-in **COM / C++** d’Excel 2010 (aucun `PaloSpreadsheetFuncs` ni `.xll` ici). L’implémentation actuelle est **JavaScript** dans **`docs/functions.js`** : **`PALO.DATAC`** → lecture HTTP **`/cell/value`** ; **`PALO.SETDATA`** → écriture **`/cell/replace`** (splash explicite).
- La **référence source** de l’ancien tableur Jedox / Palo est la bibliothèque client **`PaloSpreadsheetFuncs`** (Jedox ~5.1), décrite en **section B** et pointée depuis le **contexte** (dépôt miroir **`jedox-mirror`**, chemins type `molap/client_libraries/5.1/PaloSpreadsheetFuncs/`, fichiers **`SpreadsheetFuncs.h`** / **`SpreadsheetFuncs.cpp`**). C’est là qu’il faut chercher les signatures **`FPaloGetdata*`** / **`FPaloSetdata*`** et le comportement exact des fonctions Excel historiques.

### G.2 `PALO.DATAC` vs `PALO.GETDATA_C` vs écriture

- Dans l’**inventaire C++** (section **B**), la lecture par **noms d’éléments** (coordonnées « texte ») est surtout portée par **`FPaloGetdataC`** → convention mécanique **`PALO.GETDATA_C`**. Sous **Jedox / Palo for Excel**, le nom **`PALO.DATAC`** est en pratique l’**alias utilisateur** de cette famille **« get data by coordinate names »** (souvent avec **collecte** des appels pour un même cube — voir doc Jedox *Data functions*).
- L’**écriture** dans le cube (y compris sur consolidés / **splash**) côté client historique repose sur les entrées **`FPaloSetdata`**, **`FPaloSetdataA`**, etc. (section **B**, lignes 56–60) → exposées en **`PALO.SETDATA`** et variantes — **pas** sur une « sous-commande » documentée **à l’intérieur** de **`PALO.DATAC`** dans la documentation Jedox actuelle.

### G.3 Syntaxes spéciales (`#`, `!`, etc.) et « `<valeur> like …` »

- Les modes de **splashing** / répartition (ex. préfixes **`#`**, **`!`**, **`!!`** selon les guides Jedox) concernent la **saisie / l’écriture** et le **comportement OLAP**, documentés côté **Jedox** (splashing, saisie manuelle), à distinguer d’une syntaxe d’**arguments** `PALO.DATAC(base; cube; …)` standard.
- Une chaîne du type **`<valeur> like réalisé;2025`** (exemple utilisateur) **n’apparaît pas** dans ce dépôt ni dans le tableau **`FPalo*`** du cahier. Pistes de clarification si vous retrouvez un classeur 2010 :
  - confusion avec **`PALO.DATA`** / **`PALO.SETDATA`** ou avec une **saisie dans la cellule** (comportement « palo data » vs formule) ;
  - **`réalisé`** / **`2025`** comme **noms d’éléments** de dimensions (séparateurs `;` vs virgule selon locale Excel) ;
  - **feuille / macro** métier, ou mot **`like`** issu du **langage de règles** (hors formule tableur).

### G.4 Liens documentation Jedox (hors dépôt)

- [Jedox Data Functions](https://knowledgebase.jedox.com/jedox/planning/jedox-data-functions.htm) — rôles **`PALO.DATA`**, **`PALO.DATAC`**, **`PALO.SETDATA`**, collecte d’appels.
- [Splashing — troubleshooting](https://knowledgebase.jedox.com/jedox/planning/splashing-troubleshooting.htm) — écriture consolidée / erreurs fréquentes.
- **Ce dépôt** — comportement historique **`LIKE` / `COPY`** et flux **Action** sur **`PALO.DATAC`** : voir [`docs/palo-like-copy-datac-action.md`](./palo-like-copy-datac-action.md).

### G.5 Popup « Action » sur une cellule en formule `PALO.DATAC` (spécification cible)

Objectif : parcours guidé selon la **nature du chemin** (présence ou non d’**éléments consolidés** sur les coordonnées de la cellule) et **mode avancé** pour **`LIKE` / `COPY`** aligné sur l’add-in 2010.

**Sémantique LIKE (chemins partiels)** : comme sous Jedox / `parsePath` dans **`PaloSpreadsheetFuncs`**, l’utilisateur ne saisit que les **éléments qui diffèrent** entre la cellule **cible** (chemin complet issu de **`PALO.DATAC`**) et la cellule **source** de référence ; les **autres dimensions** reprennent les **mêmes noms d’éléments** que sur la cible. Exemple : cible `2024, Démographie, titi` et saisie **`300 like 2025`** → source **`2025, Démographie, titi`** ; l’écriture applique la **répartition type LIKE** de cette source vers la cible avec la **valeur agrégée** demandée (**300**). **Ne pas** envoyer la chaîne brute à **`/cell/replace`** (erreur de conversion) : le client doit appeler **`GET /cell/copy`** avec **`name_path`**, **`name_path_to`** et le paramètre optionnel **`value`** (voir `cell_copy.api` dans jedox-mirror). Détail : [`docs/palo-like-copy-datac-action.md`](./palo-like-copy-datac-action.md) § **5.2**.

#### G.5.1 Préalable — Identifier une consolidation sur le chemin

- À partir du **chemin résolu** de la cellule (base, cube, liste d’éléments **un par dimension**), déterminer si **au moins une** coordonnée est un **élément consolidé** (agrégat hiérarchique — aligné sur **`palo.etype` / `PALO.ETYPE`**, métadonnées dimension, ou API type liste d’éléments / info dimension).
- **Branchement** : si **aucune** consolidation → parcours **(A)** ; si **au moins une** → parcours **(B)**.

#### G.5.2 Parcours (A) — Aucune consolidation sur le chemin courant

L’utilisateur doit pouvoir :

1. **Définir une valeur** à la coordonnée courante (écriture cube ; alignement conceptuel avec **`PALO.SETDATA`** / **`/cell/replace`** selon le contrat retenu).
2. **Additionner la valeur** à la coordonnée courante lorsque la valeur saisie est **numérique** (mode cumul / add).
3. **Copier une valeur** depuis une **autre coordonnée** du même cube vers la coordonnée courante — référence client/serveur historique **`GET /cell/copy`** (client **`libpalo_ng`** : `Cube::CellCopy` dans le miroir [jedox-mirror](https://github.com/gpizzetta/jedox-mirror/tree/master/molap/client_libraries/5.1/libpalo_ng/source/Palo/Cube.cpp) ; doc API `cell_copy.api` côté serveur `molap/server/5.1/Api/`).

**Composant transversal** : un **outil de construction de path** (sélection base / cube + choix d’un élément par dimension) produisant un chemin **réutilisable** dans le complément (copie, autres volets, export texte, etc.) — **obligatoire** pour le point **(A).3** et **réutilisable** pour **(B).3**.

#### G.5.3 Parcours (B) — Au moins un élément consolidé sur le chemin

L’utilisateur doit pouvoir :

1. **Définir une valeur** et choisir un **mode de splash** : par exemple valeur sur **éléments de base** uniquement, valeur sur **consolidé** avec **répartition / division** vers la base, ou répartition pilotée comme un **path** (s’aligner sur la doc Jedox *splashing*, les paramètres **`/cell/replace`** / splash, et les notions **`path` / `path_to` / `locked_paths`** de **`/cell/copy`** lorsque le comportement est proche d’une copie ou d’un verrou de chemins).
2. **Idem avec addition** : valeur **numérique** + mode de splash + cumul (même grille de choix qu’en **(B).1**).
3. **Copier une valeur depuis un path** : comme **(A).3**, avec **`/cell/copy`** (chemins source et cible) ; gérer explicitement les **consolidés** sur la source ou la cible (ex. **`use_rules`**, **`locked_paths`** — voir `cell_copy.api` dans jedox-mirror).

**Modes splash HTTP (`?splash=0…5`)** pour **`/cell/replace`** : tableau et libellés UI dans [`palo-like-copy-datac-action.md`](./palo-like-copy-datac-action.md) § **5.1** (aligné sur `normalizeSplashMode` dans `functions.js`). En particulier **`splash=1` (default)** = splash **par défaut serveur** sur le consolidé (répartition sur les bases selon la logique Jedox — **pas** une simple division égale type `#` seul ; le comportement exact dépend de la version).

#### G.5.4 Rappels d’implémentation (hors périmètre immédiat)

- L’implémentation du popup Action sur **`PALO.DATAC`** est décrite dans [`palo-like-copy-datac-action.md`](./palo-like-copy-datac-action.md) (parcours guidé **`/cell/replace`**, mode avancé **`LIKE`/`COPY`** → **`/cell/copy`** + **`value`** pour LIKE).
- Le comportement détaillé LIKE / tokenizer côté **`PaloSpreadsheetFuncs`** (`FPaloSetdata`, `parseCopyParams`, `CellCopyWrapper`) reste la **référence** pour étendre d’autres variantes (`! like`, ordre des jetons, `WITHRULES`, etc.).

---

## F. Suivi

- [x] Référencer le code **`PaloSpreadsheetFuncs`** (5.1) — miroir [jedox-mirror](https://github.com/gpizzetta/jedox-mirror/tree/master/molap/client_libraries/5.1/PaloSpreadsheetFuncs) ; inventaire **150** entrées + helpers section **B.3**.
- [ ] **Popup Action `PALO.DATAC`** : spécification **G.5** (détection consolidation, parcours (A) / (B), outil path réutilisable, `/cell/replace` vs `/cell/copy`, modes splash).
- [ ] Valider le mapping **`PALO.*`** proposé vs noms réels de l’add-in Excel historique (si documentation ou binaire disponible) et vs routes HTTP du serveur (`PaloHttpServer`, jobs, etc.).
- [ ] Tracer une colonne « implémenté / partiel / reporté » dans une table de suivi (feuille projet ou issues GitHub).

---

*Document pour le dépôt **palo-excel-addin** ; la section B est alignée sur `SpreadsheetFuncs.h` du miroir jedox-mirror (branche `master`, chemins indiqués en tête de document).*
