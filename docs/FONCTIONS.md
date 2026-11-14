# Module des fonctions PALO

Catalogue des fonctions de classeur du complément. Le contrat historique est Jedox 5.1, fichier `molap/client_libraries/5.1/Common/Documentation.xml`. L'écriture lancée par le menu Action est décrite dans [CAHIER-ECRITURE.md](CAHIER-ECRITURE.md).

## Noms

Le manifeste déclare l'espace de noms `PALO`. Excel place cet espace devant le champ `name` de `functions.json`. La formule est donc `PALO.` suivi de ce nom.

Un nom qui commençait déjà par `PALO_` donnait une formule `PALO.PALO_ECOUNT`. Le nom historique est `ECOUNT`, et la formule est `PALO.ECOUNT`. L'identifiant passé à `CustomFunctions.associate` est le même que le champ `id`. Les fonctions JavaScript gardent leur nom interne : le calcul des fonctions déjà écrites ne change pas, seul le nom vu par Excel change.

Après rechargement du complément, une cellule qui contient encore l'ancien nom affiche `#NAME?`. Il faut réécrire la formule.

| Ancienne formule | Formule |
|---|---|
| `PALO.PALO_SETDATA` | `PALO.SETDATA` |
| `PALO.PALO_ECOUNT` | `PALO.ECOUNT` |
| `PALO.PALO_ECHILDCOUNT` | `PALO.ECHILDCOUNT` |
| `PALO.PALO_ECHILD` | `PALO.ECHILD` |
| `PALO.PALO_EPARENTCOUNT` | `PALO.EPARENTCOUNT` |
| `PALO.PALO_EPARENT` | `PALO.EPARENT` |
| `PALO.PALO_ELEVEL` | `PALO.ELEVEL` |
| `PALO.PALO_EINDENT` | `PALO.EINDENT` |
| `PALO.PALO_ETYPE` | `PALO.ETYPE` |
| `PALO.PALO_EWEIGHT` | `PALO.EWEIGHT` |
| `PALO.PALO_DATABASE_LIST_DIMENSIONS` | `PALO.DATABASE_LIST_DIMENSIONS` |
| `PALO.PALO_CUBE_LIST_DIMENSIONS` | `PALO.CUBE_LIST_DIMENSIONS` |
| `PALO.PALO_DIMENSION_LIST_CUBES` | `PALO.DIMENSION_LIST_CUBES` |
| `PALO.PALO_DIMENSION_LIST_ELEMENTS` | `PALO.DIMENSION_LIST_ELEMENTS` |

Le ruban qui dépose une écriture d'exemple écrit `PALO.SETDATA`.

## Fonctions disponibles

| Formule | Rôle |
|---|---|
| `PALO.DATAC` | Lit une valeur dans un cube Palo (servdb; cube; coord1..coordN). |
| `PALO.SETDATA` | Ecrit une valeur dans un cube Palo. |
| `PALO.ENAME` | Nom de l'élément. Le 4e argument lit cet attribut dans le cube d'attributs de la dimension. Un attribut texte vide renvoie le nom de l'élément. |
| `PALO.ECOUNT` | Donne le nombre d'éléments d'une dimension. |
| `PALO.ECHILDCOUNT` | Donne le nombre d'enfants directs d'un élément. |
| `PALO.ECHILD` | Donne le nom de l'enfant à la position indiquée (la première position est 1). |
| `PALO.EPARENTCOUNT` | Donne le nombre de parents d'un élément. |
| `PALO.EPARENT` | Donne le nom du parent à la position indiquée (la première position est 1). |
| `PALO.ELEVEL` | Donne le niveau de l'élément dans la hiérarchie. |
| `PALO.EINDENT` | Donne le retrait hiérarchique de l'élément, en partant de 1. |
| `PALO.ETYPE` | Renvoie le type de l'élément : numeric, string ou consolidated. |
| `PALO.EWEIGHT` | Renvoie le poids de l'enfant dans le parent indiqué. |
| `PALO.DATABASE_LIST_DIMENSIONS` | Liste les dimensions d'une base, une par ligne. |
| `PALO.CUBE_LIST_DIMENSIONS` | Liste les dimensions d'un cube, dans l'ordre du cube. |
| `PALO.DIMENSION_LIST_CUBES` | Liste les cubes qui utilisent une dimension. |
| `PALO.DIMENSION_LIST_ELEMENTS` | Liste les éléments d'une dimension. |

`PALO.RUNTIME_DIAG` contrôle le runtime Office. `PALO.STORAGE_DIAG` contrôle le stockage de la session. Ces deux fonctions n'existent pas dans Jedox 5.1.

Le menu Action n'est pas une formule. Il écrit dans la cellule `PALO.DATAC` active : remplacement, ajout, pourcentage, splash d'une consolidation, copie `LIKE`. Le détail est dans le cahier d'écriture.

## Fonctions en TODO

Chaque fonction historique absente est enregistrée. Elle ne contacte pas le serveur. Elle accepte les arguments de l'ancienne signature et les ignore. La cellule affiche :

```
#PALO! TODO PALO.NOM : rôle de la fonction
```

L'info-bulle Excel reprend la même phrase, avec le préfixe `TODO.`. Pour implémenter une fonction : remplacer son modèle dans `functions-core.js`, déclarer ses vrais paramètres dans `functions.json`, puis retirer sa ligne de la section correspondante.

### Lectures de cube

| Formule | Ce qu'elle doit faire |
|---|---|
| `PALO.DATA` | Lit la valeur de la cellule de données indiquée à partir du cube. |
| `PALO.DATA.AVG` | Renvoie la valeur moyenne d'une plage de cellules du cube. |
| `PALO.DATA.CNT` | Renvoie le nombre de cellules d'une plage de cellules du cube. |
| `PALO.DATA.MAX` | Renvoie la valeur maximale d'une plage de cellules du cube. |
| `PALO.DATA.MIN` | Renvoie la valeur minimale d'une plage de cellules du cube. |
| `PALO.DATA.SUM` | Renvoie la somme d'une plage de cellules du cube. |
| `PALO.DATAA` | Lit une valeur numérique en acceptant un alias à la place du nom d'élément. |
| `PALO.DATAAC` | Comme DATAA, et signale une erreur si la cellule ne peut pas être lue. |
| `PALO.DATAAT` | Lit une valeur texte en acceptant un alias à la place du nom d'élément. |
| `PALO.DATAATC` | Comme DATAAT, et signale une erreur si la cellule texte ne peut pas être lue. |
| `PALO.DATAT` | Lit la valeur texte de la cellule indiquée. Renvoie une chaîne vide si la lecture échoue. |
| `PALO.DATATC` | Lit la valeur texte de la cellule indiquée et renvoie une erreur si la lecture échoue. Le texte est limité à 255 caractères. |
| `PALO.DATAV` | Lit une cellule du cube et se place dans une formule matricielle Excel. |
| `PALO.DATAX` | Lit une cellule du cube, et ne se calcule que lorsque sa feuille est active. |

### Écritures

| Formule | Ce qu'elle doit faire |
|---|---|
| `PALO.ERROR_LOG` | Journalise une ligne d'import refusée par le serveur. |
| `PALO.SETDATAA` | Écrit une valeur en acceptant un alias à la place du nom d'élément. |
| `PALO.SETDATAAIF` | Écrit une valeur, en acceptant un alias, lorsque la condition est vraie. |
| `PALO.SETDATAA_WRAP` | Variante interne de SETDATAA : enveloppe l'écriture et journalise l'erreur. |
| `PALO.SETDATAIF` | Écrit une valeur dans une cellule du cube lorsque la condition est vraie. |
| `PALO.SETDATA_BULK` | Écrit plusieurs cellules du cube dans un seul appel, sans interpréter de préfixe. |
| `PALO.SETDATA_ERROR` | Journalise une ligne d'import refusée par le serveur. |
| `PALO.SETDATA_WRAP` | Variante interne de SETDATA : enveloppe l'écriture et journalise l'erreur au lieu de la remonter. |

### Éléments

| Formule | Ce qu'elle doit faire |
|---|---|
| `PALO.EADD` | Ajoute un élément, ou un lien de consolidation, dans une dimension. |
| `PALO.EALIAS` | Renvoie l'alias d'un élément, ou les éléments qui portent l'alias indiqué. |
| `PALO.EDELETE` | Supprime l'élément spécifié. |
| `PALO.EFIRST` | Renvoie le premier élément de la dimension, ou le premier enfant du parent indiqué. |
| `PALO.EINDEX` | Renvoie la position de l’élément indiqué dans la dimension. |
| `PALO.EISCHILD` | Vérifie si un élément consolidé contient l’élément indiqué. |
| `PALO.ELEMENT_LIST_ANCESTORS` | Listes de tous les ancêtres de l'élément spécifié |
| `PALO.ELEMENT_LIST_CHILDREN` | Liste les enfants directs d'un élément. |
| `PALO.ELEMENT_LIST_CONSOLIDATION_ELEMENTS` | Retourne une liste avec tous les éléments enfants d'un élément consolidé. |
| `PALO.ELEMENT_LIST_DESCENDANTS` | Liste tous les descendants. |
| `PALO.ELEMENT_LIST_PARENTS` | Retourne une liste de tous les parents d'un élément |
| `PALO.ELEMENT_LIST_SIBLINGS` | Retourne une liste de tous les éléments qui sont sur le même niveau hiérarchique que l’élément spécifié |
| `PALO.ENEXT` | Renvoie le nom de l'élément qui suit l'élément indiqué. |
| `PALO.EPREV` | Renvoie le nom de l'élément qui précède l'élément indiqué. |
| `PALO.ERENAME` | Renomme l’élément indiqué. |
| `PALO.ESIBLING` | Renvoie le nom de l’élément frère défini. |
| `PALO.ETOPLEVEL` | Renvoie le niveau le plus élevé que possède un élément dans la hiérarchie de consolidation de la dimension indiquée. |
| `PALO.EUPDATE` | Met à jour le type, le poids ou la position d'un élément existant. |
| `PALO.EXPANDTYPE` | Regroupe les modes d'expansion (SELF, CHILDREN, LEAVES) d'une agrégation. |

### Listes

| Formule | Ce qu'elle doit faire |
|---|---|
| `PALO.DATABASE_LIST_CUBES` | La fonction établit, dans une plage que vous aurez prédéfinie, une liste des cubes d’une base de données. |
| `PALO.DATABASE_LIST_DIMENSIONS_EXT` | Liste les dimensions d'une base avec leurs attributs : identifiant, type et nombre d'éléments. |
| `PALO.ROOT_LIST_DATABASES` | Donne une liste avec toutes les bases de données du serveur spécifié. |
| `PALO.ROOT_LIST_DATABASES_EXT` | Liste les bases du serveur avec leurs attributs. |

### Sous-ensembles

| Formule | Ce qu'elle doit faire |
|---|---|
| `PALO.AFILTER` | Le filtre attribut supprime des éléments pour lesquels les lignes attribut définies ne correspondent pas. |
| `PALO.DFILTER` | Calcule pour chaque élément d’un sous-ensemble une valeur de données. La valeur peut être numérique ou un String. |
| `PALO.HFILTER` | Filtre un sous-ensemble selon la place de l'élément dans la hiérarchie : niveau, parents, enfants ou feuilles. |
| `PALO.PICKLIST` | Peut ajouter des éléments à un sous-ensemble qui ne peuvent en aucun cas être à nouveau supprimés. En outre, il est possible de définir par ces derniers une multitude d’éléments qui feront office de filtre. |
| `PALO.SERVERSUBSET` | Ancien nom de SERVER_SUBSET : calcule un sous-ensemble d'éléments sur le serveur. |
| `PALO.SERVER_SUBSET` | Renvoie un sous-ensemble d’éléments de dimension. |
| `PALO.SORT` | Définit l'affichage des éléments du sous-ensemble, par exemple en ordre de définition ou en ordre hiérarchique. |
| `PALO.SUBCUBE` | Décrit la tranche de cube lue par DFILTER. La dimension filtrée reste vide, les autres reçoivent un ou plusieurs éléments. |
| `PALO.SUBSET` | Renvoie un sous-ensemble d’éléments de dimension. |
| `PALO.SUBSETSIZE` | Renvoie la longueur d’un extrait d’ensemble d’un élément de dimension. |
| `PALO.SUBSETX` | Renvoie un sous-ensemble d’éléments de dimension si est appelée depuis une feuille active. |
| `PALO.TFILTER` | Éléments dont les noms correspondent à des modèles définis, sont conservés. Le reste est supprimé de l’affichage. Plusieurs critères de filtre sont indiqués dans la matrice. |

### Vues et assistants

| Formule | Ce qu'elle doit faire |
|---|---|
| `PALO.AGR.CHILDREN` | Indique à une agrégation de prendre les enfants directs. |
| `PALO.AGR.LEAVES` | Indique à une agrégation de prendre les éléments de base situés sous l'élément. |
| `PALO.AGR.SELF` | Indique à une agrégation de prendre l'élément lui-même. |
| `PALO.COORD` | Regroupe des coordonnées dans une matrice pour DATA.SUM et les autres agrégats. |
| `PALO.CUBE` | Désigne la zone de cube lue par une vue Palo. |
| `PALO.DIMENSION` | Désigne une dimension dans la définition d'une vue Palo. |
| `PALO.VIEW.AREAGET` | Renvoie les valeurs de la zone d'une vue Palo. |
| `PALO.VIEW.AXISDEF` | Définit un axe (lignes ou colonnes) d'une vue Palo. |
| `PALO.VIEW.AXISGET` | Renvoie les éléments placés sur un axe de vue Palo. |
| `PALO.VIEW.AXISGETINDEX` | Renvoie la position d'un élément sur un axe de vue Palo. |
| `PALO.VIEW.SUBSETDEF` | Définit le sous-ensemble utilisé par un axe de vue Palo. |

### Connexion

| Formule | Ce qu'elle doit faire |
|---|---|
| `PALO.CONNECTIONUSER` | Retourne l'utilisateur de la connexion |
| `PALO.PING` | Vérifie que le serveur Jedox OLAP de la connexion répond. |
| `PALO.REGISTER` | Enregistre un serveur PALO pour qu'une formule puisse s'y connecter. |
| `PALO.REMOVE_CONNECTION` | Ferme la connexion nommée. |
| `PALO.SERVER_INFO` | Retourne des informations sur le serveur Jedox OLAP d'une connexion donnée. |
| `PALO.USERGROUPS` | Retourne le groupe d'utilisateurs de la connexion |

## Prochaines étapes

1. Recharger le complément pour prendre le nouveau `functions.json`. Réécrire les cellules `PALO.PALO_*` avec le nom de la colonne Formule.
2. Lectures proches de `PALO.DATAC` : `DATA`, `DATAT` pour le texte, `DATAA` et `DATAAC` pour les alias, `DATAX` pour la feuille active, puis les agrégats `DATA.SUM`, `DATA.AVG`, `DATA.CNT`, `DATA.MAX` et `DATA.MIN`.
3. Écritures en formule : `SETDATAIF` écrit si la condition est vraie, `SETDATAA` accepte un alias, `SETDATA_BULK` écrit plusieurs cellules en un appel. Le menu Action couvre déjà le splash d'une cellule `DATAC`.
4. Navigation dans une dimension : `EFIRST`, `ENEXT`, `EPREV`, `EINDEX`, `EISCHILD`, `ESIBLING`, `EALIAS`, `ETOPLEVEL`.
5. Listes : enfants, parents, frères, ancêtres, descendants, cubes d'une base, bases du serveur.
6. Administration d'éléments (`EADD`, `EDELETE`, `ERENAME`, `EUPDATE`) une fois le comportement validé : ces fonctions modifient la dimension.
7. Sous-ensembles (`SUBSET` et ses filtres), puis les vues (`VIEW.*`). Ce sont des constructeurs de requête.
8. Connexion (`REGISTER`, `PING`, `REMOVE_CONNECTION`, `SERVER_INFO`). La session est déjà tenue par le volet. Une formule n'en a besoin que si elle doit vivre sans ce volet.

Le cahier d'écriture laisse en dehors du menu Action : `WITHRULES`, `PREDICT`, les suffixes `K` et `M`, les splash 4 et 5.
