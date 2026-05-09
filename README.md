# Palo OLAP Add-in (Excel Office 365)

Add-in Excel Microsoft 365 en fichiers statiques (HTML / JS / JSON), aligne sur le cahier des charges `Exemple_addin_palo`. Publication GitHub Pages : [https://gpizzetta.github.io/palo-excel-addin/](https://gpizzetta.github.io/palo-excel-addin/).

## Structure

- **`docs/`** : seul dossier des assets du complément (taskpane, commandes, fonctions, `functions.json`, `manifest.xml`). C’est ce dossier que **GitHub Pages** sert lorsque la source du site est **`/docs`** (branche `main`).
- **`manifest.xml` à la racine** : copie du fichier `docs/manifest.xml` pour pouvoir sideloader depuis la racine du clone. Après toute modification du manifeste dans `docs/`, aligner la racine : `cp docs/manifest.xml manifest.xml`.

## MVP implemente (V1)

- Gestion des connexions Palo (ajout, suppression, selection, stockage local).
- Test de connexion via `server/login` + `server/databases`.
- Exploration de base (databases, cubes, dimensions sur la base `Demo`).
- Fonctions Excel custom:
  - `PALO.DATAC`, `PALO_SETDATA`
  - `ENAME`, `PALO_ECOUNT`, `PALO_ECHILDCOUNT`, `PALO_ECHILD`
  - `PALO_EPARENTCOUNT`, `PALO_EPARENT`, `PALO_ELEVEL`, `PALO_EINDENT`, `PALO_ETYPE`, `PALO_EWEIGHT`
  - `PALO_DATABASE_LIST_DIMENSIONS`, `PALO_CUBE_LIST_DIMENSIONS`, `PALO_DIMENSION_LIST_CUBES`, `PALO_DIMENSION_LIST_ELEMENTS`
- Ruban Excel:
  - Ouvrir taskpane
  - Tester connexion
  - Inserer formules `PALO.DATAC` / `PALO_SETDATA` / `PALO.ENAME`

## Prerequis

- Un site HTTPS pour les assets de l’add-in (GitHub Pages, CDN, ou serveur web classique).
- Le serveur Palo doit accepter les requetes `fetch()` depuis l’origine Office (CORS cote Palo), independamment de l’hebergement des fichiers de l’add-in.

## 1) Deployer (statique)

- **GitHub Pages** : branche `main`, dossier **`/docs`** — rien d’autre a copier ; pousse les changements dans `docs/`.
- **Autre hebergeur** : deployer le contenu de `docs/` (ou le sous-chemin prevu) a la racine HTTPS indiquee dans le manifeste.

## 2) Manifeste et URLs

Les deux fichiers `docs/manifest.xml` et `manifest.xml` (racine) doivent rester identiques. Ils pointent vers `https://gpizzetta.github.io/palo-excel-addin/` (taskpane, `functions.json`, assets). Pour un autre domaine ou compte GitHub, editer `docs/manifest.xml`, puis `cp docs/manifest.xml manifest.xml`.

## 3) Sideload dans Excel

1. Ouvre Excel (Office 365 desktop ou web).
2. Va dans **Insert > My Add-ins > Upload My Add-in**.
3. Charge `manifest.xml` (racine du depot), `docs/manifest.xml`, ou l’URL `https://gpizzetta.github.io/palo-excel-addin/manifest.xml` une fois le deploiement a jour.

## Notes

- `PALO.ENAME` attend **3 arguments** : `servdb`, `dimension`, **`elementId`** (identifiant numerique Palo, pas le libelle). Les arguments en trop sont ignores. Exemple : `=PALO.ENAME("NEW/DWH";"D_ANNEE";2025)` (pas besoin de `;1;""` a la fin).
- Le serveur Palo est appele directement depuis l'addin (Office.js).
- En Excel Web, l'URL Palo doit etre accessible depuis le cloud Microsoft.
- Le HTTPS est requis par Office add-ins.

