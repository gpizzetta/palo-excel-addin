# Palo OLAP Add-in (Excel Office 365)

Add-in Excel Microsoft 365 en fichiers statiques (HTML / JS / JSON), aligne sur le cahier des charges `Exemple_addin_palo`. Publication GitHub Pages : [https://gpizzetta.github.io/palo-excel-addin/](https://gpizzetta.github.io/palo-excel-addin/).

## Structure

- `manifest.xml` (racine du depot) : manifeste Office ; memes URLs que la copie `docs/manifest.xml` pour le sideload depuis Pages.
- `public/` : source des assets (taskpane, commandes, fonctions, JSON).
- `docs/` : copie synchronisee de `public/` + `manifest.xml` pour GitHub Pages lorsque la source du site est le dossier **`/docs`** (branche `main`). Apres chaque changement dans `public/`, relancer : `rsync -a --delete public/ docs/` puis `cp manifest.xml docs/manifest.xml`.

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

- **GitHub Pages** : reglage courant — branche `main`, dossier **`/docs`** (contenu = copie de `public/`). Alternative : dossier **`/public`** comme racine du site ; les URLs du manifeste restent `https://gpizzetta.github.io/palo-excel-addin/...` (sans segment `/public/` dans l’URL).
- Autre hebergeur : copier les fichiers de `public/` a la racine HTTPS du chemin prevu dans `manifest.xml`.

## 2) Manifeste et URLs

Le manifeste a la racine pointe vers `https://gpizzetta.github.io/palo-excel-addin/` (taskpane, `functions.json`, assets). Pour un autre domaine ou compte GitHub, adapter toutes les `DefaultValue` dans `manifest.xml` et la section `AppDomains`, puis resynchroniser `docs/manifest.xml` si besoin.

## 3) Sideload dans Excel

1. Ouvre Excel (Office 365 desktop ou web).
2. Va dans **Insert > My Add-ins > Upload My Add-in**.
3. Charge `manifest.xml` (racine du depot) ou, depuis le site Pages, `https://gpizzetta.github.io/palo-excel-addin/manifest.xml` une fois le deploiement a jour.

## Notes

- `PALO.ENAME` attend **3 arguments** : `servdb`, `dimension`, **`elementId`** (identifiant numerique Palo, pas le libelle). Les arguments en trop sont ignores. Exemple : `=PALO.ENAME("NEW/DWH";"D_ANNEE";2025)` (pas besoin de `;1;""` a la fin).
- Le serveur Palo est appele directement depuis l'addin (Office.js).
- En Excel Web, l'URL Palo doit etre accessible depuis le cloud Microsoft.
- Le HTTPS est requis par Office add-ins.

