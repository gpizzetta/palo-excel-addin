# Palo OLAP Add-in (Excel Office 365)

Add-in Excel Microsoft 365 en fichiers statiques (HTML / JS / JSON), aligne sur le cahier des charges `Exemple_addin_palo`. Peut etre publie sur GitHub Pages (`public/` comme racine du site).

## Structure

- `manifest.xml` : manifeste Office a sideloader dans Excel.
- `public/` : taskpane, runtime commandes/fonctions, assets JS.

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

- Copier le dossier `paloaddin/` (ou au minimum `paloaddin/public/`) sur ton hebergeur.
- Exemple d’URL apres publication :
  - `https://ton-domaine/.../public/taskpane.html`
  - `https://ton-domaine/.../public/functions.js`
  - Metadonnees des fonctions : `public/functions.json`.

## 2) Mettre a jour le manifeste

Dans `paloaddin/manifest.xml`, remplacer `https://portal-129032.berdoz.local/portal_gip/paloaddin/public`
par l'URL reelle de ton serveur.

## 3) Sideload dans Excel

1. Ouvre Excel (Office 365 desktop ou web).
2. Va dans **Insert > My Add-ins > Upload My Add-in**.
3. Charge le fichier `paloaddin/manifest.xml`.

## Notes

- `PALO.ENAME` attend **3 arguments** : `servdb`, `dimension`, **`elementId`** (identifiant numerique Palo, pas le libelle). Les arguments en trop sont ignores. Exemple : `=PALO.ENAME("NEW/DWH";"D_ANNEE";2025)` (pas besoin de `;1;""` a la fin).
- Le serveur Palo est appele directement depuis l'addin (Office.js).
- En Excel Web, l'URL Palo doit etre accessible depuis le cloud Microsoft.
- Le HTTPS est requis par Office add-ins.

