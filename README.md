# Palo OLAP Add-in (Excel Office 365)

Add-in Excel Microsoft 365 servi par PHP, aligne sur le cahier des charges `Exemple_addin_palo`.

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

- Un serveur HTTPS deja en place (certificat gere par ton infra)
- PHP 7.4+

## 1) Deployer sur ton serveur PHP

- Copier le dossier `paloaddin/` sur le serveur.
- Exposer `paloaddin/public` en HTTPS, par exemple:
  - `https://ton-domaine/paloaddin/public/taskpane.php`
  - `https://ton-domaine/paloaddin/public/functions.js`

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

