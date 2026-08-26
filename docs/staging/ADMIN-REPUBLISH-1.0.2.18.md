# Republier le manifeste prod — v1.0.2.18

Les assets GitHub Pages sont à jour. **Sans republier le manifeste côté admin M365**, Excel peut continuer à servir l’ancienne version (ruban / custom functions figés).

## URL manifeste production

https://gpizzetta.github.io/palo-excel-addin/manifest.xml

- **Version attendue** : `1.0.2.18`
- **Id Office (prod)** : `9e35717b-2d76-4b84-b9f2-c7c1df86d901`
- **DisplayName** : Palo OLAP Add-in (sans BETA)

## Centre d’administration Microsoft 365

1. Aller dans **Paramètres** → **Applications intégrées** (ou déploiement de compléments centralisé).
2. Ouvrir le complément **Palo OLAP Add-in**.
3. **Mettre à jour / republier** en pointant l’**URL du manifeste** ci-dessus (pas un XML téléchargé une seule fois).
4. Attendre la propagation (parfois quelques minutes à quelques heures selon le tenant).

## Sur chaque poste (toi + collègue Desktop)

1. Fermer Excel complètement.
2. Rouvrir Excel / Excel Online.
3. Ouvrir le volet **Connexion** : le pied de page doit afficher **1.0.2.18**.
4. Si version trop ancienne : retirer le complément, redémarrer Excel, le réinstaller depuis le catalogue / sideload.

## Checklist validation formules

Après version 1.0.2.18 visible :

1. Connexion active + test OK dans le volet.
2. `=PALO.RUNTIME_DIAG()` → chaîne de diagnostic (pas `#VALEUR!`).
3. Une formule `=PALO.DATAC(...);...` réelle → valeur ou `#PALO! …` explicite (pas `#VALEUR!` systématique).
4. `=PALO.ENAME(...)` sur un élément connu.
5. Répéter sur **Excel Online** et **Excel Desktop**.

## Vérifier le déploiement Pages (sans Excel)

```bash
curl -sS https://gpizzetta.github.io/palo-excel-addin/version.json
# {"version":"1.0.2.18",...}
```
