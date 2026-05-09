# TODO - Plugin paloaddin

## P0 - Compatibilite formules (documents existants)

- [ ] Aligner les nouvelles formules sur `exemple_addin_palo/client_libraries/5.1/PaloSpreadsheetfuncs`:
  - meme nom de fonction,
  - memes parametres,
  - meme ordre des parametres.
- [ ] Verifier que cette compatibilite permet l'ouverture et le recalcul des documents existants sans adaptation manuelle.

## P0 - Correctif critique DATAC_TEST / resolution en IDs

- [ ] Conserver `PALO.DATAC_TEST` comme fonction de diagnostic avec parametres en dur + retour de l'URL appelee.
- [ ] Dans `PALO.DATAC_TEST`, si l'appel reussit sans erreur, retourner l'URL appelee et la valeur lue.
- [ ] Corriger le flux de resolution pour que les appels Palo utilisent les IDs (base/cube/elements) et non les chaines de noms.
- [ ] Corriger l'erreur observee:
  - `#PALO! HTTP 400 sur https://palo.berdoz.local/element/info?sid=...&name_database=DWH&name_dimension=D_COMPTE&name_element=Chiffre+d%27affaire`
- [ ] Verifier apres correction que l'URL finale de lecture cellule est construite avec `databaseId`, `cubeId` et `name_path` en IDs.
- [ ] Ajouter un log explicite de controle qui affiche la conversion nom -> ID pour chaque segment.

## P1 - Ruban Excel (Ribbon)

- [ ] Avoir exactement 2 boutons dans le ruban:
  - `Connexion` pour gerer les connexions Palo (ajout, edition, suppression, selection),
  - `Exploration` pour explorer les bases de donnees, cubes et dimensions.
- [ ] Verifier que les 2 boutons sont visibles et fonctionnels sur Excel Desktop et Excel Web.

## P1 - Icones Ribbon

- [ ] Corriger les icones manquantes ou illisibles (blanc sur blanc) pour les boutons du ruban.
- [ ] Integrer les assets aux tailles Office requises et verifier le rendu reel dans Excel.

## P0 - Exploration et mappings (cache intelligent)

- [ ] Dans l'exploration de connexion, reutiliser les donnees deja memorisees (cubes, dimensions, elements de dimension) pour les mappings.
- [ ] Eviter les doublons de donnees en memoire et eviter les appels API redondants.
- [ ] Prevoir des informations de debug pour tracer l'origine des donnees (cache vs API) et faciliter le diagnostic.

## P0 - Cycle de connexion et retry des fonctions

- [ ] Pour les fonctions `PALO.DATAC` et autres, ne pas refaire une connexion complete a chaque appel si la session persistante est valide.
- [ ] Utiliser une session persistante avec regeneration/renouvellement selon timeout.
- [ ] En cas d'erreur de connexion, declencher une reconnexion, relancer la lecture du cube et reessayer une seule fois.
- [ ] Tracer clairement ce scenario en debug (appel initial, reconnexion, retry unique, resultat final).
- [ ] Centraliser ce processus de connexion (session, timeout, reconnexion, retry) dans un composant unique reutilisable par tout le plugin.
- [ ] Interdire les implementations locales du meme mecanisme dans les fonctions metier pour eviter les divergences.

## P0 - Debug et tracage global

- [ ] Pendant la phase debug, activer les logs de trace sans dependre d'options manuelles (logs disponibles par defaut en mode debug).
- [ ] Creer une fonction centrale de tracage reutilisable partout (pas de `console.log` disperses).
- [ ] Garantir la visibilite des logs dans la console de la page, y compris en contexte iframe.
- [ ] Ajouter un mecanisme de persistance/relay des traces (cookie ou stockage local) pour conserver et consulter l'historique.
- [ ] Tracer systematiquement les couples `dimension/element` pour verifier qu'ils ne sont jamais inverses.
- [ ] Tracer systematiquement l'URL appelee sur le serveur Palo pour chaque appel API.

## Notes

- Cette TODO est en cours de dictée et sera completee avec les prochains points.
