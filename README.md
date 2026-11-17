# Qualité de l'air – Base consolidée DGE / OMS Congo

Application web interactive (en français) de la base consolidée des mesures de qualité de l'air
collectées par la Direction Générale de l'Environnement avec l'OMS Congo
(atelier du 17 au 19 novembre 2026).

## Contenu de l'application

- **Tableau de bord** : filtres par polluant, secteur, département, site et qualité des données ;
  indicateurs clés ; taux de dépassement des valeurs guides OMS ; concentrations par site
  comparées à la valeur guide ; répartition par statut, département et secteur.
- **Carte des stations** : stations géoréférencées, en rouge si au moins un dépassement ;
  téléchargement KML (Google Earth) et GeoJSON (QGIS).
- **Données** : tableau des 2 516 mesures, recherche, tri, export CSV et Excel de la sélection.
- **Émissions** : inventaires TotalEnergies (t/an) et rejets cheminée Dangote vs limites SFI/IFC.
- **Anomalies** : journal des corrections et liste des mesures à vérifier (exportable).
- **Rapport et graphiques** : rapport de synthèse PDF et les 13 figures.
- **Méthode** : règles de traitement, valeurs de référence, liste des fiches.

## Structure

```
data/Base_consolidee_qualite_air_DGE_OMS.xlsx   source unique des données
scripts/build-data.mjs                          Excel -> public/data/donnees.json
scripts/vendor.mjs                              copie Chart.js, Leaflet, ExcelJS dans public/vendor
scripts/build_report.py                         génère le rapport PDF
public/                                         site publié par Netlify
netlify.toml                                    configuration Netlify
```

Le site n'utilise aucun serveur ni base de données : tout est statique. Les bibliothèques sont
servies depuis le site lui-même (pas de CDN), seul le fond de carte OpenStreetMap demande Internet.

## Mettre en ligne sur Netlify

1. Sur [app.netlify.com](https://app.netlify.com), **Add new project → Import an existing project → GitHub**.
2. Choisir le dépôt `junelstone/air-consolidation`.
3. Les réglages sont lus dans `netlify.toml` (commande `npm run build`, dossier `public`) : cliquer sur **Deploy**.

Chaque modification poussée sur la branche `main` redéploie le site automatiquement.

## Mettre à jour les données

1. Remplacer `data/Base_consolidee_qualite_air_DGE_OMS.xlsx` par la nouvelle version
   (mêmes onglets et mêmes en-têtes de colonnes), directement sur GitHub (**Add file → Upload files**) ou en local.
2. Netlify reconstruit le site : le tableau de bord, la carte et le tableau sont à jour.

Le rapport PDF, lui, n'est pas régénéré par Netlify. Pour le refaire après une mise à jour :

```bash
pip install -r scripts/requirements-report.txt
python3 scripts/build_report.py
```

puis pousser le nouveau fichier `public/rapport/Rapport_synthese_qualite_air_DGE_OMS.pdf`.
Les tableaux du rapport sont relus dans l'Excel ; les textes des constats sont dans le script.

## Travailler en local

```bash
npm install
npm run build      # régénère public/data/donnees.json
npm start          # ouvre le site sur http://localhost:3000
```
