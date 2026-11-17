# Qualité de l'air – Base consolidée DGE / OMS Congo

Application web interactive (en français) de la base consolidée des mesures de qualité de l'air
collectées par la Direction Générale de l'Environnement avec l'OMS Congo
(atelier du 17 au 19 novembre 2026).

## Contenu de l'application

- **Tableau de bord** : filtres par polluant, secteur, département, site, type d'air (air ambiant par défaut,
  air intérieur, air de site), année, qualité des données et référence benzène (OMS 1,7 ou UE 5 µg/m³) ;
  indicateurs clés ; taux de dépassement des valeurs guides OMS ; concentrations par site
  comparées à la valeur guide ; répartition par statut, département et secteur.
- **Carte des stations** : stations géoréférencées, en rouge si au moins un dépassement ;
  téléchargement KML (Google Earth) et GeoJSON (QGIS).
- **Données** : tableau des 2 516 mesures, recherche, tri, export CSV et Excel de la sélection.
- **Entreprises et émissions** : fiche par entreprise (mesures, dépassements, mesures à vérifier), inventaires
  TotalEnergies (t/an) et rejets cheminée Dangote vs limites SFI/IFC ; bouton « Imprimer la fiche ».
- **Anomalies** : journal des corrections et liste des 232 mesures à vérifier, classées en 5 catégories ;
  export Excel de travail (un onglet par catégorie, colonnes Décision / valeur corrigée / vérifié par à remplir).
- **Rapport et graphiques** : rapport de synthèse PDF et les 13 figures (version figée du 17/11/2026).
- **Méthode** : règles de traitement, références normatives, hypothèses de conversion, valeurs de référence.

Le tableau de bord affiche par défaut l'**air ambiant** : 156 dépassements sur 1 035 mesures comparées (15 %).
Le rapport PDF compte **tous les types d'air** (187 sur 1 174, 16 %) : choisir « Tous les types d'air » pour retrouver ses chiffres.

## Références normatives

| Référentiel | Organisme, année | Valeurs utilisées | Application |
|---|---|---|---|
| Lignes directrices mondiales OMS sur la qualité de l'air | OMS, 2021 | PM2.5 15 µg/m³ (24 h) et 5 (an) ; PM10 45 (24 h) et 15 (an) ; NO2 25 (24 h) et 10 (an) ; SO2 40 (24 h) ; O3 100 (8 h) ; CO 4 000 µg/m³ (24 h) | Air ambiant (référence principale) |
| Air Quality Guidelines for Europe, 2e éd. | OMS Europe, 2000 | Benzène 1,7 µg/m³ (an, risque 10⁻⁵) ; toluène 260 (1 sem.) ; H2S 150 (24 h) ; Pb 0,5 ; Cd 0,005 ; As 0,0066 ; Ni 0,025 ; Hg 1 ; Mn 0,15 (an) ; V 1 µg/m³ (24 h) | COV, H2S et métaux |
| Directive 2008/50/CE | UE, 2008 | Benzène 5 µg/m³ (an) | Information seulement |
| Directive 2004/107/CE | UE, 2004 | As 6, Cd 5, Ni 20 ng/m³ (an) | Information seulement |
| Directives EHS ciment et chaux, tableau 1 | SFI / IFC, 2007 | NOx 600, SO2 400, poussières 30 mg/Nm³ (gaz sec, 0 °C, 101,3 kPa, 10 % O2) | Cheminée Dangote |

Conversions : mg/m³ × 1 000 ; ppm × M / 24,45 × 1 000 (25 °C, 1 atm) ; % × 10 000 → ppm.
Aucune norme congolaise n'a été identifiée (à confirmer par la DGE). Détail dans l'onglet « Méthode » du site.

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
2. Netlify reconstruit le site : le tableau de bord, la carte, le tableau et les fichiers KML / GeoJSON sont à jour.

Le rapport PDF, lui, n'est pas régénéré par Netlify. Pour le refaire après une mise à jour :

```bash
pip install -r scripts/requirements-report.txt
python3 scripts/build_report.py
```

puis pousser le nouveau fichier `public/rapport/Rapport_synthese_qualite_air_DGE_OMS.pdf`.
Les tableaux du rapport sont relus dans l'Excel ; les textes des constats sont dans le script.
Les 13 figures de `public/graphiques/` ne sont pas régénérées par ce script.

## Pendant l'atelier : traiter les mesures à vérifier

1. Onglet **Anomalies → Exporter pour vérification (Excel)** : un onglet par catégorie
   (copie suspecte, métaux µg ou ng, LQ en mg/m³, unité incohérente, valeur improbable).
2. Chaque groupe remplit, ligne par ligne, *Décision* (Confirmée / Corrigée / Rejetée), *Valeur_corrigée*,
   *Unité_corrigée*, *Vérifié_par*, *Date_vérification*, à partir du rapport source. Pas de rapport source = Rejetée.
3. Une seule personne reporte les décisions dans `data/Base_consolidee_qualite_air_DGE_OMS.xlsx`
   (colonne *Statut_qualité* : « Valide » si confirmée ou corrigée, avec la valeur corrigée et le statut de conformité recalculé ;
   « Rejetée » si rejetée, avec *Statut_conformité* « Non évalué (donnée rejetée) » ; la décision est notée dans *Commentaires*), sans changer les onglets ni les en-têtes.
4. Pousser l'Excel sur `main` : le site est à jour en une à deux minutes (vérifier la date en pied de page).
5. Régénérer le rapport PDF une seule fois, à la fin de l'atelier.

Hors connexion : `npm install && npm run build && npm start` sur l'ordinateur de présentation ;
tout fonctionne sauf le fond de carte OpenStreetMap.

## Travailler en local

```bash
npm install
npm run build      # régénère public/data/donnees.json
npm start          # ouvre le site sur http://localhost:3000
```
