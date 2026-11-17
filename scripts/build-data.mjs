// Convertit la base consolidée Excel en fichiers JSON lus par l'application.
// Usage : npm run build   (ou : node scripts/build-data.mjs [chemin.xlsx])
// Pour mettre à jour le site : remplacer data/Base_consolidee_qualite_air_DGE_OMS.xlsx
// en gardant les mêmes onglets et en-têtes de colonnes, puis relancer.
import ExcelJS from "exceljs";
import { mkdir, writeFile, copyFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const src = process.argv[2] || path.join(root, "data", "Base_consolidee_qualite_air_DGE_OMS.xlsx");
const out = path.join(root, "public", "data");

function cellValue(v) {
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === "object") {
    if ("result" in v) return cellValue(v.result);
    if ("richText" in v) return v.richText.map((t) => t.text).join("");
    if ("text" in v) return v.text;
    if ("error" in v) return null;
  }
  if (typeof v === "string") {
    const s = v.trim();
    return s === "" ? null : s;
  }
  return v;
}

// Lit un onglet en tableau d'objets à partir de la ligne d'en-tête.
// maxCol limite la lecture aux colonnes du tableau principal.
function sheetToObjects(ws, { headerRow = 1, maxCol } = {}) {
  const header = [];
  const hr = ws.getRow(headerRow);
  const last = maxCol || hr.cellCount;
  for (let c = 1; c <= last; c++) header.push(cellValue(hr.getCell(c).value));
  const rows = [];
  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const o = {};
    let empty = true;
    header.forEach((h, i) => {
      if (!h) return;
      const v = cellValue(row.getCell(i + 1).value);
      if (v !== null) empty = false;
      o[h] = v;
    });
    if (!empty) rows.push(o);
  }
  return rows;
}

function round(x, d = 4) {
  if (typeof x !== "number" || !Number.isFinite(x)) return x;
  const f = 10 ** d;
  return Math.round(x * f) / f;
}

const wb = new ExcelJS.Workbook();
await wb.xlsx.readFile(src);
const sheet = (name) => {
  const ws = wb.getWorksheet(name);
  if (!ws) throw new Error(`Onglet introuvable dans l'Excel : « ${name} »`);
  return ws;
};

// 1. Mesures (colonnes utiles seulement, pour alléger le fichier)
const KEEP = [
  "ID", "Fiche_source", "Entreprise", "Site", "Secteur", "Département", "District", "Localité",
  "Type_zone", "Matrice", "Station", "Latitude", "Longitude", "Type_document", "Mode_prélèvement",
  "Durée_prélèvement", "Laboratoire", "Année", "Date", "Polluant", "Famille", "Valeur_source",
  "Unité_source", "Signe", "Valeur_µg_m3", "Valeur_guide_µg_m3", "Période_valeur_guide",
  "Référence", "Ratio_valeur_guide", "Statut_conformité", "Statut_qualité", "Commentaires",
];
const mesures = sheetToObjects(sheet("Base_consolidée")).map((r) => {
  const o = {};
  for (const k of KEEP) {
    let v = r[k] ?? null;
    if (k === "Valeur_µg_m3" || k === "Ratio_valeur_guide") v = round(v, 4);
    if (k === "Latitude" || k === "Longitude") v = round(v, 6);
    if (k === "Valeur_source" && v !== null) v = String(v);
    o[k] = v;
  }
  return o;
});

const fiches = sheetToObjects(sheet("Fiches"));
const synthesePolluants = sheetToObjects(sheet("Synthèse_polluants"));
const valeursGuides = sheetToObjects(sheet("Valeurs_guides"), { maxCol: 5 }).map((r) => ({
  Polluant: r.Polluant,
  Valeur_µg_m3: r["Valeur_µg_m3"],
  Période: r["Période"],
  Référence: r["Référence"],
}));
// Dans l'onglet, la 5e colonne porte aussi l'en-tête "Valeur_µg_m3" (valeurs d'émission en texte).
{
  const ws = sheet("Valeurs_guides");
  for (let r = 2, i = 0; r <= ws.rowCount && i < valeursGuides.length; r++) {
    const p = cellValue(ws.getRow(r).getCell(1).value);
    if (!p) continue;
    const extra = cellValue(ws.getRow(r).getCell(5).value);
    if (extra !== null && valeursGuides[i].Valeur_µg_m3 == null) valeursGuides[i].Valeur_µg_m3 = extra;
    i++;
  }
}
const emissions = sheetToObjects(sheet("Émissions_inventaire"), { maxCol: 10 });
const cheminee = sheetToObjects(sheet("Émissions_cheminée"));
const anomalies = sheetToObjects(sheet("Anomalies"));

// 2. Stations : une entrée par station géoréférencée, avec ses dépassements
const stations = new Map();
for (const m of mesures) {
  if (m.Latitude == null || m.Longitude == null) continue;
  if (m["Statut_qualité"] === "Doublon (exclu)") continue;
  const key = `${m.Site}|${m.Station}|${m.Latitude}|${m.Longitude}`;
  if (!stations.has(key)) {
    stations.set(key, {
      site: m.Site, station: m.Station, entreprise: m.Entreprise, secteur: m.Secteur,
      departement: m["Département"], lat: m.Latitude, lon: m.Longitude,
      nb: 0, nbValides: 0, depassements: new Set(), ratioMax: 0,
    });
  }
  const s = stations.get(key);
  s.nb++;
  if (m["Statut_qualité"] === "Valide") s.nbValides++;
  if (m["Statut_conformité"] === "Dépassement") {
    s.depassements.add(m.Polluant);
    if (typeof m.Ratio_valeur_guide === "number") s.ratioMax = Math.max(s.ratioMax, m.Ratio_valeur_guide);
  }
}
const stationsOut = [...stations.values()].map((s) => ({
  ...s, depassements: [...s.depassements].sort(), ratioMax: round(s.ratioMax, 2),
}));

const meta = {
  source: path.basename(src),
  genere_le: new Date().toISOString(),
  nb_mesures: mesures.length,
  nb_valides: mesures.filter((m) => m["Statut_qualité"] === "Valide").length,
  nb_a_verifier: mesures.filter((m) => m["Statut_qualité"] === "À vérifier").length,
  nb_doublons: mesures.filter((m) => m["Statut_qualité"] === "Doublon (exclu)").length,
  nb_depassements: mesures.filter((m) => m["Statut_conformité"] === "Dépassement").length,
  nb_sites: new Set(mesures.filter((m) => m["Statut_qualité"] !== "Doublon (exclu)").map((m) => m.Site)).size,
  nb_entreprises: new Set(mesures.filter((m) => m["Statut_qualité"] !== "Doublon (exclu)").map((m) => m.Entreprise)).size,
  nb_polluants: new Set(mesures.map((m) => m.Polluant)).size,
  nb_stations: stationsOut.length,
  nb_fiches: fiches.length,
};

await mkdir(out, { recursive: true });
const data = { meta, mesures, fiches, synthesePolluants, valeursGuides, emissions, cheminee, anomalies, stations: stationsOut };
await writeFile(path.join(out, "donnees.json"), JSON.stringify(data));
await copyFile(src, path.join(out, "Base_consolidee_qualite_air_DGE_OMS.xlsx"));
console.log(`OK : ${meta.nb_mesures} mesures, ${meta.nb_stations} stations, ${anomalies.length} anomalies -> public/data/donnees.json`);
