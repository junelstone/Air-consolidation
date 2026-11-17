/* Application Qualité de l'air – DGE / OMS Congo
   Données : data/donnees.json, généré depuis l'Excel par scripts/build-data.mjs */
"use strict";

const COMPARED = [
  "Conforme",
  "Conforme (< LQ ≤ valeur guide)",
  "Dépassement",
  "Non concluant (LQ > valeur guide)",
];
const ALL = "__all__";
const FIGURES = [
  ["01_depassements_par_polluant.png", "Figure 01 – Taux de dépassement des valeurs guides par polluant"],
  ["02_points_chauds_sites_polluants.png", "Figure 02 – Points chauds : dépassements par site et par polluant"],
  ["03_camemberts_secteurs_statuts.png", "Figure 03 – Répartition des mesures par secteur et par statut"],
  ["04_PM10_par_site.png", "Figure 04 – PM10 par site"],
  ["05_PM25_par_site.png", "Figure 05 – PM2.5 par site"],
  ["06_benzene_par_site.png", "Figure 06 – Benzène par site"],
  ["07_SO2_par_site.png", "Figure 07 – SO2 par site"],
  ["08_NO2_par_site.png", "Figure 08 – NO2 par site"],
  ["09_depassements_par_departement.png", "Figure 09 – Dépassements par département"],
  ["10_depassements_par_secteur.png", "Figure 10 – Dépassements par secteur d'activité"],
  ["11_emissions_totalenergies.png", "Figure 11 – Émissions déclarées TotalEnergies"],
  ["12_dangote_cheminee.png", "Figure 12 – Rejets cheminée Dangote vs limites IFC"],
  ["13_carte_stations_sud.png", "Figure 13 – Carte des stations (sud du pays)"],
];

let DATA = null;
const state = { polluant: ALL, secteur: ALL, departement: ALL, site: ALL, qualite: "valides",
  q: "", sortKey: "ID", sortDir: 1, page: 0, tab: "dashboard" };
const charts = {};
let map = null, mapLayer = null;

/* ---------- utilitaires ---------- */
const $ = (s) => document.querySelector(s);
const nf = new Intl.NumberFormat("fr-FR", { maximumFractionDigits: 2 });
const nf4 = new Intl.NumberFormat("fr-FR", { maximumSignificantDigits: 4 });
const pf = new Intl.NumberFormat("fr-FR", { style: "percent", maximumFractionDigits: 0 });
const fmt = (v) => (v == null || v === "" ? "" : typeof v === "number" ? (Math.abs(v) < 1 ? nf4.format(v) : nf.format(v)) : v);
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const uniq = (arr) => [...new Set(arr.filter((x) => x != null && x !== ""))].sort((a, b) => String(a).localeCompare(String(b), "fr"));
const css = (name) => getComputedStyle(document.documentElement).getPropertyValue(name).trim();
const median = (xs) => {
  if (!xs.length) return null;
  const s = [...xs].sort((a, b) => a - b), m = s.length >> 1;
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
};
const isCompared = (m) => COMPARED.includes(m["Statut_conformité"]);
const tagClass = (statut) => {
  if (!statut) return "neutral";
  if (statut.startsWith("Dépassement")) return "bad";
  if (statut.startsWith("Conforme") || statut === "Valide") return "ok";
  if (statut.startsWith("Non concluant") || statut.startsWith("À vérifier") || statut.startsWith("Non évalué (donnée")) return "warn";
  return "neutral";
};

/* ---------- filtres ---------- */
function qualityOk(m) {
  const q = m["Statut_qualité"];
  if (state.qualite === "valides") return q === "Valide";
  if (state.qualite === "horsdoublons") return q !== "Doublon (exclu)";
  return true;
}
function filtered({ ignorePolluant = false, ignoreQualite = false } = {}) {
  return DATA.mesures.filter((m) =>
    (ignorePolluant || state.polluant === ALL || m.Polluant === state.polluant) &&
    (state.secteur === ALL || m.Secteur === state.secteur) &&
    (state.departement === ALL || m["Département"] === state.departement) &&
    (state.site === ALL || m.Site === state.site) &&
    (ignoreQualite || qualityOk(m)));
}

function fillSelect(sel, values, current, allLabel) {
  sel.innerHTML = `<option value="${ALL}">${allLabel}</option>` +
    values.map((v) => `<option value="${esc(v)}"${v === current ? " selected" : ""}>${esc(v)}</option>`).join("");
  if (current !== ALL && !values.includes(current)) sel.value = ALL;
}
function refreshFilterOptions() {
  const ms = DATA.mesures;
  fillSelect($("#f-polluant"), uniq(ms.map((m) => m.Polluant)), state.polluant, "Tous les polluants");
  fillSelect($("#f-secteur"), uniq(ms.map((m) => m.Secteur)), state.secteur, "Tous les secteurs");
  fillSelect($("#f-departement"), uniq(ms.map((m) => m["Département"])), state.departement, "Tous les départements");
  const sites = uniq(ms.filter((m) => (state.secteur === ALL || m.Secteur === state.secteur) &&
    (state.departement === ALL || m["Département"] === state.departement)).map((m) => m.Site));
  if (state.site !== ALL && !sites.includes(state.site)) state.site = ALL;
  fillSelect($("#f-site"), sites, state.site, "Tous les sites");
  $("#f-qualite").value = state.qualite;
}
function bindFilters() {
  const map = { "#f-polluant": "polluant", "#f-secteur": "secteur", "#f-departement": "departement", "#f-site": "site", "#f-qualite": "qualite" };
  for (const [sel, key] of Object.entries(map)) {
    $(sel).addEventListener("change", (e) => { state[key] = e.target.value; state.page = 0; refreshFilterOptions(); renderAll(); saveHash(); });
  }
  $("#f-reset").addEventListener("click", () => {
    Object.assign(state, { polluant: ALL, secteur: ALL, departement: ALL, site: ALL, qualite: "valides", q: "", page: 0 });
    $("#q").value = ""; refreshFilterOptions(); renderAll(); saveHash();
  });
}
function saveHash() {
  const p = new URLSearchParams();
  p.set("onglet", state.tab);
  for (const k of ["polluant", "secteur", "departement", "site"]) if (state[k] !== ALL) p.set(k, state[k]);
  if (state.qualite !== "valides") p.set("qualite", state.qualite);
  if (state.tab === "emissions" && emState.entreprise) { p.set("entreprise", emState.entreprise); if (emState.site !== ALL) p.set("site_em", emState.site); }
  history.replaceState(null, "", "#" + p.toString());
}
function loadHash() {
  const p = new URLSearchParams(location.hash.slice(1));
  for (const k of ["polluant", "secteur", "departement", "site", "qualite"]) if (p.get(k)) state[k] = p.get(k);
  if (p.get("onglet")) state.tab = p.get("onglet");
  if (p.get("entreprise")) emState.entreprise = p.get("entreprise");
  if (p.get("site_em")) emState.site = p.get("site_em");
}

/* ---------- onglets ---------- */
function bindTabs() {
  document.querySelectorAll(".tabs button").forEach((b) => b.addEventListener("click", () => showTab(b.dataset.tab)));
}
function showTab(tab) {
  if (!document.getElementById("tab-" + tab)) tab = "dashboard";
  state.tab = tab;
  document.querySelectorAll(".tabs button").forEach((b) => b.setAttribute("aria-selected", String(b.dataset.tab === tab)));
  document.querySelectorAll(".panel").forEach((p) => p.classList.toggle("active", p.id === "tab-" + tab));
  $("#filters").classList.toggle("hidden", ["emissions", "rapport", "methode"].includes(tab));
  if (tab === "carte") { initMap(); setTimeout(() => { map.invalidateSize(); renderMap(); }, 50); }
  saveHash();
}

/* ---------- graphiques ---------- */
function chartDefaults() {
  Chart.defaults.color = css("--muted");
  Chart.defaults.borderColor = css("--border");
  Chart.defaults.font.family = getComputedStyle(document.body).fontFamily;
}
function makeChart(id, config) {
  if (charts[id]) charts[id].destroy();
  charts[id] = new Chart(document.getElementById(id), config);
  return charts[id];
}
// Ligne verticale (ou horizontale) de la valeur guide
const guideLinePlugin = {
  id: "guideLine",
  afterDatasetsDraw(chart, _a, opts) {
    if (opts == null || opts.value == null) return;
    const horizontal = chart.options.indexAxis === "y";
    const scale = horizontal ? chart.scales.x : chart.scales.y;
    const pos = scale.getPixelForValue(opts.value);
    const { ctx, chartArea: a } = chart;
    if ((horizontal && (pos < a.left || pos > a.right)) || (!horizontal && (pos < a.top || pos > a.bottom))) return;
    ctx.save();
    ctx.strokeStyle = opts.color; ctx.lineWidth = 2; ctx.setLineDash([6, 4]);
    ctx.beginPath();
    if (horizontal) { ctx.moveTo(pos, a.top); ctx.lineTo(pos, a.bottom); }
    else { ctx.moveTo(a.left, pos); ctx.lineTo(a.right, pos); }
    ctx.stroke();
    ctx.setLineDash([]); ctx.fillStyle = opts.color; ctx.font = "600 12px " + Chart.defaults.font.family;
    const label = opts.label || "";
    if (horizontal) { ctx.textAlign = pos > a.right - 160 ? "right" : "left"; ctx.fillStyle = opts.color; ctx.fillText(label, pos + (ctx.textAlign === "left" ? 6 : -6), a.bottom - 6); }
    else { ctx.textAlign = "right"; ctx.fillText(label, a.right, pos - 6); }
    ctx.restore();
  },
};

function guideFor(polluant) {
  const counts = new Map();
  for (const m of DATA.mesures) {
    if (m.Polluant !== polluant || m["Valeur_guide_µg_m3"] == null) continue;
    const k = m["Valeur_guide_µg_m3"] + "|" + (m["Période_valeur_guide"] || "");
    counts.set(k, (counts.get(k) || 0) + 1);
  }
  if (!counts.size) return null;
  const [k] = [...counts.entries()].sort((a, b) => b[1] - a[1])[0];
  const [v, p] = k.split("|");
  return { value: Number(v), periode: p };
}

function statsByPolluant(rows) {
  const g = new Map();
  for (const m of rows) {
    if (!g.has(m.Polluant)) g.set(m.Polluant, []);
    g.get(m.Polluant).push(m);
  }
  return [...g.entries()].map(([p, ms]) => {
    const detected = ms.filter((m) => typeof m["Valeur_µg_m3"] === "number" && m.Signe !== "<").map((m) => m["Valeur_µg_m3"]);
    const comp = ms.filter(isCompared);
    const dep = ms.filter((m) => m["Statut_conformité"] === "Dépassement");
    const g2 = guideFor(p);
    return {
      polluant: p, famille: ms[0].Famille, n: ms.length, sites: new Set(ms.map((m) => m.Site)).size,
      min: detected.length ? Math.min(...detected) : null, med: median(detected),
      max: detected.length ? Math.max(...detected) : null,
      comp: comp.length, dep: dep.length, nc: ms.filter((m) => m["Statut_conformité"] === "Non concluant (LQ > valeur guide)").length,
      taux: comp.length ? dep.length / comp.length : null, guide: g2,
    };
  });
}

function renderKpis(rows) {
  const comp = rows.filter(isCompared).length;
  const dep = rows.filter((m) => m["Statut_conformité"] === "Dépassement").length;
  const k = [
    [nf.format(rows.length), "mesures sélectionnées"],
    [nf.format(new Set(rows.map((m) => m.Site)).size), "sites"],
    [nf.format(new Set(rows.map((m) => m.Entreprise)).size), "entreprises"],
    [nf.format(comp), "mesures comparables à une valeur guide"],
    [nf.format(dep), "dépassements", "bad"],
    [comp ? pf.format(dep / comp) : "–", "taux de dépassement", "bad"],
  ];
  $("#kpis").innerHTML = k.map(([v, l, c]) => `<div class="kpi ${c || ""}"><div class="v">${v}</div><div class="l">${l}</div></div>`).join("");
}

function renderDepassementsChart(rows) {
  const stats = statsByPolluant(filtered({ ignorePolluant: true }))
    .filter((s) => s.comp >= 3).sort((a, b) => b.taux - a.taux || b.dep - a.dep).slice(0, 20);
  const bad = css("--bad"), ok = css("--ok"), acc = css("--accent");
  makeChart("c-depassements", {
    type: "bar",
    data: {
      labels: stats.map((s) => s.polluant),
      datasets: [{
        label: "Taux de dépassement", data: stats.map((s) => Math.round(s.taux * 1000) / 10),
        backgroundColor: stats.map((s) => (s.polluant === state.polluant ? acc : s.taux >= 0.25 ? bad : s.taux > 0 ? css("--warn") : ok)),
        borderRadius: 4,
      }],
    },
    options: {
      maintainAspectRatio: false,
      scales: { y: { beginAtZero: true, max: 100, ticks: { callback: (v) => v + " %" } } },
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: (c) => { const s = stats[c.dataIndex]; return `${s.dep} dépassements sur ${s.comp} mesures comparées (${c.formattedValue} %)`; },
          afterLabel: (c) => { const g = stats[c.dataIndex].guide; return g ? `Valeur guide : ${fmt(g.value)} µg/m³ (${g.periode})` : ""; } } },
      },
      onClick: (_e, els) => {
        if (!els.length) return;
        const p = stats[els[0].index].polluant;
        state.polluant = state.polluant === p ? ALL : p; state.page = 0;
        refreshFilterOptions(); renderAll(); saveHash();
      },
    },
  });
}

function renderSitesChart() {
  let pol = state.polluant;
  if (pol === ALL) pol = "Benzène";
  const rows = filtered({ ignorePolluant: true }).filter((m) => m.Polluant === pol && typeof m["Valeur_µg_m3"] === "number");
  const g = new Map();
  for (const m of rows) {
    const key = m.Site;
    if (!g.has(key)) g.set(key, []);
    g.get(key).push(m);
  }
  const guide = guideFor(pol);
  const items = [...g.entries()].map(([site, ms]) => {
    const vals = ms.map((m) => m["Valeur_µg_m3"]);
    return { site, max: Math.max(...vals), med: median(vals), n: ms.length, dep: ms.filter((m) => m["Statut_conformité"] === "Dépassement").length };
  }).sort((a, b) => b.max - a.max);
  $("#t-sites").textContent = `${pol} : concentrations par site comparées à la valeur guide`;
  $("#h-sites").textContent = (state.polluant === ALL ? "Polluant affiché par défaut : benzène. Choisissez un polluant dans les filtres pour en changer. " : "") +
    (guide ? `Ligne pointillée : valeur guide ${fmt(guide.value)} µg/m³ (${guide.periode}). ` : "Pas de valeur guide OMS pour ce polluant. ") +
    "Barres : maximum et médiane des mesures de chaque site.";
  const vals = items.flatMap((i) => [i.max, i.med]).filter((v) => v > 0);
  const useLog = vals.length && Math.max(...vals) / Math.max(Math.min(...vals), 1e-9) > 200;
  const bad = css("--bad"), ok = css("--ok");
  const holder = document.getElementById("c-sites").parentElement;
  holder.style.height = Math.max(320, items.length * 26 + 80) + "px";
  makeChart("c-sites", {
    type: "bar",
    data: {
      labels: items.map((i) => i.site),
      datasets: [
        { label: "Maximum", data: items.map((i) => i.max), backgroundColor: items.map((i) => (guide && i.max > guide.value ? bad : ok)), borderRadius: 3 },
        { label: "Médiane", data: items.map((i) => i.med), backgroundColor: css("--neutral"), borderRadius: 3 },
      ],
    },
    options: {
      indexAxis: "y", maintainAspectRatio: false,
      scales: { x: { type: useLog ? "logarithmic" : "linear", ticks: useLog ? { callback: (v) => (Math.abs(Math.log10(v) % 1) < 1e-9 ? fmt(v) : "") } : {}, title: { display: true, text: "µg/m³" + (useLog ? " (échelle logarithmique)" : "") } } },
      plugins: {
        guideLine: guide ? { value: guide.value, color: css("--text"), label: `Valeur guide ${fmt(guide.value)}` } : null,
        tooltip: { callbacks: { label: (c) => `${c.dataset.label} : ${fmt(c.raw)} µg/m³`,
          afterBody: (cs) => { const i = items[cs[0].dataIndex]; return `${i.n} mesure(s), ${i.dep} dépassement(s)`; } } },
      },
    },
    plugins: [guideLinePlugin],
  });
}

function renderStatuts(rows) {
  const c = new Map();
  for (const m of rows) c.set(m["Statut_conformité"], (c.get(m["Statut_conformité"]) || 0) + 1);
  const entries = [...c.entries()].sort((a, b) => b[1] - a[1]);
  const color = (s) => ({ bad: css("--bad"), ok: css("--ok"), warn: css("--warn"), neutral: css("--neutral") }[tagClass(s)]);
  const palette = entries.map(([s], i) => {
    if (s === "Conforme (< LQ ≤ valeur guide)") return "#7fc8a0";
    if (s && s.startsWith("Non comparable")) return i % 2 ? "#a9b6b0" : css("--neutral");
    return color(s);
  });
  makeChart("c-statuts", {
    type: "doughnut",
    data: { labels: entries.map((e) => e[0]), datasets: [{ data: entries.map((e) => e[1]), backgroundColor: palette, borderColor: css("--surface"), borderWidth: 2 }] },
    options: { maintainAspectRatio: false, plugins: { legend: { position: "right", labels: { boxWidth: 12, font: { size: 11 } } } } },
  });
}

function stackedBy(id, rows, key) {
  const g = new Map();
  for (const m of rows) {
    if (!isCompared(m)) continue;
    const k = m[key];
    if (!g.has(k)) g.set(k, { dep: 0, conf: 0, nc: 0 });
    const s = m["Statut_conformité"], o = g.get(k);
    if (s === "Dépassement") o.dep++; else if (s.startsWith("Conforme")) o.conf++; else o.nc++;
  }
  const items = [...g.entries()].sort((a, b) => b[1].dep - a[1].dep);
  makeChart(id, {
    type: "bar",
    data: {
      labels: items.map((i) => i[0]),
      datasets: [
        { label: "Dépassement", data: items.map((i) => i[1].dep), backgroundColor: css("--bad") },
        { label: "Conforme", data: items.map((i) => i[1].conf), backgroundColor: css("--ok") },
        { label: "Non concluant", data: items.map((i) => i[1].nc), backgroundColor: css("--warn") },
      ],
    },
    options: {
      maintainAspectRatio: false, scales: { x: { stacked: true }, y: { stacked: true, title: { display: true, text: "mesures comparées" } } },
      plugins: { tooltip: { callbacks: { footer: (cs) => { const o = items[cs[0].dataIndex][1]; const t = o.dep + o.conf + o.nc; return `Taux de dépassement : ${pf.format(o.dep / t)}`; } } } },
    },
  });
}

function renderSyntheseTable(rows) {
  const stats = statsByPolluant(rows).sort((a, b) => (b.taux ?? -1) - (a.taux ?? -1) || b.n - a.n);
  const head = ["Polluant", "Famille", "Mesures", "Sites", "Min", "Médiane", "Max", "Valeur guide", "Comparées", "Dépassements", "Non concluants", "Taux"];
  $("#t-synthese").innerHTML = `<thead><tr>${head.map((h) => `<th>${h}</th>`).join("")}</tr></thead><tbody>` +
    stats.map((s) => `<tr><td><b>${esc(s.polluant)}</b></td><td>${esc(s.famille)}</td><td class="num">${s.n}</td><td class="num">${s.sites}</td>
      <td class="num">${fmt(s.min)}</td><td class="num">${fmt(s.med)}</td><td class="num">${fmt(s.max)}</td>
      <td class="num">${s.guide ? fmt(s.guide.value) + " <small>(" + esc(s.guide.periode) + ")</small>" : "–"}</td>
      <td class="num">${s.comp}</td><td class="num">${s.dep}</td><td class="num">${s.nc}</td>
      <td class="num">${s.taux == null ? "–" : `<span class="tag ${s.taux >= 0.25 ? "bad" : s.taux > 0 ? "warn" : "ok"}">${pf.format(s.taux)}</span>`}</td></tr>`).join("") +
    "</tbody>";
}

function renderDashboard() {
  const rows = filtered();
  renderKpis(rows);
  renderDepassementsChart(rows);
  renderSitesChart();
  renderStatuts(rows);
  stackedBy("c-departements", rows, "Département");
  stackedBy("c-secteurs", rows, "Secteur");
  renderSyntheseTable(rows);
}

/* ---------- carte ---------- */
function initMap() {
  if (map) return;
  map = L.map("map", { scrollWheelZoom: true }).setView([-4.4, 12.3], 7);
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 18, attribution: "&copy; contributeurs OpenStreetMap",
  }).addTo(map);
  mapLayer = L.layerGroup().addTo(map);
}
function renderMap() {
  if (!map) return;
  mapLayer.clearLayers();
  const rows = filtered();
  const g = new Map();
  let sansCoord = 0;
  for (const m of rows) {
    if (m.Latitude == null || m.Longitude == null) { sansCoord++; continue; }
    const k = `${m.Latitude},${m.Longitude}|${m.Station}`;
    if (!g.has(k)) g.set(k, { lat: m.Latitude, lon: m.Longitude, site: m.Site, station: m.Station, entreprise: m.Entreprise,
      secteur: m.Secteur, dep: m["Département"], zone: m.Type_zone, n: 0, dps: new Map(), pols: new Set() });
    const s = g.get(k);
    s.n++; s.pols.add(m.Polluant);
    if (m["Statut_conformité"] === "Dépassement") {
      const r = m.Ratio_valeur_guide || 0;
      s.dps.set(m.Polluant, Math.max(s.dps.get(m.Polluant) || 0, r));
    }
  }
  const bad = css("--bad"), ok = css("--ok");
  const bounds = [];
  for (const s of g.values()) {
    const has = s.dps.size > 0;
    const maxR = Math.max(0, ...s.dps.values());
    const radius = has ? Math.min(16, 7 + Math.log10(Math.max(maxR, 1)) * 3) : 6;
    const mk = L.circleMarker([s.lat, s.lon], { radius, color: "#fff", weight: 1.5, fillColor: has ? bad : ok, fillOpacity: 0.9 });
    const dps = [...s.dps.entries()].sort((a, b) => b[1] - a[1])
      .map(([p, r]) => `<li>${esc(p)} : jusqu'à ${nf.format(r)} × la valeur guide</li>`).join("");
    mk.bindPopup(`<b>${esc(s.site)}</b> – station ${esc(s.station ?? "")}<br>${esc(s.entreprise)} · ${esc(s.secteur)}<br>
      ${esc(s.dep)}${s.zone ? " · zone " + esc(s.zone).toLowerCase() : ""}<br>${s.n} mesure(s), ${s.pols.size} polluant(s)<br>
      ${has ? `<b style="color:${bad}">Dépassements :</b><ul style="margin:4px 0 0;padding-left:16px">${dps}</ul>` : `<span style="color:${ok}">Aucun dépassement</span>`}`);
    mk.bindTooltip(`${s.site} (${s.station ?? ""})`);
    mk.addTo(mapLayer);
    bounds.push([s.lat, s.lon]);
  }
  if (bounds.length) map.fitBounds(bounds, { padding: [30, 30], maxZoom: 11 });
  $("#map-note").textContent = `${g.size} station(s) affichée(s) pour la sélection. ${sansCoord ? nf.format(sansCoord) + " mesure(s) sans coordonnées GPS ne sont pas cartographiées." : ""}`;
}

/* ---------- tableau des mesures ---------- */
const COLS = [
  ["ID", "ID", "num"], ["Entreprise", "Entreprise"], ["Site", "Site"], ["Station", "Station"],
  ["Département", "Département"], ["Secteur", "Secteur"], ["Matrice", "Matrice"], ["Année", "Année", "num"],
  ["Polluant", "Polluant"], ["Valeur_source", "Valeur source", "num"], ["Unité_source", "Unité"],
  ["Valeur_µg_m3", "µg/m³", "num"], ["Valeur_guide_µg_m3", "Valeur guide", "num"], ["Ratio_valeur_guide", "Ratio", "num"],
  ["Statut_conformité", "Conformité"], ["Statut_qualité", "Qualité"], ["Commentaires", "Commentaires"],
];
const PAGE = 50;
function tableRows() {
  const q = state.q.trim().toLowerCase();
  let rows = filtered();
  if (q) rows = rows.filter((m) => COLS.some(([k]) => String(m[k] ?? "").toLowerCase().includes(q)) || String(m.Fiche_source).toLowerCase().includes(q));
  const k = state.sortKey, d = state.sortDir;
  return [...rows].sort((a, b) => {
    const x = a[k], y = b[k];
    if (x == null && y == null) return 0; if (x == null) return 1; if (y == null) return -1;
    return (typeof x === "number" && typeof y === "number" ? x - y : String(x).localeCompare(String(y), "fr", { numeric: true })) * d;
  });
}
function renderTable() {
  const rows = tableRows();
  const pages = Math.max(1, Math.ceil(rows.length / PAGE));
  state.page = Math.min(state.page, pages - 1);
  const slice = rows.slice(state.page * PAGE, (state.page + 1) * PAGE);
  $("#table-count").textContent = `${nf.format(rows.length)} mesure(s) correspondant aux filtres. Cliquez sur un en-tête pour trier.`;
  $("#t-mesures").innerHTML = `<thead><tr>${COLS.map(([k, l]) => `<th data-k="${k}">${l}${state.sortKey === k ? (state.sortDir > 0 ? " ▲" : " ▼") : ""}</th>`).join("")}</tr></thead><tbody>` +
    slice.map((m) => `<tr>${COLS.map(([k, , t]) => {
      let v = m[k];
      if (k === "Statut_conformité" || k === "Statut_qualité") return `<td><span class="tag ${tagClass(v)}">${esc(v)}</span></td>`;
      if (k === "Valeur_source" && m.Signe === "<" && !String(v).includes("<")) v = "< " + v;
      if (k === "ID" || k === "Année") return `<td class="num">${esc(v)}</td>`;
      return `<td class="${t || ""}${k === "Commentaires" ? " long" : ""}">${esc(typeof v === "number" ? fmt(v) : v)}</td>`;
    }).join("")}</tr>`).join("") + "</tbody>";
  $("#t-mesures thead").onclick = (e) => {
    const th = e.target.closest("th"); if (!th) return;
    const k = th.dataset.k;
    if (state.sortKey === k) state.sortDir *= -1; else { state.sortKey = k; state.sortDir = 1; }
    renderTable();
  };
  $("#pager").innerHTML = `<button class="btn" id="pg-prev" ${state.page === 0 ? "disabled" : ""}>‹ Précédent</button>
    <span>Page ${state.page + 1} / ${pages}</span>
    <button class="btn" id="pg-next" ${state.page >= pages - 1 ? "disabled" : ""}>Suivant ›</button>`;
  $("#pg-prev").onclick = () => { state.page--; renderTable(); };
  $("#pg-next").onclick = () => { state.page++; renderTable(); };
}

/* ---------- exports ---------- */
function download(name, blob) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}
function exportCsv(rows, name) {
  if (!rows.length) return;
  const keys = Object.keys(rows[0]);
  const cell = (v) => { if (v == null) return ""; const s = typeof v === "number" ? String(v).replace(".", ",") : String(v); return /[;"\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; };
  const csv = "﻿" + [keys.join(";"), ...rows.map((r) => keys.map((k) => cell(r[k])).join(";"))].join("\r\n");
  download(name, new Blob([csv], { type: "text/csv;charset=utf-8" }));
}
let excelJsLoading = null;
function loadExcelJs() {
  if (window.ExcelJS) return Promise.resolve();
  excelJsLoading ??= new Promise((ok, ko) => {
    const sc = document.createElement("script");
    sc.src = "vendor/exceljs.min.js"; sc.onload = ok; sc.onerror = () => ko(new Error("chargement d'ExcelJS impossible"));
    document.head.appendChild(sc);
  });
  return excelJsLoading;
}
async function exportXlsx(rows, name, sheet = "Mesures") {
  if (!rows.length) return;
  await loadExcelJs();
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(sheet.slice(0, 31));
  const keys = Object.keys(rows[0]);
  ws.columns = keys.map((k) => ({ header: k, key: k, width: Math.min(40, Math.max(10, k.length + 2)) }));
  ws.addRows(rows);
  ws.getRow(1).font = { bold: true };
  ws.views = [{ state: "frozen", ySplit: 1 }];
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: keys.length } };
  const buf = await wb.xlsx.writeBuffer();
  download(name, new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
}
const stamp = () => new Date().toISOString().slice(0, 10);

/* ---------- émissions (par entreprise) ---------- */
const PAL = ["#0b6e4f", "#c0392b", "#2f6db3", "#c27c0e", "#7b4fa0", "#1f9aa8", "#8a6d3b", "#d35d8f", "#5b6b64", "#3e8e2f"];
const emState = { entreprise: null, site: ALL, polluant: null };

function companies() {
  const inv = new Set(DATA.emissions.map((e) => e.Entreprise));
  const ch = new Set(DATA.cheminee.map((e) => e.Entreprise));
  const all = uniq([...DATA.mesures.filter((m) => m["Statut_qualité"] !== "Doublon (exclu)").map((m) => m.Entreprise), ...inv, ...ch]);
  return all.map((e) => ({ name: e, inv: inv.has(e), ch: ch.has(e) }))
    .sort((x, y) => (Number(y.inv || y.ch) - Number(x.inv || x.ch)) || x.name.localeCompare(y.name, "fr"));
}

function initEmissions() {
  const list = companies();
  const sel = $("#e-entreprise");
  const withEm = list.filter((c) => c.inv || c.ch), without = list.filter((c) => !c.inv && !c.ch);
  sel.innerHTML =
    `<optgroup label="Avec émissions déclarées">${withEm.map((c) => `<option value="${esc(c.name)}">${esc(c.name)} (${c.inv ? "inventaire t/an" : "cheminée mg/Nm³"})</option>`).join("")}</optgroup>` +
    `<optgroup label="Mesures de concentration uniquement">${without.map((c) => `<option value="${esc(c.name)}">${esc(c.name)}</option>`).join("")}</optgroup>`;
  emState.entreprise = emState.entreprise && list.some((c) => c.name === emState.entreprise) ? emState.entreprise : withEm[0]?.name || list[0].name;
  sel.value = emState.entreprise;
  sel.onchange = () => { emState.entreprise = sel.value; emState.site = ALL; emState.polluant = null; renderEmissions(); saveHash(); };
  $("#e-site").onchange = (e) => { emState.site = e.target.value; renderEmissions(); saveHash(); };
  $("#e-polluant").onchange = (e) => { emState.polluant = e.target.value; renderEmissions(); };
}

function renderEmissions() {
  const ent = emState.entreprise;
  const inv = DATA.emissions.filter((e) => e.Entreprise === ent && typeof e["Émission_t_an"] === "number");
  const ch = DATA.cheminee.filter((e) => e.Entreprise === ent);
  const mes = DATA.mesures.filter((m) => m.Entreprise === ent && m["Statut_qualité"] !== "Doublon (exclu)");

  // sites de l'entreprise (émissions + mesures)
  const sites = uniq([...inv.map((e) => e.Site), ...ch.map((e) => e.Site), ...mes.map((m) => m.Site)]);
  if (emState.site !== ALL && !sites.includes(emState.site)) emState.site = ALL;
  fillSelect($("#e-site"), sites, emState.site, `Tous les sites (${sites.length})`);
  const bySite = (r) => emState.site === ALL || r.Site === emState.site;
  const invS = inv.filter(bySite), chS = ch.filter(bySite), mesS = mes.filter(bySite);

  // polluants émis
  const src = invS.length ? invS : chS;
  const pols = uniq(src.map((e) => e.Polluant));
  const psel = $("#e-polluant");
  psel.disabled = !pols.length;
  if (!pols.includes(emState.polluant)) emState.polluant = pols.find((p) => /^COV/.test(p)) || pols.find((p) => p === "NOx") || pols[0] || null;
  psel.innerHTML = pols.length ? pols.map((p) => `<option${p === emState.polluant ? " selected" : ""}>${esc(p)}</option>`).join("") : "<option>—</option>";

  // indicateurs
  const valid = mesS.filter((m) => m["Statut_qualité"] === "Valide");
  const comp = valid.filter(isCompared).length, dep = valid.filter((m) => m["Statut_conformité"] === "Dépassement").length;
  const k = [
    [sites.length, "site(s)"],
    [nf.format(mesS.length), "mesures de concentration"],
    [nf.format(dep), "dépassements de valeur guide", "bad"],
    [comp ? pf.format(dep / comp) : "–", "taux de dépassement", "bad"],
    [invS.length ? "Inventaire" : chS.length ? "Cheminée" : "Aucune", "données d'émission"],
  ];
  $("#e-kpis").innerHTML = k.map(([v, l, c]) => `<div class="kpi ${c || ""}"><div class="v">${v}</div><div class="l">${l}</div></div>`).join("");

  // graphique des émissions
  const wrap = $("#e-chart-wrap"), empty = $("#e-empty");
  if (charts["c-emissions"]) { charts["c-emissions"].destroy(); delete charts["c-emissions"]; }
  if (invS.length) {
    $("#e-title").textContent = `Émissions annuelles déclarées – ${ent} (t/an)`;
    $("#e-hint").textContent = "Inventaires d'émissions issus du reporting de l'entreprise : ce ne sont pas des concentrations, elles ne se comparent pas aux valeurs guides OMS.";
    const rows = invS.filter((e) => e.Polluant === emState.polluant);
    const years = uniq(rows.map((r) => r["Année"])).sort((a, b) => a - b);
    const ss = uniq(rows.map((r) => r.Site));
    wrap.hidden = false; empty.hidden = true;
    makeChart("c-emissions", {
      type: "line",
      data: { labels: years, datasets: ss.map((s, i) => ({
        label: s, data: years.map((y) => rows.find((r) => r.Site === s && r["Année"] === y)?.["Émission_t_an"] ?? null),
        borderColor: PAL[i % PAL.length], backgroundColor: PAL[i % PAL.length], tension: 0, spanGaps: true, pointRadius: 3,
      })) },
      options: { maintainAspectRatio: false, scales: { y: { beginAtZero: true, title: { display: true, text: `${emState.polluant} (t/an)` } } },
        plugins: { tooltip: { callbacks: { label: (c) => `${c.dataset.label} : ${fmt(c.raw)} t` } } } },
    });
  } else if (chS.length) {
    $("#e-title").textContent = `Rejets à la cheminée – ${ent} (mg/Nm³)`;
    $("#e-hint").textContent = "Concentrations à l'émission comparées aux valeurs limites d'émission SFI/IFC (directives EHS cimenterie, 2007).";
    const dr = chS.filter((c) => c.Polluant === emState.polluant).sort((a, b) => a["Année"] - b["Année"]);
    const vle = dr.find((r) => r.VLE_IFC_mg_Nm3 != null)?.VLE_IFC_mg_Nm3;
    wrap.hidden = false; empty.hidden = true;
    makeChart("c-emissions", {
      type: "bar",
      data: { labels: dr.map((r) => r["Année"]), datasets: [{ label: emState.polluant + " (mg/Nm³)", data: dr.map((r) => r.Valeur_mg_Nm3),
        backgroundColor: dr.map((r) => (vle != null && r.Valeur_mg_Nm3 > vle ? css("--bad") : css("--ok"))), borderRadius: 4 }] },
      options: { maintainAspectRatio: false, scales: { y: { beginAtZero: true, suggestedMax: vle ? vle * 1.2 : undefined, title: { display: true, text: "mg/Nm³" } } },
        plugins: { legend: { display: false }, guideLine: vle != null ? { value: vle, color: css("--text"), label: `Limite IFC ${fmt(vle)} mg/Nm³` } : null,
          tooltip: { callbacks: { afterLabel: (c) => dr[c.dataIndex].Statut || "" } } } },
      plugins: [guideLinePlugin],
    });
  } else {
    $("#e-title").textContent = `Émissions déclarées – ${ent}`;
    $("#e-hint").textContent = "";
    wrap.hidden = true; empty.hidden = false;
    empty.textContent = "Aucune donnée d'émission (inventaire annuel ou mesure à la cheminée) n'a été transmise pour cette entreprise. Seules des mesures de concentration dans l'air sont disponibles : voir ci-dessous.";
  }

  // concentrations mesurées : ratio max à la valeur guide par polluant
  $("#e-conc-title").textContent = `Concentrations mesurées autour des installations – ${ent}${emState.site !== ALL ? " – " + emState.site : ""}`;
  const g = new Map();
  for (const m of valid) {
    if (!isCompared(m) || typeof m.Ratio_valeur_guide !== "number") continue;
    const o = g.get(m.Polluant) || { r: 0, n: 0, dep: 0, max: null, guide: m["Valeur_guide_µg_m3"], per: m["Période_valeur_guide"] };
    o.n++; if (m["Statut_conformité"] === "Dépassement") o.dep++;
    if (m.Ratio_valeur_guide > o.r) { o.r = m.Ratio_valeur_guide; o.max = m["Valeur_µg_m3"]; }
    g.set(m.Polluant, o);
  }
  const items = [...g.entries()].filter(([, o]) => o.r > 0).sort((a, b) => b[1].r - a[1].r);
  const cwrap = $("#e-conc-wrap"), cempty = $("#e-conc-empty");
  if (charts["c-conc"]) { charts["c-conc"].destroy(); delete charts["c-conc"]; }
  if (!items.length) {
    cwrap.hidden = true; cempty.hidden = false;
    cempty.textContent = "Aucune mesure de concentration valide comparable à une valeur guide pour cette sélection.";
  } else {
    cwrap.hidden = false; cempty.hidden = true;
    cwrap.style.height = Math.max(260, items.length * 30 + 90) + "px";
    makeChart("c-conc", {
      type: "bar",
      data: { labels: items.map((i) => i[0]), datasets: [{ label: "Maximum / valeur guide", data: items.map((i) => i[1].r),
        backgroundColor: items.map((i) => (i[1].r > 1 ? css("--bad") : css("--ok"))), borderRadius: 3 }] },
      options: { indexAxis: "y", maintainAspectRatio: false,
        scales: { x: { type: "logarithmic", min: Math.min(0.1, 10 ** Math.floor(Math.log10(Math.min(...items.map((i) => i[1].r))))), max: Math.max(10, 10 ** Math.ceil(Math.log10(Math.max(...items.map((i) => i[1].r))))), ticks: { callback: (v) => (Math.abs(Math.log10(v) % 1) < 1e-9 ? fmt(v) + " ×" : "") },
          title: { display: true, text: "rapport à la valeur guide (échelle logarithmique)" } } },
        plugins: { legend: { display: false }, guideLine: { value: 1, color: css("--text"), label: "1 × valeur guide" },
          tooltip: { callbacks: { label: (c) => { const o = items[c.dataIndex][1]; return `Max ${fmt(o.max)} µg/m³ = ${fmt(o.r)} × la valeur guide (${fmt(o.guide)} µg/m³, ${o.per})`; },
            afterLabel: (c) => { const o = items[c.dataIndex][1]; return `${o.dep} dépassement(s) sur ${o.n} mesure(s) comparée(s)`; } } } } },
      plugins: [guideLinePlugin],
    });
  }

  // tableau des sites / fiches
  const fiches = DATA.fiches.filter((f) => f.Entreprise === ent && (emState.site === ALL || f.Site === emState.site));
  const fk = ["Fiche_source", "Site", "Secteur", "Département", "Type_document", "Années", "Nb_mesures", "Nb_valides", "Statut_fiche"];
  $("#t-entreprise").innerHTML = `<thead><tr>${fk.map((x) => `<th>${esc(x.replace(/_/g, " "))}</th>`).join("")}</tr></thead><tbody>` +
    (fiches.length ? fiches.map((f) => `<tr>${fk.map((x) => x === "Statut_fiche" ? `<td><span class="tag ${String(f[x]).startsWith("Retenue") ? "ok" : "warn"}">${esc(f[x])}</span></td>` : `<td class="${typeof f[x] === "number" ? "num" : ""}">${esc(f[x])}</td>`).join("")}</tr>`).join("")
      : `<tr><td colspan="${fk.length}">Aucune fiche pour cette sélection.</td></tr>`) + "</tbody>";
}

/* ---------- anomalies ---------- */
function renderAnomalies() {
  const types = uniq(DATA.anomalies.map((a) => a.Type));
  const sel = $("#a-type");
  if (!sel.options.length) {
    sel.innerHTML = `<option value="${ALL}">Tous les types (${DATA.anomalies.length})</option>` +
      types.map((t) => `<option>${esc(t)}</option>`).join("");
    sel.onchange = renderAnomalies;
  }
  const list = DATA.anomalies.filter((a) => sel.value === ALL || a.Type === sel.value);
  $("#t-anomalies").innerHTML = `<thead><tr><th>Fiche source</th><th>Type</th><th>Détail</th></tr></thead><tbody>` +
    list.map((a) => `<tr><td><b>${esc(a.Fiche_source)}</b></td><td><span class="tag ${a.Type === "Note de traitement" ? "neutral" : "warn"}">${esc(a.Type)}</span></td><td>${esc(a["Détail"])}</td></tr>`).join("") + "</tbody>";

  const v = verifRows();
  $("#verif-count").textContent = `${nf.format(v.length)} mesure(s) marquées « À vérifier » pour la sélection (filtres polluant, secteur, département et site). Elles sont exclues des comparaisons tant que le rapport source n'a pas été revu.`;
  $("#t-verif").innerHTML = `<thead><tr><th>ID</th><th>Fiche</th><th>Site</th><th>Station</th><th>Année</th><th>Polluant</th><th>Valeur source</th><th>Unité</th><th>µg/m³</th><th>Motif</th></tr></thead><tbody>` +
    v.map((m) => `<tr><td class="num">${m.ID}</td><td>${esc(m.Fiche_source)}</td><td>${esc(m.Site)}</td><td>${esc(m.Station)}</td><td class="num">${esc(m["Année"])}</td>
      <td>${esc(m.Polluant)}</td><td class="num">${esc(m.Valeur_source)}</td><td>${esc(m["Unité_source"])}</td><td class="num">${fmt(m["Valeur_µg_m3"])}</td><td>${esc(m.Commentaires)}</td></tr>`).join("") + "</tbody>";
}
const verifRows = () => filtered({ ignoreQualite: true }).filter((m) => m["Statut_qualité"] === "À vérifier");

/* ---------- rapport, méthode ---------- */
function renderStatic() {
  $("#gallery").innerHTML = FIGURES.map(([f, cap]) =>
    `<figure><a href="graphiques/${f}" target="_blank" rel="noopener"><img src="graphiques/${f}" alt="${esc(cap)}" loading="lazy"></a>
     <figcaption>${esc(cap)} · <a href="graphiques/${f}" download>télécharger</a></figcaption></figure>`).join("");
  $("#t-guides").innerHTML = `<thead><tr><th>Polluant</th><th>Valeur</th><th>Période</th><th>Référence</th></tr></thead><tbody>` +
    DATA.valeursGuides.map((g) => `<tr><td><b>${esc(g.Polluant)}</b></td><td class="num">${esc(typeof g["Valeur_µg_m3"] === "number" ? fmt(g["Valeur_µg_m3"]) + " µg/m³" : g["Valeur_µg_m3"])}</td><td>${esc(g["Période"])}</td><td>${esc(g["Référence"])}</td></tr>`).join("") + "</tbody>";
  const fk = ["Fiche_source", "Entreprise", "Site", "Secteur", "Département", "Type_document", "Années", "Nb_mesures", "Nb_valides", "Nb_à_vérifier", "Statut_fiche", "Lacunes"];
  $("#t-fiches").innerHTML = `<thead><tr>${fk.map((k) => `<th>${esc(k.replace(/_/g, " "))}</th>`).join("")}</tr></thead><tbody>` +
    DATA.fiches.map((f) => `<tr>${fk.map((k) => k === "Statut_fiche" ? `<td><span class="tag ${f[k] === "Retenue" ? "ok" : "warn"}">${esc(f[k])}</span></td>` : `<td class="${typeof f[k] === "number" ? "num" : ""}">${esc(f[k])}</td>`).join("")}</tr>`).join("") + "</tbody>";
  const d = new Date(DATA.meta.genere_le);
  $("#version").textContent = `${nf.format(DATA.meta.nb_mesures)} mesures, ${DATA.meta.nb_fiches} fiches`;
  $("#gen").textContent = ` Données générées le ${d.toLocaleDateString("fr-FR")} depuis ${DATA.meta.source}.`;
}

function renderAll() {
  renderDashboard();
  renderTable();
  renderAnomalies();
  if (state.tab === "carte") renderMap();
}

/* ---------- démarrage ---------- */
async function start() {
  try {
    const res = await fetch("data/donnees.json", { cache: "no-cache" });
    if (!res.ok) throw new Error("HTTP " + res.status);
    DATA = await res.json();
  } catch (e) {
    $("#loading").textContent = "Impossible de charger les données (" + e.message + ").";
    return;
  }
  chartDefaults();
  loadHash();
  refreshFilterOptions();
  bindFilters();
  bindTabs();
  $("#q").addEventListener("input", (e) => { state.q = e.target.value; state.page = 0; renderTable(); });
  $("#exp-csv").onclick = () => exportCsv(tableRows(), `mesures_qualite_air_${stamp()}.csv`);
  $("#exp-xlsx").onclick = () => exportXlsx(tableRows(), `mesures_qualite_air_${stamp()}.xlsx`);
  $("#exp-verif").onclick = () => exportXlsx(verifRows(), `mesures_a_verifier_${stamp()}.xlsx`, "À vérifier");
  renderStatic();
  renderAll();
  initEmissions();
  renderEmissions();
  showTab(state.tab);
  $("#loading").classList.add("hidden");
  window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", () => { chartDefaults(); renderAll(); renderEmissions(); });
}
start();
