// Copie les bibliothèques front (Chart.js, Leaflet, ExcelJS) dans public/vendor,
// pour que le site fonctionne sans CDN (utile hors connexion en atelier).
import { cp, mkdir } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const nm = (p) => path.join(root, "node_modules", p);
const out = path.join(root, "public", "vendor");
await mkdir(out, { recursive: true });
await cp(nm("chart.js/dist/chart.umd.js"), path.join(out, "chart.umd.js"));
await cp(nm("leaflet/dist/leaflet.js"), path.join(out, "leaflet/leaflet.js"));
await cp(nm("leaflet/dist/leaflet.css"), path.join(out, "leaflet/leaflet.css"));
await cp(nm("leaflet/dist/images"), path.join(out, "leaflet/images"), { recursive: true });
await cp(nm("exceljs/dist/exceljs.min.js"), path.join(out, "exceljs.min.js"));
console.log("OK : bibliothèques copiées dans public/vendor");
