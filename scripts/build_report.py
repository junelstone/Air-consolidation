#!/usr/bin/env python3
"""Génère le rapport de synthèse PDF (public/rapport/Rapport_synthese_qualite_air_DGE_OMS.pdf).

Usage : pip install reportlab openpyxl  puis  python3 scripts/build_report.py

Les tableaux chiffrés (synthèse par polluant, fiches, anomalies, valeurs guides)
sont relus dans l'Excel de data/. Le texte des constats reprend le projet de plan
validé en atelier : à mettre à jour ici si les données changent fortement.
Les figures sont celles de public/graphiques/.
"""
from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

import openpyxl
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate, Frame, Image, KeepTogether, ListFlowable, ListItem, NextPageTemplate,
    PageBreak, PageTemplate, Paragraph, Spacer, Table, TableStyle,
)
from reportlab.platypus.tableofcontents import TableOfContents

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / "data" / "Base_consolidee_qualite_air_DGE_OMS.xlsx"
FIG = ROOT / "public" / "graphiques"
OUT = ROOT / "public" / "rapport" / "Rapport_synthese_qualite_air_DGE_OMS.pdf"

# ---------- polices (accents, µ, ≤, ³) ----------
FONT_DIRS = [Path("/usr/share/fonts/truetype/dejavu"), Path("/Library/Fonts"), Path("C:/Windows/Fonts")]
def _font(name):
    for d in FONT_DIRS:
        if (d / name).exists():
            return str(d / name)
    return None
reg, bold = _font("DejaVuSans.ttf"), _font("DejaVuSans-Bold.ttf")
if reg and bold:
    pdfmetrics.registerFont(TTFont("Body", reg))
    pdfmetrics.registerFont(TTFont("Body-Bold", bold))
    pdfmetrics.registerFontFamily("Body", normal="Body", bold="Body-Bold", italic="Body", boldItalic="Body-Bold")
    F, FB = "Body", "Body-Bold"
else:
    F, FB = "Helvetica", "Helvetica-Bold"

GREEN = colors.HexColor("#0b6e4f")
LIGHT = colors.HexColor("#eef2f0")
GREY = colors.HexColor("#5b6b64")
RED = colors.HexColor("#c0392b")

ss = getSampleStyleSheet()
S = {
    "body": ParagraphStyle("body", parent=ss["Normal"], fontName=F, fontSize=9.6, leading=13.6, spaceAfter=6),
    "small": ParagraphStyle("small", parent=ss["Normal"], fontName=F, fontSize=7.6, leading=9.6),
    "smallb": ParagraphStyle("smallb", parent=ss["Normal"], fontName=FB, fontSize=7.6, leading=9.6, textColor=colors.white),
    "cap": ParagraphStyle("cap", parent=ss["Normal"], fontName=F, fontSize=8.2, leading=10.5, textColor=GREY, alignment=TA_CENTER, spaceBefore=3, spaceAfter=12),
    "h1": ParagraphStyle("h1", parent=ss["Heading1"], fontName=FB, fontSize=16, leading=20, textColor=GREEN, spaceBefore=4, spaceAfter=10),
    "h2": ParagraphStyle("h2", parent=ss["Heading2"], fontName=FB, fontSize=12, leading=15, textColor=GREEN, spaceBefore=10, spaceAfter=6),
    "title": ParagraphStyle("title", parent=ss["Title"], fontName=FB, fontSize=24, leading=30, textColor=GREEN, alignment=0),
    "subtitle": ParagraphStyle("subtitle", parent=ss["Normal"], fontName=F, fontSize=13, leading=18, textColor=GREY),
    "toc1": ParagraphStyle("toc1", fontName=F, fontSize=10.5, leading=16, leftIndent=0),
    "toc2": ParagraphStyle("toc2", fontName=F, fontSize=9, leading=13, leftIndent=16, textColor=GREY),
    "note": ParagraphStyle("note", parent=ss["Normal"], fontName=F, fontSize=9, leading=12.5, textColor=GREY, backColor=LIGHT, borderPadding=6, spaceBefore=4, spaceAfter=10),
}


class Doc(BaseDocTemplate):
    def __init__(self, path, **kw):
        super().__init__(str(path), pagesize=A4, leftMargin=2 * cm, rightMargin=2 * cm, topMargin=2 * cm, bottomMargin=2 * cm,
                         title="Rapport de synthèse – Qualité de l'air (mission DGE-OMS)",
                         author="Direction Générale de l'Environnement – OMS Congo", subject="Qualité de l'air, République du Congo", **kw)
        frame = Frame(self.leftMargin, self.bottomMargin, self.width, self.height, id="f")
        self.addPageTemplates([PageTemplate("cover", [frame], onPage=self._cover), PageTemplate("normal", [frame], onPage=self._page)])

    def _cover(self, c, d):
        c.saveState()
        c.setFillColor(GREEN)
        c.rect(0, A4[1] - 1.2 * cm, A4[0], 1.2 * cm, stroke=0, fill=1)
        c.rect(0, 0, A4[0], 0.6 * cm, stroke=0, fill=1)
        c.restoreState()

    def _page(self, c, d):
        c.saveState()
        c.setFont(F, 7.5)
        c.setFillColor(GREY)
        c.drawString(2 * cm, 1.2 * cm, "Rapport de synthèse – Qualité de l'air – Mission DGE / OMS Congo – Projet à valider en atelier")
        c.drawRightString(A4[0] - 2 * cm, 1.2 * cm, f"Page {d.page}")
        c.setStrokeColor(GREEN)
        c.setLineWidth(0.6)
        c.line(2 * cm, A4[1] - 1.5 * cm, A4[0] - 2 * cm, A4[1] - 1.5 * cm)
        c.restoreState()

    def afterFlowable(self, f):
        if isinstance(f, Paragraph) and f.style.name in ("h1", "h2"):
            level = 0 if f.style.name == "h1" else 1
            text = f.getPlainText()
            key = f"h{id(f)}"
            self.canv.bookmarkPage(key)
            self.canv.addOutlineEntry(text, key, level=level, closed=level > 0)
            self.notify("TOCEntry", (level, text, self.page, key))


# ---------- helpers ----------
def P(t, st="body"):
    return Paragraph(t, S[st])

def bullets(items, numbered=False):
    return ListFlowable([ListItem(P(i), leftIndent=12) for i in items], bulletType="1" if numbered else "bullet",
                        start="1" if numbered else None, bulletFontName=F, bulletFontSize=9, leftIndent=14, spaceAfter=6)

def fig(num, caption, max_h=17 * cm):
    path = next(FIG.glob(f"{num:02d}_*.png"))
    from PIL import Image as PI
    w, h = PI.open(path).size
    width = 17 * cm
    height = width * h / w
    if height > max_h:
        height = max_h
        width = height * w / h
    return KeepTogether([Image(str(path), width=width, height=height), P(f"<b>Figure {num:02d}</b> – {caption}", "cap")])

def table(rows, widths, header=True, zebra=True, font_size=7.6):
    data = []
    for i, r in enumerate(rows):
        st = S["smallb"] if header and i == 0 else S["small"]
        data.append([Paragraph(str(c) if c is not None else "", st) for c in r])
    t = Table(data, colWidths=widths, repeatRows=1 if header else 0)
    style = [
        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#d9e1dd")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    if header:
        style.append(("BACKGROUND", (0, 0), (-1, 0), GREEN))
    if zebra:
        for i in range(1 if header else 0, len(rows)):
            if i % 2 == 0:
                style.append(("BACKGROUND", (0, i), (-1, i), LIGHT))
    t.setStyle(TableStyle(style))
    return t

def fr(x, d=2):
    if x is None or x == "":
        return "–"
    if isinstance(x, (int, float)):
        if isinstance(x, float) and abs(x) < 1 and x != 0:
            s = f"{x:.4g}"
        else:
            s = f"{x:,.{d}f}".rstrip("0").rstrip(".") if isinstance(x, float) else f"{x:,}"
        return s.replace(",", " ").replace(".", ",")
    return str(x)

def pct(x):
    return "–" if x is None else f"{round(x * 100):d} %"


# ---------- données ----------
def load():
    wb = openpyxl.load_workbook(XLSX, read_only=True, data_only=True)
    def rows(name, ncol=None):
        it = wb[name].iter_rows(values_only=True)
        head = next(it)
        if ncol:
            head = head[:ncol]
        out = []
        for r in it:
            r = r[: len(head)]
            if any(v not in (None, "") for v in r):
                out.append(dict(zip(head, r)))
        return out
    return {
        "synth": rows("Synthèse_polluants"),
        "fiches": rows("Fiches"),
        "anom": rows("Anomalies"),
        "guides": [list(r[:5]) for r in wb["Valeurs_guides"].iter_rows(min_row=2, values_only=True) if r[0]],
        "base": rows("Base_consolidée"),
    }


def build():
    D = load()
    base = D["base"]
    n = len(base)
    valid = sum(1 for r in base if r["Statut_qualité"] == "Valide")
    verif = sum(1 for r in base if r["Statut_qualité"] == "À vérifier")
    hors = [r for r in base if r["Statut_qualité"] != "Doublon (exclu)"]
    nsites = len({r["Site"] for r in hors})
    nent = len({r["Entreprise"] for r in hors})
    npol = len({r["Polluant"] for r in base})
    nc = sum(1 for r in base if r["Statut_conformité"] == "Non concluant (LQ > valeur guide)")

    doc = Doc(OUT)
    st = []

    # ---- couverture ----
    st += [Spacer(1, 3.2 * cm),
           P("RÉPUBLIQUE DU CONGO", "subtitle"),
           P("Direction Générale de l'Environnement (DGE) – Organisation mondiale de la Santé (OMS Congo)", "subtitle"),
           Spacer(1, 1.6 * cm),
           P("Rapport technique de synthèse", "title"),
           P("Qualité de l'air issue des évaluations environnementales et sociales", "title"),
           Spacer(1, 0.6 * cm),
           P("Analyse situationnelle des données de la mission de collecte documentaire : départements de Pointe-Noire, "
             "du Kouilou, du Niari, de la Bouenza, de Brazzaville et de la Cuvette.", "subtitle"),
           Spacer(1, 2.2 * cm),
           table([["Atelier", "Du 17 au 19 novembre 2026, OMS Congo"],
                  ["Version", f"Projet v1 du {dt.date(2026, 11, 17).strftime('%d/%m/%Y')}, à valider en atelier"],
                  ["Données", f"{fr(n)} mesures issues de {len(D['fiches'])} fiches de collecte ; {fr(valid)} mesures valides"],
                  ["Livrables associés", "Base consolidée (Excel), application web interactive, fichiers cartographiques KML / GeoJSON"]],
                 [4 * cm, 13 * cm], header=False, zebra=False),
           NextPageTemplate("normal"), PageBreak()]

    # ---- sommaire ----
    toc = TableOfContents()
    toc.levelStyles = [S["toc1"], S["toc2"]]
    st += [Paragraph("Sommaire", ParagraphStyle("tocT", parent=S["h1"])), toc, PageBreak()]

    # ---- 1. Résumé exécutif ----
    st += [P("1. Résumé exécutif", "h1"),
           P(f"La mission de collecte documentaire DGE / OMS a permis de rassembler {len(D['fiches'])} fiches et "
             f"{fr(n)} mesures de concentration de polluants dans l'air, provenant de {nsites} sites et {nent} entreprises "
             f"entre 2016 et 2025. Après dédoublonnage, harmonisation des unités et contrôle de plausibilité, "
             f"{fr(valid)} mesures sont valides et {verif} restent à vérifier sur les rapports sources."),
           P("<b>Le benzène et les particules PM10 sont les deux problèmes principaux.</b> Le benzène dépasse la valeur "
             "guide OMS (1,7 µg/m³) dans 58 % des mesures valides, les PM10 dépassent 45 µg/m³ sur 24 h dans 43 % des cas. "
             "Le cadmium (36 %), le nickel (31 %) et les PM2.5 (23 %) suivent. Les gaz de combustion (NO2, SO2, CO, O3) "
             "dépassent rarement, sauf dans la zone pétrolière de Mboundi."),
           P("<b>Zones prioritaires :</b> dépôts pétroliers SCLOG (Mpila et Port autonome), zone ENI de Mboundi, "
             "Litchendjili et OPT (Kouilou), fonderie SIDECO de Vindoulou et plateformes offshore (air intérieur)."),
           P("<b>Cinq recommandations clés :</b>"),
           bullets(["Adopter des valeurs limites nationales de qualité de l'air ambiant alignées sur l'OMS 2021.",
                    "Imposer un cahier des charges type pour le volet air des EIES, AES et rapports de surveillance.",
                    "Centraliser les résultats dans un format numérique unique (gabarit de la base consolidée).",
                    "Inspecter et contrôler en priorité les points chauds identifiés.",
                    "Installer avec l'OMS des stations de fond à Pointe-Noire et Brazzaville."], numbered=True),
           P("Ces résultats sont indicatifs : la plupart des prélèvements durent de 1 à 14 jours, alors que plusieurs "
             "valeurs guides sont annuelles.", "note"),
           fig(1, "Taux de dépassement des valeurs guides par polluant (mesures valides)", 11 * cm),
           PageBreak()]

    # ---- 2. Contexte ----
    st += [P("2. Contexte et objectifs", "h1"),
           P("Les promoteurs de projets publics et privés produisent régulièrement des documents techniques intégrant des "
             "analyses de la qualité de l'air : études d'impact environnemental et social (EIES), audits environnementaux "
             "et sociaux (AES) et rapports périodiques de surveillance. Ces informations restent dispersées dans les dossiers "
             "individuels et largement sous-exploitées, faute d'un mécanisme d'harmonisation et d'analyse transversale."),
           P("Conformément à ses missions de prévention des pollutions, d'inspection des installations classées et de suivi "
             "de l'état de l'environnement, la Direction Générale de l'Environnement a conduit, avec l'OMS, une mission de "
             "collecte documentaire auprès des entreprises de Pointe-Noire, du Kouilou, du Niari, de la Bouenza et de Brazzaville. "
             "L'objectif est de transformer ces « données dormantes » en outil de pilotage, en vue d'un plan d'action national "
             "d'amélioration de la qualité de l'air."),
           P("L'atelier du 17 au 19 novembre 2026 vise à :"),
           bullets(["traiter et valider la base de données (doublons, harmonisation des unités) ;",
                    "identifier les polluants à analyser par secteur d'activité ;",
                    "analyser les résultats par comparaison avec les normes de qualité de l'air ;",
                    "cartographier les zones à forte pollution et les secteurs d'activité ;",
                    "formuler des recommandations à l'administration et aux promoteurs ;",
                    "présenter les grandes lignes du rapport de la mission."]),
           P("Trois livrables sont attendus : une base de données consolidée, une feuille de route opérationnelle et le "
             "présent rapport technique de synthèse.")]

    # ---- 3. Méthodologie ----
    st += [P("3. Méthodologie", "h1"),
           P("3.1 Mission de collecte documentaire", "h2"),
           P("Les données proviennent de rapports d'EIES, d'AES, de rapports d'analyse et de surveillance transmis par les "
             "entreprises des secteurs pétrolier (amont et aval), gazier, minier, cimentier, métallurgique, agroalimentaire et "
             "du traitement des déchets. Chaque document a été saisi dans une fiche de collecte normalisée (une fiche par site "
             "ou campagne)."),
           P("3.2 Fiche de collecte et base de données", "h2"),
           P("Les 54 onglets du classeur de collecte ont été consolidés en une base unique : une ligne par mesure "
             "(polluant × station × date), avec la valeur et l'unité d'origine, la valeur convertie, la valeur guide de "
             "référence, le statut de conformité et le statut qualité. La liste des fiches figure en annexe."),
           P("3.3 Traitement des données", "h2"),
           bullets(["<b>Harmonisation des unités</b> : conversion en µg/m³ (mg/m³ × 1000 ; ppm × M/24,45 × 1000 à 25 °C et 1 atm).",
                    "<b>Valeurs inférieures à la limite de quantification</b> : conservées avec la LQ. Si la LQ est inférieure ou égale à la "
                    "valeur guide, la mesure est conforme ; sinon elle est « non concluante ».",
                    "<b>Doublons</b> exclus : LGN 2 (= LGN 1), Luyane Mpili (= Luyan Surveillance), Feuil40 (= Saris Raffinerie), "
                    "Good Luck Mining (= GPL Louvakou, doublon suspect). Les fiches ENI Mboundi J1 à J5 sont des jours distincts et sont conservées.",
                    "<b>Contrôle de plausibilité</b> : valeurs physiquement improbables, unités incohérentes ou copies suspectes marquées "
                    "« À vérifier » et exclues des comparaisons.",
                    "<b>Air intérieur et air de site</b> (plateformes offshore, dépôts SCLOG, usine Dangote) : comparés aux valeurs d'air ambiant à titre indicatif."]),
           P("3.4 Référentiels de comparaison", "h2"),
           P("Faute de texte congolais fixant des valeurs limites d'air ambiant (à confirmer par la DGE), les mesures sont "
             "comparées aux lignes directrices OMS 2021 et, pour les métaux, le benzène, le toluène et l'H2S, aux valeurs "
             "OMS 2000 (Europe). La valeur de court terme (24 h, 8 h ou hebdomadaire) est utilisée quand elle existe. "
             "Les rejets à la cheminée sont comparés aux directives EHS de la SFI / IFC."),
           table([["Polluant", "Valeur", "Période", "Référence"]] +
                 [[g[0], (fr(g[1]) + " µg/m³") if isinstance(g[1], (int, float)) else (g[4] or "–"), g[2], g[3]] for g in D["guides"]],
                 [4.2 * cm, 2.6 * cm, 2.6 * cm, 7.6 * cm]),
           PageBreak()]

    # ---- 4. Description des données ----
    by_dep = {}
    for r in hors:
        by_dep.setdefault(r["Département"], set()).add(r["Site"])
    dep_txt = ", ".join(f"{k} {len(v)}" for k, v in sorted(by_dep.items(), key=lambda kv: -len(kv[1])))
    st += [P("4. Description des données", "h1"),
           P(f"La base retient {fr(len(hors))} mesures hors doublons, dont {fr(valid)} valides, issues de {nsites} sites "
             f"et {nent} entreprises sur la période 2016–2025."),
           table([["Indicateur", "Valeur"],
                  ["Fiches de collecte", f"{len(D['fiches'])} (dont 4 doublons, 1 vide, 6 fiches d'émissions)"],
                  ["Mesures de concentration extraites", fr(n)],
                  ["Mesures valides", fr(valid)],
                  ["Mesures à vérifier sur le rapport source", fr(verif)],
                  ["Polluants distincts", str(npol)],
                  ["Sites par département", dep_txt],
                  ["Secteurs", "Pétrole et gaz amont 69 % des mesures, cimenterie 8 %, gaz 6 %, pétrole aval 5 %, mines 4 %, "
                               "agroalimentaire 4 %, déchets 3 %, fonderie 1 %"],
                  ["Stations géoréférencées", "50"]],
                 [6 * cm, 11 * cm]),
           Spacer(1, 6),
           P("Le secteur pétrolier amont domine l'échantillon : les conclusions sur les autres secteurs reposent sur peu de mesures."),
           fig(3, "Répartition des mesures par secteur d'activité et par statut de conformité", 7 * cm),
           PageBreak()]

    # ---- 5. Résultats par famille ----
    synth = sorted([s for s in D["synth"] if s.get("Nb_comparées")], key=lambda s: -(s.get("%_dépassement") or 0))
    st += [P("5. Résultats par famille de polluants", "h1"),
           P("Le tableau ci-dessous résume, pour chaque polluant comparable à une valeur guide, les mesures valides, "
             "les concentrations observées et le nombre de dépassements."),
           table([["Polluant", "Mesures", "Sites", "Médiane", "Max", "Valeur guide", "Comparées", "Dépass.", "Taux"]] +
                 [[s["Polluant"], s["Nb_mesures_valides"], s["Nb_sites"], fr(s["Médiane_µg_m3"]), fr(s["Max_µg_m3"]),
                   f"{fr(s['Valeur_guide_µg_m3'])} ({s['Période']})" if s.get("Valeur_guide_µg_m3") is not None else "–",
                   s["Nb_comparées"], s["Nb_dépassements"], pct(s.get("%_dépassement"))] for s in synth],
                 [3.3 * cm, 1.4 * cm, 1.1 * cm, 1.6 * cm, 1.6 * cm, 3.0 * cm, 1.6 * cm, 1.4 * cm, 1.4 * cm]),
           P("Concentrations en µg/m³. Taux = dépassements / mesures comparées.", "cap"),
           P("5.1 Particules (PM10, PM2.5)", "h2"),
           P("Les PM10 dépassent la valeur guide de 45 µg/m³ (24 h) dans 44 des 102 mesures valides (43 %). Le maximum, "
             "122,5 µg/m³, est relevé au centre emplisseur GPL de Louvakou. Dans la zone de Mboundi, 26 des 42 mesures "
             "journalières de la campagne 2023 dépassent 45 µg/m³ ; la fonderie SIDECO de Vindoulou dépasse 7 jours sur 8, "
             "en zone résidentielle. Les PM2.5 dépassent 15 µg/m³ dans 23 % des cas, surtout en air intérieur sur les "
             "plateformes offshore (jusqu'à 108,9 µg/m³ à Moho Nord)."),
           fig(4, "Concentrations de PM10 par site et valeur guide OMS 24 h", 11 * cm),
           fig(5, "Concentrations de PM2.5 par site et valeur guide OMS 24 h", 7 * cm),
           P("5.2 Gaz inorganiques (NO2, SO2, O3, CO, H2S)", "h2"),
           P("Les gaz de combustion dépassent rarement : SO2 6 fois sur 139 (maximum 270 µg/m³ à Mboundi), NO2 4 fois sur 140 "
             "(310 µg/m³ à Mboundi), CO 3 fois sur 125 et O3 2 fois sur 103 (290 µg/m³ à Litchendjili). La zone de Mboundi, "
             "Litchendjili et OPT est le seul secteur où plusieurs gaz dépassent ensemble."),
           fig(7, "Concentrations de SO2 par site", 14 * cm),
           fig(8, "Concentrations de NO2 par site", 14 * cm),
           P("5.3 COV et BTEX", "h2"),
           P("Le benzène est le polluant le plus préoccupant : 71 dépassements sur 123 mesures comparées (58 %) par rapport à "
             "la valeur OMS de 1,7 µg/m³ (risque de cancer de 1 sur 100 000). Les valeurs les plus élevées sont mesurées "
             "dans l'enceinte des dépôts SCLOG (jusqu'à 4 100 µg/m³ à Mpila), sur des prélèvements de 8 h plus proches d'une "
             "exposition professionnelle. Le toluène dépasse 260 µg/m³ (semaine) dans 8 % des cas, avec un maximum de "
             "11 300 µg/m³ au même dépôt."),
           fig(6, "Concentrations de benzène par site (échelle logarithmique)", 15 * cm),
           P("5.4 Métaux sur particules", "h2"),
           P("Le cadmium (14 dépassements sur 39, 36 %) et le nickel (14 sur 45, 31 %) dépassent de façon récurrente sur les "
             "permis SNPC, la raffinerie CORAF et la base ENI ; les maxima (0,5 et 0,15 µg/m³) sont relevés au permis Kindji. "
             f"Pour l'arsenic, le cadmium et le nickel, de nombreuses limites de quantification sont supérieures aux valeurs guides "
             f"({nc} résultats non concluants au total). Les teneurs de zinc de 300 à 400 µg/m³ relevées dans plusieurs rapports SGS "
             "font douter de l'unité utilisée."),
           PageBreak()]

    # ---- 6. Secteurs et départements ----
    st += [P("6. Analyse par secteur d'activité et par département", "h1"),
           P("<b>Par département</b> : le Kouilou cumule le plus de dépassements (PM10 59 %, benzène 68 %, nickel 43 %). "
             "Pointe-Noire est surtout touché par le benzène (38 %) et les métaux. Les autres départements comptent peu de mesures."),
           fig(9, "Dépassements par département et par polluant", 10 * cm),
           P("<b>Par secteur</b> : le pétrole aval et l'amont concentrent les dépassements de benzène ; la fonderie et l'amont "
             "pétrolier concentrent les dépassements de PM10. La cimenterie Dangote reste sous les valeurs guides en PM10 et PM2.5 "
             "dans l'usine, mais ses NOx à la cheminée ont dépassé la limite IFC de 600 mg/Nm³ en 2022 (685 mg/Nm³)."),
           fig(10, "Dépassements par secteur d'activité et par polluant", 10 * cm),
           P("Les polluants à suivre en priorité par secteur, proposés à partir de ces résultats, figurent au chapitre 10.")]

    # ---- 7. Cartographie ----
    st += [P("7. Cartographie des zones à forte pollution", "h1"),
           P("Les 50 stations géoréférencées sont fournies aux formats KML (Google Earth) et GeoJSON (QGIS), et consultables "
             "sur la carte interactive de l'application web. Les points chauds sont les suivants :"),
           bullets(["<b>Dépôts pétroliers SCLOG</b> (Mpila et Port autonome) : benzène jusqu'à 2 400 fois la valeur guide et toluène "
                    "jusqu'à 43 fois, sur des prélèvements de 8 h dans l'enceinte.",
                    "<b>Zone de Mboundi, Litchendjili et OPT</b> (ENI, Kouilou) : seul secteur où plusieurs gaz dépassent ensemble "
                    "(benzène, SO2, NO2, O3, CO, H2S) ; PM10 au-dessus de 45 µg/m³ dans 26 des 42 mesures journalières de 2023.",
                    "<b>Fonderie SIDECO</b> (Vindoulou) : PM10 au-dessus de 45 µg/m³ 7 jours sur 8, en zone résidentielle.",
                    "<b>Plateformes offshore TotalEnergies</b> : PM2.5 en air intérieur jusqu'à 7 fois la valeur guide (fumoir, cuisines, cabines).",
                    "<b>Métaux (cadmium, nickel)</b> : dépassements récurrents sur les permis SNPC, la raffinerie CORAF et la base ENI."]),
           fig(2, "Points chauds : rapport à la valeur guide par site et par polluant", 15 * cm),
           fig(13, "Carte des stations de mesure dans le sud du pays", 10 * cm),
           PageBreak()]

    # ---- 8. Émissions ----
    st += [P("8. Émissions déclarées", "h1"),
           P("Les inventaires de TotalEnergies sont des émissions annuelles (t/an) et non des concentrations : ils ne se comparent "
             "pas aux valeurs guides OMS. Le terminal de Djéno émet de 2 400 à 6 400 t/an de COV. Les NOx de Moho Nord ont baissé "
             "de 4 060 t en 2016 à 830 t en 2025. Le SO2 de PNGF Nord Yanga est remonté à 553 t en 2025."),
           fig(11, "Émissions annuelles déclarées par TotalEnergies E&amp;P Congo", 11 * cm),
           P("Les mesures à la cheminée de la cimenterie Dangote (mg/Nm³) sont comparées aux valeurs limites d'émission de la SFI / IFC "
             "pour les cimenteries. Les NOx ont dépassé 600 mg/Nm³ en 2022 ; le SO2, déclaré à 0 toutes les années, doit être "
             "réellement mesuré."),
           fig(12, "Rejets à la cheminée de Dangote Cement comparés aux limites SFI / IFC", 7 * cm),
           PageBreak()]

    # ---- 9. Limites ----
    st += [P("9. Limites de l'étude et qualité des données", "h1"),
           P("Les résultats sont indicatifs : la plupart des mesures sont ponctuelles (1 à 14 jours), alors que plusieurs valeurs "
             "guides sont annuelles."),
           bullets(["<b>Durées non comparables</b> : tubes passifs de 3 à 18 jours, pompes 8 h, analyseurs en continu. Une valeur sur "
                    "7 jours comparée à une valeur guide sur 24 h ou annuelle ne prouve pas un dépassement réglementaire.",
                    f"<b>Unités incohérentes</b> : NO2 et SO2 en ppm à des niveaux invraisemblables (AOGC, SLE), métaux en mg/m³ chez "
                    f"Saris et Green Services, H2S en mg/m³ chez ENI. {verif} valeurs sont marquées « à vérifier ».",
                    "<b>Copies entre fiches</b> : 4 doublons exclus, et des valeurs de Luyuan identiques à celles de SNPC Mayombe.",
                    f"<b>Limites de quantification trop hautes</b> : {nc} résultats « &lt; LQ » ne permettent pas de conclure, "
                    "surtout pour l'arsenic, le cadmium et le nickel.",
                    "<b>Métadonnées manquantes</b> : plus de 25 fiches sans coordonnées GPS, année absente sur 3 fiches, année par défaut "
                    "du modèle recopiée sur d'autres.",
                    "<b>Couverture déséquilibrée</b> : 69 % des mesures viennent du pétrole amont, une seule fiche pour Brazzaville et "
                    "pour la Cuvette, aucune donnée de fond urbain.",
                    "<b>Air intérieur et enceintes industrielles</b> : les mesures offshore et SCLOG ne décrivent pas l'exposition de la population."]),
           ]

    # ---- 10. Recommandations ----
    st += [P("10. Recommandations", "h1"),
           P("Cinq priorités ressortent : fixer des normes nationales, standardiser les mesures, et agir d'abord sur le benzène "
             "des dépôts et la zone de Mboundi."),
           P("10.1 À l'administration (DGE)", "h2"),
           bullets(["Adopter par décret des valeurs limites nationales de qualité de l'air ambiant, alignées sur les valeurs guides "
                    "OMS 2021 avec des cibles intermédiaires.",
                    "Imposer un cahier des charges type pour le volet air des EIES, AES et rapports de surveillance : polluants minimaux "
                    "par secteur, durée de prélèvement, unités en µg/m³, LQ inférieures aux valeurs guides, coordonnées GPS et dates obligatoires.",
                    "Exiger la transmission des résultats dans un format numérique unique (gabarit Excel de la base consolidée) et tenir à jour cette base.",
                    "Diligenter une inspection et une campagne de contrôle indépendante sur les points chauds : dépôts SCLOG, zone Mboundi, "
                    "Litchendjili et OPT, fonderie de Vindoulou.",
                    "Mettre en place, avec l'OMS, des stations de fond à Pointe-Noire et Brazzaville (PM2.5, PM10, NO2, O3) pour mesurer "
                    "l'exposition de la population."], numbered=True),
           P("10.2 Aux promoteurs", "h2"),
           bullets(["Dépôts et raffinage : récupération des vapeurs au chargement, contrôle des fuites, suivi du benzène en limite de "
                    "propriété et protection des travailleurs.",
                    "Amont pétrolier : réduction du torchage et des évents, surveillance continue en villages riverains (PM10, SO2, NO2, benzène, H2S).",
                    "Fonderie et carrières : abattement des poussières (capotage, arrosage, filtration), suivi des métaux sur PM10.",
                    "Plateformes offshore : améliorer la ventilation et appliquer l'interdiction de fumer hors zones dédiées.",
                    "Cimenteries : maintenir les NOx sous 600 mg/Nm³ à la cheminée et mesurer réellement le SO2 (déclaré à 0)."], numbered=True),
           P("10.3 Polluants à suivre en priorité par secteur (à valider en atelier)", "h2"),
           table([["Secteur", "Polluants à suivre en priorité"],
                  ["Pétrole et gaz amont", "PM10, PM2.5, NO2, SO2, CO, H2S, benzène et BTEX"],
                  ["Pétrole aval (dépôts, raffinerie)", "Benzène et BTEX, COV totaux, H2S, SO2"],
                  ["Gaz (GPL, GNL)", "COV, benzène, NO2, PM10"],
                  ["Cimenterie", "PM10, PM2.5, NOx, SO2, poussières à la cheminée"],
                  ["Fonderie et métallurgie", "PM10, PM2.5, plomb, cadmium, zinc, nickel"],
                  ["Mines et carrières", "PM10, PM2.5, métaux (arsenic, plomb, mercure, manganèse)"],
                  ["Traitement des déchets", "PM10, NO2, SO2, CO, COV, métaux (incinération)"],
                  ["Agroalimentaire", "PM10, NO2, SO2, CO (chaudières)"]],
                 [6 * cm, 11 * cm]),
           PageBreak()]

    # ---- 11. Feuille de route ----
    st += [P("11. Feuille de route opérationnelle", "h1"),
           P("La feuille de route (livrable 2) sera finalisée en atelier. Le tableau ci-dessous propose une base de discussion "
             "à partir des recommandations ; les échéances et responsables sont à arrêter par les participants.", "note"),
           table([["Action", "Responsable proposé", "Horizon"],
                  ["Vérifier les 232 valeurs « à vérifier » sur les rapports sources et mettre à jour la base", "DGE", "1 à 3 mois"],
                  ["Confirmer l'existence ou non de normes nationales ; choisir la référence benzène (OMS 1,7 ou UE 5 µg/m³)", "DGE, ministère", "3 mois"],
                  ["Diffuser le gabarit de saisie et le cahier des charges du volet air des EIES / AES", "DGE", "6 mois"],
                  ["Campagne de contrôle indépendante sur les points chauds (SCLOG, Mboundi, Vindoulou)", "DGE, OMS, laboratoire agréé", "6 à 12 mois"],
                  ["Installer des stations de fond à Pointe-Noire et Brazzaville", "DGE, OMS", "12 à 24 mois"],
                  ["Projet de décret fixant les valeurs limites nationales", "Ministère de l'Environnement", "12 à 24 mois"],
                  ["Mise à jour annuelle de la base et du tableau de bord", "DGE", "Annuel"]],
                 [9.5 * cm, 4.5 * cm, 3 * cm]),
           P("12. Annexes", "h1"),
           P("A. Questions à trancher en atelier", "h2"),
           bullets(["Existe-t-il un texte congolais fixant des valeurs limites d'air ambiant ? Aucun n'a été trouvé en ligne.",
                    "Good Luck Mining : les valeurs identiques à GPL Louvakou sont-elles une erreur de saisie ?",
                    "Luyuan des Mines : les résultats correspondent-ils au rapport de Luyuan ou à celui de SNPC Mayombe ?",
                    "Saris Raffinerie : quelle est l'année des mesures (2019 ou 2021) et l'unité réelle des métaux ?",
                    "AOGC et SLE : NO2 et SO2 sont-ils en ppm, en ppb ou en µg/m³ ?",
                    "Zinc mesuré à 300–400 µg/m³ (ENI, CORAF, SNPC Kindji) : µg/m³ ou ng/m³ ?",
                    "ENI Base Ya Sika et Agri Hub : O3 à 2 000 µg/m³ et benzène à 1 800 µg/m³, à vérifier sur le rapport.",
                    "Faut-il retenir la valeur OMS 2000 (1,7 µg/m³) ou la valeur UE (5 µg/m³) comme référence pour le benzène ?"]),
           PageBreak(),
           P("B. Liste des fiches de collecte", "h2"),
           table([["Fiche", "Entreprise", "Secteur", "Dépt.", "Années", "Mes.", "Val.", "Statut"]] +
                 [[f["Fiche_source"], f["Entreprise"], f["Secteur"], f["Département"], f["Années"], f["Nb_mesures"],
                   f["Nb_valides"], f["Statut_fiche"]] for f in D["fiches"]],
                 [3.6 * cm, 3.0 * cm, 3.0 * cm, 1.8 * cm, 1.6 * cm, 1.0 * cm, 1.0 * cm, 2.0 * cm]),
           PageBreak(),
           P("C. Journal des anomalies", "h2"),
           table([["Fiche", "Type", "Détail"]] + [[a["Fiche_source"], a["Type"], a["Détail"]] for a in D["anom"]],
                 [3.8 * cm, 2.8 * cm, 10.4 * cm]),
           P("D. Dictionnaire de la base", "h2"),
           P("Statut qualité : Valide | À vérifier | Doublon (exclu) | Non détecté / non renseigné. "
             "Statut de conformité : Conforme | Conforme (&lt; LQ ≤ valeur guide) | Dépassement | Non concluant (LQ &gt; valeur guide) | "
             "Non comparable | Non évalué. Le détail des colonnes figure dans l'onglet LISEZMOI de la base consolidée."),
           ]

    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc.multiBuild(st)
    print(f"OK : {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    sys.exit(build())
