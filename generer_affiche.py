#!/usr/bin/env python3
"""Génère l'affiche mensuelle des événements à partir d'un fichier Excel.

Usage :
    python generer_affiche.py evenements.xlsx [config.yaml] [-o sortie.pdf]

Colonnes attendues dans l'Excel : Date | Evenement | Debut | Fin | Prix
Plusieurs lignes avec la même date sont regroupées dans la même case.

Sorties : le PDF (reportlab, aucune dépendance native) et le SVG équivalent
(retouchable dans Inkscape). Dépendances : pyyaml, openpyxl, pillow,
reportlab, fonttools.
"""

import argparse
import base64
import os
import sys
from collections import defaultdict
from datetime import date, datetime, time
from pathlib import Path
from xml.sax.saxutils import escape

import yaml
from openpyxl import load_workbook
from PIL import Image, ImageFont

MM_PAR_PT = 25.4 / 72.0
PT_PAR_MM = 72.0 / 25.4
MESURE_ECHELLE = 4  # précision de mesure Pillow (px par pt)

JOURS = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
MOIS = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet",
        "Août", "Septembre", "Octobre", "Novembre", "Décembre"]


# ---------------------------------------------------------------- polices

def _dossiers_polices() -> list[Path]:
    if os.name == "nt":
        dossiers = [Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts",
                    Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft" / "Windows" / "Fonts"]
    else:
        home = Path.home()
        dossiers = [Path("/usr/share/fonts"), Path("/usr/local/share/fonts"),
                    home / ".fonts", home / ".local" / "share" / "fonts"]
    return [d for d in dossiers if d.is_dir()]


_INDEX_POLICES = None


def index_polices() -> dict:
    """Scanne les polices du système : {famille: {sous-famille: (chemin, nom, style)}}."""
    global _INDEX_POLICES
    if _INDEX_POLICES is not None:
        return _INDEX_POLICES
    from fontTools.ttLib import TTCollection, TTFont
    _INDEX_POLICES = {}
    for dossier in _dossiers_polices():
        for f in dossier.rglob("*"):
            if f.suffix.lower() not in (".ttf", ".otf", ".ttc"):
                continue
            try:
                fontes = (TTCollection(str(f), lazy=True).fonts
                          if f.suffix.lower() == ".ttc" else [TTFont(str(f), lazy=True)])
            except Exception:
                continue
            for ft in fontes:
                try:
                    noms = ft["name"]
                except Exception:
                    continue
                n1 = noms.getDebugName(1)          # famille "affichée"
                if not n1:
                    continue
                n2 = noms.getDebugName(2) or "Regular"
                n16 = noms.getDebugName(16) or n1  # famille typographique
                n17 = noms.getDebugName(17) or n2
                for fam, sub in {(n1, n2), (n16, n17)}:
                    (_INDEX_POLICES.setdefault(fam.lower(), {})
                     .setdefault(sub.lower(), (str(f), n1, n2)))
    return _INDEX_POLICES


GRAISSES = {
    "extrabold": ["extrabold", "extra bold", "ultrabold", "heavy", "bold"],
    "semibold":  ["semibold", "semi bold", "demibold", "medium", "bold"],
    "bold":      ["bold"],
    "normal":    ["regular", "normal", "book"],
}


class Police:
    """Résout une catégorie de police du config vers un fichier TTF/OTF."""

    def __init__(self, spec: dict):
        famille = spec["famille"]
        graisse = str(spec.get("graisse", "normal")).lower()
        self.taille_pt = float(spec["taille_pt"])
        self.taille_mm = self.taille_pt * MM_PAR_PT
        self.echelle_y = float(spec.get("echelle_verticale", 1.0))
        self.hauteur_mm = self.taille_mm * self.echelle_y  # hauteur de ligne
        self.couleur = spec.get("couleur", "#000000")
        self.italique = str(spec.get("style", "normal")).lower().startswith("ital")

        idx = index_polices()
        fam = idx.get(famille.lower())
        if fam is None:
            repli = "arial" if os.name == "nt" else "dejavu sans"
            print(f"⚠ Police « {famille} » introuvable, repli sur {repli}.",
                  file=sys.stderr)
            fam = idx.get(repli) or next(iter(idx.values()))

        candidats = GRAISSES.get(graisse, []) + ["regular", "normal"]
        if self.italique:  # préférer une vraie italique si elle existe
            candidats = [f"{c} italic" for c in candidats] + ["italic"] + candidats
        trouve = next((c for c in candidats if c in fam), None)
        if trouve is None:
            trouve = next(iter(fam))
            print(f"⚠ « {famille} » : aucune variante {graisse} trouvée, "
                  f"utilisation de « {trouve} » (vérifiez avec -v).", file=sys.stderr)
        self.chemin, nom, style = fam[trouve]
        self.svg_famille = nom
        self.svg_graisse = "bold" if "bold" in style.lower() else "normal"
        self.italique_reel = "italic" in style.lower()
        self.description = f"{nom} / {style} — {self.chemin}"
        self._pil = ImageFont.truetype(self.chemin, int(self.taille_pt * MESURE_ECHELLE))

    def largeur_mm(self, texte: str) -> float:
        l, t, r, b = self._pil.getbbox(texte)
        return (r / MESURE_ECHELLE) * MM_PAR_PT


# ---------------------------------------------------------------- dessin
# Couche de primitives : une seule mise en page, deux sorties (PDF + SVG).
# Coordonnées en mm, origine en haut à gauche, y de texte = ligne de base.

class Dessin:
    def __init__(self, largeur: float, hauteur: float):
        self.W, self.H, self.ops = largeur, hauteur, []

    def rect(self, x, y, l, h, rayon=0, fond="#FFFFFF", ombre=False, rotation=None):
        self.ops.append(("rect", x, y, l, h, rayon, fond, ombre, rotation))

    def ligne(self, x1, y1, x2, y2, couleur="#111111", epaisseur=0.45):
        self.ops.append(("ligne", x1, y1, x2, y2, couleur, epaisseur))

    def texte(self, x, y, contenu, police: Police, ancre="debut", rotation=None):
        self.ops.append(("texte", x, y, contenu, police, ancre, rotation))

    def chemin(self, segments, fond, opacite=1.0):
        """segments : liste de ("M",x,y) | ("L",x,y) | ("C",x1,y1,x2,y2,x,y) | ("Z",)"""
        self.ops.append(("chemin", segments, fond, opacite))

    def image(self, fichier, x, y, l, h=None, deformer=False):
        if h is None:
            with Image.open(fichier) as im:
                h = l * im.size[1] / im.size[0]
        self.ops.append(("image", str(fichier), x, y, l, h, deformer))

    # ---------------------------------------------------------- sortie SVG

    def vers_svg(self) -> str:
        s = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{self.W}mm" '
             f'height="{self.H}mm" viewBox="0 0 {self.W} {self.H}">',
             '<defs><filter id="ombre" x="-20%" y="-20%" width="140%" height="140%">'
             '<feDropShadow dx="0.5" dy="0.8" stdDeviation="0.8" flood-opacity="0.25"/>'
             '</filter></defs>']
        for op in self.ops:
            genre = op[0]
            if genre == "rect":
                _, x, y, l, h, r, fond, ombre, rot = op
                el = (f'<rect x="{x:.2f}" y="{y:.2f}" width="{l:.2f}" height="{h:.2f}" '
                      f'rx="{r}" fill="{fond}"{" filter=\"url(#ombre)\"" if ombre else ""}/>')
                s.append(self._svg_rot(el, rot))
            elif genre == "ligne":
                _, x1, y1, x2, y2, c, ep = op
                s.append(f'<line x1="{x1:.2f}" y1="{y1:.2f}" x2="{x2:.2f}" y2="{y2:.2f}" '
                         f'stroke="{c}" stroke-width="{ep}"/>')
            elif genre == "texte":
                _, x, y, contenu, p, ancre, rot = op
                a = {"debut": "start", "milieu": "middle", "fin": "end"}[ancre]
                attrs = (f'font-family="{escape(p.svg_famille)}" '
                         f'font-size="{p.taille_mm:.3f}" fill="{p.couleur}"')
                if p.svg_graisse == "bold":
                    attrs += ' font-weight="bold"'
                if p.italique:
                    attrs += ' font-style="italic"'
                if p.echelle_y != 1.0:  # étirement autour de la ligne de base
                    el = (f'<text x="0" y="0" text-anchor="{a}" {attrs} transform='
                          f'"translate({x:.2f} {y:.2f}) scale(1 {p.echelle_y})">'
                          f'{escape(contenu)}</text>')
                else:
                    el = (f'<text x="{x:.2f}" y="{y:.2f}" text-anchor="{a}" {attrs}>'
                          f'{escape(contenu)}</text>')
                s.append(self._svg_rot(el, rot))
            elif genre == "chemin":
                _, segments, fond, opa = op
                d = " ".join(seg[0] + " " + " ".join(f"{v:.1f}" for v in seg[1:])
                             for seg in segments)
                s.append(f'<path d="{d}" fill="{fond}" opacity="{opa}"/>')
            elif genre == "image":
                _, fichier, x, y, l, h, deformer = op
                chemin = Path(fichier)
                mime = "image/png" if chemin.suffix.lower() == ".png" else "image/jpeg"
                data = base64.b64encode(chemin.read_bytes()).decode()
                par = "none" if deformer else "xMidYMid meet"
                s.append(f'<image x="{x}" y="{y}" width="{l}" height="{h:.2f}" '
                         f'href="data:{mime};base64,{data}" preserveAspectRatio="{par}"/>')
        s.append("</svg>")
        return "\n".join(s)

    @staticmethod
    def _svg_rot(element: str, rot) -> str:
        if not rot:
            return element
        angle, cx, cy = rot
        return f'<g transform="rotate({angle} {cx:.1f} {cy:.1f})">{element}</g>'

    # ---------------------------------------------------------- sortie PDF

    def dessiner_pdf(self, c):
        from reportlab.lib.colors import Color, HexColor
        from reportlab.lib.units import mm as MM
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont

        Y = lambda y: (self.H - y) * MM  # origine PDF en bas à gauche

        def nom_pdf(p: Police) -> str:
            if p.chemin not in _POLICES_PDF:
                nom = f"F{len(_POLICES_PDF)}"
                try:
                    pdfmetrics.registerFont(TTFont(nom, p.chemin))
                except Exception as e:
                    sys.exit(f"Impossible d'embarquer la police {p.chemin} dans le "
                             f"PDF ({e}).\nS'il s'agit d'un OTF à contours "
                             f"PostScript/CFF, convertissez-le une fois en TTF :\n"
                             f"    pip install otf2ttf && otf2ttf police.otf")
                _POLICES_PDF[p.chemin] = nom
            return _POLICES_PDF[p.chemin]

        def rotation(rot):
            if rot:
                angle, cx, cy = rot
                c.translate(cx * MM, Y(cy))
                c.rotate(-angle)  # y inversé → sens horaire inversé
                c.translate(-cx * MM, -Y(cy))

        for op in self.ops:
            genre = op[0]
            if genre == "rect":
                _, x, y, l, h, r, fond, ombre, rot = op
                c.saveState()
                rotation(rot)
                if ombre:  # ombre douce approximée par couches translucides
                    for delta, alpha in ((1.4, 0.05), (1.0, 0.07), (0.6, 0.10)):
                        c.setFillColor(Color(0, 0, 0, alpha=alpha))
                        c.roundRect((x + 0.5) * MM - delta, Y(y + h + 0.8) - delta,
                                    l * MM + 2 * delta, h * MM + 2 * delta,
                                    r * MM, stroke=0, fill=1)
                c.setFillColor(HexColor(fond))
                c.roundRect(x * MM, Y(y + h), l * MM, h * MM, r * MM, stroke=0, fill=1)
                c.restoreState()
            elif genre == "ligne":
                _, x1, y1, x2, y2, coul, ep = op
                c.setStrokeColor(HexColor(coul))
                c.setLineWidth(ep * MM)
                c.line(x1 * MM, Y(y1), x2 * MM, Y(y2))
            elif genre == "texte":
                _, x, y, contenu, p, ancre, rot = op
                dx = {"debut": 0.0, "milieu": -0.5, "fin": -1.0}[ancre] \
                    * p.largeur_mm(contenu)
                c.saveState()
                rotation(rot)
                c.setFillColor(HexColor(p.couleur))
                c.setFont(nom_pdf(p), p.taille_pt)
                oblique = p.italique and not p.italique_reel
                if p.echelle_y != 1.0 or oblique:
                    c.translate((x + dx) * MM, Y(y))  # transforms sur la ligne de base
                    if p.echelle_y != 1.0:
                        c.transform(1, 0, 0, p.echelle_y, 0, 0)
                    if oblique:
                        c.transform(1, 0, 0.21, 1, 0, 0)
                    c.drawString(0, 0, contenu)
                else:
                    c.drawString((x + dx) * MM, Y(y), contenu)
                c.restoreState()
            elif genre == "chemin":
                _, segments, fond, opa = op
                col = HexColor(fond)
                c.setFillColor(Color(col.red, col.green, col.blue, alpha=opa))
                p = c.beginPath()
                for seg in segments:
                    if seg[0] == "M":
                        p.moveTo(seg[1] * MM, Y(seg[2]))
                    elif seg[0] == "L":
                        p.lineTo(seg[1] * MM, Y(seg[2]))
                    elif seg[0] == "C":
                        p.curveTo(seg[1] * MM, Y(seg[2]), seg[3] * MM, Y(seg[4]),
                                  seg[5] * MM, Y(seg[6]))
                    elif seg[0] == "Z":
                        p.close()
                c.drawPath(p, stroke=0, fill=1)
            elif genre == "image":
                _, fichier, x, y, l, h, deformer = op
                c.drawImage(fichier, x * MM, Y(y + h), l * MM, h * MM,
                            mask="auto", preserveAspectRatio=not deformer)
        c.showPage()


_POLICES_PDF = {}


def ecrire_pdf(dessins: list, sortie: str):
    from reportlab.lib.units import mm as MM
    from reportlab.pdfgen import canvas as rl
    c = rl.Canvas(sortie, pagesize=(dessins[0].W * MM, dessins[0].H * MM),
                  pageCompression=1)
    for d in dessins:
        d.dessiner_pdf(c)
    c.save()


# ---------------------------------------------------------------- données

def fmt_heure(h) -> str:
    if isinstance(h, datetime):
        h = h.time()
    if isinstance(h, time):
        if h == time(0, 0):
            return "MINUIT"
        return f"{h.hour:02d}H{h.minute:02d}" if h.minute else f"{h.hour:02d}H"
    return str(h).strip()  # déjà formaté en texte


def fmt_prix(p) -> str | None:
    if p is None or p == "":
        return None
    if isinstance(p, (int, float)):
        if float(p).is_integer():
            return f"{int(p)}€"
        return f"{p:.2f}".replace(".", ",") + "€"
    return str(p).strip()


def lire_evenements(chemin: str):
    """Retourne (liste de (date, [événements]), (mois, année))."""
    ws = load_workbook(chemin, data_only=True).active
    par_jour = defaultdict(list)
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        d, nom, debut, fin, prix = row[:5]
        if isinstance(d, datetime):
            d = d.date()
        par_jour[d].append({
            "nom": str(nom).strip(),
            "heures": f"{fmt_heure(debut)}-{fmt_heure(fin)}",
            "prix": fmt_prix(prix),
        })
    jours = sorted(par_jour.items())
    if not jours:
        sys.exit("Aucun événement trouvé dans le fichier Excel.")
    d0 = jours[0][0]
    return jours, (MOIS[d0.month - 1], d0.year)


def libelle_jour(d: date) -> str:
    return f"{JOURS[d.weekday()]} {d.day:02d} {MOIS[d.month - 1]}"


# ---------------------------------------------------------------- mise en page

def couper_lignes(texte: str, police: Police, largeur_max: float) -> list[str]:
    mots, lignes, courante = texte.split(), [], ""
    for mot in mots:
        essai = f"{courante} {mot}".strip()
        if courante and police.largeur_mm(essai) > largeur_max:
            lignes.append(courante)
            courante = mot
        else:
            courante = essai
    if courante:
        lignes.append(courante)
    return lignes


class Case:
    """Une case = un jour. Prépare son contenu et connaît sa hauteur."""

    def __init__(self, d, evenements, cfg, polices):
        c = cfg["cases"]
        self.pad = c["rembourrage_mm"]
        self.largeur = c["largeur_mm"]
        interne = self.largeur - 2 * self.pad
        il = c.get("interligne", 1.15)
        self.p = polices

        self.titre = libelle_jour(d)
        self.evts = []
        h = self.pad + polices["jour"].hauteur_mm * il + 1.2  # titre + soulignement
        for i, e in enumerate(evenements):
            if i:
                h += 2.4  # séparateur entre événements
            lignes = couper_lignes(e["nom"], polices["evenement"], interne)
            prix_sur_ligne = False
            if e["prix"]:
                reste = interne - polices["evenement"].largeur_mm(lignes[-1])
                prix_sur_ligne = reste > polices["prix"].largeur_mm(e["prix"]) + 3
            h += polices["heure"].hauteur_mm * il
            h += len(lignes) * polices["evenement"].hauteur_mm * il
            if e["prix"] and not prix_sur_ligne:
                h += polices["prix"].hauteur_mm * il
            self.evts.append({**e, "lignes": lignes, "prix_sur_ligne": prix_sur_ligne})
        self.hauteur = h + self.pad * 0.6
        self.il = il

    def dessiner(self, dessin: Dessin, x, y, cfg):
        c = cfg["cases"]
        p = self.p
        dessin.rect(x, y, self.largeur, self.hauteur, c["rayon_mm"],
                    c["fond"], ombre=c.get("ombre", False))
        xg, xd = x + self.pad, x + self.largeur - self.pad
        cy = y + self.pad + p["jour"].hauteur_mm * 0.85
        dessin.texte(xg, cy, self.titre, p["jour"])
        cy += p["jour"].hauteur_mm * (self.il - 0.85) + 0.6
        dessin.ligne(xg, cy, xd, cy)
        cy += 0.6
        for i, e in enumerate(self.evts):
            if i:
                cy += 1.2
                dessin.ligne(xg, cy, xd, cy)
                cy += 1.2
            cy += p["heure"].hauteur_mm * self.il
            dessin.texte(xd, cy - p["heure"].hauteur_mm * 0.25, e["heures"],
                         p["heure"], ancre="fin")
            for j, ligne in enumerate(e["lignes"]):
                cy += p["evenement"].hauteur_mm * self.il
                dessin.texte(xg, cy - p["evenement"].hauteur_mm * 0.25, ligne,
                             p["evenement"])
                if e["prix"] and e["prix_sur_ligne"] and j == len(e["lignes"]) - 1:
                    dessin.texte(xd, cy - p["evenement"].hauteur_mm * 0.25,
                                 e["prix"], p["prix"], ancre="fin")
            if e["prix"] and not e["prix_sur_ligne"]:
                cy += p["prix"].hauteur_mm * self.il
                dessin.texte(xd, cy - p["prix"].hauteur_mm * 0.25, e["prix"],
                             p["prix"], ancre="fin")


def paginer(cases, cfg):
    """Répartit les cases en pages de deux colonnes (J à gauche, J+1 à droite…).

    Une page est pleine quand une colonne ne peut plus accueillir sa prochaine
    case avec l'écart minimal (la colonne droite compte en plus son décalage
    initial sous la première case gauche).
    """
    c, z = cfg["cases"], cfg["zone_cases"]
    bande = z["bas_mm"] - z["haut_mm"]
    e = c.get("ecart_min_mm", 5)

    def tient(colonne, case, droite):
        k = len(colonne) + 1
        total = sum(x.hauteur for x in colonne) + case.hauteur
        return total + (k - 1 + (1 if droite else 0)) * e <= bande

    pages, G, D = [], [], []
    for case in cases:
        a_gauche = (len(G) + len(D)) % 2 == 0
        colonne = G if a_gauche else D
        if tient(colonne, case, not a_gauche):
            colonne.append(case)
        elif not (G or D):
            print(f"⚠ La case « {case.titre} » dépasse à elle seule la zone : "
                  f"réduisez les tailles de police.", file=sys.stderr)
            G.append(case)
        else:
            pages.append((G, D))
            G, D = [case], []  # nouvelle page, on repart à gauche
    pages.append((G, D))
    return pages, e


def positions_page(G, D, cfg, e):
    """Positions (x, y) des cases d'une page.

    Colonne gauche : première case en haut de zone (toujours la plus haute),
    répartition régulière sur la bande. Colonne droite : chaque case est
    idéalement centrée entre les deux cases gauches qu'elle sépare
    chronologiquement, puis remontée ou descendue juste assez pour respecter
    la zone et l'écart minimal `e` (qui sert aussi de décalage minimal sous
    la première case gauche)."""
    c, z = cfg["cases"], cfg["zone_cases"]
    haut, bas = z["haut_mm"], z["bas_mm"]
    bande = bas - haut

    # Colonne gauche : répartition régulière depuis le haut
    if len(G) > 1:
        g = (bande - sum(x.hauteur for x in G)) / (len(G) - 1)
        g = max(min(g, c.get("ecart_max_mm", float("inf"))), e)
    else:
        g = 0
    yG, y = [], haut
    for case in G:
        yG.append(y)
        y += case.hauteur + g

    # Colonne droite : position idéale au centre de la paire gauche encadrante
    yD = []
    for i, case in enumerate(D):
        if i + 1 < len(G):
            milieu = (yG[i] + yG[i + 1] + G[i + 1].hauteur) / 2
        else:  # dernière case, après la dernière case gauche
            milieu = yG[-1] + G[-1].hauteur + g / 2
        yD.append(milieu - case.hauteur / 2)
    if yD:
        # passe descendante : chaque case assez haute pour laisser passer les suivantes
        yD[-1] = min(yD[-1], bas - D[-1].hauteur)
        for i in range(len(D) - 2, -1, -1):
            yD[i] = min(yD[i], yD[i + 1] - e - D[i].hauteur)
        # passe montante : zone et écart minimal
        yD[0] = max(yD[0], haut + e)
        for i in range(1, len(D)):
            yD[i] = max(yD[i], yD[i - 1] + D[i - 1].hauteur + e)

    largeur_page = cfg["page"]["largeur_mm"]
    x_gauche = c["marge_gauche_mm"]
    x_droite = largeur_page - c["marge_droite_mm"] - c["largeur_mm"]
    return ([(case, x_gauche, y) for case, y in zip(G, yG)]
            + [(case, x_droite, y) for case, y in zip(D, yD)])


# ---------------------------------------------------------------- affiche

def dessiner_decor(dessin: Dessin, cfg):
    d = cfg.get("decor", {})
    if not d.get("actif"):
        return
    c = d.get("couleurs", ["#45C4E9", "#F0A830", "#2E7FC1"])
    W, H = dessin.W, dessin.H
    dessin.chemin([("M", W*0.30, 0), ("C", W*0.55, H*0.05, W*0.60, H*0.18, W*0.52, H*0.30),
                   ("C", W*0.44, H*0.44, W*0.10, H*0.50, 0, H*0.42),
                   ("L", 0, 0), ("Z",)], c[0], 0.85)
    dessin.chemin([("M", W, H*0.30), ("C", W*0.75, H*0.36, W*0.62, H*0.52, W*0.70, H*0.68),
                   ("C", W*0.78, H*0.85, W*0.55, H, W*0.30, H),
                   ("L", W, H), ("Z",)], c[1], 0.9)
    dessin.chemin([("M", 0, H*0.55), ("C", W*0.18, H*0.50, W*0.30, H*0.62, W*0.24, H*0.75),
                   ("C", W*0.18, H*0.88, W*0.05, H*0.92, 0, H*0.90), ("Z",)],
                  c[2] if len(c) > 2 else c[0], 0.75)

def preparer_fond(chemin: Path, cfg_page: dict, largeur_mm: float) -> Path:
    """Rééchantillonne le fond à `fond_dpi` (et en JPEG si `fond_qualite`)."""
    dpi = cfg_page.get("fond_dpi")
    if not dpi:
        return chemin
    import tempfile
    with Image.open(chemin) as im:
        cible = round(largeur_mm / 25.4 * dpi)
        if im.size[0] <= cible:
            return chemin
        im = im.resize((cible, round(im.size[1] * cible / im.size[0])),
                       Image.LANCZOS)
        qualite = cfg_page.get("fond_qualite")
        tmp = Path(tempfile.gettempdir())
        if qualite:  # JPEG : encore plus léger, pour fonds sans transparence
            sortie = tmp / f"{chemin.stem}_{dpi}dpi.jpg"
            im.convert("RGB").save(sortie, quality=qualite)
        else:
            sortie = tmp / f"{chemin.stem}_{dpi}dpi.png"
            im.save(sortie, optimize=True)
        return sortie
    
def dessiner_gabarit(dessin: Dessin, cfg, polices, mois, annee, dossier: Path):
    """Éléments communs à chaque page : fond, décor, titre, pastille, pied, images."""
    W, H = dessin.W, dessin.H
    dessin.rect(0, 0, W, H, fond=cfg["page"].get("fond", "#FFFFFF"))
    fond = cfg["page"].get("fond_image")
    if fond:
        chemin_fond = preparer_fond((dossier / fond).resolve(), cfg["page"], W)
        dessin.image(chemin_fond, 0, 0, W, H, deformer=True)
    dessiner_decor(dessin, cfg)

    a = cfg["affiche"]
    pt, pst = polices["titre"], polices["sous_titre"]
    x0, y0 = 10, 12 + pt.hauteur_mm
    dessin.texte(x0, y0, a["titre"], pt)
    dessin.texte(x0 + pt.largeur_mm(a["titre"]) + 4, y0, a["sous_titre"], pst)

    pm, texte_mois = polices["mois"], f"{mois} {annee}"
    pcfg = cfg["pastille_mois"]
    lm = pm.largeur_mm(texte_mois) + 12
    hm = pm.hauteur_mm + 5
    xm = W - pcfg.get("marge_droite_mm", 12) - lm
    ym = pcfg.get("y_mm", y0 + 6) + 10
    rot = (-3, xm + lm / 2, ym + hm / 2)
    dessin.rect(xm, ym, lm, hm, rayon=hm / 2, fond=pcfg["fond"], rotation=rot)
    dessin.texte(xm + lm / 2, ym + hm * 0.72, texte_mois, pm, ancre="milieu",
                 rotation=rot)

    pp = polices["pied"]
    yp = cfg["pied_de_page"]["haut_mm"] + pp.hauteur_mm + 5
    for ligne in a.get("pied", []):
        dessin.texte(cfg["cases"]["marge_gauche_mm"], yp, ligne, pp)
        yp += pp.hauteur_mm * 1.3

    for img in cfg.get("images", []) or []:
        chemin = (dossier / img["chemin"]).resolve()
        if not chemin.exists():
            print(f"⚠ Image introuvable : {chemin}", file=sys.stderr)
            continue
        dessin.image(chemin, img["x_mm"], img["y_mm"], img["largeur_mm"],
                     img.get("hauteur_mm"))


def generer(xlsx: str, config: str, sortie: str, verbeux: bool = False):
    cfg = yaml.safe_load(Path(config).read_text(encoding="utf-8"))
    polices = {nom: Police(spec) for nom, spec in cfg["polices"].items()}
    if verbeux:
        for nom, p in polices.items():
            print(f"  {nom:<12} → {p.description}")
    jours, (mois, annee) = lire_evenements(xlsx)

    cases = [Case(d, evts, cfg, polices) for d, evts in jours]
    pages, e = paginer(cases, cfg)

    W, H = cfg["page"]["largeur_mm"], cfg["page"]["hauteur_mm"]
    dossier = Path(config).parent
    dessins = []
    for G, D in pages:
        dessin = Dessin(W, H)
        dessiner_gabarit(dessin, cfg, polices, mois, annee, dossier)
        for case, x, y in positions_page(G, D, cfg, e):
            case.dessiner(dessin, x, y, cfg)
        dessins.append(dessin)

    ecrire_pdf(dessins, sortie)
    for i, dessin in enumerate(dessins):
        suffixe = "" if i == 0 else f"_{i + 1}"
        chemin_svg = Path(sortie).with_suffix("")
        Path(f"{chemin_svg}{suffixe}.svg").write_text(dessin.vers_svg(),
                                                      encoding="utf-8")
    print(f"✓ {sortie} généré ({len(cases)} jours, "
          f"{sum(len(c.evts) for c in cases)} événements, "
          f"{len(pages)} page{'s' if len(pages) > 1 else ''})")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("xlsx", help="Fichier Excel des événements")
    ap.add_argument("config", nargs="?", default="config.yaml")
    ap.add_argument("-o", "--sortie", default="affiche.pdf")
    ap.add_argument("-v", "--verbeux", action="store_true",
                    help="affiche la police résolue pour chaque catégorie")
    args = ap.parse_args()
    generer(args.xlsx, args.config, args.sortie, args.verbeux)
