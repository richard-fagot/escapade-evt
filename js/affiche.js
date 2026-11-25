/* Génération de l'affiche mensuelle (PDF) à partir des lignes Excel.
 * Dépend de pdf-lib, fontkit, js-yaml et Excel (js/excel.js).
 *
 * Coordonnées en mm, origine en haut à gauche, y de texte = ligne de base
 * (conversion vers le repère PDF, en points et origine en bas, au dessin). */
"use strict";

const Affiche = (() => {
  const { PDFDocument, rgb, pushGraphicsState, popGraphicsState,
          concatTransformationMatrix } = PDFLib;

  const PT_PAR_MM = 72 / 25.4;
  const MM_PAR_PT = 25.4 / 72;
  const JOURS = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi"];
  const MOIS = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet",
                "Août", "Septembre", "Octobre", "Novembre", "Décembre"];

  // ------------------------------------------------------------ ressources

  async function charger(url, type = "arrayBuffer") {
    let rep;
    try {
      rep = await fetch(url, { cache: "no-cache" });
    } catch (e) {
      throw new Error(`Impossible de charger « ${url} ». Ouvrez la page via un serveur web ` +
                      `(GitHub Pages, ou « python -m http.server » dans le dossier) ` +
                      `plutôt qu'en double-cliquant sur index.html.`);
    }
    if (!rep.ok) throw new Error(`Fichier « ${url} » introuvable (${rep.status}).`);
    return rep[type]();
  }

  async function chargerConfig(url = "config.yaml") {
    return jsyaml.load(await charger(url, "text"));
  }

  function couleur(hex) {
    const h = String(hex).replace("#", "");
    const n = parseInt(h.length === 3 ? h.replace(/./g, "$&$&") : h, 16);
    return rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255);
  }

  class Police {
    constructor(spec, fonte) {
      this.fonte = fonte;
      this.taillePt = Number(spec.taille_pt);
      this.tailleMm = this.taillePt * MM_PAR_PT;
      this.echelleY = Number(spec.echelle_verticale ?? 1);
      this.hauteurMm = this.tailleMm * this.echelleY;  // hauteur de ligne
      this.couleur = couleur(spec.couleur ?? "#000000");
      this.italique = String(spec.style ?? "normal").toLowerCase().startsWith("ital");
    }
    largeurMm(texte) {
      return this.fonte.widthOfTextAtSize(texte, this.taillePt) * MM_PAR_PT;
    }
  }

  /** Ré-échantillonne le fond à `fond_dpi` et le convertit en JPEG si
   *  `fond_qualite` est défini (sinon PNG). Retourne { octets, jpeg }. */
  async function preparerFond(octets, cfgPage, largeurMm) {
    const dpi = cfgPage.fond_dpi;
    const qualite = cfgPage.fond_qualite;
    const estPng = octets[0] === 0x89 && octets[1] === 0x50;
    if (!dpi && !qualite) return { octets, jpeg: !estPng };
    const bitmap = await createImageBitmap(new Blob([octets]));
    let l = bitmap.width, h = bitmap.height;
    if (dpi) {
      const cible = Math.round(largeurMm / 25.4 * dpi);
      if (cible < l) { h = Math.round(h * cible / l); l = cible; }
    }
    const canvas = document.createElement("canvas");
    canvas.width = l; canvas.height = h;
    const ctx = canvas.getContext("2d");
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = "high";
    if (qualite) { ctx.fillStyle = "#FFFFFF"; ctx.fillRect(0, 0, l, h); }
    ctx.drawImage(bitmap, 0, 0, l, h);
    const blob = await new Promise(ok => canvas.toBlob(ok, qualite ? "image/jpeg" : "image/png",
                                                       qualite ? qualite / 100 : undefined));
    return { octets: new Uint8Array(await blob.arrayBuffer()), jpeg: !!qualite };
  }

  // ------------------------------------------------------------ données

  function deuxChiffres(n) { return String(n).padStart(2, "0"); }

  function dateExcel(v, date1904) {
    if (typeof v === "number") {
      const origine = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
      const d = new Date(origine + Math.round(v * 86400) * 1000);
      return { a: d.getUTCFullYear(), m: d.getUTCMonth() + 1, j: d.getUTCDate() };
    }
    const t = String(v).trim();
    let r = t.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);  // 12/06/2026
    if (r) return { a: r[3].length === 2 ? 2000 + Number(r[3]) : Number(r[3]), m: Number(r[2]), j: Number(r[1]) };
    r = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);                    // 2026-06-12
    if (r) return { a: Number(r[1]), m: Number(r[2]), j: Number(r[3]) };
    return null;
  }

  function fmtHeure(h) {
    if (h === null || h === undefined || h === "") return "";
    if (typeof h === "number") {
      const minutes = Math.round((h - Math.floor(h)) * 1440) % 1440;
      if (minutes === 0) return "MINUIT";
      const hh = Math.floor(minutes / 60), mm = minutes % 60;
      return mm ? `${deuxChiffres(hh)}H${deuxChiffres(mm)}` : `${deuxChiffres(hh)}H`;
    }
    return String(h).trim();  // déjà formaté en texte
  }

  function fmtPrix(p) {
    if (p === null || p === undefined || p === "") return null;
    if (typeof p === "number") {
      return Number.isInteger(p) ? `${p}€` : p.toFixed(2).replace(".", ",") + "€";
    }
    return String(p).trim() || null;
  }

  /** Retourne { jours: [{date, evenements}], mois, annee }. */
  function lireEvenements(buffer) {
    const { lignes, date1904 } = Excel.lire(buffer);
    const parJour = new Map();
    lignes.slice(1).forEach((row, i) => {
      if (!row || row[0] === null || row[0] === undefined || row[0] === "") return;
      const [d, nom, debut, fin, prix] = row;
      const date = dateExcel(d, date1904);
      if (!date) throw new Error(`Ligne ${i + 2} : date « ${d} » non reconnue.`);
      const cle = date.a * 10000 + date.m * 100 + date.j;
      if (!parJour.has(cle)) parJour.set(cle, { date, evenements: [] });
      const heures = [fmtHeure(debut), fmtHeure(fin)].filter(Boolean).join("-");
      parJour.get(cle).evenements.push({
        nom: String(nom ?? "").trim(), heures, prix: fmtPrix(prix),
      });
    });
    const jours = [...parJour.entries()].sort((a, b) => a[0] - b[0]).map(e => e[1]);
    if (!jours.length) throw new Error("Aucun événement trouvé dans le fichier Excel.");
    const d0 = jours[0].date;
    return { jours, mois: MOIS[d0.m - 1], annee: d0.a };
  }

  function libelleJour({ a, m, j }) {
    const jourSemaine = new Date(Date.UTC(a, m - 1, j)).getUTCDay();
    return `${JOURS[jourSemaine]} ${deuxChiffres(j)} ${MOIS[m - 1]}`;
  }

  // ------------------------------------------------------------ dessin

  /** Couche de primitives au-dessus d'une page pdf-lib, en mm depuis le haut. */
  class Dessin {
    constructor(page, largeur, hauteur) {
      this.page = page; this.W = largeur; this.H = hauteur;
    }
    // Transformation (mm, repère haut-gauche) → points PDF
    X(x) { return x * PT_PAR_MM; }
    Y(y) { return (this.H - y) * PT_PAR_MM; }

    avecRotation(rot, fn) {
      if (!rot) return fn();
      const [angle, cx, cy] = rot;
      const a = -angle * Math.PI / 180;  // y inversé → sens inversé
      const cos = Math.cos(a), sin = Math.sin(a);
      const px = this.X(cx), py = this.Y(cy);
      this.page.pushOperators(pushGraphicsState(),
        concatTransformationMatrix(cos, sin, -sin, cos,
                                   px - cos * px + sin * py, py - sin * px - cos * py));
      fn();
      this.page.pushOperators(popGraphicsState());
    }

    /** Chemin SVG d'un rectangle arrondi, en mm repère haut-gauche. */
    static cheminRect(x, y, l, h, r) {
      r = Math.min(r, l / 2, h / 2);
      if (r <= 0) return `M ${x} ${y} H ${x + l} V ${y + h} H ${x} Z`;
      const k = 0.5523 * r;
      return `M ${x + r} ${y} H ${x + l - r} C ${x + l - r + k} ${y} ${x + l} ${y + r - k} ${x + l} ${y + r} ` +
             `V ${y + h - r} C ${x + l} ${y + h - r + k} ${x + l - r + k} ${y + h} ${x + l - r} ${y + h} ` +
             `H ${x + r} C ${x + r - k} ${y + h} ${x} ${y + h - r + k} ${x} ${y + h - r} ` +
             `V ${y + r} C ${x} ${y + r - k} ${x + r - k} ${y} ${x + r} ${y} Z`;
    }

    chemin(d, fond, opacite = 1) {
      // drawSvgPath inverse l'axe y depuis (x, y) : origine = haut de page
      this.page.drawSvgPath(d, { x: 0, y: this.H * PT_PAR_MM, scale: PT_PAR_MM,
                                 color: couleur(fond), opacity: opacite, borderWidth: 0 });
    }

    rect(x, y, l, h, { rayon = 0, fond = "#FFFFFF", ombre = false, rotation = null } = {}) {
      this.avecRotation(rotation, () => {
        if (ombre) {  // ombre douce approximée par couches translucides
          for (const [delta, alpha] of [[0.5, 0.05], [0.35, 0.07], [0.2, 0.10]]) {
            this.chemin(Dessin.cheminRect(x + 0.5 - delta, y + 0.8 - delta, l + 2 * delta,
                                          h + 2 * delta, rayon + delta), "#000000", alpha);
          }
        }
        this.chemin(Dessin.cheminRect(x, y, l, h, rayon), fond);
      });
    }

    ligne(x1, y1, x2, y2, coul = "#111111", epaisseur = 0.45) {
      this.page.drawLine({ start: { x: this.X(x1), y: this.Y(y1) },
                           end: { x: this.X(x2), y: this.Y(y2) },
                           thickness: epaisseur * PT_PAR_MM, color: couleur(coul) });
    }

    texte(x, y, contenu, p, { ancre = "debut", rotation = null } = {}) {
      const dx = { debut: 0, milieu: -0.5, fin: -1 }[ancre] * p.largeurMm(contenu);
      this.avecRotation(rotation, () => {
        // transformations appliquées autour de la ligne de base
        this.page.pushOperators(pushGraphicsState(),
          concatTransformationMatrix(1, 0, 0, 1, this.X(x + dx), this.Y(y)));
        if (p.echelleY !== 1) this.page.pushOperators(concatTransformationMatrix(1, 0, 0, p.echelleY, 0, 0));
        if (p.italique) this.page.pushOperators(concatTransformationMatrix(1, 0, 0.21, 1, 0, 0));
        this.page.drawText(contenu, { x: 0, y: 0, font: p.fonte, size: p.taillePt, color: p.couleur });
        this.page.pushOperators(popGraphicsState());
      });
    }

    image(img, x, y, l, h) {
      if (h === undefined || h === null) h = l * img.height / img.width;
      this.page.drawImage(img, { x: this.X(x), y: this.Y(y + h),
                                 width: l * PT_PAR_MM, height: h * PT_PAR_MM });
    }
  }

  // ------------------------------------------------------------ mise en page

  function couperLignes(texte, police, largeurMax) {
    const lignes = [];
    let courante = "";
    for (const mot of texte.split(/\s+/).filter(Boolean)) {
      const essai = courante ? `${courante} ${mot}` : mot;
      if (courante && police.largeurMm(essai) > largeurMax) {
        lignes.push(courante);
        courante = mot;
      } else {
        courante = essai;
      }
    }
    if (courante) lignes.push(courante);
    return lignes.length ? lignes : [""];
  }

  /** Une case = un jour. Prépare son contenu et connaît sa hauteur. */
  class Case {
    constructor(jour, cfg, polices) {
      const c = cfg.cases;
      this.pad = c.rembourrage_mm;
      this.largeur = c.largeur_mm;
      this.il = c.interligne ?? 1.15;
      this.p = polices;
      const interne = this.largeur - 2 * this.pad;
      const il = this.il;

      this.titre = libelleJour(jour.date);
      this.evts = [];
      let h = this.pad + polices.jour.hauteurMm * il + 1.2;  // titre + soulignement
      jour.evenements.forEach((e, i) => {
        if (i) h += 2.4;  // séparateur entre événements
        const lignes = couperLignes(e.nom, polices.evenement, interne);
        let prixSurLigne = false;
        if (e.prix) {
          const reste = interne - polices.evenement.largeurMm(lignes[lignes.length - 1]);
          prixSurLigne = reste > polices.prix.largeurMm(e.prix) + 3;
        }
        h += polices.heure.hauteurMm * il;
        h += lignes.length * polices.evenement.hauteurMm * il;
        if (e.prix && !prixSurLigne) h += polices.prix.hauteurMm * il;
        this.evts.push({ ...e, lignes, prixSurLigne });
      });
      this.hauteur = h + this.pad * 0.6;
    }

    dessiner(dessin, x, y, cfg) {
      const c = cfg.cases, p = this.p, il = this.il;
      dessin.rect(x, y, this.largeur, this.hauteur,
                  { rayon: c.rayon_mm, fond: c.fond, ombre: !!c.ombre });
      const xg = x + this.pad, xd = x + this.largeur - this.pad;
      let cy = y + this.pad + p.jour.hauteurMm * 0.85;
      dessin.texte(xg, cy, this.titre, p.jour);
      cy += p.jour.hauteurMm * (il - 0.85) + 0.6;
      dessin.ligne(xg, cy, xd, cy);
      cy += 0.6;
      this.evts.forEach((e, i) => {
        if (i) {
          cy += 1.2;
          dessin.ligne(xg, cy, xd, cy);
          cy += 1.2;
        }
        cy += p.heure.hauteurMm * il;
        dessin.texte(xd, cy - p.heure.hauteurMm * 0.25, e.heures, p.heure, { ancre: "fin" });
        e.lignes.forEach((ligne, j) => {
          cy += p.evenement.hauteurMm * il;
          const base = cy - p.evenement.hauteurMm * 0.25;
          dessin.texte(xg, base, ligne, p.evenement);
          if (e.prix && e.prixSurLigne && j === e.lignes.length - 1) {
            dessin.texte(xd, base, e.prix, p.prix, { ancre: "fin" });
          }
        });
        if (e.prix && !e.prixSurLigne) {
          cy += p.prix.hauteurMm * il;
          dessin.texte(xd, cy - p.prix.hauteurMm * 0.25, e.prix, p.prix, { ancre: "fin" });
        }
      });
    }
  }

  /** Répartit les cases en pages de deux colonnes (J à gauche, J+1 à droite…).
   *  Une page est pleine quand une colonne ne peut plus accueillir sa prochaine
   *  case avec l'écart minimal (la colonne droite compte en plus son décalage
   *  initial sous la première case gauche). */
  function paginer(cases, cfg, avertir) {
    const z = cfg.zone_cases;
    const bande = z.bas_mm - z.haut_mm;
    const e = cfg.cases.ecart_min_mm ?? 5;
    const somme = col => col.reduce((s, x) => s + x.hauteur, 0);
    const tient = (col, cas, droite) =>
      somme(col) + cas.hauteur + (col.length + (droite ? 1 : 0)) * e <= bande;

    const pages = [];
    let G = [], D = [];
    for (const cas of cases) {
      const aGauche = (G.length + D.length) % 2 === 0;
      const col = aGauche ? G : D;
      if (tient(col, cas, !aGauche)) {
        col.push(cas);
      } else if (!G.length && !D.length) {
        avertir(`La case « ${cas.titre} » dépasse à elle seule la zone : réduisez les tailles de police.`);
        G.push(cas);
      } else {
        pages.push([G, D]);
        G = [cas]; D = [];  // nouvelle page, on repart à gauche
      }
    }
    pages.push([G, D]);
    return { pages, e };
  }

  /** Positions (x, y) des cases d'une page.
   *  Colonne gauche : première case en haut de zone, répartition régulière.
   *  Colonne droite : chaque case est idéalement centrée entre les deux cases
   *  gauches qu'elle sépare, puis décalée juste assez pour respecter la zone
   *  et l'écart minimal. */
  function positionsPage(G, D, cfg, e) {
    const c = cfg.cases, z = cfg.zone_cases;
    const haut = z.haut_mm, bas = z.bas_mm, bande = bas - haut;

    let g = 0;
    if (G.length > 1) {
      g = (bande - G.reduce((s, x) => s + x.hauteur, 0)) / (G.length - 1);
      g = Math.max(Math.min(g, c.ecart_max_mm ?? Infinity), e);
    }
    const yG = [];
    let y = haut;
    for (const cas of G) { yG.push(y); y += cas.hauteur + g; }

    const yD = D.map((cas, i) => {
      const milieu = i + 1 < G.length
        ? (yG[i] + yG[i + 1] + G[i + 1].hauteur) / 2
        : yG[yG.length - 1] + G[G.length - 1].hauteur + g / 2;
      return milieu - cas.hauteur / 2;
    });
    if (yD.length) {
      const n = D.length;
      yD[n - 1] = Math.min(yD[n - 1], bas - D[n - 1].hauteur);
      for (let i = n - 2; i >= 0; i--) yD[i] = Math.min(yD[i], yD[i + 1] - e - D[i].hauteur);
      yD[0] = Math.max(yD[0], haut + e);
      for (let i = 1; i < n; i++) yD[i] = Math.max(yD[i], yD[i - 1] + D[i - 1].hauteur + e);
    }

    const xGauche = c.marge_gauche_mm;
    const xDroite = cfg.page.largeur_mm - c.marge_droite_mm - c.largeur_mm;
    return [...G.map((cas, i) => [cas, xGauche, yG[i]]),
            ...D.map((cas, i) => [cas, xDroite, yD[i]])];
  }

  // ------------------------------------------------------------ affiche

  function dessinerDecor(dessin, cfg) {
    const d = cfg.decor || {};
    if (!d.actif) return;
    const c = d.couleurs || ["#45C4E9", "#F0A830", "#2E7FC1"];
    const W = dessin.W, H = dessin.H;
    dessin.chemin(`M ${W*0.30} 0 C ${W*0.55} ${H*0.05} ${W*0.60} ${H*0.18} ${W*0.52} ${H*0.30} ` +
                  `C ${W*0.44} ${H*0.44} ${W*0.10} ${H*0.50} 0 ${H*0.42} L 0 0 Z`, c[0], 0.85);
    dessin.chemin(`M ${W} ${H*0.30} C ${W*0.75} ${H*0.36} ${W*0.62} ${H*0.52} ${W*0.70} ${H*0.68} ` +
                  `C ${W*0.78} ${H*0.85} ${W*0.55} ${H} ${W*0.30} ${H} L ${W} ${H} Z`, c[1], 0.9);
    dessin.chemin(`M 0 ${H*0.55} C ${W*0.18} ${H*0.50} ${W*0.30} ${H*0.62} ${W*0.24} ${H*0.75} ` +
                  `C ${W*0.18} ${H*0.88} ${W*0.05} ${H*0.92} 0 ${H*0.90} Z`, c[2] || c[0], 0.75);
  }

  /** Éléments communs à chaque page : fond, décor, titre, pastille, pied, images. */
  function dessinerGabarit(dessin, cfg, polices, mois, annee, res) {
    const W = dessin.W, H = dessin.H;
    dessin.rect(0, 0, W, H, { fond: cfg.page.fond || "#FFFFFF" });
    if (res.fond) dessin.image(res.fond, 0, 0, W, H);
    dessinerDecor(dessin, cfg);

    const a = cfg.affiche;
    const pt = polices.titre, pst = polices.sous_titre;
    const x0 = 10, y0 = 12 + pt.hauteurMm;
    dessin.texte(x0, y0, a.titre, pt);
    dessin.texte(x0 + pt.largeurMm(a.titre) + 4, y0, a.sous_titre, pst);

    const pm = polices.mois, texteMois = `${mois} ${annee}`;
    const pc = cfg.pastille_mois;
    const lm = pm.largeurMm(texteMois) + 12;
    const hm = pm.hauteurMm + 5;
    const xm = W - (pc.marge_droite_mm ?? 12) - lm;
    const ym = (pc.y_mm ?? y0 + 6) + 10;
    const rotation = [-3, xm + lm / 2, ym + hm / 2];
    dessin.rect(xm, ym, lm, hm, { rayon: hm / 2, fond: pc.fond, rotation });
    dessin.texte(xm + lm / 2, ym + hm * 0.72, texteMois, pm, { ancre: "milieu", rotation });

    const pp = polices.pied;
    let yp = cfg.pied_de_page.haut_mm + pp.hauteurMm + 5;
    for (const ligne of a.pied || []) {
      dessin.texte(cfg.cases.marge_gauche_mm, yp, ligne, pp);
      yp += pp.hauteurMm * 1.3;
    }

    for (const { img, spec } of res.images) {
      dessin.image(img, spec.x_mm, spec.y_mm, spec.largeur_mm, spec.hauteur_mm);
    }
  }

  async function embarquerImage(pdf, octets) {
    const png = octets[0] === 0x89 && octets[1] === 0x50;
    return png ? pdf.embedPng(octets) : pdf.embedJpg(octets);
  }

  /** Génère le PDF. Retourne { octets, nomFichier, resume, avertissements }. */
  async function generer(bufferExcel, cfg) {
    const avertissements = [];
    const avertir = m => avertissements.push(m);
    const { jours, mois, annee } = lireEvenements(bufferExcel);

    const pdf = await PDFDocument.create();
    pdf.registerFontkit(fontkit);
    pdf.setTitle(`${cfg.affiche.titre} ${cfg.affiche.sous_titre} — ${mois} ${annee}`);
    pdf.setCreator("Générateur d'affiche L'Escapade");

    // Polices : un fichier n'est embarqué qu'une fois même s'il sert à plusieurs catégories
    const fontes = new Map();
    const polices = {};
    for (const [nom, spec] of Object.entries(cfg.polices)) {
      if (!spec.fichier) throw new Error(`Police « ${nom} » : paramètre « fichier » manquant dans config.yaml.`);
      if (!fontes.has(spec.fichier)) {
        const octets = await charger(spec.fichier);
        fontes.set(spec.fichier, await pdf.embedFont(octets, { subset: true }));
      }
      polices[nom] = new Police(spec, fontes.get(spec.fichier));
    }

    const W = cfg.page.largeur_mm, H = cfg.page.hauteur_mm;
    const res = { fond: null, images: [] };
    if (cfg.page.fond_image) {
      const brut = new Uint8Array(await charger(cfg.page.fond_image));
      const { octets, jpeg } = await preparerFond(brut, cfg.page, W);
      res.fond = jpeg ? await pdf.embedJpg(octets) : await pdf.embedPng(octets);
    }
    for (const spec of cfg.images || []) {
      try {
        res.images.push({ img: await embarquerImage(pdf, new Uint8Array(await charger(spec.chemin))), spec });
      } catch (e) {
        avertir(`Image ignorée : ${spec.chemin} (${e.message})`);
      }
    }

    const cases = jours.map(j => new Case(j, cfg, polices));
    const { pages, e } = paginer(cases, cfg, avertir);
    for (const [G, D] of pages) {
      const dessin = new Dessin(pdf.addPage([W * PT_PAR_MM, H * PT_PAR_MM]), W, H);
      dessinerGabarit(dessin, cfg, polices, mois, annee, res);
      for (const [cas, x, y] of positionsPage(G, D, cfg, e)) cas.dessiner(dessin, x, y, cfg);
    }

    const nbEvts = cases.reduce((s, c) => s + c.evts.length, 0);
    const slug = mois.normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
    return {
      octets: await pdf.save(),
      nomFichier: `affiche_${slug}_${annee}.pdf`,
      resume: `${cases.length} jours, ${nbEvts} événements, ${pages.length} page${pages.length > 1 ? "s" : ""}`,
      avertissements,
    };
  }

  return { chargerConfig, generer };
})();
