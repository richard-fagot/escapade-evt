/* Lecture minimale d'un classeur .xlsx (zip + XML) : première feuille,
 * valeurs brutes des cellules. Dépend de fflate (lib/fflate.min.js). */
"use strict";

const Excel = (() => {
  const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  function xml(fichiers, chemin) {
    const brut = fichiers[chemin];
    if (!brut) return null;
    return new DOMParser().parseFromString(fflate.strFromU8(brut), "application/xml");
  }

  function texte(el) {  // concatène les <t> (gère le texte enrichi)
    return Array.from(el.getElementsByTagNameNS(NS, "t")).map(t => t.textContent).join("");
  }

  function colonne(ref) {  // "C12" → 2
    let n = 0;
    for (const ch of ref.replace(/\d+$/, "")) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }

  /** Retourne { lignes: [[valeur, ...], ...], date1904: bool }.
   *  Nombres → Number, textes → String, cellules vides → null. */
  function lire(buffer) {
    let fichiers;
    try {
      fichiers = fflate.unzipSync(new Uint8Array(buffer));
    } catch (e) {
      throw new Error("Ce fichier n'est pas un classeur Excel .xlsx valide.");
    }
    const classeur = xml(fichiers, "xl/workbook.xml");
    if (!classeur) throw new Error("Classeur Excel illisible (xl/workbook.xml absent).");
    const pr = classeur.getElementsByTagNameNS(NS, "workbookPr")[0];
    const date1904 = !!pr && ["1", "true"].includes(pr.getAttribute("date1904"));

    // Première feuille → fichier via les relations
    const feuille = classeur.getElementsByTagNameNS(NS, "sheet")[0];
    let chemin = "xl/worksheets/sheet1.xml";
    const rels = xml(fichiers, "xl/_rels/workbook.xml.rels");
    if (feuille && rels) {
      const id = feuille.getAttributeNS(NS_REL, "id") || feuille.getAttribute("r:id");
      const rel = Array.from(rels.getElementsByTagName("Relationship"))
        .find(r => r.getAttribute("Id") === id);
      if (rel) {
        const cible = rel.getAttribute("Target");
        chemin = cible.startsWith("/") ? cible.slice(1) : "xl/" + cible;
      }
    }

    const sst = xml(fichiers, "xl/sharedStrings.xml");
    const partages = sst ? Array.from(sst.getElementsByTagNameNS(NS, "si")).map(texte) : [];

    const doc = xml(fichiers, chemin);
    if (!doc) throw new Error("Feuille Excel introuvable dans le classeur.");
    const lignes = [];
    for (const row of doc.getElementsByTagNameNS(NS, "row")) {
      const valeurs = [];
      for (const c of row.getElementsByTagNameNS(NS, "c")) {
        const type = c.getAttribute("t");
        const v = c.getElementsByTagNameNS(NS, "v")[0];
        let val = null;
        if (type === "s") val = v ? partages[Number(v.textContent)] : null;
        else if (type === "inlineStr") val = texte(c);
        else if (type === "str" || type === "e") val = v ? v.textContent : null;
        else if (type === "b") val = v ? v.textContent === "1" : null;
        else if (v && v.textContent !== "") val = Number(v.textContent);
        valeurs[colonne(c.getAttribute("r") || "A")] = val;
      }
      lignes[Number(row.getAttribute("r")) - 1 || lignes.length] = Array.from(valeurs, x => x ?? null);
    }
    return { lignes: Array.from(lignes, l => l || []), date1904 };
  }

  return { lire };
})();
