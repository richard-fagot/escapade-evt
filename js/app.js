/* Interface : glisser-déposer du fichier Excel → génération et téléchargement du PDF. */
"use strict";

(() => {
  const zone = document.getElementById("zone");
  const entree = document.getElementById("fichier");
  const etat = document.getElementById("etat");
  const listeAvert = document.getElementById("avertissements");
  const actions = document.getElementById("actions");
  const lien = document.getElementById("telecharger");
  const apercu = document.getElementById("apercu");
  let urlPdf = null;

  function afficher(message, classe = "") {
    etat.textContent = message;
    etat.className = "etat " + classe;
  }

  function telecharger(url, nom) {
    const a = document.createElement("a");
    a.href = url; a.download = nom;
    document.body.appendChild(a);
    a.click();
    a.remove();
  }

  async function traiter(fichier) {
    if (!fichier) return;
    if (!/\.xlsx$/i.test(fichier.name)) {
      afficher(`« ${fichier.name} » n'est pas un fichier .xlsx.`, "erreur");
      return;
    }
    zone.classList.add("occupe");
    listeAvert.replaceChildren();
    afficher(`Génération de l'affiche à partir de « ${fichier.name} »…`);
    try {
      // config relue à chaque fois : une modification est prise en compte sans recharger la page
      const cfg = await Affiche.chargerConfig();
      const r = await Affiche.generer(await fichier.arrayBuffer(), cfg);
      if (urlPdf) URL.revokeObjectURL(urlPdf);
      urlPdf = URL.createObjectURL(new Blob([r.octets], { type: "application/pdf" }));
      telecharger(urlPdf, r.nomFichier);
      lien.href = urlPdf;
      lien.download = r.nomFichier;
      lien.textContent = `Télécharger à nouveau ${r.nomFichier}`;
      actions.hidden = false;
      apercu.src = urlPdf;
      apercu.hidden = false;
      afficher(`✓ ${r.nomFichier} généré (${r.resume}).`, "ok");
      for (const m of r.avertissements) {
        const li = document.createElement("li");
        li.textContent = "⚠ " + m;
        listeAvert.appendChild(li);
      }
    } catch (e) {
      console.error(e);
      afficher(e.message || String(e), "erreur");
    } finally {
      zone.classList.remove("occupe");
      entree.value = "";
    }
  }

  entree.addEventListener("change", () => traiter(entree.files[0]));
  zone.addEventListener("keydown", e => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); entree.click(); }
  });
  for (const ev of ["dragenter", "dragover"]) {
    zone.addEventListener(ev, e => { e.preventDefault(); zone.classList.add("survol"); });
  }
  for (const ev of ["dragleave", "drop"]) {
    zone.addEventListener(ev, () => zone.classList.remove("survol"));
  }
  zone.addEventListener("drop", e => {
    e.preventDefault();
    traiter(e.dataTransfer.files[0]);
  });
  // Évite que le navigateur ouvre le fichier si on le lâche à côté de la zone
  window.addEventListener("dragover", e => e.preventDefault());
  window.addEventListener("drop", e => e.preventDefault());
})();
