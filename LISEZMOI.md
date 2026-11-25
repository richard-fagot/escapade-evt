# Générateur d'affiche mensuelle — L'Escapade

Page web qui génère l'affiche des activités du mois : on glisse-dépose le
fichier Excel des événements et le PDF se télécharge aussitôt (avec un aperçu
dans la page). Tout se passe dans le navigateur : aucune donnée n'est envoyée
sur internet, et rien n'est à installer.

## Utilisation

La page doit être servie par un serveur web : en ouvrant directement
`index.html` par double-clic, le navigateur refuse de lire `config.yaml`, le
fond et les polices.

- **En ligne (recommandé)** : activer GitHub Pages sur le dépôt
  (Settings → Pages → branche principale, dossier racine) puis ouvrir l'adresse
  donnée par GitHub.
- **En local** : dans le dossier du projet, lancer

  ```bash
  python -m http.server 8000
  ```

  puis ouvrir http://localhost:8000.

Ensuite, glisser le fichier `.xlsx` sur la zone (ou cliquer dessus pour le
choisir). Le fichier `affiche_<mois>_<année>.pdf` est téléchargé.

## Organisation des fichiers

| Fichier / dossier | Rôle |
|---|---|
| `index.html`, `style.css` | la page |
| `js/app.js` | glisser-déposer, téléchargement, aperçu |
| `js/affiche.js` | lecture des événements, mise en page, dessin du PDF |
| `js/excel.js` | lecture du fichier `.xlsx` |
| `config.yaml` | **tous les paramètres** : textes, couleurs, tailles, marges, polices |
| `fond.png` | fond d'affiche plein format A4 (logos et badge compris) |
| `polices/` | fichiers de polices TTF (+ leurs licences) |
| `lib/` | bibliothèques JavaScript, copiées localement (pdf-lib, fontkit, js-yaml, fflate) |

`config.yaml` est relu à chaque génération : après une modification, il suffit
de redéposer le fichier Excel (pensez à recharger la page si le navigateur
garde l'ancienne version en cache).

## Fichier Excel

Colonnes : `Date | Evenement | Debut | Fin | Prix` (première ligne = en-têtes,
première feuille du classeur).

- Une ligne par événement ; les lignes de même date sont regroupées dans une case.
- `Date` : date Excel, ou texte `12/06/2026`.
- `Debut`/`Fin` : heures Excel (`14:00`) ou texte libre (`"14H"`). Une fin à
  `00:00` s'affiche `MINUIT`.
- `Prix` : nombre (10 → `10€`, 7,5 → `7,50€`), texte libre ou vide.
- L'ordre des lignes n'importe pas, tout est trié par date.
- Le mois de la pastille est déduit automatiquement des dates.

## Polices

Les polices de l'affiche d'origine (Cy Grotesk Grand Dark et Futura Condensed)
sont commerciales : elles ne peuvent pas être redistribuées dans ce dépôt. Le
projet utilise à la place des polices libres (licence SIL OFL, voir
`polices/*-OFL.txt`), générées en graisses statiques depuis Google Fonts :

- **Archivo Expanded** (Regular, ExtraBold, Black) à la place de Cy Grotesk ;
- **Jost** (Medium, SemiBold), une Futura libre, à la place de Futura.

Pour utiliser une autre police, y compris Cy Grotesk/Futura si vous en avez
la licence : copier son fichier `.ttf` ou `.otf` dans `polices/` et indiquer
son chemin dans `config.yaml` (`polices → <catégorie> → fichier`).

Pour regénérer une graisse depuis une police variable :

```bash
pip install fonttools
fonttools varLib.instancer "Archivo[wdth,wght].ttf" wdth=125 wght=900 --update-name-table -o polices/Archivo-Expanded-Black.ttf
```

## Configuration (config.yaml)

- `affiche` : titre, sous-titre et lignes du pied de page.
- `page` : format, fond (`fond_image`), résolution du fond embarqué
  (`fond_dpi`) et compression JPEG (`fond_qualite`, à retirer pour un fond
  avec transparence).
- `polices` : pour chaque catégorie (titre, sous-titre, mois, jour,
  événement, heure, prix, pied) : `fichier`, `taille_pt`, `style`
  (`normal`/`italique`, l'italique est simulé par inclinaison), `couleur` et
  `echelle_verticale` optionnelle (étirement des glyphes en hauteur,
  interlignes compris).
- `cases` : largeur, marges, arrondi, ombre, et deux réglages de colonnes :
  `ecart_min_mm` (espace minimal entre deux cases d'une colonne, qui sert
  aussi de décalage minimal de la colonne droite sous la première case
  gauche) et `ecart_max_mm` optionnel (plafonne l'écart de la colonne gauche
  quand le mois est peu chargé).
- `zone_cases` : bande verticale des colonnes. Les jours alternent gauche
  (J), droite (J+1)… La première case (gauche) est toujours la plus haute ;
  la colonne gauche est répartie régulièrement sur la bande ; chaque case de
  droite est centrée entre les deux cases gauches qu'elle sépare
  chronologiquement, et remonte juste ce qu'il faut quand le mois est chargé,
  sans jamais descendre sous `ecart_min_mm` ni empiéter sur le pied de page.
  Si tout ne tient pas sur une page, des pages supplémentaires sont ajoutées
  au PDF avec le même gabarit.
- `pastille_mois` : couleur et position de la pastille du mois.
- `decor` : formes de fond colorées (désactivées quand un fond image est utilisé).
- `images` : logos ou décorations PNG/JPG positionnés librement (voir
  l'exemple commenté en fin de fichier).

Un avertissement s'affiche sous la zone de dépôt si une case *seule* dépasse
la zone (police trop grande) ; sinon le contenu excédentaire passe sur la page
suivante.
