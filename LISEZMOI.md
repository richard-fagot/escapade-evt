# Générateur d'affiche mensuelle — L'Escapade

Ce programme permet de générer l'affiche des activités proposées par l'Escapade chaque moi.
Il suffit de donner un fichier excel qui contient la liste des événements et hop, ça génère l'affiche.


## Installation

```bash
pip install reportlab pyyaml openpyxl pillow fonttools
```

La police par défaut est **Baloo 2** (Google Fonts, gratuite, très proche de celle
de l'affiche d'origine). Pour l'installer sous Linux : télécharger le TTF variable,
générer les graisses statiques puis rafraîchir fontconfig :

```bash
curl -sL -o Baloo2.ttf "https://raw.githubusercontent.com/google/fonts/main/ofl/baloo2/Baloo2%5Bwght%5D.ttf"
for w in 400 600 700 800; do
  fonttools varLib.instancer Baloo2.ttf wght=$w --update-name-table -o ~/.fonts/Baloo2-$w.ttf -q
done
fc-cache -f ~/.fonts
```

Sous Windows : télécharger les 4 TTF (ou les générer avec les mêmes commandes
fonttools), puis clic droit → Installer sur chacun. Toute autre police installée
sur le système fonctionne : le script scanne directement les dossiers de polices
(C:\Windows\Fonts et les polices utilisateur sous Windows, /usr/share/fonts et
~/.fonts sous Linux) et retrouve la famille par son nom.

## Utilisation

```bash
python generer_affiche.py evenements.xlsx config.yaml -o affiche_juin_2026.pdf
```

Produit le PDF (généré directement avec reportlab, aucune bibliothèque native
requise) et le SVG correspondant (même nom, extension `.svg`), éditable dans
Inkscape si besoin d'une retouche ponctuelle.

## Exécutable autonome

```bash
pip install pyinstaller
pyinstaller --onefile --name affiche generer_affiche.py
```

Produit `dist/affiche.exe` (Windows) ou `dist/affiche` (Linux) — à construire
sur l'OS cible. L'exécutable se lance seul :

```
affiche.exe evenements.xlsx config.yaml -o affiche.pdf
```

`config.yaml`, `evenements.xlsx` et `fond.png` restent des fichiers externes à
côté de l'exécutable : c'est voulu, ce sont eux que l'on modifie chaque mois.
Les polices sont lues depuis celles installées sur le système.

## Fichier Excel

Colonnes : `Date | Evenement | Debut | Fin | Prix`

- Une ligne par événement ; les lignes de même date sont regroupées dans une case.
- `Debut`/`Fin` : heures Excel (`14:00`) ou texte libre (`"14H"`). Une fin à
  `00:00` s'affiche `MINUIT`.
- `Prix` : nombre (10 → `10€`, 7,5 → `7,50€`) ou vide.
- L'ordre des lignes n'importe pas, tout est trié par date.
- Le mois de la pastille est déduit automatiquement des dates.

## Fond d'affiche

Le fond fourni (`fond.png`, plein format A4 avec logos et badge intégrés) est
référencé par `page.fond_image` dans le config. Le script génère par-dessus :
titre, pastille du mois, cases et pied de page. `pastille_mois.marge_droite_mm`
et `y_mm` permettent de la caler à côté du badge. Pour changer de fond, il
suffit de remplacer le PNG (même ratio A4).

## Configuration (config.yaml)

- `polices` : famille, taille (pt), style (`normal`/`italique`), graisse
  (`normal`/`semibold`/`bold`/`extrabold`), couleur et `echelle_verticale`
  optionnelle (étirement des glyphes en hauteur, interlignes compris), pour
  chaque catégorie : titre, sous-titre, mois, jours, événements, heures, prix,
  pied de page.

Le drapeau `-v` affiche le fichier de police réellement utilisé pour chaque
catégorie — utile si le rendu semble étrange (variante Condensed prise en
repli, police variable installée à la place des graisses statiques…).
- `cases` : largeur, marges, arrondi, ombre, et deux réglages de colonnes :
  `ecart_min_mm` (espace minimal entre deux cases d'une colonne, qui sert
  aussi de décalage minimal de la colonne droite sous la première case
  gauche) et `ecart_max_mm` optionnel (plafonne l'écart de la colonne gauche
  quand le mois est peu chargé).
- `zone_cases` : bande verticale des colonnes. Les jours alternent gauche
  (J), droite (J+1)… La première case (gauche) est toujours la plus haute ;
  la colonne gauche est répartie régulièrement sur la bande ; chaque case de
  droite est centrée entre les deux cases gauches qu'elle sépare
  chronologiquement, et remonte de ce milieu juste ce qu'il faut quand le
  mois est chargé, sans jamais descendre sous `ecart_min_mm` ni empiéter sur
  le pied de page. Si tout ne tient pas sur une page, des pages
  supplémentaires sont créées automatiquement avec le même gabarit (fond,
  titre, pastille, pied) ; les SVG sont alors numérotés (`affiche.svg`,
  `affiche_2.svg`, …).
- `decor` : formes de fond colorées (désactivables avec `actif: false`).
- `images` : logos ou décorations PNG/JPG positionnés librement (voir l'exemple
  commenté en fin de fichier).

Un avertissement n'apparaît que si une case *seule* dépasse la zone (police
trop grande) ; sinon le contenu excédentaire passe sur la page suivante.
