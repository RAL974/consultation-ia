# Travaux parsers — devis reçus le 27/08/2026 (brief Claude Code)

Contexte : le pipeline Hermes « flux demandes d'achat » a classé les devis
reçus par mail et généré 10 comparatifs (voir CLAUDE.md, section « Flux
demandes d'achat »). Plusieurs devis n'ont pas été lus correctement. Ce
fichier liste précisément les anomalies, avec les PDF réels comme fixtures.

Règle d'or du projet (cf. CLAUDE.md) : aucune primitive n'invente une règle
de parsing — chacune reproduit exactement un mécanisme observé dans un vrai
PDF, verrouillé par `tests/test_parsers.py` + une fixture dans
`tests/fixtures/`. Ne pas corriger sur un seul exemple sans fixture.

Pour chaque anomalie : copier le PDF dans `tests/fixtures/`, étendre le
gabarit (`moteur/fournisseurs/<fournisseur>.py`) ou le détecteur
(`moteur/detecteur.py`), ajouter un test, puis VÉRIFIER en relançant
`py -3 main.py "<affaire>"` : attendu « PDF lu(s) X sur Y » avec 0 anomalie
et les lignes complètes (totaux exacts).

---

## 1. ISHOP Saint-Denis (consultations/URGENT - ISHOP Saint-Denis/devis/)

- `DEVIS N° 821409.pdf` — fournisseur DEM reconnu mais 0 article extrait.
- `e.s ISHOP Saint-Denis.pdf` et `e.s ISHOP Saint-Denis (2).pdf` — 109
  DISTRIBUTION reconnu mais 0 article ; message : « Total HT du PDF
  (120,00 €) != somme des lignes extraites (0,00 €) » → **3e variante de
  gabarit 109 jamais vue** (les 2 variantes connues sont dans dist109.py).
- Attendu : la consultation portait sur 100 m de câble 3G1,5 sans halogène
  (Besoin URGENT - ISHOP Saint-Denis.txt) — 1 ligne par devis, total exact.

## 2. Réglettes - Rico Carpaye (consultations/Réglettes - Rico Carpaye/devis/)

- `Luminaire Etanche PORTO 120LM-PL0120Sen.pdf` — **fournisseur non
  reconnu** (expéditeur : idriss@artdeco.re → ART DECO / artdeco.re, à
  ajouter au détecteur + parser).
- `Etanche LED Malaga  (1).pdf` — **non reconnu** (expéditeur DEM).
- `DEVIS N  617004507 ELECTRICITE SERVICES REUNION .pdf` — **non reconnu**
  (expéditeur : idriss@artdeco.re ; entité ELECTRICITE SERVICES REUNION).
- `reglette detecteur.pdf` — **non reconnu** (expéditeur COREDIME ; gabarit
  Coredime sans en-tête de tableau ?).
- `DEVIS N° 821416.pdf` — DEM reconnu mais 0 article.
- `e.s Réglettes - Rico Carpaye.pdf` — 109 reconnu mais 0 article.
- Attendu : 14 réglettes LED 36 W 4000 K avec détecteurs intégrés
  (Besoin Réglettes - Rico Carpaye.txt).

## 3. BT - Floe (consultations/BT - Floe/devis/)

- `D1109436.pdf` — ELECTRIC PLUS (GMR) reconnu mais 0 article extrait.
  (Le même gabarit passe pour d'autres PDF GMR : comparer avec D1109458.pdf
  qui fonctionne — voir DICOM.)

## 4. R2V 3G1.5 - Rico Carpaye (consultations/R2V 3G1.5 - Rico Carpaye/devis/)

- `D1109369.pdf` — ELECTRIC PLUS reconnu mais 0 article extrait.

## 5. Non bloquant — Luminaires - Lagourgue

- `FT-STELLAR.pdf` — fiche technique (pas un devis), ignorée à juste titre.
  Rien à faire, sauf si l'on veut classifier les FT séparément.

---

## Rappels utiles pour les tests

- Les PDF réels restent dans `consultations/<affaire>/devis/` — copier dans
  `tests/fixtures/` avec un nom explicite (ex. `devis_dem_ishop_sans_halo.pdf`,
  `devis_109_ishop_3e_variante.pdf`, `devis_artdeco_porto.pdf`).
- Fournisseurs identifiés depuis les expéditeurs : Louis Maryse = RAVATE
  (maryse.louis@ravate.com) ; idriss = ART DECO (idriss@artdeco.re) ;
  COLTRAT Jimmy = COREDIME/Sonepar (jimmy.coltrat@sonepar.fr) ; Jeanne
  BRENNUS = ELECTRICITE SERVICES REUNION (ESR) ; Audrey CHAN-KUI = STAND 64.
- Interpréteur : toujours `py -3` (jamais `python` seul), voir CLAUDE.md.
- Après correction : relancer `py -3 main.py "<affaire>"` et vérifier le
  résumé (PDF lus, lignes, anomalies) + `py -3 -m pytest tests/ -q`.
