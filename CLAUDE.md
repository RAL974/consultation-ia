# Consultation AI

Outil Python local qui automatise la consultation fournisseurs d'Électricité
Services Réunion : besoin chantier → lecture des devis PDF → comparatif →
choix du mieux-disant. L'utilisateur est acheteur, pas développeur : toute
session Claude Code sur ce projet fait le travail technique elle-même,
explique en clair, et ne demande jamais à l'utilisateur d'écrire du code.

## Interpréteur Python réel

Le projet auto-installe ses dépendances (`moteur/dependances.py`) dans
l'interpréteur qui l'exécute. **Toujours lancer avec `py -3`, jamais juste
`python`**, sauf si tu as vérifié que ça pointe au même endroit : `python`
seul peut pointer vers un tout autre interpréteur, sans les dépendances du
projet (source de confusion historique, voir les notes de version dans
`installer.py`). `lancer_gui.bat` applique déjà cette règle (il préfère
`py -3`, se rabat sur `python` seulement si `py` est absent) et journalise
l'interpréteur réellement utilisé dans `lancement.log` à chaque lancement —
c'est le fichier à regarder pour savoir quel `python.exe` exact tourne sur
UN poste donné (il diffère d'un poste à l'autre ; ne pas supposer un chemin
particulier).

## Lancement

- Ligne de commande, une consultation : `py -3 main.py <NomConsultation>`
  (traite `consultations/<NomConsultation>/`, écrit dans son sous-dossier
  `resultats/`). Sans argument, traite le seul dossier existant sous
  `consultations/` s'il n'y en a qu'un, sinon liste les choix possibles.
- Sans souris ni VS Code : double-clic sur `lancer_gui.bat` (ouvre `gui.py`,
  une interface Tkinter : on choisit le dossier de consultation, puis 2
  boutons, « Générer le comparatif » puis, une fois la décision prise dans
  Excel, « Générer le panier »).
- Audit de la base d'équivalences : `py -3 audit.py` → `resultats/Audit_BDD.xlsx`.

## Architecture

```
main.py                    point d'entrée CLI (dossier de consultation en argument, ou auto si un seul)
gui.py                     interface graphique (Tkinter), appelle moteur/pipeline.py
moteur/
  pipeline.py               orchestration commune à main.py et gui.py
  consultation.py            résout le "dossier de consultation" à traiter (consultations/<nom>/), voir section dédiée
  journal.py                 dédouble la sortie (print) vers resultats/journal_....log de la consultation, voir section dédiée
  besoin.py                 lecture du fichier besoin (.xlsx / .txt)
  lecture_pdf.py            lit les PDF de devis/, détecte le fournisseur, appelle le parser, autocontrôle — tolérant aux pannes (voir section dédiée)
  detecteur.py              texte du PDF -> nom du fournisseur (regex)
  parsers.py                registre : découvre automatiquement les parsers de fournisseurs/
  fournisseurs/
    _gabarit.py              MOTEUR partagé (scan_ancre, scan_regex, table disponibilités)
    <fournisseur>.py          GABARIT + appel au moteur (ou logique procédurale si irréductible)
  autocontrole.py            qté × prix_net ≈ montant -> signale les lignes suspectes
  normalisation.py           code interne (câbles, ampoules, disjoncteurs...) + conversion prix -> unité de base
  base.py                    base d'équivalences SQLite (base/consultation.db, reconstruite depuis base/BDD_articles.csv) — repli historique du comparateur
  referentiel.py             référentiel articles SQLite (moteur/articles.db) — clé normalisée + alias appris + composés, voir section dédiée
  comparateur.py             rapproche les offres entre fournisseurs (référentiel en priorité, base.py en repli), consolide par cœur numérique
  excel.py                   génère Comparatif.xlsx (document de décision, voir bandeau du fichier)
  panier.py                  génère resultats/Panier_<chantier>_<date>.xlsx à partir des décisions de l'acheteur, voir section dédiée
  audit.py                   génère Audit_BDD.xlsx (qualité de la base d'équivalences)
consultations/
  <NomConsultation>/          une consultation = un dossier autonome et rejouable, voir section dédiée
    Besoin ....txt|xlsx        fichier besoin (facultatif)
    devis/*.pdf                 devis PDF de cette consultation
    resultats/                  Comparatif/Panier générés + journal_....log (créé au premier traitement)
base/
  BDD_articles.csv           LA base articles achats (source de vérité, ~10 000 réf.) : Référence;Désignation;Fournisseur;Fabricant;Catégorie;Tarif approximatif;CONCAT;Clé_Réf;... — lue à la fois par base.py et referentiel.py
  familles.csv                équivalences manuelles (Famille;Référence), facultatif
referentiel/
  composes.csv                composés manuels (Cle_besoin;Membre;Quantite), facultatif, vide par défaut
  A_confirmer.xlsx            généré à chaque exécution : rapprochements proposés en attente de décision (voir section dédiée)
  exports/                    généré à chaque exécution : graine CSV du futur référentiel Appro-Tracker
1.3.0.1. Suivi commandes - <année>.xlsx   fichier de l'acheteur (pas du dépôt, gitignore), export de la feuille "Commandes" utilisé par panier.py pour caler colonnes/fournisseurs/chantiers — voir section dédiée
installer.py                bundle de distribution autonome (copie base64 du projet) — outil de déploiement, PAS le moteur ; devient obsolète à chaque évolution du code, à régénérer séparément si besoin (hors périmètre courant)
README.md                   prise en main non-développeur (installer, utiliser, légende des couleurs, journal en cas de souci)
requirements.txt / requirements-dev.txt   filet standard en plus de l'auto-installation de moteur/dependances.py
tests/
  fixtures/                   VRAIS PDF de devis (jamais de PDF synthétique)
  test_parsers.py              Ravate/Coredime/DEM/109 Distribution/Electric Plus/Cominter
  test_parsers_edoi.py, test_parsers_cominter_mayotte.py,
  test_parsers_clareo_importer_stand64.py, test_parsers_imporelec.py
  test_detecteur.py, test_comparateur.py, test_autocontrole.py, test_referentiel.py
  test_besoin.py, test_normalisation.py, test_panier.py, test_consultation.py, test_lecture_pdf.py
```

## Dossiers de consultation (moteur/consultation.py)

Une consultation = un dossier autonome sous `consultations/<nom>/` (nom
libre, ex. `2026-08_Doujani` ou juste `Doujani`) : le fichier besoin
(facultatif) directement dedans, un sous-dossier `devis/` avec les PDF, et
un sous-dossier `resultats/` (Comparatif/Panier générés + journal). Chaque
consultation reste ainsi rangée et rejouable indépendamment des autres —
remplace l'ancien schéma `besoins/` + `devis/<Chantier>/` +
`resultats/` partagé à la racine (migré cette session, voir git log).

`resoudre_dossier_consultation(dossier_projet, cible=None)` : `cible` peut
être un nom de sous-dossier de `consultations/`, un chemin direct, ou
`None` (auto-sélectionne s'il n'y a qu'un seul dossier ; sinon lève
`ConsultationIntrouvable` avec la liste des dossiers disponibles — jamais
de choix deviné au hasard). Utilisé par `main.py` (argument CLI optionnel)
et `gui.py` (sélecteur de dossier unique).

**`dossier_projet` reste TOUJOURS la racine du projet** (contient `base/`,
`moteur/`, `referentiel/`) dans `generer_comparatif()`/`generer_panier()` —
jamais le dossier de consultation. C'est `dossier_resultats` (paramètre
optionnel de ces deux fonctions, et de `exporter_comparatif()`) qui pointe
vers `consultations/<nom>/resultats/` ; par défaut (non fourni)
`dossier_projet/"resultats"`, conservé pour compatibilité et verrouillé par
`tests/test_panier.py::test_pipeline_complet_besoin_vers_panier`.

## Tolérance aux pannes (moteur/lecture_pdf.py)

Un PDF illisible (corrompu, protégé par mot de passe...) ou une erreur du
parser sur UN devis ne doit jamais interrompre le traitement des autres.
`analyser_devis()` entoure chaque étape (ouverture, détection+parsing) d'un
`try/except` local par PDF, journalise une raison en clair, et passe au
suivant. Un rapport de synthèse est imprimé en fin de lecture (X PDF lus
sur N, Y lignes extraites, Z anomalies avec raison par fichier) — voir
`tests/test_lecture_pdf.py`.

## Journal (moteur/journal.py)

Toute la sortie du moteur passe par `print()` (voir bandeau de
`moteur/pipeline.py`) ; `capturer_journal()` dédouble `sys.stdout` le temps
d'une génération (comparatif ou panier), vers sa destination habituelle
(terminal ou file d'attente du GUI) ET vers
`<dossier_resultats>/journal_<comparatif|panier>_<AAAAMMJJ_HHMMSS>.log` —
c'est le fichier à envoyer en cas de problème (voir README.md). Écrit même
si une exception remonte (bloc `finally`).

## Comment ajouter un fournisseur

1. **Récupère un vrai PDF de devis** de ce fournisseur — jamais de règle de
   parsing écrite sans PDF réel à l'appui (règle d'or, voir plus bas).
2. Copie-le dans `tests/fixtures/<fournisseur>.pdf`.
3. Crée `moteur/fournisseurs/<fournisseur>.py` : un bandeau `# --- GABARIT ---`
   avec les constantes (marqueur d'ancrage, regex, offsets), puis soit
   `scan_ancre()`/`scan_regex()` de `_gabarit.py` si la structure s'y prête
   (un marqueur répété + champs à position fixe, ou un regex par ligne), soit
   une boucle explicite si la structure est vraiment procédurale (comme
   Cominter ou 109 Distribution — documenté en tête de fichier, voir
   `_gabarit.py`). Termine par :
   ```python
   FOURNISSEURS = ["NOM_DETECTE"]
   parse = parse_<fournisseur>
   ```
   (découverte automatique par `moteur/parsers.py`, aucune autre modification
   nécessaire).
4. Ajoute le motif de détection dans `moteur/detecteur.py` si besoin.
5. Écris le test dans `tests/test_parsers.py` : lis le PDF fixture, exécute
   le parser, **fige** la sortie complète (liste d'`Article`, champ par
   champ) — c'est ce test qui protège toute refonte future.
6. Lance `py -3 -m pytest` : tout doit être vert.
7. **Valide sur le PDF ouvert à côté** : chaque champ extrait doit
   correspondre à ce que tu vois dans le vrai document.

## Référentiel articles (moteur/referentiel.py, moteur/articles.db)

Rôle : faire tomber sur la même ligne de comparatif des références qui
désignent le même article mais s'écrivent différemment selon le fournisseur
(411651 / L411651 / LEG411651...). S'appuie sur le référentiel achats réel
de l'entreprise (`base/BDD_articles.csv`, colonnes CONCAT / Clé_Réf), pas sur
une logique de normalisation inventée. Complète `moteur/base.py`, qui reste
en place comme repli (le comparateur essaie d'abord le référentiel ; s'il ne
sait pas répondre avec certitude, il retombe sur `base.groupe()` comme
avant — rien de l'existant n'a été retiré).

**Préfixes marque déduits, pas codés en dur** (`deduire_prefixes()`, à
partir des couples réels CONCAT/Clé_Réf/Fabricant) : sur les données de ce
projet, seuls 4 fabricants sur ~300 utilisent un préfixe — Legrand → `LEG`,
L'Ebénoïd → `EBE`, Planet Wattohm → `PW`, Legrand - Bticino → `BT`. Tous les
autres (Schneider Electric, BLM, Finsecur, Philips...) n'en ont aucun.
Aucun SUFFIXE marque (ex. `411651-LEG`) n'a été observé dans les données
réelles. **Attention** : le cas classique "411651 / L411651 / LEG411651 /
411651-LEG" était déjà résolu par `moteur/base.coeur_numerique()` (extrait
les chiffres, peu importe où est le préfixe) — la vraie valeur ajoutée du
référentiel est ailleurs : les références à cœur NON numérique
(`MELV429338T`...), et surtout l'apprentissage/la confirmation.

**Trois statuts de résolution** (`Referentiel.resoudre()`) :
- `"connu"` — alias exact (importé de la BDD, ou confirmé par l'acheteur) :
  utilisé directement comme clé de regroupement.
- `"propose"` — un candidat plausible existe (préfixe marque retiré, ou
  cœur numérique, + désignation suffisamment proche) mais N'EST PAS fusionné
  automatiquement : il est écrit dans `referentiel/A_confirmer.xlsx` en fin
  d'exécution.
- `"nouveau"` — rien trouvé : la référence devient sa propre clé, ajoutée au
  référentiel (source='devis'), pas besoin de confirmation.

**Workflow de confirmation, pour l'acheteur** (fichier Excel aller-retour —
pas de question en console, le mode GUI n'a pas de console) :
1. Après une génération, s'il existe des rapprochements proposés, ouvre
   `referentiel/A_confirmer.xlsx`.
2. Pour chaque ligne, compare "Référence détectée"/"Désignation détectée"
   (ce qui vient du devis) à "Clé proposée"/"Désignation de la clé proposée"
   (ce que le référentiel pense que c'est), PDF ouvert à côté si besoin.
3. Remplis la colonne Décision : `OUI` (confirme le rapprochement), `NON`
   (rejette — la référence restera "nouveau" et ne sera plus reproposée),
   ou colle directement la bonne clé si le référentiel s'est trompé.
4. Enregistre le fichier. **À la PROCHAINE génération**, ces décisions sont
   appliquées avant de comparer les nouveaux devis (pas besoin de relancer
   tout de suite — le fichier est relu au début de chaque exécution).

**Composés** (`referentiel/composes.csv`, facultatif, VIDE par défaut —
aucune donnée de composé n'existe dans la BDD achats source, ce fichier
n'est peuplé qu'à la main par l'acheteur) : une ligne de besoin qui s'y
trouve déclarée se décompose automatiquement en plusieurs lignes du
comparatif (ex. un "coffret" = coffret nu + porte), chacune avec sa propre
meilleure offre.

**Export d'apprentissage** (`referentiel/exports/*.csv`, régénéré à chaque
exécution, format documenté en tête de chaque fichier) — graine du futur
référentiel Appro-Tracker : `aliases_appris.csv` (variantes de référence
apprises en consultation), `composes.csv` (composés manuels + appris),
`articles_nouveaux.csv` (références rencontrées en devis absentes de la
BDD achats).

## Panier de commande (moteur/panier.py)

Étape 4, après décision de l'acheteur sur le Comparatif.xlsx (étape 3).

**Fournisseur retenu = texte libre, jamais une formule.** Dans
Comparatif.xlsx, cette colonne est pré-remplie au mieux-disant mais reste
modifiable à la main dans Excel (liste déroulante limitée aux fournisseurs
ayant une offre sur la ligne). Les colonnes "Réf. fournisseur"/"N°
devis"/"Disponibilité"/"Montant retenu" sont des formules INDEX/MATCH qui
suivent ce choix (pour que l'acheteur voie tout de suite, dans Excel, à
quoi correspond son choix) — mais `panier.py` ne les relit JAMAIS : une
formule ne se recalcule que si le fichier a été rouvert dans Excel, alors
que "valider tel quel" (sans jamais ouvrir le fichier) doit aussi
fonctionner. `panier.py` relit uniquement le TEXTE de "Fournisseur retenu"
et reconstruit l'offre réelle en repartant du besoin + des devis (mêmes
fichiers qu'à la génération du comparatif — recalculés depuis le même
dossier de consultation, c'est pourquoi le GUI ne redemande qu'une seule
fois ce dossier, réutilisé pour les deux boutons).

**Colonnes calées sur le vrai Suivi commandes.** `trouver_fichier_suivi()`
cherche un fichier `*Suivi commandes*.xlsx` à la RACINE du projet (celui de
l'acheteur, pas un fichier du dépôt — nom horodaté par année, ex.
`1.3.0.1. Suivi commandes - 2026.xlsx`, à remplacer par le prochain export
quand l'année change) et lit l'ordre EXACT des en-têtes de sa feuille
"Commandes" : le Panier reprend ces mêmes en-têtes, dans le même ordre,
pour un copier-coller direct. Seules les colonnes que Consultation AI peut
remplir avec certitude sont renseignées (Référence, Désignation, Qté
commandée, Tarif convenu, Devis associé, Chantier, Fournisseur, Type de
commande = "Chantier") ; le reste (mode de livraison, statut, dates,
facturation...) appartient au suivi de la commande APRÈS passation et
reste vide. Si aucun export n'est trouvé, un jeu de 8 colonnes de repli est
utilisé (le panier reste utilisable, juste pas calé sur l'ordre du Suivi).

**Nom de fournisseur canonique** (`MAPPING_FOURNISSEURS` dans panier.py) :
les noms internes de détection (`RAVATE`, `SAGEES`, `109 DISTRIBUTION`...)
ne sont PAS toujours ceux de la liste "Fournisseurs" du Suivi (ex. `RAVATE`
→ `RAVATE ELEC`, `SAGEES` → `Sagees`, `ELECTRIC PLUS` → `GMR`). Un
fournisseur absent du mapping est recopié tel quel et signalé (jamais un
nom inventé) — `tests/test_panier.py` verrouille que tout fournisseur de
`moteur/detecteur.py` (sauf LEGRAND, non pertinent) y a une entrée.

**Code chantier** : le nom déduit du fichier besoin (ex. "Doujani") est
recherché dans la liste des chantiers du Suivi ("Listes Paramètres",
colonne Chantier, ex. "137 ZAC Doujani"). Si la recherche est AMBIGÜE
(ex. "Doujani" trouve à la fois "137 ZAC Doujani" et "155 Ecole Doujani",
cas réel) ou ne trouve rien, le nom brut est gardé tel quel et signalé —
jamais de code deviné au hasard.

**Rien perdu en silence** : les lignes de besoin sans offre, ou dont le
"Fournisseur retenu" est vide/ne correspond à aucune offre (faute de
frappe), n'entrent pas dans le panier commandable mais sont listées dans
l'onglet "Non commandées" avec le motif.

## Tests

    py -3 -m pytest          # tout le socle
    py -3 -m pytest -v       # détail test par test

Tous les tests de parsers tournent sur de VRAIS PDF (`tests/fixtures/`), pas
sur du texte inventé — c'est ce qui les rend dignes de confiance.

## Règle d'or

**Jamais de règle de parsing (regex, offset, mot-clé) inventée sans un vrai
PDF à l'appui.** Un format supposé "logique" peut très bien ne pas
correspondre au PDF réel du fournisseur (voir 109 Distribution ci-dessous).
L'acheteur valide chaque extraction sur le PDF ouvert à côté avant de
considérer un parser fiable.

## Fournisseurs couverts (verrouillés par un vrai PDF + tests)

| Fournisseur | Gabarit | Statut | Particularité apprise |
|---|---|---|---|
| RAVATE | `_gabarit.scan_ancre` | ✅ couvert | Toujours la Réf. FNR, jamais la réf interne |
| RAVATE PRO | délègue à `ravate.py` | ✅ couvert (confirmé par l'acheteur) | Construction identique à Ravate Elec (confirmé directement, pas de PDF Ravate Pro dédié disponible) |
| COREDIME | `_gabarit.scan_regex` | ✅ couvert | Garde-fou qté×prix propre en plus de l'autocontrôle global |
| DEM | `_gabarit.scan_regex` | ✅ couvert | Prix affichés AU CENT (/C) ; prix_net dérivé de montant/qté |
| ELECTRIC PLUS (alias GMR) | `_gabarit.scan_ancre` | ✅ couvert | "GMR" = marque publique du canal Electric Plus, même gabarit |
| 109 DISTRIBUTION | `_gabarit.scan_ancre`, 2 variantes essayées | ✅ couvert (corrigé cette session) | **2 structures réelles différentes** chez ce fournisseur (réf avant ou après le bloc chiffré) — voir "Points fragiles" |
| COMINTER | procédural (2 formats v1/v2) | ✅ couvert (v1 confirmé réel ; v2 toujours non confronté à un PDF réel) | — |
| COMINTER MAYOTTE | procédural, module dédié `cominter_mayotte.py` | ✅ couvert | **Entité et gabarit DIFFÉRENTS de Cominter Réunion** (répond à la question posée en début de session) — voir "Points fragiles" |
| EDOI | `_gabarit.scan_regex` | ✅ couvert (3 PDF réels) | Basé à Mamoudzou (Sonepar) ; dispo = "DISPO" ou délai en semaines ("12sem") |
| CLAREO | code préexistant, confirmé cette session | ✅ couvert (1 PDF réel, 27 lignes, total exact) | — |
| IMPORTER | code préexistant, confirmé cette session | ✅ couvert (1 PDF réel, 5 lignes, total exact) | — |
| STAND 64 | code préexistant, confirmé cette session | ✅ couvert (1 PDF réel, 33 lignes, total exact) | — |
| IMPORELEC | code préexistant, confirmé cette session | ✅ couvert avec un bug connu (1 PDF réel, 195 lignes) | Voir "Points fragiles" — sous-totaux intermédiaires du PDF mal gérés sur 6 lignes/195 |
| TELENCO | `_gabarit`-style procédural, module dédié | ✅ couvert (3 PDF réels, totaux exacts) | Référence = entier SANS virgule (se distingue des champs chiffrés qui ont toujours 2 décimales) |
| COMPTOIR DU CABLING | procédural, module dédié | ✅ couvert (4 PDF réels ; 3/4 totaux exacts) | Structure confirmée avec un 2e PDF plus simple après un 1er essai abandonné — voir "Points fragiles" pour le PDF encore incomplet |
| SAGEES | procédural, module dédié, **1 format sur 3 couvert** | ⚠️ partiellement couvert (3 PDF réels sur le format "V0", totaux exacts) | **Fournisseur découvert cette session** (absent de la liste initiale), ajouté à la demande de l'acheteur. Pas de colonne Référence sur ce format. 2 autres formats réels identifiés mais NON couverts — voir "Points fragiles" |
| LEGRAND | code préexistant, jamais vérifié | ❌ non couvert (probablement non pertinent) | **Précision de l'acheteur** : "Legrand" n'est pas un fournisseur distinct — certains distributeurs transmettent tel quel, sans reformatage, le devis que Legrand leur a fourni. Le motif de détection existant restera donc rarement déclenché à raison ; pas de PDF dédié à chercher |

## Points fragiles connus

- **109 Distribution : DEUX structures réelles coexistent** chez ce même
  fournisseur (voir `moteur/fournisseurs/dist109.py`) : la référence vient
  tantôt AVANT le bloc chiffré ("Commande client n°..."), tantôt APRÈS
  ("Devis n°..."). Les deux sont essayées sur chaque bloc, celle dont la
  référence "a la bonne forme" est retenue. Sur le 1er PDF vu, 5 lignes/38
  (câbles HO7VU 1.5mm²) ont un Total 2 % plus bas que Qté × P.U.Net affiché
  — jamais expliqué par une colonne visible. `prix_net` est donc calculé
  par Total/Qté (toujours exact), pas recopié du "P.U.Net" affiché (gardé
  à titre indicatif dans `prix_brut`).
- **Cominter Mayotte : 1 ligne/19 non extraite sur le PDF vu**
  (`moteur/fournisseurs/cominter_mayotte.py`) : le dernier article d'une
  page a ses valeurs numériques extraites AVANT sa référence dans le flux
  PyMuPDF (jamais vu ailleurs) — signalé "bloc incomplet" plutôt que
  rattaché au hasard (règle d'or : un seul exemple, pas de règle générale
  fiable à en tirer). Total du devis 446,35€ plus élevé que la somme
  extraite pour cette raison, vérifié à la main.
- **Imporelec : sous-totaux intermédiaires mal gérés** (bug PRÉEXISTANT,
  non corrigé cette session, voir bandeau de `moteur/fournisseurs/imporelec.py`) :
  ce PDF affiche périodiquement un sous-total cumulé au milieu du tableau ;
  le parser le prend à tort pour la Référence de l'article suivant (6
  lignes/195 touchées). Les montants restent corrects (donc le comparatif
  prix n'est pas faussé), seule la Référence de ces 6 lignes est
  inexploitable pour le rapprochement. Verrouillé tel quel par
  `tests/test_parsers_imporelec.py::test_parse_imporelec_bug_sous_total_connu`
  pour ne pas le perdre de vue.
- **Comptoir du Cabling : 1 PDF sur 4 encore incomplet**
  (`moteur/fournisseurs/comptoir_cabling.py`) : un premier essai (1 seul
  PDF) avait été abandonné (structure mal comprise, 30 lignes pour 335€ au
  lieu de 2 090€ réels). Avec 4 PDF réels, la structure fixe (Référence,
  TVA%, Montant, P.U., Qté, Désignation) a pu être confirmée — 3 PDF sur 4
  retombent exactement sur leur "Total HT remisé". Le 4e
  (`comptoir_cabling_1.pdf`) reste incomplet de 55,51€/2 089,76€ : quelques
  références longues sont coupées sur 2 lignes par l'extraction PDF (ex.
  "PASSEFILME"/"T" au lieu de "PASSEFILMET"), non recollées (pas assez
  d'exemples pour une règle fiable, règle d'or).
- **Sagees : 2 formats réels sur 3 NON couverts.** Ce fournisseur envoie
  au moins 3 gabarits de devis différents selon le contexte/l'agence :
  1. Format "V0" (`moteur/fournisseurs/sagees.py`, COUVERT, 3 PDF réels,
     totaux exacts) : pas de colonne Référence, articles identifiés par
     désignation seule.
  2. Format avec colonne Référence + Conditionnement
     (`devis/Pour DB/ESPACE SOLEIL CLINIQUE MAYOTTE.pdf`,
     `tests/fixtures/sagees_4.pdf`) : NON couvert.
  3. Format "bordereau de prix" en lettre de couverture, prix suffixés "€"
     (`devis/Pour DB/ML20260615 ES.pdf`, `tests/fixtures/sagees_5.pdf`) :
     NON couvert.
  Les fixtures 4 et 5 sont déjà en place dans `tests/fixtures/` pour une
  prochaine session — pas de gabarit écrit dessus cette fois (temps de
  session, et éviter de bâcler un 2e/3e format après le format 1).
- **Coredime a son propre garde-fou qté×prix silencieux** (ligne écartée
  sans message si l'écart dépasse 5 % ou 1 €), maintenant redondant avec
  l'autocontrôle global (`moteur/autocontrole.py`, signale sans écarter).
  Laissé tel quel pour ne pas changer quelles lignes sont gardées ; à
  nettoyer/harmoniser dans une session dédiée si souhaité.
- **Fixtures réelles supplémentaires disponibles mais pas encore
  exploitées** (déposées dans `devis/Pour DB/` cette session) : 2 PDF
  Coredime, 2 PDF DEM, 3 PDF Electric Plus, 2 PDF Ravate de plus que les
  fixtures déjà verrouillées — utiles pour une future session de
  durcissement (plus de remises/reliquats/multi-page par fournisseur déjà
  couvert), pas indispensables (les gabarits sont déjà réputés fiables).
- `installer.py` embarque une copie base64 du projet à une date donnée : elle
  est déjà en retard sur le code actuel (normal, c'est un outil de
  déploiement figé, pas un miroir live) — à régénérer si une nouvelle
  installation autonome est nécessaire.
- **Deux bases SQLite coexistent** : `base/consultation.db` (moteur/base.py,
  historique) et `moteur/articles.db` (moteur/referentiel.py, nouveau). Le
  comparateur essaie le référentiel en premier, retombe sur `base.py` sinon
  — volontaire cette session (risque le plus bas), mais une fusion des deux
  mécanismes est envisageable plus tard une fois le référentiel éprouvé en
  usage réel.
- Le premier import de `moteur/articles.db` (création du fichier) prend
  ~3s pour les 9988 lignes ; les imports suivants (fichier déjà là, comme à
  chaque lancement normal de l'app) prennent ~0,15s — largement sous la
  contrainte de quelques secondes.
- **Bug corrigé cette session — désynchronisation référentiel selon l'ordre
  de résolution** (`moteur/referentiel.py Referentiel.resoudre()`) : une
  référence de devis inconnue de la BDD s'auto-enregistre comme alias vers
  elle-même (pour ne pas rechercher un candidat à chaque rencontre). Le
  souci : au sein d'une même exécution, l'ARTICLE (boucle articles de
  `comparer()`, traitée en premier) la voyait "nouveau" -> repli sur
  `base.groupe()`, puis la LIGNE DE BESOIN correspondante (boucle besoin,
  après) la voyait déjà "connu" (auto-alias) -> clé préfixée `"REF:"` par
  `comparateur._cle()`. Deux clés différentes pour la MÊME référence : la
  ligne retombait à tort en "AUCUNE OFFRE". Symptomatique surtout des
  références à moins de 4 chiffres (le filet de sécurité par cœur numérique
  de `comparer()` ne s'applique pas). Fix : un alias d'origine `'devis'`
  (auto-enregistré, pas confirmé) est désormais TOUJOURS rapporté comme
  `"nouveau"`, jamais `"connu"` — seuls les alias `'import'`/`'confirme'`
  restent prioritaires. Régression verrouillée par
  `tests/test_referentiel.py::test_comparateur_reference_inconnue_sans_coeur_numerique_matche_quand_meme`.
- **Bug corrigé cette session — ligne d'en-tête fantôme dans un besoin
  .txt** (`moteur/besoin.py`) : `consultations/Doujani/Besoin Doujani.txt`
  (fichier réel, alors dans `besoins/` avant la migration vers
  `consultations/`) commence par une ligne d'en-têtes collée depuis Excel
  ("Désignation⇥Référence⇥Qté") ; elle était lue comme une ligne de besoin
  à part entière (qté 0, sans référence, "AUCUNE OFFRE" dans le
  Comparatif). `_est_ligne_entete()` détecte et ignore désormais cette
  première ligne si elle ressemble à un en-tête. Voir `tests/test_besoin.py`.
- Le fichier `1.3.0.1. Suivi commandes - <année>.xlsx` change de nom chaque
  année (export renouvelé) : `panier.py` le retrouve par motif
  (`*Suivi commandes*.xlsx`, le plus récent si plusieurs), jamais par nom
  figé — rien à toucher au changement d'année, juste déposer le nouvel
  export à la racine du projet.
- **`devis/Pour DB/` et `resultats/` (racine) sont un reliquat volontaire.**
  `devis/Pour DB/` reste un dépôt de PDF bruts pas encore rattachés à une
  consultation (fixtures potentielles pour de futurs parsers, voir plus
  haut) — jamais scanné automatiquement par `main.py`/`gui.py`.
  `resultats/` (racine, vide depuis la migration vers `consultations/`)
  n'est plus utilisé qu'en repli technique (valeur par défaut de
  `dossier_resultats` quand une fonction du moteur est appelée sans en
  préciser un, ex. tests). Les deux sont dans `.gitignore`.

## Évolutions envisagées (pas commencées)

- **Ingestion des devis par email** : aujourd'hui l'acheteur dépose les PDF
  à la main dans `consultations/<nom>/devis/` ; une boîte mail dédiée (ou
  un dossier surveillé) pourrait alimenter ce dossier automatiquement.
- **Bascule vers Appro-Tracker** quand son module Commandes existera :
  `referentiel/exports/*.csv` (alias appris, composés, articles nouveaux)
  est déjà pensé comme graine pour cette migration — voir section
  "Référentiel articles".
- **Enrichissement continu du référentiel partagé** : aujourd'hui
  `moteur/articles.db` et les confirmations (`A_confirmer.xlsx`) sont
  propres à ce poste ; un référentiel partagé entre plusieurs postes
  (Xavier compris) permettrait de capitaliser les alias confirmés une
  seule fois pour tout le monde.
