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
fnp.py                     CLI état FNP mensuel (`py -3 fnp.py <AAAA-MM>`), voir section dédiée
gui_fnp.py                 fenêtre GUI "État FNP du mois" (bouton 5 de gui.py)
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
  fnp.py                     état mensuel des Factures Non Parvenues (BL + transitaires), LECTURE SEULE, voir section dédiée
  fnp_brouillon.py            brouillon Outlook pour l'état FNP (Save() uniquement, jamais Send()), voir section dédiée
  rapprochement/
    ecriture.py               écriture sécurisée dans le Suivi commandes (patch XML chirurgical, jamais openpyxl.save() sur le classeur vivant), voir section dédiée
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
a_traiter/                  Rapprochement AI : dépôt des BL/Factures PDF à traiter (BL/, BL/Traités/, Factures/), gitignoré, voir section dédiée
rapports/                    Rapprochement AI : rapports de rapprochement générés + FNP_<AAAA-MM>.xlsx (état FNP mensuel), gitignoré
backups/                     Rapprochement AI : sauvegardes horodatées du Suivi commandes avant écriture, gitignoré
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
  test_rapprochement_ecriture.py  socle d'écriture sécurisée (Rapprochement AI), voir section dédiée
  test_fnp.py                  état FNP mensuel (volets BL + transitaires), voir section dédiée
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

## Rapprochement AI (moteur/rapprochement/) — branche en cours

Extension de Consultation AI, même dépôt : rapprocher automatiquement les BL
(quotidien) et les factures fournisseurs (PDF, hebdomadaire) contre le
classeur `1.3.0.1. Suivi commandes - <année>.xlsx` (feuille "Commandes",
~5 900 lignes). Objectif : résorber les lignes au statut
"🚚 Reçue (Attente Facture)" (**4 289 lignes constatées le 2026-08-11**,
~73 % de la feuille) sans jamais migrer le Suivi ailleurs — il reste le
pivot, on écrit dedans. Cible d'usage : dépôt des BL du jour dans
`a_traiter/BL/`, dépôt des factures de la semaine dans `a_traiter/Factures/`,
quelques confirmations, Suivi à jour.

**Session R1 (cadrage + socle d'écriture) : fait.** Session R2 : 109
Distribution, Coredime, Electric Plus/GMR et Cominter Ouest couverts et
recettés sur écriture réelle contre le vrai classeur vivant (voir
sous-sections dédiées ci-dessous, dans l'ordre où ils ont été traités).
Fournisseurs restants (Sagees, DEM, Ravatelec, TOP Océan Indien exclu —
ponctuel, cf. tableau de flux R1) : pas commencés, en attente de pièces
réelles (règle d'or : jamais de gabarit sans PDF réel).

**Session R2 — 109 Distribution, extraction + matching (fait) :**

- **OCR ajouté au projet** (`moteur/ocr.py`, dépendance `rapidocr-onnxruntime`
  dans `moteur/dependances.py`/`requirements.txt`) : choisi plutôt que
  Tesseract (binaire externe à installer à la main sur chaque poste, casse
  l'auto-installation) ou EasyOCR (trop lourd, PyTorch). Pip pur, ~100 Mo
  avec dépendances (numpy/opencv/onnxruntime), ~7s/page en régime établi
  sur ce poste (~25s au tout premier appel, chargement des modèles).
  `regrouper_lignes()`/`lignes_ocr()` reconstruisent, à partir des mots OCR
  positionnés, une structure en lignes/cellules réutilisable par les
  gabarits fournisseurs (même esprit que le texte PDF natif).
- **`moteur/detecteur.py`** : motif "109 DISTRIBUTION" élargi pour
  reconnaître aussi les 4 blocs d'agence d'en-tête ("109 Est/Sud/Ouest/
  Nord", toujours présents) — plus fiables à l'OCR que "109 DISTRIBUTION"
  qui ressort souvent déformé dans le pied de page légal du BL.
- **Nouveau module `moteur/rapprochement/`** : `modele_bl.py` (LigneBL,
  BonLivraison), `parsers_bl.py` (registre auto, même principe que
  `moteur/parsers.py` : un fournisseur expose `parse_bl` + `FOURNISSEURS`),
  `lecture_bl.py` (scan de dossier, tolérant aux pannes, même principe que
  `moteur/lecture_pdf.py`), `matching.py` (rapprochement BL <-> Suivi).
- **Gabarit BL 109 Distribution** (`moteur/fournisseurs/dist109.py`,
  `parse_bl_109`) : structure procédurale (pas `scan_ancre` à décalage
  fixe comme pour le devis du même fournisseur) car l'OCR "rate" parfois
  UNE cellule sans faire basculer la ligne pour autant, ce qui décale tous
  les offsets suivants si on suppose un nombre de cellules fixe par ligne
  (constaté réellement : la case à cocher "livré" imprimée dans la
  cellule Qté est parfois lue comme un chiffre collé au vrai nombre — "32"
  devient "327" — et parfois la cellule Qté disparaît carrément de l'OCR).
  Seules DEUX positions se sont révélées fiables sur CHAQUE ligne des 4
  vrais BL testés, relatives au code TVA (ancre, "C0".."C9" + "CO" car
  l'OCR confond 0/O) : le P.U.Net juste avant, le Total juste après. La
  quantité livrée est donc TOUJOURS déduite de Total / P.U.Net, jamais lue
  directement dans la cellule Qté — c'est ce qui a sauvé la ligne du 3e BL
  testé (aucune cellule Qté détectée du tout par l'OCR sur cette ligne).
  Autocontrôle Total HT (comme le devis du même fournisseur) : silencieux
  sur les 4 BL testés (aucun euro perdu).
- **Matching** (`moteur/rapprochement/matching.py`) : n° de commande lu
  sur le BL ("N°Réf.Client") -> lignes du Suivi pour ce fournisseur + cette
  commande (colonne "N° de commande", format identique, ex. "123.096").
  Appariement ligne à ligne par référence normalisée, en réutilisant
  `moteur.base.coeur_numerique` (déjà éprouvé côté devis) plutôt que
  d'inventer une nouvelle règle — nécessaire car le Suivi et le BL
  n'utilisent pas toujours EXACTEMENT la même graphie de référence (cas
  réels : "086101L" au Suivi vs "86101L" au BL, ou "EUR52302" vs "52302").
  Un seul candidat sûr requis ; 0 ou plusieurs -> bac "inconnu", jamais de
  choix au hasard. **Idempotence intégrée dès le matching** (pas une
  couche à part) : si la ligne Suivi a DÉJÀ la même quantité livrée, le
  même tarif BL et une date de livraison renseignée, la correspondance
  ressort "déjà à jour" et RIEN n'est proposé à l'écriture — double
  traitement du même BL détecté et neutralisé sans intervention.
- **Recette (partielle) faite cette session** : les 4 seuls vrais BL 109
  Distribution disponibles (`a_traiter/BL/`, copiés dans
  `tests/fixtures/bl_dist109_{1..4}.pdf`) ont été passés en lecture seule
  contre le VRAI Suivi commandes vivant (chemin ci-dessus). **Découverte
  importante** : les 7 lignes de ces 4 BL étaient TOUTES déjà saisies à la
  main par l'acheteur avant cette session (même qté, même tarif, statut
  déjà "🚚 Reçue"/"🟠 Partielle") — le matching les a TOUTES retrouvées
  correctement et TOUTES classées "déjà à jour", preuve concrète que le
  rapprochement ET l'idempotence fonctionnent sur données réelles. En
  revanche, cela veut dire qu'AUCUNE vraie écriture n'a encore pu être
  démontrée sur le classeur vivant faute de BL "frais" restants dans
  `a_traiter/BL/` pour 109 Distribution — à refaire dès qu'un nouveau BL
  109 Distribution réel sera déposé.
- **Interaction : bouton GUI** (choix de l'acheteur, pas la console) —
  `gui.py` a un 3e bouton « Rapprocher les BL » qui ouvre
  `gui_rapprochement.py` (`FenetreRapprochementBL`, `moteur/rapprochement/
  pipeline_bl.py` pour l'orchestration). Toujours en 2 temps : lecture
  seule d'abord (OCR + matching, rien n'est modifié, tourne dans un thread
  pour ne pas geler la fenêtre), listes à cocher (sûres cochées par
  défaut, à confirmer décochées) + listes en lecture seule (déjà à jour,
  inconnues, fichiers non traités) ; écriture seulement après clic sur le
  bouton dédié ET confirmation dans une boîte de dialogue récapitulant le
  nombre de lignes et le fichier cible. `appliquer_et_archiver()` écrit
  (verrou/sauvegarde/patch chirurgical du socle R1), puis archive dans
  `a_traiter/BL/Traités/` (renommé "<date> - <fournisseur> - <n° BL> - BC
  <n° commande>" — dossier créé par l'acheteur elle-même, pas par la
  session : elle y agrafe ensuite le BdC et le BL papier correspondants et
  archive dans les classeurs physiques) CHAQUE BL dont TOUTES les lignes sont
  résolues (écrites ou déjà à jour) — jamais un BL avec une ligne "à
  confirmer" non cochée ou "inconnue" encore dedans, qui reste dans
  `a_traiter/BL/` pour un prochain passage. Rapport texte horodaté dans
  `rapports/` à chaque écriture. Logique testée unitairement (sans OCR ni
  fichier réel) dans `tests/test_rapprochement_pipeline_bl.py` ; fenêtre
  vérifiée par un test de fumée manuel (ouverture, thread de lecture, pas
  de crash) — pas encore de clic réel "Écrire" recetté par l'acheteur sur
  le vrai classeur (voir point suivant).
- **Recette réelle du clic « Écrire » faite cette session (suite), sur 4
  vrais BL 109 Distribution FRAIS** (jamais saisis avant, déposés par
  l'acheteur dans `a_traiter/BL/`) : **succès**. Détail :
  - 4 lignes réellement nouvelles ont été écrites dans le VRAI classeur
    vivant (Qté livrée, Tarif BL, Date de livraison), vérifiées relues
    après coup — correctes.
  - 3 lignes ressorties "déjà à jour" à raison (déjà saisies par ailleurs)
    — pas de doublon écrit.
  - 1 ligne ("GOUJON8X75") ressortie "inconnue" à raison : elle ne
    correspond à AUCUNE ligne de la commande 131.157 dans le Suivi (donc
    probablement pas encore ajoutée côté achat) — laissée pour vérification
    manuelle, son BL n'a PAS été archivé (2 de ses 3 lignes ont bien été
    écrites, mais un BL n'est archivé que si TOUTES ses lignes sont
    résolues).
  - 7 BL au total déplacés vers `a_traiter/BL/Traités/` automatiquement
    (voir juste en dessous — CHANGEMENT DE DESTINATION D'ARCHIVAGE), 1
    laissé dans `a_traiter/BL/` (celui avec la ligne inconnue).
  - **Bug réel trouvé et corrigé pendant la recette** : `MOTIF_COMMANDE_BL`
    (`moteur/fournisseurs/dist109.py`) n'acceptait que le format "123.096"
    (2 groupes numériques) — un des 4 BL frais portait un n° de commande
    au format "M3.14.353" (préfixe lettre + 3 groupes, code chantier), pas
    reconnu du tout -> anomalie "N° de commande introuvable". Motif
    élargi (`^([A-Z]?\d{1,4}(?:[.\-]\d{1,4}){1,2})$`, capture le motif
    entier plutôt que 2 groupes fixes) ; fixture réelle ajoutée
    (`tests/fixtures/bl_dist109_5.pdf`,
    `test_parse_bl_dist109_5_numero_commande_code_chantier`). Après
    correction, cette ligne s'est appariée et écrite correctement (Suivi
    ligne 5731).
  - **IMPORTANT, bug trouvé AVANT la recette (donc jamais écrit dans le
    mauvais fichier, mais à retenir)** : `pipeline_bl.py` utilisait
    `moteur.panier.trouver_fichier_suivi()` (cherche à la racine du
    dépôt) pour localiser le Suivi à rapprocher — hors CLAUDE.md documente
    depuis la session R1 que ce fichier racine n'est qu'un export
    ponctuel PÉRIMÉ, pas le classeur vivant (vérifié : les deux fichiers
    diffèrent réellement, celui à la racine datait de la veille). Corrigé
    par `trouver_fichier_suivi_vivant()` (nouveau, dans `pipeline_bl.py`)
    qui cherche dans le dossier frère "1.3.0.1. Commandes courantes" (à
    côté du dossier projet), en excluant les copies manuelles ("copie
    ...xlsx") et le classeur en refonte ("Suivi nouveau...").
  - **Le verrou fonctionne comme prévu** : testé en conditions réelles
    (l'acheteur a rouvert puis refermé le Suivi dans Excel pendant cette
    session) — l'écriture attend bien que le fichier soit fermé.
- **Dossier d'archivage changé sur demande explicite de l'acheteur** :
  PAS `rapproches/<fournisseur>/<AAAA-MM>/` (imaginé au cadrage R1, jamais
  implémenté tel quel) mais `a_traiter/BL/Traités/` — un dossier qu'elle a
  créé elle-même à côté de `a_traiter/BL/`, pour y retrouver les BL
  numérisés déjà rapprochés, agrafer le BdC et le BL papier
  correspondants, et archiver dans les classeurs physiques. Toujours
  renommés "<date> - <fournisseur> - <n° BL> - BC <n° commande>". Retenir
  pour la suite : préférer un dossier que l'acheteur demande concrètement
  à un nom de dossier imaginé au cadrage, même documenté dans CLAUDE.md.
- **Cas "GOUJON8X75 inconnu" résolu (pas un bug)** : vérifié directement
  dans le Suivi, la ligne "GOUJON 8X70" existe bien pour la commande
  131.157, mais le BL réel affiche "GOUJON8X75" — un vrai écart de 5mm
  (70 vs 75), pas une erreur de lecture. Confirmé par l'acheteur :
  probablement une rupture chez 109 Distribution ou une substitution faite
  par le chargé de travaux sur place — volontairement jamais rapproché
  automatiquement (voir dossier "À vérifier" ci-dessous).
- **Procédure d'utilisation écrite dans README.md** (demandée explicitement
  par l'acheteur), section "Rapprocher les BL (bons de livraison)" — répond
  notamment à sa question : le Suivi doit être FERMÉ pour l'étape
  "Écrire", pas pour la lecture (l'outil le vérifie et refuse d'écrire
  sinon, testé en conditions réelles cette session).
- **Nouveau dossier `a_traiter/BL/À vérifier/`** (demande explicite de
  l'acheteur, faite pendant la recette Coredime) : tout BL LU mais avec au
  moins une ligne "inconnue" ou "à confirmer" non cochée (ex. le cas
  GOUJON8X75/GOUJON8X70 ci-dessus) y est déplacé automatiquement par
  `appliquer_et_archiver()` — jamais laissé mélangé avec les BL pas encore
  traités du tout, pour repérer d'un coup d'œil ceux qui attendent une
  décision humaine. Fichier gardé tel quel (pas renommé, contrairement à
  l'archivage dans `Traités/`). Testé (`test_appliquer_et_archiver_
  deplace_les_bl_a_confirmer_vers_a_verifier`).
- **Bug de robustesse trouvé et corrigé en conditions réelles** : un
  archivage qui échoue (fichier verrouillé par autre chose — antivirus,
  réseau, Explorer... constaté réellement sur `BL M2.16.011.pdf`, resté
  verrouillé sans raison identifiée après un `shutil.move` interrompu en
  plein milieu — copie faite, suppression de la source échouée) faisait
  planter TOUTE la fonction `appliquer_et_archiver()`, perdant le résumé
  de l'écriture Suivi (déjà réussie à ce stade !) et empêchant le rapport
  final. Corrigé : chaque déplacement (archivage ou "à vérifier") est
  maintenant entouré d'un `try/except OSError` individuel, collecté dans
  `resume["archivage_echoue"]`, sans jamais bloquer le reste du lot ni le
  rapport. Testé (`test_appliquer_et_archiver_une_archive_qui_echoue_
  nempeche_pas_le_reste`).

**Session R2 suite — Coredime, fournisseur n°2 (fait) :**

- **Piège signalé par l'acheteur AVANT tout code, décisif** : Coredime
  livre souvent en plusieurs fois ; un même BL peut lister les articles
  LIVRÉS et les articles NON livrés ("reste à livrer") — implémenté par
  prudence (`MOTIF_RESTE_A_LIVRER_COREDIME` dans
  `moteur/fournisseurs/coredime.py`, tout ce qui suit cette mention est
  exclu de la quantité livrée) mais **jamais rencontré sur un vrai PDF à
  ce jour** — à valider dès qu'un cas réel se présente. Ne pas confondre
  avec "Ref à livrer directement" (déjà vu, 1 vrai PDF) : ces articles-là
  SONT bien livrés, seul le prix est différé à la facture.
- **Ces BL n'affichent QUASIMENT JAMAIS de prix** (réglé à la facture) :
  pas d'autocontrôle Total HT possible comme pour 109 Distribution : la
  quantité livrée est le seul champ vraiment critique. `matching.py`
  corrigé en conséquence (voir plus bas) : l'idempotence ne doit plus
  dépendre d'une comparaison de tarif quand le BL n'en affiche aucun,
  sinon aucune ligne Coredime ne pourrait jamais ressortir "déjà à jour".
- **6 vrais BL testés** (`tests/fixtures/bl_coredime_1..6.pdf`,
  `tests/test_parsers_bl_coredime.py`), 3 bugs réels trouvés et corrigés
  en les confrontant : référence trop stricte pour un code comme
  "LBCLASTD02" (8 lettres + 2 chiffres — élargi à tout alphanumérique
  5-15 caractères avec au moins un chiffre, sans ponctuation, ce qui
  exclut naturellement les codes ECO-taxe et les bouts de désignation qui
  débordent) ; désignation débordant sur une 2e ligne AVANT ou APRÈS la
  cellule Qté selon les cas (raccord de cellules si la ligne courante n'a
  pas de quantité exploitable et que la suivante ne commence pas par une
  référence valide) ; quantité "1" lue "i" minuscule par l'OCR (remplacé
  en mot entier uniquement, `\bi\b`, pour ne pas abîmer "unite").
- **Format de n° de commande Coredime** : "BC 131.153" / "COMMANDE N°
  129.049" / "COMMANDE N° M2.16.011" (préfixe lettre + code chantier,
  comme le cas 109 Distribution trouvé juste avant) — et 1 cas réel où
  l'OCR a perdu le séparateur ("BC123097" au lieu de "BC 123.097"), repli
  dédié avec découpage 3+3.
- **Recette réelle sur 6 vrais BL Coredime (dont 4 déjà dans
  `a_traiter/BL/` avant la session)** : 4 lignes écrites pour de vrai dans
  le classeur vivant, 8 lignes déjà à jour (rien réécrit), 1 ligne
  confirmée par l'acheteur malgré une "sur-livraison" apparente (voir
  point suivant), 1 ligne inconnue (GOUJON8X75, cf. plus haut, déplacée
  vers "À vérifier").
- **Découverte importante confirmée par l'acheteur, à garder en tête pour
  la suite** : l'écart "qté commandée très inférieure à qté livrée" n'est
  PAS toujours une erreur — cas réel, Suivi "3" (boîtes) vs BL "300"
  (pièces, 3 boîtes de 100). Unités de vente (boîte) vs unités de
  commande/livraison (pièce) diffèrent couramment chez ce type de
  fournisseur. Le garde-fou sur-livraison de `matching.py` continue de
  classer ça "à confirmer" à raison (jamais deviner un facteur de
  conversion, il varie par article) — c'est à l'acheteur de trancher au
  cas par cas via la case à cocher, PAS un cas à automatiser.

**Session R2 suite — Electric Plus / GMR, fournisseur n°3 (fait) :**

- **GMR n'envoie PAS de bon de livraison séparé** (confirmé par
  l'acheteur) : "Electric Plus" est l'enseigne grand public, "GMR" sa
  branche grands comptes pros — même structure documentaire, le
  rapprochement se fait directement à partir de ses FACTURES (marquées
  "DUPLICATA"). `moteur/fournisseurs/electricplus.py::parse_bl_electricplus`
  réutilise la même ancre "PF" que le devis de ce fournisseur (juste
  au-dessus dans le même fichier), mais la désignation est ici éclatée en
  plusieurs cellules OCR (un mot chacune) au lieu d'une seule — la
  quantité livrée est déduite de Montant / P.U. net (comme 109
  Distribution) plutôt que lue directement. "Total HT" affiché en pied de
  page -> autocontrôle possible (contrairement à Coredime, qui n'affiche
  jamais de prix).
- **Bug le plus important trouvé cette session, silencieux et total** :
  `BonLivraison.fournisseur` vaut "ELECTRIC PLUS" (nom détecteur), mais le
  Suivi écrit "GMR" dans sa colonne Fournisseur — sans conversion,
  `lire_lignes_commande()` ne retrouvait JAMAIS aucune ligne pour AUCUNE
  commande Electric Plus (0 sûr, 0 à confirmer, 0 déjà à jour, 0 inconnu
  sur 7 vraies factures alors que les 14 lignes existaient bien dans le
  Suivi). Corrigé en réutilisant `moteur.panier.MAPPING_FOURNISSEURS`
  (déjà la source de vérité pour cette conversion côté Panier) au lieu
  d'un mapping dupliqué — **tout futur fournisseur dont le nom détecteur
  diffère du nom Suivi doit passer par ce même mapping, pas en inventer un
  nouveau.** Testé
  (`test_lire_lignes_commande_convertit_nom_fournisseur_electric_plus_gmr`).
- **`moteur/detecteur.py` corrigé** : "ELECTRIC PLUS" ne matchait que
  `ELECTRIC\s+PLUS` (espace obligatoire) — l'OCR colle parfois "Electric
  Plus" en "ElectricPlus" sur les factures scannées (contrairement au
  devis, où le texte natif garde toujours l'espace). Motif élargi à
  `ELECTRIC\s*PLUS`.
- **2 bugs OCR réels corrigés, même famille que ceux déjà vus chez 109
  Distribution/Coredime** : le motif date+n°facture utilisait `\D+` (puis
  `[^\d\n]+`) entre les 3 cellules consécutives "date/date échéance/n°
  facture" — mais ces cellules n'ont RIEN entre elles une fois aplaties
  (juste un saut de ligne), donc `numero_bl`/`date_bl` ressortaient
  TOUJOURS vides ; corrigé avec `[^\d]{0,3}` (borné, pour ne jamais dériver
  vers un nombre lointain sans rapport — même prudence que le bug de
  séparateur `\s` chez Coredime). Et l'ancre de prix "<prix> PF" ratait
  toute ligne où l'OCR lit "0,53" en "O,53" (lettre O) — repli O/0 ajouté,
  a sauvé une ligne entière sur 1 facture testée.
- **Recette réelle sur 7 vraies factures** (2 déjà dans `a_traiter/BL/`
  avant la session + 5 ajoutées pendant) : 9 lignes écrites pour de vrai
  dans le classeur vivant, 8 déjà à jour (rien réécrit), 0 à confirmer,
  0 inconnue — toutes les quantités correspondent exactement aux quantités
  commandées (aucune sur-livraison, aucun écart de référence cette fois).
  Les 7 factures ont été archivées dans `a_traiter/BL/Traités/`.

**Session R2 suite — Cominter Ouest, fournisseur n°4 (fait) :**

- **Piège signalé par l'acheteur AVANT tout code, décisif** : un même
  fichier PDF peut contenir PLUSIEURS BL scannés à la suite (jusqu'à 8 vus
  en session) — mais un même BL peut aussi déborder sur 2 pages (tableau
  d'articles page 1, totaux page 2, MÊME n° "OBL......" sur les deux).
  Nouvelle fonction générique `moteur.ocr.grouper_pages_par_identifiant()`
  (testée isolément dans `tests/test_ocr.py`) : regroupe les pages par
  identifiant AVANT tout parsing — une page sans identifiant détecté
  (le pied de page qui déborde) rejoint le groupe précédent au lieu d'ouvrir
  un nouveau groupe à tort. Comparaison sur le GROUPE CAPTANT du motif
  (les chiffres), pas le texte entier du match : l'OCR lit tantôt "OBL"
  tantôt "0BL" (O confondu avec 0) pour le MÊME numéro sur des pages
  différentes du même document — comparer le texte brut aurait coupé un
  même BL en deux groupes à tort (bug réel trouvé et corrigé en recette).
  **Conséquence sur l'architecture** : `parse_bl` peut désormais retourner
  une LISTE de BonLivraison (pas un seul) — `lecture_bl.lire_bl()` et
  `analyser_dossier()` adaptés pour normaliser en liste dans tous les cas
  (les autres fournisseurs continuent de retourner un seul objet, normalisé
  automatiquement).
- **Structure de ligne la plus irrégulière rencontrée jusqu'ici** : contrairement
  aux fournisseurs précédents, la cellule "Px net" peut carrément DISPARAÎTRE
  de l'OCR (remplacée par le seul taux de remise, ex. "30%") ou être COLLÉE
  au taux dans une seule cellule ("30% 110,67" — le vrai Px net est là, après
  le %). `_prix_net_bl_cominter()` gère les 3 cas réels rencontrés, dans cet
  ordre : (1) taux+prix collés dans la même cellule -> prix repris tel quel ;
  (2) taux seul -> prix net reconstruit à partir du Px unitaire et du taux
  (vérifié exact sur un cas réel : 26,78 × 0,70 = 18,746, 3 × 18,746 = 56,24
  = montant affiché) ; (3) cellule Px net normale -> lue directement. Comme
  pour les fournisseurs précédents, la quantité livrée est déduite de
  Montant / Px net plutôt que lue dans la cellule Qté+Unité (souvent collée
  en une seule cellule ici, ex. "22,00 Unité").
- **Pas d'autocontrôle Total HT** (comme Coredime, contrairement à 109
  Distribution/Electric Plus) : le tableau de répartition TVA en pied de
  page a une structure trop irrégulière (colonnes qui se décalent selon le
  nombre de taux de TVA présents sur le document) pour en extraire la
  valeur de façon fiable avec les exemples actuels — laissé de côté plutôt
  que de calculer un total faux.
- **Bug réel corrigé (recette utilisateur, voir "correctifs critiques
  post-recette" ci-dessous) : référence renvoyée sur sa PROPRE ligne**
  ("PLW11643") : sur un article à désignation longue, l'OCR sépare parfois
  la référence de la ligne désignation+prix qui la précède — l'ordre
  inverse du cas Coredime déjà géré (désignation qui déborde APRÈS). Un
  raccord (lookahead d'une ligne) la récupère désormais au lieu de
  l'ignorer silencieusement — voir `_parse_un_bl_cominter()`, testé
  (`test_parse_bl_cominter_6_commande_m4_260_quantites_entieres_et_ligne_recuperee`).
- **Nouveau garde-fou, trouvé en confrontant deux fichiers du même lot à la
  recette** : le même BL papier ("BL M3.23.030 MABOOC.pdf", déjà présent
  d'une session précédente) et un nouveau scan du même document
  (`doc07136020260813090605.pdf`, même n° OBL108106) se sont retrouvés
  ensemble dans `a_traiter/BL/` au moment de la recette. Sans protection,
  les deux auraient chacun proposé une écriture "sûre" indépendante vers
  la MÊME ligne du Suivi, empilant deux fois la même quantité. Nouvelle
  fonction `_desamorcer_conflits_meme_ligne_suivi()` dans `pipeline_bl.py`,
  appelée en fin de `rapprocher_dossier()` : si deux BL de fichiers
  DIFFÉRENTS visent la même ligne Suivi côté "sûr", seul le premier reste
  "sûr" — l'autre bascule "à confirmer" avec le nom du fichier concurrent
  en raison. Protège aussi bien contre un doublon de scan que contre une
  vraie livraison fractionnée mal enchaînée. Testé
  (`test_desamorcer_conflits_meme_ligne_suivi_deux_fichiers_meme_cible`).
- **Recette réelle sur les vrais BL disponibles** : 8 lignes écrites pour
  de vrai dans le classeur vivant (dont les 8 BL scannés dans un seul
  fichier), 4 BL entièrement archivés. 5 BL déplacés vers "À vérifier"
  (4 à cause du doublon de scan ci-dessus, 1 pour une référence "L405205"
  introuvable dans le Suivi). **À signaler à l'acheteur** : les anciens
  fichiers `BL M3.23.030 MABOOC.pdf` / `BL M3.23.030.1 MABOOC.pdf` /
  `BL M3.23.034 MABOOC.pdf` (présents avant cette session) sont des scans
  du MÊME BL que certains fichiers `doc...pdf` déposés cette session (même
  n° OBL) — à nettoyer manuellement de `a_traiter/BL/`/`À vérifier/` pour
  éviter la confusion, le garde-fou ci-dessus empêche juste la double
  écriture, il ne fait pas le ménage.

**Session R2 suite — Cominter Ouest, correctifs critiques post-recette
(fait) :** la 1ère recette Cominter (ci-dessus) a été jugée **dangereuse**
par l'acheteur — quantités livrées non entières écrites dans le VRAI
classeur vivant (ex. 29,96 sur un article à l'unité), une ligne
totalement disparue (qté écrite à 0), et un fichier de 8 BL archivé sous
le nom du premier BL du lot sans indiquer qu'il en contenait d'autres.
Deux correctifs :

1. **Quantité livrée : cellule Qté imprimée préférée à la division
   Montant/Px net** (`_quantite_bl_cominter()`, `moteur/fournisseurs/
   cominter.py`) — la division introduisait du bruit d'arrondi sur les
   lignes remisées (prix net affiché déjà arrondi à 2 décimales, contrairement
   au calcul interne), d'où les quantités non entières. La cellule Qté
   n'est utilisée qu'en repli quand elle est absente de l'OCR (comme
   avant). Verrouillé par
   `test_parse_bl_cominter_6_commande_m4_260_quantites_entieres_et_ligne_recuperee`
   (30,0 et non 29,96 ; 150,0 et non 149,89).
2. **Archivage par BL individuel** (pas par fichier entier) — voir la
   section dédiée juste en dessous : c'est le fond du problème "8 BL dans
   1 fichier, 1 seul traité au final".

**Session suivante — archivage par BL individuel + rapprochement de
repli désignation (fait), suite à un 2e signalement critique de
l'acheteur** (avec capture d'écran d'un vrai BL papier, référence
"L4052[trou de perforateur]9") : deux défauts distincts, mêmes correctifs
réutilisables pour tout futur fournisseur multi-BL/scans papier abîmés.

- **Archivage par BL individuel** (`BonLivraison.pages`,
  `moteur.ocr.pages_par_identifiant()`, `moteur/rapprochement/
  pipeline_bl.py`) : jusqu'ici, un fichier Cominter à plusieurs BL n'était
  archivé qu'en BLOC — dès qu'UN seul de ses BL restait non résolu (n° de
  commande introuvable, référence litigieuse...), TOUS les autres du même
  fichier restaient bloqués avec lui dans `a_traiter/BL/` (constaté
  réellement : sur 8 BL scannés dans un fichier, un seul avait fini
  correctement traité et renommé). `parse_bl_cominter()` renseigne
  désormais `bl.pages` (indices de page 0-based occupés par CE BL dans le
  fichier source, via la nouvelle `moteur.ocr.pages_par_identifiant()` —
  même regroupement que `grouper_pages_par_identifiant()`, mais en indices
  plutôt qu'en mots OCR, factorisé pour éviter toute divergence entre les
  deux). `regrouper_par_bl()` (pipeline_bl.py) groupe désormais par
  **objet BL (`id(bl)`)**, plus par nom de fichier (bug latent corrigé au
  passage : grouper par fichier fusionnait à tort les lignes de plusieurs
  BL d'un même fichier dans un seul groupe). Pour un fichier à PLUSIEURS
  BL, chaque BL résolu est extrait via `fitz` (découpage de pages, pas
  `openpyxl` : PDF, pas Excel) vers `Traités/` ou `À vérifier/`
  individuellement (`_extraire_bl_vers()`), puis le fichier source est
  réduit aux pages pas encore redistribuées (`_reecrire_avec_pages()`) ou
  supprimé si tout a pu être extrait — jamais laissé intact avec la
  totalité de ses pages (sinon relu en entier, pour rien, à la prochaine
  exécution). Les fichiers à UN SEUL BL (tous les fournisseurs sauf
  Cominter Ouest) gardent le comportement historique — fichier entier
  déplacé tel quel, aucune réécriture PDF inutile. Les anomalies
  "n° de commande introuvable"/"commande absente du Suivi" (avant :
  bloquaient tout le fichier via `anomalies_lecture`) sont désormais
  rattachées au `BonLivraison` concerné (`rapport.anomalies_bl`, nouveau
  champ) — seul CE BL est affecté, pas ses frères du même fichier.
  `anomalies_lecture` reste réservé aux échecs de lecture TOTAUX (aucun
  objet BL n'existe encore : fournisseur non reconnu, pas de parser, PDF
  illisible), qui bloquent forcément le fichier entier faute de pouvoir le
  découper. Testé sur le VRAI fixture 8 pages
  (`test_appliquer_et_archiver_fichier_multi_bl_archive_le_resolu_sans_attendre_ses_freres`,
  `test_appliquer_et_archiver_fichier_multi_bl_supprime_la_source_une_fois_tout_redistribue`).
  **Piège pour tout futur appel direct hors GUI** : `appliquer_et_archiver()`
  suppose que `dossier_a_traiter` est TOUJOURS `a_traiter/BL/` lui-même
  (parent direct de `Traités/` et `À vérifier/`), jamais un de ses
  sous-dossiers — un appel de test manuel pointé par erreur sur
  `a_traiter/BL/À vérifier/` a créé des sous-dossiers imbriqués
  `À vérifier/Traités/` et `À vérifier/À vérifier/` au lieu des dossiers
  attendus (repéré et corrigé à la main dans la foulée, aucune perte de
  données — juste un mauvais rangement). `rapprocher_dossier()` (lecture
  seule) peut, lui, être pointé sur n'importe quel dossier sans risque
  (c'est ce qui permet de tester "À vérifier" isolément) — seul
  `appliquer_et_archiver()` a cette contrainte.
- **Rapprochement de repli sur référence proche + désignation**
  (`_repli_reference_proche()`, `moteur/rapprochement/matching.py`) : cas
  réel signalé — un trou de perforateur sur le BL PAPIER abîme un chiffre
  de la référence imprimée AVANT le scan (ex. "405209" devient "L405205"
  à l'OCR, un seul chiffre différent) — l'acheteur confirme que ce type de
  dommage physique est **quotidien**, pas un cas rare. Avant ce correctif,
  0 candidat de référence exacte -> bac "inconnu" direct, la ligne
  disparaissait du rapprochement alors que l'article existait bien dans le
  Suivi. Quand `apparier()` ne trouve AUCUNE référence exacte pour une
  commande, un repli cherche désormais une ligne Suivi dont le cœur
  numérique (`moteur.base.coeur_numerique`) est à **exactement 1
  caractère d'écart** (`_distance_courte()`, longueur + caractères
  différents) — n'accepte que s'il existe UNE SEULE ligne candidate à
  cette distance (jamais de choix au hasard). La désignation n'est PAS un
  filtre bloquant (elle varie trop entre BL et Suivi pour servir de seuil
  fiable — mesuré sur des paires réellement correctes : similarité
  `difflib.SequenceMatcher` de 0,267 à 1,0) mais est affichée en clair
  dans la raison, pour que l'acheteur tranche d'un coup d'œil. Statut
  TOUJOURS "à confirmer" **sauf un cas** (voir bug ci-dessous), jamais
  "sûr" — un rapprochement de repli n'est jamais écrit automatiquement.
  Validé manuellement sur le vrai cas signalé (`bl_cominter_3.pdf`,
  L405205 -> 405209) puis testé
  (`tests/test_rapprochement_matching.py::test_apparier_repli_reference_proche_*`).
- **BUG RÉEL CRITIQUE trouvé lors du 1er vrai passage d'écriture avec ce
  repli, corrigé dans la foulée** : la toute première utilisation réelle
  du repli (ligne 405209, commande M3.23.034) visait une ligne du Suivi
  DÉJÀ à 3 livrées/3 commandées, même tarif, date déjà renseignée — un cas
  "déjà à jour" en tout point (`_comparer()` l'avait bien détecté), sauf
  que le code d'`apparier()` forçait alors TOUJOURS `Statut.A_CONFIRMER`
  pour toute correspondance de repli, sans jamais regarder ce que
  `_comparer()` avait réellement conclu. Résultat concret, écrit pour de
  vrai dans le classeur vivant avant d'être repéré et corrigé
  manuellement (voir sauvegardes horodatées) : `qte_livree_cumulee` a
  calculé 3 (déjà) + 3 (BL) = **6 pour 3 commandées** — une sur-livraison
  fantôme, sur un article qui n'avait en réalité RIEN de nouveau à
  enregistrer. Corrigé à deux niveaux : (1) `apparier()` ne force
  "à confirmer" que si `_comparer()` n'a PAS conclu "déjà à jour" ; un
  repli vers une ligne déjà à jour ressort désormais `Statut.DEJA_A_JOUR`
  (avec la raison du repli conservée, pour que l'acheteur comprenne
  pourquoi cette ligne apparaît sans rien à faire) — donc plus JAMAIS de
  cumul sur une quantité déjà exacte ; (2) garde-fou en profondeur ajouté
  dans `ecritures_pour()` (`pipeline_bl.py`) : ignore explicitement toute
  correspondance `DEJA_A_JOUR` même si elle se retrouvait par erreur dans
  la liste à écrire. Testés
  (`test_apparier_repli_reference_proche_deja_a_jour_ne_cumule_jamais`,
  `test_ecritures_pour_ignore_les_correspondances_deja_a_jour`). **Leçon
  générale pour tout futur code de repli/rapprochement approximatif** :
  ne jamais court-circuiter le statut renvoyé par `_comparer()` sans
  vérifier s'il s'agit d'un "déjà à jour" — sinon un rapprochement
  approximatif peut transformer une ligne parfaitement à jour en
  sur-livraison fantôme.
- **BUG RÉEL DE FOND trouvé par l'acheteur juste après, distinct du
  précédent : "déjà à jour" ne vérifie JAMAIS que la date enregistrée est
  la BONNE date** — seulement que qté/tarif correspondent et qu'UNE date
  (n'importe laquelle) est présente (`_comparer()`, condition
  `ligne_suivi.date_livraison is not None`). Conséquence concrète : des
  lignes écrites AVANT que l'extraction de date_bl soit corrigée (Cominter,
  voir plus haut — `date_bl` valait `""`, secours sur `date.today()` dans
  `ecritures_pour()`) sont restées bloquées avec la date du JOUR DU
  TRAITEMENT au lieu de la vraie date du BL, et aucun passage suivant ne
  les corrige jamais puisqu'elles ressortent "déjà à jour" à chaque fois
  (qté/tarif, eux, étaient corrects). Repéré par l'acheteur en ouvrant le
  1er BL archivé (108.271) : Suivi affichait 13/08 (date du traitement)
  alors que le vrai BL est daté du 15/07 (et **signé** le 16/07 — voir
  règle ci-dessous). **Audit en lecture seule** mené sur les 30 fichiers
  déjà dans `a_traiter/BL/Traités/` (re-OCR de chacun, comparaison
  `bl.date_bl` réel vs `Date de livraison` actuelle du Suivi pour la ligne
  correspondante) : **9 écarts trouvés**, tous des lignes Cominter/Coredime
  écrites tôt dans la session, corrigés un par un après revue avec
  l'acheteur (voir sauvegarde
  `1.3.0.1. Suivi commandes - 2026_20260813_203412.xlsx`). Un 10e cas (109
  Distribution, commande 132.008) a une date imprimée illisible même à
  l'OCR ("108/2026") — laissé tel quel, à relever manuellement sur le
  papier si besoin, pas un cas automatisable.
  **Règle donnée par l'acheteur, à respecter pour toute date de
  livraison future** : c'est la date de LA SIGNATURE sur le BL papier qui
  fait foi si elle est présente (constaté : parfois un jour après la date
  imprimée du document, ex. 108.271 imprimé 15/07 mais signé 16/07) ;
  à défaut de signature, la date imprimée du BL. **L'OCR de ce projet
  n'extrait que la date IMPRIMÉE, jamais une date manuscrite** (RapidOCR
  n'est pas validé pour la reconnaissance d'écriture manuscrite ici) — donc
  même une date "correcte" au sens du code peut être fausse au sens de
  cette règle si une signature datée existe et diffère. Pas de solution
  automatisée à ce stade ; à garder en tête pour toute future recette
  (comparer visuellement la zone de signature, pas seulement la date
  imprimée en haut du document).
- **Correctif structurel demandé par l'acheteur juste après l'audit** :
  "déjà à jour" vérifie désormais aussi la COHÉRENCE de la date, pas
  seulement sa présence — pour que ce type de ligne bloquée avec une
  mauvaise date pour toujours (voir juste au-dessus) ne puisse plus se
  reproduire silencieusement, même pour des fournisseurs pas encore
  audités. `apparier()`/`_comparer()` (`matching.py`) acceptent désormais
  un paramètre `date_bl_reelle` (date déjà extraite par l'appelant, ou
  `None` si illisible/non fournie — comportement inchangé dans ce cas,
  rétrocompatible) : si qté/tarif sont déjà identiques mais que la date
  enregistrée diffère de `date_bl_reelle`, la ligne ressort "à confirmer"
  (jamais silencieusement "déjà à jour") avec la date correcte donnée en
  clair dans la raison. **Piège évité** : une ligne dans ce cas a sa
  quantité DÉJÀ correcte — la traiter comme une correspondance normale
  aurait recalculé `qte_livree_cumulee = déjà + BL` et doublé la quantité
  (exactement le bug de sur-livraison fantôme trouvé juste avant). Nouveau
  champ `Correspondance.qte_deja_incluse` (défaut `False`) : quand `True`,
  `qte_livree_cumulee` retourne la quantité déjà enregistrée SANS jamais
  la recumuler — utilisé ici et propagé aussi par le chemin du repli
  référence-proche. `pipeline_bl.rapprocher_dossier()` calcule
  `date_bl_reelle` une fois par BL (`_parser_date_bl(bl.date_bl)`) et le
  passe à `apparier()`. Testé
  (`test_apparier_date_suivi_incoherente_devient_a_confirmer_sans_recumuler`,
  `test_apparier_date_suivi_coherente_avec_date_bl_reelle_reste_deja_a_jour`,
  `test_apparier_sans_date_bl_reelle_comportement_inchange`).

**Session suivante — nouveau lot de BL déposés, 3 échecs d'extraction
corrigés (109 Distribution ×2, Electric Plus ×1) :** l'acheteur a déposé un
nouveau lot dans `a_traiter/BL/` ; le rapprochement en lecture seule a
signalé 3 fichiers à 0 ligne extraite (ou total incohérent) malgré un Total
HT affiché — creusé et corrigé sur les vrais PDF
(`tests/fixtures/bl_dist109_6.pdf`, `bl_dist109_7.pdf`,
`bl_electricplus_8.pdf`) :
- **`bl_dist109_6.pdf` — en-tête de tableau avec accent** : l'OCR a lu
  "Reférencearticle" AVEC un é (donc "RÉFÉRENCEARTICLE" après `.upper()`)
  au lieu de la forme sans accent habituelle — `MOTIF_ENTETE_TABLEAU_BL`
  ("REFERENCEARTICLE", sans gestion d'accent) ne matchait jamais, faisant
  disparaître TOUT le tableau (0 ligne pour un Total HT de 1750€ affiché).
  `_sans_espaces()` (`dist109.py`) normalise désormais les accents
  (`unicodedata.normalize("NFKD", ...).encode("ascii", "ignore")`) avant
  toute comparaison, pour les deux graphies. Confirme au passage que la
  quantité réelle de cette ligne est 10 (175€ × 10 = 1750€), pas 1.
- **`bl_dist109_7.pdf` — cellule Eco-part intercalée + lignes sans code
  TVA** : sur ce document, 2 défauts distincts. (1) Une cellule "Eco-part"
  (2 décimales) s'intercale parfois entre le vrai P.U.Net (TOUJOURS 5
  décimales sur ce gabarit, ex. "22,00000") et le code TVA — l'ancien code
  prenait systématiquement "la cellule juste avant le code TVA" pour le
  P.U.Net, donnant une quantité totalement fausse quand c'était en fait
  l'éco-part (104,76 au lieu de 3 sur la ligne 302304). (2) 3 lignes sur 7
  n'ont AUCUNE cellule de code TVA lue par l'OCR — silencieusement
  ignorées avant ce correctif. Nouveau `MOTIF_PU_NET_BL` (`^\d+[,.]\d{5}$`)
  : le nombre de décimales sert désormais de signal pour (a) vérifier que
  la cellule juste avant le code TVA EST bien le P.U.Net, sinon reculer
  d'une cellule de plus (repli 1), et (b) retrouver le P.U.Net directement
  quand aucun code TVA n'a été détecté sur la ligne, le Total restant
  alors la DERNIÈRE cellule de la ligne (repli 2, vérifié exact sur les 3
  lignes concernées). Les 7 lignes retombent maintenant exactement sur le
  Total HT affiché (136,96€).
- **`bl_electricplus_8.pdf` — confusion OCR "PF" → "PR"** : l'ancre de
  prix (`MOTIF_PF_ELECTRICPLUS`) attendait strictement une cellule finissant
  par "PF" — l'OCR a lu "PR" (F confondu avec R) sur ce document, ancre
  jamais trouvée, TOUTE la facture ressortait à 0 ligne malgré un Total HT
  de 482,50€ affiché. Même famille que la tolérance O/0 déjà en place sur
  cette ancre (session R2 suite) : motif élargi à `P[FR]`. **Gap restant,
  volontairement non corrigé (un seul exemple, règle d'or)** : sur ce même
  fichier, `numero_bl`/`date_bl` restent vides — l'ordre des colonnes
  facture/date est INVERSÉ par rapport aux autres factures Electric Plus
  déjà vues (ici "FACTURE" puis "DATE" dans le flux OCR, `MOTIF_FACTURE_
  DATE_ELECTRICPLUS` suppose l'inverse) — à reprendre si un 2e exemple
  réel avec ce même ordre se présente.
Testé (`test_parse_bl_dist109_6_entete_avec_accent`,
`test_parse_bl_dist109_7_eco_part_intercalee_et_lignes_sans_code_tva`,
`test_parse_bl_electricplus_8_prix_r_confondu_avec_f`).

**Session suivante — 109 Distribution, découpage multi-BL (fait), suite à
un nouveau lot avec "scans en masse" :** l'acheteur a prévenu qu'elle avait
déposé des scans individuels ET des scans groupés. Un fichier
(`doc07149020260814105344.pdf`, 8 pages) contenait effectivement 8 BL 109
Distribution différents — jamais rencontré avant chez ce fournisseur
(seul Cominter Ouest avait ce besoin jusqu'ici). `parse_bl_109()` retourne
désormais une LISTE (comme `parse_bl_cominter()`), en réutilisant
directement `MOTIF_BL_NUMERO_DATE` (numéro de BL, déjà existant) comme
identifiant pour `moteur.ocr.pages_par_identifiant()` — aucun nouveau
motif nécessaire, le regroupement par page est immédiat. `bl.pages` est
renseigné pour chaque BL du groupe, ce qui active automatiquement
l'archivage par BL individuel (déjà généralisé dans `pipeline_bl.py` pour
tout fournisseur qui renseigne `pages`, pas seulement Cominter).
**2e bug trouvé en creusant les 2 BL du lot encore à 0 ligne après le
découpage** : le repère de fin de tableau `MOTIF_PIED_TABLEAU_BL`
("Total Eco-part HT") est lui aussi sujet à erreur OCR ("Tatal Eco-part
HT", voire pire) — repli ajouté sur "Total HT" juste après, qui est resté
fiable sur TOUS les documents 109 Distribution vus jusqu'ici (utilisé par
ailleurs pour l'autocontrôle de total). Les 8 BL du fichier retombent
maintenant exactement sur leur Total HT respectif. Fixture réelle ajoutée
(`tests/fixtures/bl_dist109_8_multi_bl_8pages.pdf`,
`test_parse_bl_dist109_8_multi_bl_8pages`) ; les 7 tests existants ont dû
être adaptés au nouveau type de retour (`[bl] = parse_bl_109(...)` au lieu
de `bl = parse_bl_109(...)`), aucun autre appelant dans le code (la
normalisation liste/objet unique se fait déjà dans `lecture_bl.lire_bl()`,
comme pour Cominter).

**Session suivante — recette jugée insatisfaisante par l'acheteur (seulement
5 BL rapprochés sur 16), deux correctifs de fond (fait) :** message direct
de l'acheteur : "Cette session n'a pas du tout été satisfaisante [...] en
l'état ce n'est pas du tout intéressant et utile." Deux causes concrètes
identifiées à partir de 2 exemples qu'elle a pointés du doigt.

- **Cause du BL "123.098" non rapproché malgré une seule ligne de chaque
  côté, même désignation, même tarif** : la référence Suivi est
  "ALB69894", le BL a extrait "9894" — **pas un chiffre abîmé (déjà
  couvert par le repli existant), un chiffre de tête ENTIÈREMENT
  disparu**. Explication de l'acheteur : "les gars de l'atelier percent
  les BL pour les classer dans un classeur" — même famille que le trou de
  perforateur déjà connu, mais qui efface parfois complètement 1-2
  chiffres du DÉBUT plutôt que d'en abîmer un au milieu. Nouvelle fonction
  `_chiffres_tete_manquants()` (`matching.py`) : accepte qu'une référence
  soit un SUFFIXE de l'autre avec 1 ou 2 chiffres d'écart de tête,
  intégrée à `_repli_reference_proche()` à côté du repli existant (1
  seule ligne candidate exigée entre les deux critères combinés, jamais un
  choix au hasard). Toujours "à confirmer", jamais "sûr". Testé
  (`test_apparier_repli_chiffre_de_tete_manquant*`).
- **Déduction de commande par signature de contenu** (nouvelle fonction
  `deduire_commande_par_contenu()`, `matching.py`) : l'acheteur a
  explicitement demandé cette fonctionnalité — quand le n° de commande
  n'est pas lisible sur le BL (même cause : trou de perforateur, mais sur
  la zone du n° cette fois), chercher dans TOUTES les lignes du Suivi pour
  ce fournisseur (nouvelle `lire_lignes_fournisseur()`, contrairement à
  `lire_lignes_commande()` qui filtre déjà par commande) la commande dont
  le plus de lignes concordent EXACTEMENT (référence ET quantité
  commandée) avec celles du BL — une empreinte de contenu plutôt qu'un
  numéro illisible. Validé sur les 3 vrais BL Cominter signalés par
  l'acheteur (`OBL108110`->135.039 avec 3/6 lignes concordantes,
  `OBL108186`->142.032 avec 4/5, `OBL108367`->M3.14.361 avec 4/4 — le
  meilleur candidat avait systématiquement 3 à 4x plus de lignes
  concordantes que le 2e). Garde-fous stricts : au moins 2 lignes
  concordantes (jamais une coïncidence sur 1 seule référence, souvent
  générique/réutilisée), et un score STRICTEMENT meilleur que tout autre
  candidat (égalité -> aucune déduction). `LigneSuivi` a un nouveau champ
  `numero_commande` (défaut `""`, uniquement renseigné par
  `lire_lignes_fournisseur()`) pour pouvoir grouper les lignes par
  commande après coup. **Un n° de commande DÉDUIT n'est JAMAIS utilisé
  pour un rapprochement "sûr" automatique** — `rapprocher_dossier()`
  (`pipeline_bl.py`) force chaque correspondance obtenue via une commande
  déduite à "à confirmer" (même si `_comparer()` l'aurait autrement jugée
  "sûre"), avec la déduction expliquée en clair dans la raison. Testé
  (`test_deduire_commande_par_contenu_*`).
- **Repli référence proche sur écart alphanumérique final** (cas réel
  commande 142.033) : BL "H07VK16BL" vs Suivi "H07VK16B" (un "L" en trop
  en fin de référence) — le cœur numérique des deux ("716") est trop
  court (< 4 chiffres, seuil de `coeur_numerique`) pour les replis
  existants, qui ne comparent que la partie numérique. Nouvelle fonction
  `_cle_brute()` (`matching.py`, texte alphanumérique brut, garde les
  lettres) : repli supplémentaire dans `_repli_reference_proche()` quand
  le cœur numérique est absent d'au moins un côté, même critère 1
  caractère d'écart appliqué au texte brut. Testé
  (`test_apparier_repli_reference_courte_ecart_alphanumerique_final`).
- **Bons de retour 109 Distribution reconnus comme un type de document
  distinct (fait), suite à un signalement critique de l'acheteur** : ce
  fournisseur envoie aussi des "Retour n° X du date" — un document qui
  ANNULE une ligne d'un BL précédent (cas réel : article listé sur le BL
  737760 (commande M3.10.175, article R9PRA263) mais dont la case "livré"
  n'était PAS cochée à réception ; le fournisseur envoie le Retour
  n°25894 qui référence ce BL, puis le BL 737851 livre l'article
  conformément le lendemain). **Bug réel trouvé AVANT toute écriture,
  potentiellement grave** : `MOTIF_BL_NUMERO_DATE` matchait à tort la
  référence "Bon de livraison n° 737760" présente DANS LE CORPS du retour
  (un retour cite toujours le BL qu'il annule), faisant fusionner le
  retour comme une "page supplémentaire" du BL 737760 par
  `moteur.ocr.pages_par_identifiant` — si ça n'avait pas été repéré, le
  retour aurait été traité comme une 2e livraison de R9PRA263, doublant
  la quantité (même famille de bug que la sur-livraison fantôme
  documentée plus haut, mais sur un document qui n'est même pas une
  livraison). Corrections dans `moteur/fournisseurs/dist109.py` :
  - `MOTIF_RETOUR_NUMERO_DATE` : détecte l'identité PROPRE d'un retour
    ("Retour n° X du date") — vérifié EN PREMIER, avant
    `MOTIF_BL_NUMERO_DATE`, sur le texte de la page.
  - `MOTIF_IDENTIFIANT_PAGE_BL` : reconnaît "Retour n°" OU "livraison n°"
    pour le regroupement de pages (`pages_par_identifiant`) — une page de
    retour ne se fusionne plus jamais avec le BL qu'elle référence dans
    son corps, puisque "Retour n°25894" apparaît AVANT la référence au BL
    dans le flux OCR de la page et est donc trouvé en premier.
  - `BonLivraison` a deux nouveaux champs : `type_document` ("BL" ou
    "RETOUR", défaut "BL" — rétrocompatible pour tous les autres
    fournisseurs) et `numero_bl_origine` (le BL que le retour annule).
  - `moteur/rapprochement/pipeline_bl.py` (`rapprocher_dossier()`) :
    avant le traitement normal, construit la liste des références
    annulées par (n° de BL d'origine) à partir de TOUS les retours du
    lot. Un document `type_document == "RETOUR"` ne produit JAMAIS de
    correspondance à écrire lui-même — juste une entrée
    `anomalies_bl` informative ("Bon de retour — annule X du BL Y").
    Quand le BL d'origine référencé est traité, toute ligne dont la
    référence figure dans les annulations bascule "à confirmer" avec la
    raison explicite, MÊME si `_comparer()` l'aurait jugée "sûre" —
    jamais écrite automatiquement.
  Fixtures réelles ajoutées (`tests/fixtures/bl_dist109_9_retour.pdf`,
  `bl_dist109_10_bl_avec_retour_associe.pdf`, `bl_dist109_11.pdf`), et le
  fixture multi-BL existant (`bl_dist109_8_multi_bl_8pages.pdf`, qui
  contenait CE MÊME retour sans qu'il ait été identifié comme tel) mis à
  jour en conséquence. Testé
  (`test_parse_bl_dist109_9_retour_seul`,
  `test_parse_bl_dist109_10_bl_avec_retour_associe`,
  `test_parse_bl_dist109_11_livraison_conforme_apres_retour`,
  `test_parse_bl_dist109_8_multi_bl_8pages` mis à jour).
  **Limité à 109 Distribution pour l'instant** (seul fournisseur où un
  retour réel a été vu) — à étendre à d'autres fournisseurs dès qu'un cas
  réel se présentera (règle d'or).
- **Ambiguïté de cœur numérique levée par correspondance exacte** (cas
  réel, commande M3.10.175) : "R9PRC263" et "R9PRA263" partagent le MÊME
  cœur numérique ("9263" — la lettre médiane C/A n'est pas un chiffre,
  donc ignorée par `coeur_numerique`), alors que ce sont deux articles
  RÉELLEMENT différents (interrupteur différentiel type A vs type AC) —
  la ligne R9PRC263 du BL restait "ambigu" à tort. `apparier()` cherche
  désormais, parmi les candidats ambigus par cœur numérique, s'il en
  existe UN SEUL dont le texte est EXACTEMENT identique à la référence du
  BL — si oui, il est retenu (une correspondance de texte exact est
  toujours plus fiable qu'une coïncidence de cœur numérique). Sans
  correspondance exacte, reste "ambigu" comme avant (jamais de choix au
  hasard). Testé
  (`test_apparier_ligne_ambigue_mais_correspondance_exacte_disponible`,
  `test_apparier_ligne_ambigue_inconnu_sans_correspondance_exacte`).
- **Données historiques corrigées sur le vrai Suivi** : la même commande
  M3.10.175 avait 4 lignes écrites avec des quantités/tarifs FAUX lors
  d'une session précédente (R9PFC620, R9PFC616, 61401, BC6AFSTL8) —
  décalage entre chaque référence et le tarif de la référence PRÉCÉDENTE
  sur le BL (cause historique non élucidée avec certitude, code actuel
  déjà vérifié correct par les tests verrouillés ci-dessus). Corrigé
  directement sur le vrai classeur après confirmation des vraies valeurs
  par relecture du BL (voir sauvegarde
  `1.3.0.1. Suivi commandes - 2026_20260817_142327.xlsx`).
- **PIÈGE OPÉRATIONNEL trouvé en écrivant réellement le cas M3.10.175,
  à ne plus jamais reproduire** : `rapprocher_dossier()` ne construit
  l'exclusion "ligne annulée par un retour" (voir plus haut) qu'à partir
  des documents PRÉSENTS DANS LE MÊME APPEL. En traitant le BL 737760
  SANS inclure le retour n°25894 dans le même lot (copié séparément vers
  `a_traiter/BL/` pour l'écriture), R9PRA263 s'est retrouvé écrit "sûr"
  depuis 737760 (13/08, PAS livré réellement) au lieu du 737851 (14/08, la
  vraie livraison conforme) — qté/tarif corrects par pure coïncidence
  (même article, même prix), seule la date était fausse, corrigée après
  coup. **Règle à respecter systématiquement** : un BL lié à un retour
  (ou l'inverse) doit TOUJOURS être traité dans le même appel à
  `rapprocher_dossier()`/`appliquer_et_archiver()` que ce retour, jamais
  séparément — sans quoi l'exclusion ne peut structurellement pas
  s'appliquer, même si le code lui-même est correct.

**Session suivante — dossier par commande dans Traités/ (BC + BL + retours),
demande explicite de l'acheteur** : "dans traités, il faudra créer un
dossier pour chaque commande [...] dedans on y met ce bon de commande, tous
les BL et bons de retours associés, et dans un temps à venir nous
ajouterons la facture. Ainsi nous aurons tout le flux commande-BL-facture
facilement consultable, et nous pourrons très facilement repérer les
écarts de facturation."

- **`a_traiter/BL/Traités/<n° de commande>/`** (`moteur.rapprochement.
  pipeline_bl._dossier_pour_commande()`) remplace l'ancien archivage à
  plat dans `Traités/` — chaque BL résolu (écrit ou déjà à jour) rejoint le
  sous-dossier de SA commande, `archiver_bl()` et `_extraire_bl_vers()`
  (fichiers multi-BL) y déposant désormais leur cible. Un numéro de
  commande absent retombe sur "Commande inconnue" (jamais un dossier vide
  ou une erreur).
- **Copie automatique du bon de commande** (`trouver_bon_de_commande()`,
  `_copier_bon_de_commande_si_absent()`) : cherche, dans l'archive externe
  des BC (`X:\...\1.3.0.1. Commandes courantes\Commandes\<année>\`,
  arborescence RÉELLEMENT mixte — fichiers à plat ET rangés par
  sous-dossier de chantier, motif de nom "<Chantier> - BC <numéro> -
  <fournisseur>.<pdf|xlsx>"), le BC correspondant au numéro de commande, et
  le copie (`"BC - <nom original>"`) dans le dossier de la commande dès le
  premier BL archivé pour elle — jamais une 2e fois (idempotent, vérifie
  qu'aucun fichier `"BC - *"` n'y est déjà). Recherche RÉCURSIVE
  (`rglob`), jamais un chemin de sous-dossier supposé. **Ne retourne un
  résultat que s'il y a EXACTEMENT UN candidat** (règle d'or : jamais un
  choix au hasard entre plusieurs BC ambigus) — silencieusement aucun si 0
  ou plusieurs, ne bloque JAMAIS l'archivage du BL lui-même (l'absence de
  BC est un simple manque, pas une erreur). Migration réelle faite cette
  session sur les 54 fichiers alors à plat dans `Traités/` : 29 bons de
  commande retrouvés et copiés, ~25 commandes sans BC trouvé (majoritairement
  des commandes très récentes pas encore filées par l'acheteur au moment de
  la migration — comportement normal, pas un bug de la recherche, vérifié
  sur un échantillon de 11 commandes avant la migration : celles trouvées
  l'ont été de façon fiable, celles absentes le sont restées après
  vérification manuelle).
- **Un bon de retour (109 Distribution) est désormais TOUJOURS considéré
  "résolu"** (`_est_resolu()`) et rejoint directement
  `Traités/<commande>/` — avant ce correctif, il restait bloqué
  indéfiniment dans "à vérifier" à cause de l'anomalie purement
  informative que `rapprocher_dossier()` lui attache systématiquement
  ("rien à écrire depuis ce document"), alors qu'un retour n'a par nature
  jamais rien à écrire et ne doit donc jamais être traité comme "en
  attente d'une décision". Nommage dédié dans `_nom_archive_bl()`
  ("RETOUR" au lieu du nom du fournisseur, avec le BL qu'il annule entre
  parenthèses) pour le repérer sans ouvrir le fichier. Cas réel validé
  cette session : le retour n°25894 (annule R9PRA263 du BL 737760, voir
  plus haut) traînait depuis sa création dans "à vérifier" sous un nom
  ambigu ("...737760... (2).pdf", un doublon de nom avec le VRAI BL 737760
  — jamais archivé lui-même faute de ce correctif) ; retraité isolément
  cette session, il a rejoint `Traités/M3.10.175/` sous le nom
  `"2026-08-14 - RETOUR - 25894 (annule BL 737760) - BC M3.10.175.pdf"`.
- **"À vérifier" reste À PLAT** (pas de sous-dossier par commande) — ce
  sont des BL qui attendent encore une décision humaine, pas un flux
  consultable au sens de la demande ci-dessus.
- **Nettoyage du backlog "à vérifier" fait cette session** (suite à la
  demande "il faut que tu supprimes les fichiers maintenant traités du
  dossier 'à vérifier'") : un diagnostic en lecture seule
  (`rapprocher_dossier` pointé sur `À vérifier/`) a distingué, parmi 34
  fichiers, 6 devenus "résolus" — 5 doublons de scan vérifiés (taille
  identique au fichier déjà archivé, ou toutes leurs lignes "déjà à jour")
  et le retour n°25894 ci-dessus (le seul à avoir réellement besoin d'un
  archivage, pas d'une suppression). **Choix délibéré : pas de suppression
  définitive par la session elle-même** (les instructions système de ce
  projet interdisent explicitement à toute session de supprimer des
  données de façon permanente, même sur demande explicite) — les 5
  doublons confirmés ont été déplacés dans
  `a_traiter/BL/À vérifier/Doublons confirmés (à supprimer)/`, à purger
  par l'acheteur elle-même quand elle le souhaite. **Restent réellement en attente dans "à vérifier"** (pas du
  bruit à nettoyer, du vrai travail) : 17 fichiers avec une anomalie de
  rapprochement authentique (sur-livraisons à investiguer, commandes
  introuvables/ambiguës, écarts de référence/tarif à trancher) et 11
  fichiers de fournisseurs sans parser BL pour l'instant (RAVATE, STAND
  64, DEM, SAGEES — voir tableau "Flux réel" plus bas) ou non reconnus par
  l'OCR.
- **Correctif immédiat, précision de l'acheteur** : "les commandes récentes
  sont générées dans `.../Commandes courantes/Commandes/BdCPDF/`, le
  nouveau dossier que nous avons créé pour la génération automatique des
  BdC" — un second emplacement (à plat, sans sous-dossier chantier),
  frère de l'archive historique par année ("2026/"), sous le même parent
  "Commandes/". Explique une bonne partie des commandes sans BC trouvé au
  premier passage (M3.10.175, 142.033, 142.034, M3.14.360/361, M2.17.005,
  M3.23.037 — toutes de vraies commandes récentes, simplement générées
  dans ce nouveau dossier plutôt que l'archive par année). `trouver_dossier_
  commandes()` (renommée depuis `trouver_dossier_commandes_annee()`) pointe
  désormais sur la racine "Commandes/" elle-même plutôt que sur un
  sous-dossier "<année>/" figé — `trouver_bon_de_commande()` étant déjà
  récursif (rglob), ça couvre "2026/" ET "BdCPDF/" (et tout futur
  sous-dossier de la même famille) sans coder son nom en dur. Backfill fait
  cette session sur les 46 dossiers `Traités/<commande>/` déjà migrés : 6
  BC supplémentaires retrouvés et copiés (142.036, M2.17.005, M3.10.175,
  M3.14.360, M3.14.361, M3.23.037) — 0 commande sans BC après ce backfill.

**Session suivante — lot de 7 BL, bug réel de fusion de lignes OCR (109
Distribution), reconstruction manuelle RAVATE, règle générale câbles/fils
au mètre confirmée par l'acheteuse (fait) :**

- **BUG RÉEL TROUVÉ — `moteur.ocr.regrouper_lignes()` peut fusionner DEUX
  rangées physiquement distinctes d'un tableau en une seule "ligne
  visuelle"** quand les rangées sont très rapprochées verticalement : la
  fonction utilise une MOYENNE MOBILE (`y_ref`) comme référence de
  tolérance (12px), qui peut DÉRIVER progressivement d'un mot à l'autre —
  chaque écart individuel reste sous les 12px mais la chaîne complète
  dérive de 30-35px, au-delà d'une rangée de tableau entière. Cas réel
  (`BL M3.13.332.pdf`, 109 Distribution) : la référence "R2V5G16ECC" (1re
  rangée) et la référence "7431256" (2e rangée, désignation "MANCHON
  JONCTION XGT7T16") se sont retrouvées fusionnées avec les prix/quantité
  de la 1re rangée dans UN SEUL groupe — `_ligne_bl_vers_article()`
  (dist109.py) a alors attribué la quantité de R2V5G16ECC (40) à la
  référence 7431256 (qui ne commandait que 5), et la DÉSIGNATION de
  7431256 ("MANCHON JONCTION XGT7T16") s'est retrouvée traitée comme une
  fausse "référence" à part entière par le rapprochement — même schéma
  sur la ligne suivante (16042708/59210). Le rapprochement automatique a
  donc classé ce BL à 0% correct : 2 "sur-livraisons" fantômes (40 pour 5
  commandé, 30 pour 1 commandé) et 2 "références inconnues" fantômes
  (en réalité des bouts de désignation). **Détecté par l'acheteuse qui a
  contesté l'analyse initiale** ("M3.10.186 n'a pas de problème pour moi"
  — un cas voisin qui a motivé une relecture plus attentive de tout le
  lot) : reconstruit à la main en recoupant chaque quantité candidate
  contre la quantité commandée ET contre le Total HT imprimé (518,00€) —
  les 4 vraies valeurs (R2V5G16ECC=40/11,75€, 7431256=5/0,70€,
  16042708=30/0,85€, 59210=1/19,00€) retombent EXACTEMENT sur les 2
  signaux à la fois. **Pas de correctif de code** (un seul exemple à ce
  jour, `regrouper_lignes()` est partagé par TOUS les fournisseurs et
  déjà éprouvé sur des dizaines de vrais documents — la modifier
  risquerait de régresser des fixtures déjà validées ; règle d'or) —
  traité comme une correction manuelle ponctuelle. **À surveiller** : tout
  futur BL avec des rangées de tableau anormalement rapprochées pourrait
  reproduire ce symptôme (quantité/référence qui ne matche ni le total ni
  la commande, alors que le Total HT global reste cohérent une fois
  reconstruit à la main) — le réflexe est de recouper chaque ligne contre
  le Total HT imprimé ET contre la quantité commandée avant de conclure à
  une vraie sur-livraison ou une vraie référence inconnue.
- **RAVATE (`BL M3.27.005.jpg`), scan de qualité moyenne, cellules très
  éclatées** — 1 seule ligne extraite automatiquement (915210, qté=30)
  pour un Total HT de 584,82€ contre 162,00€ extraits. Reconstruit à la
  main en recoupant chaque bloc de cellules contre sa quantité commandée
  et contre le solde du Total HT : la quantité "30" appartenait en
  réalité à la référence **915205** (qte_cmd=30, 30×5,40=162,00€ pile),
  pas à 915210 (qte_cmd=9) — la ligne automatique avait donc le mauvais
  couple référence/quantité. Les 2 lignes restantes (915210=9/6,98€,
  062525=9/40,00€) déduites par élimination : 62,82+162,00+360,00 =
  584,82€ = Total HT imprimé exactement. Aucune des 3 lignes n'était donc
  une vraie sur-livraison — encore un cas de mauvaise association
  référence/quantité, cette fois côté RAVATE (scan visiblement plus
  dégradé que les autres BL du même lot, déjà connu comme point fragile
  documenté de ce fournisseur).
- **RÈGLE GÉNÉRALE confirmée par l'acheteuse, à appliquer chaque fois
  qu'un article de type câble/fil est livré en COURONNE (unité imprimée
  "COU" chez RAVATE, distincte de "UN")** : la quantité à enregistrer dans
  le Suivi est en MÈTRES LINÉAIRES, pas en nombre de couronnes — "de
  manière générale, tous les prix des câbles et fils s'entendent en
  mètres linéaires de notre côté." Cas réel (`BL M3.10.187.jpg`,
  référence RAVATE tronquée par l'OCR "R2V3G1.5C1", vraie référence
  probable "R2V3G1.5C100" d'après la désignation "...3G1.5C1O0" = C100) :
  1 couronne de 100m livrée, Suivi "R2V3G1.5T1" attend qte_cmd=100 —
  enregistré qté=100 (pas 1) et tarif ramené au mètre (70,00€/couronne →
  0,70€/m), jamais deviné seule (même famille que le cas boîte/pièce
  Coredime déjà documenté : conversion d'unité toujours confirmée par
  l'acheteuse, jamais automatisée — ici confirmée explicitement).
  **Portée** : ce n'est PAS qu'un cas RAVATE — vérifier l'unité imprimée
  (ou son équivalent) sur tout futur article câble/fil, quel que soit le
  fournisseur, avant d'enregistrer une quantité qui semble anormalement
  petite face à la commande.
- **Fichiers déposés au mauvais endroit** (`a_traiter/` au lieu de
  `a_traiter/BL/`) — déplacés manuellement avant traitement ; à vérifier
  systématiquement si le dossier `a_traiter/BL/` semble vide après un
  dépôt annoncé par l'acheteuse.
- **Recette réelle sur les 7 fichiers** : 21 lignes écrites (14 sûres
  automatiquement + 7 reconstruites à la main), 7 BL archivés avec leur
  BC.

**Session suivante — traitement du reste du backlog "à vérifier", fix
Electric Plus factures multiples (fait)** : suite directe de la session
précédente, l'acheteur a traité plusieurs cas un par un.

- **109 Distribution 131.157** : l'acheteur a corrigé elle-même la
  référence dans le Suivi (le cas GOUJON8X70→75 documenté plus haut,
  "trou de perforateur" mais en réalité une vraie substitution) — une fois
  la référence Suivi alignée sur celle du BL, la ligne GOUJON8X75 devient
  une correspondance EXACTE (SUR), plus besoin de repli. Les deux autres
  lignes du même BL (rondelles, boulons) avaient une légère sur-livraison
  (500/300 et 4/3) — écrites avec les autres sur confirmation implicite de
  l'acheteur ("normalement plus de pb pour traiter ce BL").
- **Cominter M2.22.084** : commande déduite confirmée par l'acheteur — 12
  lignes SUR/DEJA_A_JOUR écrites, puis les 3 lignes de repli restantes
  (L3332→033325, QU152502645→QUI52502645, L8005→080052) également
  confirmées et écrites dans un 2e passage. Seule 'SY00ZU51' reste
  bloquée : corruption OCR trop sévère (probablement '033325'... non,
  probablement une désignation de SY0029651 — écart trop important pour
  un repli fiable, aucune règle codée sur un seul cas, voir règle d'or) —
  le fichier reste dans "à vérifier" pour cette seule ligne.
- **Cas GMR→COMINTER (commande M3.15.397), fournisseur substitué en
  rupture de stock** : cas réel signalé par l'acheteur — "GMR était en
  rupture, nous avons pris chez COMINTER en utilisant le même n° de BdC."
  Un BL Cominter réel (OBL107273) ne trouvait aucune correspondance parce
  que le Suivi avait cette commande enregistrée sous fournisseur "GMR" —
  la recherche `lire_lignes_commande(fichier_suivi, bl.fournisseur, ...)`
  filtre STRICTEMENT par fournisseur, comme il se doit (jamais de choix au
  hasard entre deux fournisseurs différents). Comme la colonne
  "Fournisseur" du Suivi n'est PAS dans `COLONNES_MODIFIABLES` (liste
  blanche volontairement restreinte, voir plus haut — "jamais de
  compromis"), ce n'est PAS un cas que Rapprochement AI peut corriger
  lui-même : l'acheteur a changé "GMR" en "COMINTER" à la main dans
  Excel, après quoi le rapprochement a immédiatement reconnu la ligne
  comme "déjà à jour" (la quantité et le tarif y étaient déjà, seul le
  fournisseur bloquait le rapprochement) — rien de plus à écrire, juste
  à archiver. **Leçon pour la suite** : un n° de commande introuvable sous
  le fournisseur détecté sur le BL vaut la peine d'être cherché aussi sous
  d'AUTRES fournisseurs avant de conclure "commande introuvable" — un
  fournisseur peut légitimement livrer via un concurrent en cas de rupture,
  avec le même n° de commande.
- **BUG RÉEL CORRIGÉ — Electric Plus : plusieurs FACTURES distinctes dans
  un même fichier PDF** (signalé directement par l'acheteur : "article
  11527 sur la commande, PLA11527 sur le BL : quel est le souci ?").
  `doc07149220260814105422.pdf` (2 pages) contenait en réalité DEUX
  factures Electric Plus sans rapport (commande/n° facture/date différents
  par page — page 0 = commande M4.263, facture 1207019 ; page 1 = commande
  142.036, facture 1207127). Avant ce correctif, `parse_bl_electricplus()`
  traitait tout le fichier comme UN seul document : le n° de commande et
  le Total HT venaient de la page 0, mais les LIGNES d'articles
  n'étaient JAMAIS filtrées par page — un vrai risque de mélanger deux
  commandes sous un seul numéro. Deux correctifs dans
  `moteur/fournisseurs/electricplus.py` :
  1. **Découpage par page** (même principe que 109 Distribution/Cominter,
     voir `moteur.ocr.pages_par_identifiant`) — `parse_bl_electricplus()`
     retourne désormais une LISTE. **Piège spécifique à ce fournisseur** :
     l'identifiant de regroupement ne peut PAS réutiliser
     `MOTIF_FACTURE_DATE_ELECTRICPLUS` (motif à 3 cellules adjacentes
     date/date-échéance/n°facture) car `pages_par_identifiant` cherche son
     motif dans le texte BRUT de la page (mots OCR simplement joints par
     un espace dans leur ORDRE DE LECTURE OCR — PAS réordonnés en lignes
     visuelles comme le fait `regrouper_lignes()`, utilisé lui pour
     l'extraction des champs). Sur ce document réel, les 3 cellules
     n'apparaissent jamais adjacentes dans le flux brut → le motif ne
     matchait JAMAIS, donc les 2 pages fusionnaient toujours en un seul
     groupe. Nouveau `MOTIF_IDENTIFIANT_PAGE_ELECTRICPLUS` : le n° de
     facture SEUL (`\b\d{6,7}\b`), qui apparaît bien comme un token isolé
     fiable dans le flux brut sur ce document (vérifié qu'aucun autre
     nombre de 6-7 chiffres — siret, téléphone, code postal — ne peut
     matcher par coïncidence, tous plus longs ou collés à des lettres sans
     transition \w/\W).
  2. **Repli positionnel pour les lignes sans suffixe "PF"** : la ligne
     PLA11527 de la page 0 n'avait AUCUNE cellule "<prix>PF" (contrairement
     à toutes les lignes vues jusqu'ici chez ce fournisseur) — silencieusement
     ignorée avant ce correctif. Repli dans `_ligne_vers_article_electricplus`
     quand l'ancre "PF" est absente : lit les 4 dernières cellules de la
     ligne dans l'ordre du tableau (QTE, PRIX UNIT.HT, P.U.NET HT, MONTANT
     HT), jamais depuis le début (le nombre de cellules de désignation
     varie). Validé uniquement par cohérence arithmétique (60 × 2,68 =
     160,80€, correspond exactement au Total HT affiché de la page) — un
     seul exemple réel, à surveiller si un 2e cas se présente (règle d'or).
  Fixture réelle ajoutée (`tests/fixtures/bl_electricplus_9_multi_facture_
  2pages.pdf`), testé (`test_parse_bl_electricplus_9_deux_factures_dans_un_
  seul_fichier`). Recette réelle : une fois corrigé, M4.263 (PLA11527, qté
  60) écrit et archivé, 142.036 confirmé déjà entièrement à jour (doublon
  d'une facture déjà traitée).
- **Coredime 131.161** : sur-livraison 200 livrées pour 2 commandées sur
  une référence dont la désignation dit littéralement "BOITE DE 100" —
  même schéma boîte/pièce déjà documenté et confirmé par l'acheteur
  (session précédente) — écrit sans nouvelle confirmation, le motif est
  déjà établi.
- **Coredime 142.035** : 2 BL du même lot (VAUBAN.pdf et VAUBAN1.pdf)
  ciblaient la MÊME ligne Suivi (LEG406771) — le garde-fou anti-doublon
  (`_desamorcer_conflits_meme_ligne_suivi`) a fonctionné comme prévu, seul
  le premier fichier a été écrit, le second reste "à confirmer" pour cette
  ligne. L'AUTRE ligne de VAUBAN1.pdf (227060122 vs Suivi 227060124, repli
  1 caractère mais similarité de désignation faible ~65% — "CHEV..." vs
  "ANC...") reste elle aussi en attente, la correspondance semble moins
  fiable que les autres cas de repli déjà validés cette session.
- **Electric Plus 155.009 : commande introuvable, y compris sous GMR**
  (vérifié — contrairement au cas M3.15.397 ci-dessus, aucune ligne ne
  correspond ni sous ELECTRIC PLUS ni sous GMR) — l'acheteur a confirmé
  que cette commande avait été passée ORALEMENT (jamais saisie dans le
  Suivi avant cette livraison) et l'a ajoutée elle-même.
- **BUG RÉEL CORRIGÉ — Cominter : remise et Px net collés SANS espace**
  (signalé directement par l'acheteur, très fermement — "nous avons le
  prix unitaire, puis la remise, puis le prix net, aucun écart. Je crois
  que tu as un problème avec ton parser Cominter", à raison). Distinct du
  cas déjà connu et testé ("30% 110,67", AVEC un espace, voir
  `test_parse_bl_cominter_4`) : sur BL ANZEMBERG.pdf/OBL108110 (commande
  135.039), la cellule remise+Px net est collée SANS AUCUN espace
  ("30%435,94", "30%106,50") — `_prix_net_bl_cominter()`
  (`moteur/fournisseurs/cominter.py`) exigeait `\s+` (au moins un espace)
  entre le "%" et le prix dans son motif combo ; sans espace, le motif ne
  matchait JAMAIS, et la boucle de recherche retombait SILENCIEUSEMENT sur
  le Px UNITAIRE (brut, avant remise) de la cellule précédente — aucune
  anomalie levée, car montant et quantité restaient parfaitement cohérents
  avec ce mauvais prix pris isolément. Fix : `\s+` → `\s*` dans le motif
  combo (accepte 0 espace comme avant). **Deux lignes réelles touchées sur
  ce seul document** : CAELK2766 (622,77€ lu au lieu de 435,94€, jamais
  encore écrite dans le Suivi — corrigée en écrivant la bonne valeur) ET
  CAEACB4V (152,14€ au lieu de 106,50€, **déjà écrite dans le Suivi lors
  d'une session antérieure** — corrigée directement sur le classeur vivant
  après découverte). **Deuxième confirmation indépendante du même bug**,
  trouvée en creusant `doc07148920260814105315.pdf` (une 2e numérisation
  du même BL physique, commande M3.10.171/OBL108154) : cette ligne
  (AEA9Y13625/MEA9Y13625, remise "30%158,10" collée sans espace sur CE
  scan précis) avait donné le même symptôme (tarif lu 158,10€ au lieu de
  110,67€ déjà correctement enregistré depuis l'AUTRE scan de la même
  facture, `bl_cominter_4.pdf`, où l'espace était présent) — confirme que
  le bug touche plusieurs documents réels, pas un cas isolé. Fixture
  réelle ajoutée (`tests/fixtures/bl_cominter_7_remise_sans_espace.pdf`),
  testé (`test_parse_bl_cominter_7_remise_et_px_net_colles_sans_espace`).
  **Point de vigilance pour la suite** : ce document a aussi 2 lignes
  (CAEMMCPF1U4CROG, CAECORD6ASF005MSH) où l'OCR colle la référence ET la
  désignation dans une seule cellule — reste non extrait, limitation
  préexistante déjà connue (voir "Points fragiles", cœur du sujet des
  "3/6 lignes concordantes" documenté à l'origine de ce BL), pas partie de
  ce correctif.
- **M3.23.030 : cas laissé EN SUSPENS après correction directe de
  l'acheteur, à ne pas re-flaguer comme sur-livraison sans réexamen** —
  deux documents réels (`doc07136020260813090605.pdf`/OBL108106, daté
  06/08/2026, ligne 405209 qté 10 ; `doc07136120260813090616.pdf`/
  OBL108251, ligne tronquée "L4052"→405209 probable qté 5) avaient été
  signalés à tort comme "sur-livraison" (10+5=15 de plus par-dessus les
  15 déjà enregistrés dans le Suivi, daté 11/08/2026 — soit 25 pour 15
  commandés). L'acheteur a fermement contredit : "aucune sur-livraison...
  il manque 2 unités d'un article" — correspond exactement à la ligne
  BLM680527 de cette même commande (28 livrées pour 30 commandées dans le
  Suivi, un manque réel et accepté, pas une erreur). **Hypothèse non
  vérifiée à ce jour** : les deux documents 06/08 (10+5=15, exactement le
  total déjà enregistré au 11/08) documentent peut-être la MÊME livraison
  déjà comptée sous une date différente, plutôt qu'un ajout — mais les
  dates ne correspondent pas (06/08 vs 11/08 déjà enregistré), donc pas
  confirmé avec certitude. Aucune écriture faite pour ces deux documents
  (ni la version "sur-livraison" refusée par l'acheteur, ni aucune autre
  hypothèse non confirmée) — laissés tels quels dans "à vérifier" pour un
  réexamen ultérieur, plutôt que de deviner (règle d'or).

**Session suivante — RAVATE, nouveau fournisseur BL (fait), sur demande
explicite de l'acheteur** (choix entre RAVATE/STAND 64/DEM/SAGEES) : les 6
seules pièces réelles disponibles (2 initiales + 4 déposées en cours de
session) sont TOUTES des "ravatelec" (branding visible sur le document) —
RAVATE PRO reste NON couvert côté BL, contrairement au devis où l'acheteur
avait confirmé une structure identique (aucune pièce PRO disponible pour
vérifier côté BL, l'acheteur a explicitement signalé la distinction en
cours de session — pas supposé sans confirmation).

- **Structure découverte** : chaque article est réparti sur PLUSIEURS
  lignes visuelles OCR, avec DEUX codes différents qui cohabitent — le
  "Code Art" (interne Ravate, TOUJOURS numérique pur, préfixe "100"
  [9 chiffres] ou "44200" [8 chiffres] sur toutes les pièces vues) et la
  "Référence fournisseur" (le vrai code métier, ex. "VK16BT", "R2V5G10",
  "404926") — **leur position l'un par rapport à l'autre VARIE selon le
  scan** (Code Art tantôt sur la ligne désignation, tantôt en tête de la
  ligne chiffrée ; Référence fournisseur tantôt isolée sur sa propre
  ligne, tantôt en tête de la ligne chiffrée à la place du Code Art).
  Comme pour le devis de ce fournisseur ("toujours la Réf. FNR, jamais la
  réf interne", règle métier déjà confirmée), `LigneBL` utilise TOUJOURS
  la Référence fournisseur — repérée en EXCLUANT explicitement la forme
  du Code Art, jamais l'inverse (la Réf. fournisseur n'a pas de forme
  unique fiable). **Validé sur données réelles** : 14 des 16 lignes des 4
  pièces fraîches ont matché EXACTEMENT une référence déjà présente dans
  le Suivi (statut SUR immédiat, aucun repli nécessaire) — forte
  confirmation que l'extraction cible la bonne colonne.
- La ligne CHIFFRÉE est repérée par ses 4 DERNIÈRES cellules — Px Brut |
  Px Net | Remises (montant €, pas un pourcentage) | Montant HT, toujours
  dans cet ordre — plutôt que par un nombre de cellules fixe (5 ou 7 selon
  que qté+unité soient inline ou sur une ligne à part) : vérifié par
  cohérence arithmétique EXACTE sur les 6 pièces (ex. Remises =
  (Px Brut − Px Net) × Qté, à l'euro près). Quantité livrée TOUJOURS
  déduite de Montant / Px net (même logique que 109 Distribution/
  Cominter/Electric Plus), jamais lue dans une cellule Qté dont la
  position varie.
- **3 bugs réels trouvés et corrigés en confrontant les 4 pièces fraîches**
  (les 2 pièces initiales seules avaient donné un faux sentiment de
  fiabilité — 0 ligne extraite sur les 4 nouvelles avant ces correctifs) :
  1. En-tête "Reference fournisseur" parfois éclaté en plusieurs cellules
     OCR adjacentes ("Reference" | "fournisseur") au lieu d'une seule —
     la recherche de zone de tableau doit joindre la ligne visuelle
     ENTIÈRE avant de chercher le motif, jamais cellule par cellule.
  2. Le motif "BC n°... AU <date> <commande>" doit tolérer un collage
     total (aucun espace nulle part, y compris entre "AU" et la date, et
     entre la date et la commande — ex. "BC000312608CC0405AU18/08/2026
     M3.10.182" en une seule cellule OCR).
  3. Une virgule de prix peut être lue "*" par l'OCR ("99*91" au lieu de
     "99,91", confirmé par cohérence arithmétique : 130 × 6,75 = 877,50,
     cohérent avec un Px Brut ≈ 99,91 avant remise) — toléré comme
     séparateur décimal au même titre que "," et ".".
- **Recette réelle sur les 4 pièces fraîches** : 14 lignes SUR écrites et
  archivées (`Traités/M3.23.033/`, `Traités/M3.10.182/`), 1 ligne à
  confirmer (date enregistrée 07/08 vs date réelle du BL 10/08, qté/tarif
  déjà corrects — repli date-incohérente déjà existant, voir plus haut) et
  1 ligne inconnue (`XVR1IISTI`, borne de recharge — référence absente du
  Suivi pour cette commande, plausible pour un article vraiment nouveau).
Fixtures réelles ajoutées (`tests/fixtures/bl_ravate_1.jpg` à
`bl_ravate_6_huit_lignes.pdf`), testé (`tests/test_parsers_bl_ravate.py`,
8 tests).

**2 bugs réels de "bout de chaîne" trouvés par l'acheteur juste après,
tous deux corrigés — leçon générale : le rapprochement + l'écriture ne
suffisent pas, il faut vérifier le RANGEMENT jusqu'au bout à chaque
session, "sinon nous ne pouvons pas valider" (mot de l'acheteur) :**
- **BL ANZEMBERG.pdf (Cominter, OBL108110/135.039) jamais archivé** :
  les 2 corrections de prix (voir bug remise/Px net ci-dessus) avaient été
  écrites directement dans le Suivi via un script ad hoc, mais SANS passer
  par `appliquer_et_archiver()` — le fichier restait donc oublié dans
  "à vérifier" alors que ses 7 lignes extractibles étaient déjà toutes à
  jour. Retraité via le pipeline officiel : archivé dans
  `Traités/135.039/`. **Leçon** : toute correction manuelle (script ad hoc
  hors du flux normal) doit être suivie d'un passage par
  `appliquer_et_archiver()` sur le fichier source, jamais laissée à mi-chemin.
- **BUG RÉEL CORRIGÉ — `trouver_bon_de_commande()` traitait deux copies
  IDENTIQUES du même BC comme une ambiguïté** (signalé par l'acheteur :
  "M3.23.033 est rangé tout seul, sans son BdC"). Le même BC se retrouve
  couramment archivé à la fois dans `Commandes/<année>/` (ex.
  "Maintenance/") ET dans `Commandes/BdCPDF/` (filé après coup dans
  l'archive historique) — deux fichiers de MÊME NOM ET MÊME TAILLE, donc
  PAS une vraie ambiguïté (contrairement à deux BC de contenu réellement
  différent, où ne rien choisir reste la bonne règle). `trouver_bon_de_
  commande()` déduplique désormais les candidats par (nom, taille) avant
  d'exiger l'unicité. Backfill fait sur tous les dossiers `Traités/`
  existants après le correctif : 1 commande retrouvée (M3.23.033) — 0
  restante sans BC. Testé
  (`test_trouver_bon_de_commande_deux_copies_identiques_pas_ambigu`).
- **Nouveau repli de rapprochement — confusion OCR "1"/"I"** (signalé par
  l'acheteur : "XVR1IISTI, exactement le même article que sur le bon de
  commande et dans le tableau suivi, quel est le souci ?"). Suivi
  "XVR111STI" ("WittyOne 11kW") lu "XVR1IISTI" par l'OCR — DEUX des trois
  "1" confondus avec la lettre majuscule "I" dans la MÊME référence
  (distance de 2 caractères, au-delà du seuil de `_distance_courte` qui ne
  tolère qu'1 caractère d'écart). Nouvelle fonction `_cle_normalisee_i_un()`
  (`matching.py`) : normalise "I"→"1" des deux côtés puis compare par
  ÉGALITÉ (pas une distance à 1 près — plusieurs confusions 1/I peuvent
  survenir dans la même référence, comme ici). Intégré comme 3e critère de
  `_repli_reference_proche()`, indépendant du cœur numérique (ici les deux
  cœurs sont sous le seuil de 4 chiffres : "1" seul côté BL, "111" côté
  Suivi). Toujours "à confirmer", jamais "sûr". Écrit et archivé sur le cas
  réel après validation. Testé
  (`test_apparier_repli_confusion_ocr_1_i`,
  `test_apparier_repli_confusion_ocr_1_i_ignore_si_deja_identique`).
- **BUG RÉEL CORRIGÉ — RAVATE : Référence fournisseur en LOOK-AHEAD, pas
  seulement look-behind** (trouvé en traitant pour de vrai "BL 123.095
  MORANE.jpg", jamais écrit avant cette session malgré avoir servi de
  fixture de test dès le début). Confirmé que le Suivi attend
  "R2V3G2.5T1" (pas "44200019" le Code Art) pour cette commande — cette
  Référence fournisseur se trouve UNIQUEMENT sur la ligne SUIVANTE ("
  R2V3G2.5T1 | :MT:100.00"), jamais avant, cas non couvert par le repli
  déjà en place (qui ne mémorisait qu'une référence isolée PRÉCÉDENTE).
  `parse_bl_ravate()` matérialise désormais la zone de tableau en liste
  (au lieu d'itérer un générateur) pour pouvoir regarder UNE ligne en
  avant quand une ligne chiffrée retombe sur un Code Art faute de
  référence isolée précédente trouvée — jamais plus loin qu'une ligne
  (pas de "presque"). A aussi corrigé `bl_ravate_2.jpg` (même schéma,
  référence "069864" trouvée en avant plutôt qu'en Code Art "100151156") —
  tests mis à jour en conséquence, comportement plus correct qu'avant.

**Recette réelle finale de cette session (Cominter + Ravate)** : chaque
fichier RÉELLEMENT écrit ou déjà à jour cette session a été vérifié
individuellement dérouler jusqu'au bout — rapprochement ET rangement
(`Traités/<commande>/` + son BC) — pas seulement l'écriture Suivi (leçon
de la session, voir "2 bugs réels de bout de chaîne" ci-dessus : un
rapprochement sans rangement n'est PAS considéré terminé). Nettoyage
associé : ~13 fichiers identifiés comme doublons de contenu déjà archivé
(vérifiés par taille de fichier avant déplacement, jamais par nom seul)
déplacés dans `À vérifier/Doublons confirmés (à supprimer)/`, jamais
supprimés directement par la session (voir plus haut, "pas de suppression
définitive par la session elle-même").

**Session suivante — gros lot de BL du jour, fichier multi-fournisseur
découvert et traité (fait)** : l'acheteur a déposé 11 nouveaux fichiers
d'un coup ("tous les nouveaux BL du jour").

- **BUG RÉEL CORRIGÉ — RAVATE, tiret dans la commande** : "131-162"
  (commande collée directement après l'année sans espace, "2026131-162")
  contient un TIRET — la classe de caractères du motif de capture de
  commande (`MOTIF_BC_COMMANDE_BL_RAVATE`) ne l'incluait pas, donc la
  capture s'arrêtait juste avant et le `\s*$` final ne matchait plus DU
  TOUT : la ligne entière échouait, commande introuvable ET la ligne
  "BC n°..." elle-même se retrouvait accumulée à tort en tête de la
  désignation du 1er article (même motif réutilisé pour l'exclusion).
  "-" ajouté à la classe de caractères.
- **BUG RÉEL CORRIGÉ — RAVATE, confusion OCR "8"/"B" dans la commande**
  (commande M3.10.182, confirmée par le nom de fichier du BL) : "18/0B/2026
  M3.10.1B2" — le jour et le mois de la date, ET la commande elle-même,
  peuvent contenir des "B" à la place de "8". Motif élargi à `[\dB]`,
  substitution "B"→"8" après capture.
- **BUG RÉEL CORRIGÉ — RAVATE, qté+unité ET Px Net+Remises collés,
  simultanément** (même document M3.10.182) : la cellule qté+unité
  peut être collée en une seule cellule juste avant Px Brut
  (":MT:130,00:"), décalant la fenêtre des "4 dernières cellules" —
  repli sur les 3 dernières (Px Brut, Px Net, Montant) quand les 4
  échouent. Et Px Net + Remises peuvent être collés SANS espace dans
  UNE cellule ("6,75:1288,30", à ne pas confondre avec le cas déjà connu
  "%prix" — ici ce sont deux MONTANTS complets séparés par ":") — repli
  combo dédié, extrait le 1er des deux. Qté déduite de Montant/Px Net =
  877,50/6,75 = 130 pile, cohérent avec le "130,00" imprimé (jamais lue
  directement).
  Fixtures réelles ajoutées (`bl_ravate_7_tiret_dans_commande.jpg`,
  `bl_ravate_8_confusion_ocr_8_b_dans_commande.jpg`), testé (10 tests
  au total pour ce fournisseur).
- **DÉCOUVERTE MAJEURE — un fichier PDF peut mélanger PLUSIEURS
  FOURNISSEURS différents, pas seulement plusieurs BL du même
  fournisseur** (`doc07178620260819170752.pdf`, 10 pages : RAVATE p0,
  COMINTER p1/p2/p7/p9, 109 DISTRIBUTION p3/p6/p8, ELECTRIC PLUS p4,
  COREDIME p5 — probablement un lot de BL papier scannés ensemble en une
  seule fois). AUCUN parser mono-fournisseur ne peut traiter un tel
  fichier : `lire_bl()` détecte UN SEUL fournisseur pour tout le fichier
  (le premier trouvé) et applique CE SEUL parser aux 10 pages, produisant
  0 ligne exploitable pour 9 des 10 pages. **Pas d'architecture générale
  construite pour ce cas** (un seul exemple à ce jour, règle d'or) : traité
  manuellement cette session (script ad hoc, pas dans le code du projet) —
  découpage par page, bon parser appelé sur chaque tranche de 1 page,
  rapprochement et archivage individuels via les primitives existantes
  (`_extraire_bl_vers`, `_dossier_pour_commande`,
  `_copier_bon_de_commande_si_absent`). **Si ce schéma se reproduit**
  (lot de BL scannés ensemble), envisager une détection par-page du
  fournisseur dans `lecture_bl.py` plutôt que par fichier entier.
  Résultat de cette page multi-fournisseur : 4 pages entièrement résolues
  et archivées (M3.10.177, 142.037, M4.266, 131.171), 6 pages vers
  "à vérifier" (dont 3 juste extraites sans rapprochement — page RAVATE
  à l'OCR trop dégradé pour être fiable, et 2 commandes Cominter
  (130.636, M4.262) introuvables dans le Suivi même sous GMR).
- **BUG RÉEL CONFIRMÉ (2e occurrence) — regroupement multi-BL 109
  Distribution rate des identifiants différents mal lus** (`BL 131.169
  LAGOURGUE.pdf`, 2 pages, 2 BL réels distincts 739166/739111) :
  `pages_par_identifiant` a fusionné les 2 pages en UN SEUL groupe au
  lieu de 2, parce que le 2e numéro était lu "739 111" (espace au
  milieu, cassant le motif `\d{4,7}` qui exige une suite CONTIGÜE de
  chiffres) — même famille que le bug déjà documenté pour Ravate (Q à la
  place de 0). Contourné cette session en appelant `parse_bl_109()`
  page par page individuellement plutôt que sur le fichier entier — pas
  de correctif général du motif tenté (un seul exemple de plus, pas
  assez pour une règle fiable sur CE fournisseur précis).
- **2 corrections manuelles motivées (pas un repli automatique standard,
  écart trop important) sur ce même fichier 131.169** :
  1. `CAP493450` : Total HT lu "40,00" au lieu de "140,00" (chiffre "1"
     de tête disparu) — confirmé par l'autocontrôle GLOBAL du document
     (965,00€ affiché vs 865,00€ extrait, exactement 100€ d'écart) ET
     par la cohérence locale (Qté imprimée "200" × Px Net 0,70 = 140,00,
     PAS 40,00) ET par le Suivi (qte_cmd=200, exactement). Trois signaux
     indépendants convergents → corrigé manuellement (Montant=140,
     Qté recalculée=200).
  2. `'100/'` → `'53041'` : la vraie référence était sur la ligne de
     désignation PRÉCÉDENTE (désignation trop longue, déborde sur 2
     lignes — dist109 n'a pas de repli de raccord comme Coredime/Ravate
     pour ce cas), la ligne chiffrée ne portait qu'un fragment illisible.
     Qté/prix/montant déjà corrects (confirmé par l'autocontrôle DE CETTE
     SEULE PAGE : 150,00€ affiché = 150,00€ extrait).
- **2 corrections manuelles motivées sur le fichier multi-fournisseur**
  (page 2, COMINTER M3.10.177) :
  1. `'CACA/4040B'` → `'CACAI4040B'` : le Suivi a DEUX références au même
     cœur numérique ("4040"), ambiguïté non résolue automatiquement ;
     la désignation Suivi ("Angle Intérieur pour 40/40") correspond
     EXACTEMENT à la désignation BL ("AngleInterieurpour40/40") pour
     CETTE ligne — "/" est probablement un "I" mal lu par l'OCR.
  2. 2e ligne du même BL jamais extraite DU TOUT (référence coupée en 2
     cellules OCR "CAL" + "040B", ne formant ni "CACAI4040B" ni
     "CACAP4040B" reconnaissable) — ajoutée manuellement :
     `'CACAP4040B'` qté=1, prix_net=1,68€ (30% de remise sur 2,40€,
     cohérent avec le document), montant=1,68€ — correspond exactement
     à `qte_cmd=1` au Suivi pour cette référence.
  (page 8, 109 DISTRIBUTION M4.265) : `'243454'` → `'6133461'` — la
  désignation BL contient LITTÉRALEMENT "(6133461)" en suffixe
  ("JONCTIONDEPLAFONDCOFRALIS(6133461)"), qui est la vraie référence
  catalogue du Suivi ; qté=4 correspond exactement à `qte_cmd=4`.
- **Panne réseau transitoire pendant la session, résolue** : le lecteur
  X: (partage réseau) est devenu temporairement inaccessible (VPN coupé
  côté acheteur) — tous les outils shell (Bash ET l'exécution Python)
  ont échoué silencieusement (code de sortie 1, aucune sortie) pendant
  cette fenêtre. Diagnostiqué via PowerShell (qui gardait un accès
  fonctionnel), confirmé par l'acheteur (VPN reconnecté), le shell Bash
  s'est rétabli seul sans action corrective supplémentaire. Aucune
  donnée perdue ni corrompue (les écritures en cours n'avaient pas
  encore démarré au moment de la coupure).
- **2 fournisseurs restent introuvables au Suivi malgré vérification**
  (`M2.23.057`/109D, ref "300007" vs Suivi "73934" — aucun rapport
  évident ; `M3.23.040`/RAVATE, ref "089281" vs Suivi "EUR52301" —
  idem) — laissés en l'état dans "à vérifier", pas de règle inventée
  sans preuve suffisante.

**Seules 3 colonnes de la feuille "Commandes" sont de vraies données
saisies** (vérifié cellule par cellule sur le vrai classeur, pas supposé) :
"Date de livraison", "Qté livrée", "Tarif BL" — plus "Note" (texte libre
utilisé comme valeur "magique" par plusieurs formules : "Rupture
fournisseur", "Reliquat soldé", "Commande annulée"). **Tout le reste de la
feuille est une FORMULE calculée à partir de ces colonnes** — notamment
"Statut commande" (XLOOKUP/IF en cascade), "Reliquat", "RAL", "Soldé",
"Reste à facturer", "Facturé BL" (qui, malgré son nom, est le MONTANT
facturé calculé = Tarif BL × Qté livrée, pas un indicateur "facture
reçue"). Conséquence directe pour tout code futur de cette branche :
**n'écrire QUE dans ces 3+1 colonnes** ; le Statut/Reliquat/Soldé se
recalculent seuls à la prochaine ouverture Excel. Comment le rapprochement
factures (étape hebdomadaire) doit se traduire concrètement sur cette
feuille — puisqu'il n'existe aucune colonne "facture reçue" explicite à
cocher — reste une question ouverte pour l'acheteur (voir session R1,
tableau de flux) : probablement une correction de "Tarif BL" si la facture
diffère du BL, à confirmer.
*Mise à jour P1 (2026-09-04, voir « Feuille Pièces — modèle, socles,
migration ») : les 5 colonnes facture créées le 2026-09-01 ne sont PLUS
des colonnes saisies mais des formules ; la feuille « Pièces » (une ligne
par ligne de document) est ÉCRITE PAR L'OUTIL, pas saisie — seule sa
colonne « Commentaire » est humaine. Les 3+1 colonnes BL ci-dessus restent
les seules cellules de Commandes que l'outil écrit.*

**Écriture sécurisée (`moteur/rapprochement/ecriture.py`)** : openpyxl
(`load_workbook()` + `save()`) a été essayé en premier sur une COPIE du
vrai Suivi commandes et écarté — même sans toucher qu'à une seule cellule,
la réécriture complète fait disparaître `xl/calcChain.xml`,
`xl/metadata.xml`, `customXml/*`, les `printerSettings/*.bin`, et dégrade
des validations de données (avertissement obtenu sur le classeur réel :
*"Data Validation extension is not supported and will be removed"*) —
inacceptable sur un fichier vivant riche en formules, 16 Excel Tables et
validations. Le module patche donc directement la partie XML de la feuille
visée DANS LE ZIP (.xlsx = zip OOXML), sans passer par le sérialiseur
openpyxl : chaque autre partie du zip (styles, tableaux, validations,
calcChain, sharedStrings, customXml, printerSettings...) est recopiée
octet pour octet. Prouvé sur une copie du vrai classeur
(`tests/test_rapprochement_ecriture.py::test_ecriture_chirurgicale_sur_le_vrai_suivi_commandes`,
ignoré si le fichier est absent du poste, même pattern que
`test_panier.py`) : seule la partie `xl/worksheets/sheet1.xml` change,
tout le reste est identique bit à bit.

Trois garde-fous dans `appliquer()`, dans cet ordre :
1. **Verrou Excel** (`est_verrouille()`) : présence d'un fichier `~$<nom>`
   à côté du classeur -> `ClasseurVerrouille` levée, jamais d'écriture en
   force.
2. **Sauvegarde horodatée** (`sauvegarder()`) dans `backups/` AVANT toute
   écriture, rotation à 30 jours (`RETENTION_BACKUPS_JOURS`).
3. **Liste blanche de colonnes** (`COLONNES_MODIFIABLES`) : toute
   `Ecriture` visant une colonne hors de cette liste lève
   `ColonneNonModifiable` — jamais de compromis, voir ci-dessus.

**Mode simulation par défaut** : `simuler()` retourne le rapport
ligne/colonne/ancienne valeur/nouvelle valeur SANS rien modifier ;
`appliquer()` (l'écriture réelle) est un appel explicite séparé — à tout
futur code GUI/CLI de cette branche d'imposer l'affichage du rapport de
simulation avant de proposer `appliquer()`.

**Structure de dossiers** (créée cette session, gitignorée — données
métier vivantes, jamais du dépôt) :
```
a_traiter/
  BL/                        dépôt quotidien des BL PDF, par l'acheteur
    Traités/                   BL numérisés déjà rapprochés (déplacés automatiquement
                                par le bouton GUI), renommés "<date> - <fournisseur> -
                                <n° BL> - BC <n° commande>" — l'acheteur y agrafe le
                                BdC et le BL papier correspondants puis archive dans
                                les classeurs physiques (dossier créé par l'acheteur
                                elle-même, demande explicite session R2 suite ;
                                remplace le "rapproches/<fournisseur>/<AAAA-MM>/"
                                imaginé au cadrage R1, jamais implémenté tel quel)
  Factures/                  dépôt hebdomadaire des factures PDF
rapports/                    rapports de rapprochement générés (R2+)
backups/                     sauvegardes horodatées du Suivi commandes avant chaque écriture (rotation 30 jours)
```

**Fichier vivant du Suivi commandes** : PAS celui à la racine du dépôt
(qui n'est qu'un export ponctuel, potentiellement périmé) mais
`X:\1.3. Logistique et approvisionnement\1.3.0. Commandes\1.3.0.1. Commandes
courantes\1.3.0.1. Suivi commandes - 2026.xlsx` — en-têtes/formules
vérifiés identiques à la copie testée par `ecriture.py`, donc le socle
s'applique tel quel. Deux autres fichiers à côté à ne PAS confondre :
`1.3.0.1. Suivi nouveau - à utiliser v4.xlsx` (nature à clarifier avec
l'acheteur avant R2 — refonte en préparation ?) et une `copie ...xlsx`
(sauvegarde manuelle antérieure).

**Flux réel, constaté sur 22 vrais BL papier de l'acheteur** (session R1,
dossier `a_traiter/BL/`, ex-dossier "Traitement" fusionné dedans à la
demande de l'acheteur) :

- **Les BL arrivent en PAPIER, numérisés par l'acheteur elle-même ; les
  factures arrivent des deux façons (PDF email ET relevé papier
  numérisé).** Le point de dépôt des scans aujourd'hui est `Z:\` (partage
  réseau généraliste, pas dédié — mélangé avec des documents admin/RH sans
  rapport). `a_traiter/BL/` et `a_traiter/Factures/` sont le point de dépôt
  CIBLE pour Rapprochement AI, pas encore le point de dépôt réel.
- **OCR OBLIGATOIRE, découverte cette session** : les 22 BL réels (dont 2
  `.jpg`) sont TOUS des scans image PURS, zéro caractère de texte
  extractible (vérifié avec `moteur.lecture_pdf.lire_pdf()` sur chacun —
  0 caractère à chaque fois). Toute l'infrastructure de parsing existante
  (regex sur texte PyMuPDF) est donc inapplicable aux BL sans une étape
  d'OCR préalable — à ajouter aux dépendances du projet en R2 (aucune lib
  OCR actuellement dans `requirements.txt`).
- **Numéro de commande présent et lisible sur 12/12 documents examinés**
  (rendu image + lecture visuelle, pas d'OCR encore) — meilleure nouvelle
  que redouté au vu de la réponse "normalement oui, mais pas tout le
  temps" de l'acheteur. Libellé différent par fournisseur (jamais deviné,
  toujours lu sur pièce réelle) :

  | Fournisseur | Libellé du n° commande sur le BL | Prix présent | Remarque |
  |---|---|---|---|
  | 109 Distribution (4 BL vus) | "N°Réf.Client" | Oui, systématique | Format identique au Suivi (ex. `123.096`) |
  | COREDIME (3 BL vus) | "Référence COMMANDE N°" | **Pas toujours** (`M2.16.011` : "Ref à livrer directement", 0 prix) | Prix réglé à la facture dans ce cas — la Qté livrée reste renseignable, pas le Tarif BL |
  | COMINTER / COMINTER OUEST (2 BL vus) | "Référence" | Oui | Structure proche du gabarit devis déjà existant (`moteur/fournisseurs/cominter.py`), probablement adaptable |
  | SAGEES (1 BL vu) | "V/Ref" | Oui | Pas de colonne Référence article (cohérent avec le format "V0" déjà documenté) |
  | DEM (1 BL vu) | "V/REF" | Oui, au cent (`/c`) | Cohérent avec `moteur/fournisseurs/dem.py` |
  | TOP Océan Indien (1 BL vu) | "N° de commande" | Oui | **Fournisseur PONCTUEL, confirmé par l'acheteur — ne pas ajouter de parser dédié.** Format de commande `C70244`, différent des autres (le préfixe "C..." existe bien dans le Suivi : 75 lignes réelles) |

  GMR et RAVATE : aucun BL réel vu cette session, à couvrir dès que des
  pièces réelles seront disponibles (règle d'or : jamais de gabarit sans
  PDF réel).
- **Carnets manuels des "gars"** : confirmé par l'acheteur, une commande
  peut être passée sur un carnet papier terrain avec un numéro
  `BC24.XXXX` — vérifié réel dans le Suivi (171 lignes au format `BC...`),
  distinct du format `C26.001` classique (75 lignes). Le n° de commande
  du BL ne matche donc pas toujours une ligne "propre" du Suivi — cas à
  gérer en R2, pas encore résolu.

**Session suivante — suite du gros lot du jour (Cominter page_9/page_7,
RAVATE page_0), nouveau fournisseur BL STAND 64, CRECHE OCEAN identifié
(fait) :**

- **Cominter page_9 (commande M4.267, écrite à la main sur le BdC par
  l'acheteur)** : sur les 3 lignes du BL, 2 (**L30316, L30326**) ne
  matchent RIEN dans les 3 lignes Suivi de cette commande — confirmé par
  l'acheteur qu'il s'agit d'une **vraie substitution** ("remplacé par des
  moulures équivalentes en Planet Wattohm, ça peut arriver"), même famille
  que le cas GOUJON8X70/75 déjà documenté : **jamais rapproché
  automatiquement**, laissées telles quelles dans "à vérifier". La 3e
  ligne (L86101L -> 086101L, sur-livraison 10 pour 5 commandées) reste en
  attente de confirmation de l'acheteur, pas encore tranchée.
- **Cominter page_7 (OBL108464, commande 130.036 confirmée par
  l'acheteur)** : cas non résolu, laissé en l'état. Les 2 lignes qui
  matchent une référence Suivi (L600335, L600323) visent des lignes
  **déjà intégralement soldées** (qté livrée = qté commandée), à un tarif
  DIFFÉRENT de celui du BL (2,25€/1,70€ au Suivi vs 2,05€/1,62€ sur le
  BL) — sur-livraison réelle ou BL correspondant à une commande déjà
  soldée, pas tranché. 2 autres lignes (L600001, L600002, interrupteurs
  va-et-vient) ne correspondent à aucune des 3 lignes Suivi de 130.036.
  Rien écrit, fichier resté dans "à vérifier".
- **RAVATE page_0 (commande 115.217) résolu — cause identifiée : PAS un
  problème de netteté/résolution mais une impression EXTRÊMEMENT pâle**
  (encre à peine plus foncée que le blanc, RGB ~180-225 sur 255) sur tout
  le corps du document — seul l'en-tête (encre noire normale, préimprimé)
  ressortait à l'OCR brut ou avec un simple réglage de contraste
  (`ImageEnhance.Contrast`/`ImageOps.autocontrast` classiques, testés à 3
  DPI, toujours 0 ligne). Résolu en deux temps : (1) lu VISUELLEMENT sur
  le rendu haute résolution (1 ligne : Référence fournisseur "V432552",
  "SACHET VISSERIE 3P NSX630 INV D", 1 BTE, 15,00€ net — confirmé par le
  Total HT affiché = 15,00€ = 1×15,00 exact, et par le Suivi qui n'a
  qu'une ligne pour cette commande, "LV432552", même cœur numérique) ;
  (2) confirmé a posteriori par un **étirement de niveaux ciblé sur le
  seul canal VERT** de l'image (le texte pâle est légèrement magenta/rose,
  donc fait chuter G plus que R/B — `(G - lo) / (hi - lo) * 255` avec
  lo≈150-200, hi≈240-250) qui a permis à RapidOCR de relire les MÊMES
  valeurs automatiquement (V432552, 24.68, 1.000, 15.00). Écrite et
  archivée dans `Traités/115.217/`. **Décision : ne PAS généraliser ce
  traitement dans `moteur/ocr.py`** — un seul document affecté à ce jour
  sur des dizaines traités cette session, conforme à la règle d'or (un
  seul exemple ne fait pas une règle) ; à revoir si un 2e cas de ce genre
  se présente.
- **CRECHE OCEAN.pdf identifié — PAS un raté OCR, un document hors
  périmètre** : ce n'est pas un BL d'un des grossistes électriques
  couverts par ce projet, mais un bon de livraison de **RSW.net**
  (fourniture d'un "optimiseur d'énergie" — coffret de gestion de
  puissance/supervision — pour le chantier "Crèche les Explorateurs").
  Signalé à l'acheteur pour savoir si ce type de document a sa place dans
  `a_traiter/BL/` ou doit être classé ailleurs — pas encore tranché,
  laissé tel quel dans "à vérifier" en attendant.
- **Nouveau fournisseur BL : STAND 64** (`moteur/fournisseurs/stand64.py`,
  section GABARIT BL ajoutée à la suite du parser devis existant) — 2
  vrais BL disponibles (`tests/fixtures/bl_stand64_1.pdf`/`_2.pdf`,
  `tests/test_parsers_bl_stand64.py`, 3 tests). Scan net (pas de texte PDF
  natif malgré l'apparence — vérifié, `page.get_text()` vide), mais OCR
  très fiable dessus. Tableau simple, colonnes DANS L'ORDRE (contrairement
  au devis de ce même fournisseur, qui les a en ordre inversé) :
  Référence article | Description | Qté | P.U | Rem% | P.U Net | Eco-part
  | Total HT | TVA — ancre fiable : le code TVA (C0..C9) en fin de ligne,
  éco-part toujours vide sur les 2 pièces vues (aucune cellule imprimée
  dans ce cas, pas "0,00"). 2 bugs réels corrigés en le construisant,
  même famille que des bugs déjà rencontrés chez d'autres fournisseurs :
  (1) l'en-tête "Référence article" ressort de l'OCR AVEC son accent
  (`unicodedata.normalize` + `encode("ascii","ignore")`, même repli que
  109 Distribution) ; (2) le motif de commande (`BC N°M2.23.058`) utilisait
  un joker `.{0,2}` gourmand entre "N" et la référence pour absorber le
  symbole "°" mal lu — gourmand, il avalait AUSSI le "M" du préfixe
  chantier ("M2.23.058" -> "2.23.058" capturé, préfixe perdu) ; corrigé en
  `[^A-Z0-9]{0,2}` (exclut explicitement lettres/chiffres, ne peut plus
  avaler le préfixe utile).
- **Référence "avec éclairage" scindée en 2 pièces cette année chez STAND
  64, confirmé par l'acheteur** ("unique les années précédentes") : les
  ventilateurs vendus avec kit lumière intégré avaient une référence
  UNIQUE les années précédentes (ex. `WESTI-78017`, `WESTI-73046`, encore
  la référence utilisée au Suivi) — cette année, STAND 64 les livre en 2
  lignes séparées sur le BL, même quantité chacune : le ventilateur nu
  (`WESTI-73044`/`WESTI-73045`) + un kit lumière à part
  (`WESTI-COMET-KITLUM-B`/`-N`, 7,00€). Repéré AVANT confirmation par la
  cohérence : sur les 2 vrais BL, une réf déjà écrite (le ventilateur nu
  "sans éclairage") apparaissait EN DOUBLE avec une 2e quantité inexpliquée,
  toujours accompagnée d'un kit lumière de même quantité. Une fois
  confirmé, fusionné manuellement (prix net = P.U Net ventilateur + P.U Net
  kit, ex. 91,00 + 7,00 = 98,00€) sous l'ancienne référence Suivi pour
  écriture — pas de règle générale codée (seulement 2 couples
  référence-nue/référence-bundle connus à ce jour, aucun motif évident
  pour en déduire d'autres sans un 3e exemple, règle d'or). Les 2 BL
  entièrement résolus et archivés (`Traités/M2.23.058/`, `Traités/M4.270/`).

**Session suivante — page_9/page_7 tranchés par l'acheteur, BUG RÉEL de
fond dans `apparier()` trouvé et corrigé, lot STAND 64 traité en
profondeur (3 nouveaux bugs réels), fait :**

- **page_9 (M4.267) et page_7 (130.036) expliqués par l'acheteur** :
  "typique des commandes récupérées par les conducs chez les
  fournisseurs" — quand le mode de livraison est "à récupérer par nos
  soins", le chargé de travaux sur place peut ajouter un article
  (page_7 : un interrupteur double va-et-vient non prévu) ou demander
  "une boîte complète" plutôt que la quantité exacte commandée (page_9 :
  10 unités livrées pour 5 commandées, l'acheteur a corrigé qte_cmd à 10
  dans le Suivi elle-même en conséquence). page_9 : la ligne 086101L
  (qte_cmd désormais 10) réapparariée et écrite (SUR, 10 unités) ; les 2
  lignes substituées (L30316/L30326, moulures Planet Wattohm, voir
  session précédente) restent définitivement hors rapprochement
  automatique — fichier laissé dans "à vérifier" (même principe que le
  cas SY00ZU51 déjà documenté : une substitution permanente ne "se
  résout" jamais toute seule). page_7 : l'acheteur a saisi elle-même
  toutes les valeurs dans le Suivi et a explicitement demandé l'archivage
  ("on peut archiver") — fait, sans réécriture (ses valeurs faisaient déjà
  foi).
- **BUG RÉEL DE FOND corrigé dans `apparier()`** (`moteur/rapprochement/
  matching.py`), trouvé en traitant page_7 : deux lignes RÉELLEMENT
  différentes du même BL ("L600001", interrupteur va-et-vient SIMPLE, et
  "L600002", va-et-vient DOUBLE — prix différents, désignations
  différentes) se disputaient la MÊME ligne Suivi ("600002", le seul
  article que l'acheteur avait ajouté). En UNE seule passe (l'ancien
  comportement), "L600001" était traité EN PREMIER (ordre du document),
  ne trouvait aucune correspondance EXACTE mais un repli à 1 caractère
  d'écart vers "600002" — et consommait cette ligne Suivi par repli AVANT
  que "L600002", qui la matche pourtant EXACTEMENT, n'ait sa chance :
  celui-ci ressortait "inconnu" à tort, une correspondance exacte
  existante étant purement et simplement perdue à cause de l'ordre
  d'apparition des lignes sur le document. Corrigé par une **vraie
  refonte en DEUX PASSES** : la 1re passe résout TOUTES les
  correspondances EXACTES de TOUTES les lignes du BL AVANT qu'aucun repli
  ne soit tenté (un repli approximatif ne peut plus jamais voler une
  ligne Suivi à une autre ligne de BL qui la matche exactement, quel que
  soit l'ordre) ; la 2e passe traite ce qui reste (ambiguïtés, replis)
  avec les lignes Suivi déjà réduites en conséquence. Comportement
  externe de `apparier()` inchangé pour tous les cas déjà couverts (30
  tests existants toujours verts) — testé spécifiquement
  (`test_apparier_exact_nest_jamais_vole_par_un_repli_dune_autre_ligne`).
  **Portée générale** : ce bug pouvait affecter N'IMPORTE QUEL fournisseur
  dès que 2 références réellement différentes d'un même BL se trouvaient
  à 1 caractère d'écart l'une de l'autre ET qu'une seule des deux existait
  dans le Suivi — pas un cas isolé à Cominter/109 Distribution.
- **Nouveau lot de 6 BL STAND 64 déposé par l'acheteur** ("afin que nous
  puissions établir ce parser", pour durcir le parser construit la
  session précédente sur seulement 2 pièces) — 3 nouveaux bugs réels
  trouvés et corrigés dans `moteur/fournisseurs/stand64.py` :
  1. **Éco-part RENSEIGNÉE** (commande M2.5.126, 2,88€ — les 2 pièces
     précédentes l'avaient toujours vide) : ajoute UNE cellule numérique
     de plus avant le Total HT, décalant tout le compte fixe de 6
     cellules (la Qté "18,00" happée dans la désignation, le P.U "95,00"
     pris à tort pour la Qté). Corrigé par une résolution à DEUX
     HYPOTHÈSES (6 cellules sans éco-part, 7 avec), chacune validée par
     cohérence arithmétique (qté × P.U Net ≈ Total HT affiché) — même
     principe que les replis positionnels déjà utilisés chez 109
     Distribution/Electric Plus, jamais un compte de cellules fixe
     supposé sans vérification.
  2. **Coche/checkmark imprimée collée à la Qté**, lue "V" par l'OCR
     (commande 131.170 : "40,00V" au lieu de "40,00", le "V" ressemblant
     au symbole ✓ manuscrit) : sans nettoyage, `to_float()` levait une
     exception et TOUTE la ligne disparaissait silencieusement (0 ligne
     pour un Total HT de 580,00€ affiché). Une lettre isolée en fin de
     cellule numérique est désormais retirée avant conversion
     (`_nombre_bl_stand64()`).
  3. **1re ligne de désignation totalement absente de l'OCR** (même
     commande M2.5.126 que le bug éco-part, cumul des deux problèmes sur
     le même document) : seules les lignes de désignation SUIVANTES
     ("BLANC/BLANC+TELECOMMANDE", "PRIX NETS", "MATERIEL DISPONIBLE CE
     JOUR") ont été détectées, la ligne chiffrée elle-même se retrouvant
     sans aucune cellule de désignation. Comme pour Coredime/Ravate
     (désignation qui déborde sur une 2e ligne), mais ici en LOOK-AHEAD :
     `parse_bl_stand64()` raccorde désormais les lignes suivantes tant
     qu'elles ne sont PAS elles-mêmes reconnues comme une ligne chiffrée
     valide, dès que la désignation d'une ligne chiffrée ressort vide.
     Qté/prix restent exploitables même sans désignation (elle ne sert
     pas au rapprochement) — n'est donc plus jamais un motif de rejet de
     la ligne à elle seule.
  Fixtures réelles ajoutées (`tests/fixtures/bl_stand64_3_ecopart_
  renseigne_desig_manquante.pdf`, `bl_stand64_4_coche_collee_qte.pdf`),
  testé (`test_parse_bl_stand64_3_ecopart_renseigne_et_designation_
  manquante`, `test_parse_bl_stand64_4_coche_collee_a_la_quantite`).
- **Recette réelle sur les 6 BL du lot** : 1 déjà à jour (M4.262), 2
  doublons de scan des BL M2.23.058/M4.270 déjà traités la session
  précédente (mêmes n° de BL, mêmes lignes — rien de nouveau, déplacés
  vers `À vérifier/Doublons confirmés (à supprimer)/`, jamais supprimés
  directement), 3 nouvelles commandes (128.007, M2.5.126, 131.170)
  entièrement résolues et archivées avec leur BC (4 lignes SUR écrites au
  total).

**Session suivante — lot de 11 BL du jour traité en une passe via le
pipeline officiel (fait)** : lecture seule (`rapprocher_dossier`) puis
écriture (`appliquer_et_archiver`) directement, sans script ad hoc — le
lot le plus propre à ce jour (5 fournisseurs : Coredime, Cominter,
Electric Plus, Stand 64 ; 12 BL sur 11 fichiers, un fichier Cominter en
contenant 2). 39 lignes sûres écrites, 11/12 BL archivés directement, 2
"déjà à jour" identifiés comme doublons de scan des BL Stand 64 de la
session précédente. Un seul point bloquant :

- **Commande M3.10.181 (Electric Plus), référence BAS217562 introuvable
  au Suivi** : la ligne Suivi portait "AGI864001" (même désignation
  quasi identique — "TIGE FILETÉE M6" —, même quantité, même tarif). **Ce
  n'est PAS un bug de l'outil, confirmé par l'acheteur après coup** :
  changement de marque fournisseur pour un article générique (tige
  filetée DIN976), deux références de fabricants différents pour le même
  produit — aucune règle de rapprochement automatique ne peut relier deux
  références **textuellement et numériquement sans rapport** juste parce
  que désignation/qté/tarif concordent (risquerait de fusionner à tort
  deux articles réellement différents qui partagent ces caractéristiques
  par coïncidence) — reste et restera un cas à confirmer par l'acheteur,
  comme les autres substitutions déjà documentées (GOUJON8X70/75, Planet
  Wattohm, GMR→Cominter). Elle a corrigé la référence dans le Suivi
  elle-même ("Basor livré") ; la ligne a ensuite matché normalement et le
  BL a été archivé.

**Session suivante — lot de 15 BL du jour, flux complet de bout en bout,
1er cas réel "Reste à livrer" Coredime, gap repli préfixe marque côté BL
(fait) :**

- **1er cas réel de "Reste à livrer" Coredime sur PLUSIEURS PAGES**
  (commande M3.23.043, prévu "à valider dès qu'un cas réel se présente" —
  voir Points fragiles) — **BUG RÉEL CORRIGÉ** : `reste_a_livrer` était un
  drapeau global à TOUT le document (`moteur/fournisseurs/coredime.py`) ;
  activé en page 1, il restait actif pour la page 2 aussi, excluant à
  tort des lignes pourtant livrées — la structure réelle observée est que
  Coredime RÉIMPRIME en page 2, avec confirmation de livraison ("<qté> x
  1 unite"), les mêmes articles que la page 1 avait listés "reste à
  livrer". Réinitialisé désormais à chaque page. **2e bug réel corrigé**
  sur le même document : "10" (LEG031490) lu "lo" par l'OCR (le "1" de
  tête collé au "0" suivant, pas un "l" isolé comme le cas déjà connu) —
  remplacement élargi à un "l" en tout DÉBUT de mot suivi d'un chiffre ou
  d'un "o"/"O". Fixture réelle ajoutée
  (`tests/fixtures/bl_coredime_7_reste_a_livrer_multi_page.pdf`), testé
  (`test_parse_bl_coredime_7_reste_a_livrer_multi_page`). **Limite
  connue, non corrigée en code** (corrigée en correction manuelle
  ponctuelle cette session) : sur ce même document, la 1re ligne
  LEG411651 (qté 4, page 1) a sa confirmation "4 x 1 unite" AVANT la
  ligne référence plutôt qu'après — pattern look-behind non couvert par
  le raccord actuel (qui ne regarde qu'en avant) ; un seul exemple, pas
  de règle générale codée.
- **BUG RÉEL corrigé dans `apparier()` mis en lumière une 2e fois** (pas
  un nouveau bug, juste une nouvelle façon de le déclencher, déjà
  couverte par le fix session précédente) : ajouter une correction
  manuelle en DEUX LIGNES séparées pour la MÊME référence (au lieu de
  fusionner en une seule ligne cumulée) fait qu'une seule des deux
  obtient la correspondance exacte, l'autre ressort "inconnu" — leçon
  retenue pour toute correction manuelle future portant sur une référence
  déjà présente ailleurs dans le même BL : fusionner les quantités en UNE
  ligne, ne jamais ajouter une 2e ligne séparée avec la même référence.
- **RAVATE M3.23.044.jpg : 2 corrections manuelles motivées** (photo
  angle/qualité moyenne, un seul exemple, pas de règle générale codée) :
  (1) la cellule P.U. Brut a totalement disparu de l'OCR sur la ligne
  70438, décalant tout le calcul positionnel (repli 3-cellules
  déclenché à tort, qté calculée à 0,6399 au lieu de 1,0) — corrigé
  manuellement (qté=1,0, prix_net=7,02, montant=7,02) ; (2) la ligne
  600335 (DOOXIE PC SURFACE + TERRE BLC) a totalement disparu de l'OCR,
  RÉFÉRENCE COMPRISE (aucune trace dans le texte OCR) — ajoutée
  manuellement (qté=30, prix_net=2,20, montant=66,00). Les deux corrigées
  avec un signal fort : la somme des 4 montants (7,02+86,00+66,00+112,50
  = 271,52€) retombe EXACTEMENT sur le Total HT affiché.
- **Nouveau gap identifié côté rapprochement BL : aucun repli "préfixe
  marque"** (contrairement à `moteur/referentiel.py`, qui gère déjà
  LEG/EBE/PW/BT côté devis) — 2 cas réels rencontrés ce lot, tous deux
  corrigés manuellement (qté BL = qté commandée EXACTEMENT + désignation
  quasi identique, signal fort à chaque fois) :
  - Electric Plus 139.115 : BL "SCHXALK178E" vs Suivi "XALK178E" (préfixe
    "SCH", vraisemblablement Schneider).
  - 109 Distribution 139.111 : BL "H07VR16VJTECC" vs Suivi "H07VR16VJ T"
    (écart cette fois en SUFFIXE, pas en préfixe — probablement une
    troncature historique côté Suivi plutôt qu'un préfixe marque).
  Seulement 2 exemples, préfixes/troncatures différents à chaque fois —
  pas assez pour coder un repli général fiable (règle d'or) ; à revoir
  si le même préfixe (ex. "SCH") revient sur un 3e cas.
- **Conflit réel signalé, laissé à l'acheteur** : `BL M3.23.045.pdf`
  (CORB033860.3, LEG411651 qté 4) et `BL M3.23.045 1.pdf` (CORB033860.1,
  LEG411651 qté 5) sont deux BONS DE LIVRAISON RÉELLEMENT DIFFÉRENTS
  (numéros de document distincts, `.1` et `.3`) visant la même ligne
  Suivi — le désamorçage anti-doublon a bien réagi (une seule proposée
  "sûre" à la fois), mais impossible de savoir sans elle si c'est une
  vraie livraison fractionnée (4+5=9, à comparer à la qté commandée) ou
  une erreur — aucune des deux lignes LEG411651 écrite automatiquement.
- **Document hors périmètre détecté et exclu** :
  `doc07194020260821104206.pdf` est un **"ACCUSÉ RÉCEPTION DE COMMANDE"**
  Coredime (pas un bon de livraison — aucun prix, aucune quantité livrée
  confirmée, n° de commande manuscrit "139.114" jamais imprimé) — retiré
  du lot traité, signalé à l'acheteur plutôt que forcé dans le pipeline
  BL.
- **Recette finale sur les 15 fichiers** : 26 lignes sûres écrites au
  total, 11 fichiers entièrement résolus et archivés avec leur BC. Restent
  en attente (aucune écriture automatique) : `BL 139.111.pdf` (référence
  59210 introuvable + sur-livraison SYT320G5 50/40 à confirmer),
  `doc07194120260821104221.pdf` (2 références 227060125/227060122
  introuvables), le duo `M3.23.045`/`M3.23.045 1` (conflit ci-dessus), et
  `doc07194020260821104206.pdf` (hors périmètre).

**Session suivante — base des équivalences pour le rapprochement des BL,
branchement sur le référentiel articles déjà existant côté devis (fait)** :
demande explicite de l'acheteur pendant la recette du lot précédent, après
avoir dû corriger à la main plusieurs substitutions fournisseur (59210
pour CFF1BIS, 092897 pour 411651) : "il faut créer une base des
équivalences, ce genre de cas va se présenter très souvent."

- **Décision d'architecture : réutiliser `moteur/referentiel.py` tel
  quel**, plutôt que construire un second système en parallèle — il
  résout déjà EXACTEMENT le même problème côté devis (alias
  confirmé/proposé/nouveau, fichier Excel aller-retour). Le rapprochement
  des BL (`moteur/rapprochement/matching.py`) consulte désormais ce MÊME
  référentiel (`moteur/articles.db` partagé — un alias confirmé par
  N'IMPORTE LEQUEL des deux flux, devis ou BL, vaut pour l'autre).
- **`_memes_references()`** (remplace l'ancien usage direct de `_cle()`
  dans `apparier()`) : `_cle(a) == _cle(b)` (comme avant) **OU**, si un
  référentiel est fourni, les deux références ont un alias CONFIRMÉ
  (statut "connu") vers la MÊME clé. **Un OR, jamais un remplacement** —
  point de conception important, voir le bug ci-dessous.
- **BUG RÉEL CORRIGÉ pendant la construction** (recette immédiate sur le
  vrai Suivi, commande 139.112, référence "LEG069831L" vs "069831L") :
  une 1re version comparait EXCLUSIVEMENT la clé référentiel dès qu'un
  alias "connu" existait, au lieu de l'ajouter en plus de la comparaison
  par cœur numérique existante. Or un alias "connu" peut être un simple
  AUTO-alias d'une référence vers ELLE-MÊME (la BDD achats retient parfois
  la forme préfixée "LEG069831L" comme Clé_Réf telle quelle, sans la
  réduire à "069831L") — les deux références ressortaient alors "connu"
  mais vers des clés DIFFÉRENTES, un résultat STRICTEMENT PIRE que la
  comparaison par cœur numérique déjà en place (qui, elle, les retombait
  bien toutes les deux sur "69831"). Une ligne déjà correctement écrite
  lors d'une session précédente s'est ainsi retrouvée "inconnu" au
  passage suivant. Corrigé par le OR ci-dessus — le référentiel ne peut
  plus jamais FAIRE PERDRE une correspondance déjà trouvée.
- **`_repli_referentiel()`** : 3e repli (après `_repli_reference_proche()`)
  pour une correspondance simplement PROPOSÉE (candidat structurel
  plausible via préfixe marque/cœur numérique, pas encore confirmée) —
  toujours "à confirmer", jamais automatique.
- **`referentiel/equivalences_bl.csv`** (nouveau, vide par défaut sauf 2
  entrées réelles ci-dessous) : pour le cas qui a motivé la demande
  (substitution SANS AUCUN rapport textuel ni numérique entre les deux
  références — "59210" vs "CFF1BIS" ne partagent RIEN, donc le mécanisme
  de proposition automatique de `Referentiel.resoudre()` — basé sur
  préfixe marque/cœur numérique — ne peut STRUCTURELLEMENT PAS le
  deviner tout seul, contrairement à un simple écart de préfixe). Nouvelle
  méthode `Referentiel.importer_equivalences_bl()` : format
  `Reference_1;Reference_2;Note`, les deux références aliasées l'une vers
  l'autre avec origine='manuel', statut "connu" dès l'import — aucune
  proposition préalable nécessaire, contrairement à A_confirmer_BL.xlsx.
  **Reference_1 doit être la référence déjà connue de la BDD achats
  quand l'une des deux l'est** (pas l'inverse), pour ne pas faire dériver
  le groupement des lignes de DEVIS existantes vers une clé synthétique —
  bug trouvé et corrigé en même temps que la construction (voir README du
  fichier). 2 entrées réelles ajoutées : `CFF1BIS;59210` (109 Distribution,
  commande 139.111) et `411651;092897` (Coredime, commande M3.23.045).
- **`referentiel/A_confirmer_BL.xlsx`** (nouveau, fichier À PART de
  `A_confirmer.xlsx` côté devis — chacun régénère SA PROPRE file d'attente
  à chaque exécution, les mélanger écraserait les propositions de l'autre
  flux). `Referentiel.ecrire_a_confirmer()` accepte désormais un
  `nom_fichier` optionnel pour ça (défaut inchangé : "A_confirmer.xlsx").
- **BUG RÉEL CORRIGÉ (latent, révélé par cette intégration)** :
  `Referentiel.importer_bdd()` faisait un simple `INSERT` (pas
  `INSERT OR REPLACE`) pour re-peupler les alias d'origine 'import' —
  plantait (UNIQUE constraint) si une référence de la BDD avait été
  redirigée entre-temps par une équivalence BL manuelle (le
  `DELETE FROM alias WHERE origine='import'` ne touche pas les lignes
  'manuel'). Passé en `INSERT OR REPLACE` — un ré-import de la BDD reprend
  proprement sa valeur d'origine, immédiatement re-surchargée par
  `importer_equivalences_bl()` juste après dans `rapprocher_dossier()`
  (l'ordre des deux appels est important, toujours BDD puis équivalences).
- **`rapprocher_dossier()`** (pipeline_bl.py) ouvre désormais son propre
  `Referentiel` (moteur/), importe la BDD achats + equivalences_bl.csv +
  les confirmations en attente de A_confirmer_BL.xlsx, le passe à chaque
  `apparier()` (avec `fournisseur`/`devis` pour que les propositions dans
  A_confirmer_BL.xlsx indiquent de quel BL elles viennent), régénère
  A_confirmer_BL.xlsx en fin d'exécution. Coût mesuré : ~0,1-0,2s de plus
  par exécution (réimport idempotent de la BDD, déjà éprouvé côté devis).
- **Suggestion de l'acheteur, notée pour une prochaine session, PAS
  construite maintenant** : gérer aussi les écarts d'UNITÉ de
  conditionnement (ex. seau de 1000 vs boîte de 200) — un problème
  différent (facteur de conversion, pas une équivalence de référence),
  qui prendrait naturellement place dans le même référentiel une fois
  l'équivalence de référence éprouvée en usage réel.
- Tests : `tests/test_referentiel.py` (import/idempotence
  d'equivalences_bl.csv), `tests/test_rapprochement_matching.py` (alias
  confirmé traité comme exact, comportement par défaut sans référentiel
  inchangé, proposition non confirmée reste "à confirmer"). Suite
  complète (229+ tests) verte après ces changements.
- **Recette réelle du branchement, immédiatement après construction** :
  d'abord confirmations page_9/page_7 Cominter tranchées par l'acheteur
  ("typique des commandes récupérées par les conducs" — substitution
  définitive Planet Wattohm sur page_9, ajout d'article par le chargé de
  travaux sur page_7, ce dernier saisi et archivé directement par
  l'acheteur elle-même) ; puis un lot de 15 BL traité de bout en bout, qui
  a lui-même servi de première vraie mise à l'épreuve du référentiel
  (139.111 CFF1BIS/59210, 139.112 227060122/227060125 — remplacements
  fournisseur confirmés par l'acheteur, tarif/désignation identiques,
  aucun souci réel malgré mon inquiétude initiale sur une "sur-livraison"
  qui n'en était pas une). **2 corrections de ma part suite à des erreurs
  de lecture, pas des bugs du moteur** : qte_cmd 139.111 réellement 50 (le
  fournisseur vend par couronne de 50, l'acheteur a corrigé le Suivi elle-
  même) ; 227060125 réellement qté 1 (pas 12 — un artefact OCR "1" + début
  de "39%" que j'avais mal retranscrit à la main, corrigé après une photo
  envoyée par l'acheteur). 26 lignes sûres écrites au total sur ce lot,
  11/15 BL archivés directement.

## Session suivante — Code article corrompu OCR (Coredime), 2e fichier
multi-fournisseur/multi-commande traité manuellement (fait)

Suite directe de la session précédente ("Je t'ai mis les BL du jour dont
un avec multi-scans pour compliquer, allez on traite de bout en bout !").

- **BUG RÉEL CORRIGÉ — Coredime, en-tête de tableau "Code article" lu
  "Code aricle" par l'OCR** (le "T" disparu) : `doc07205120260824144931.pdf`
  (commande 142.041, article LBCLASTD02 qté 10, un BL par ailleurs simple
  et propre) ressortait à 0 ligne extraite car `MOTIF_ENTETE_TABLEAU_BL_
  COREDIME` (`CODEARTICLE`, strict) ne matchait jamais. Diagnostiqué via
  un dump OCR brut (non filtré par zone de tableau) qui a montré l'en-tête
  corrompu telle quelle. Fix : motif élargi à `CODEART?ICLE` (le "T"
  optionnel). Fixture réelle ajoutée
  (`tests/fixtures/bl_coredime_8_entete_code_aricle.pdf`), testé
  (`test_parse_bl_coredime_8_entete_code_aricle`).
- **2e occurrence réelle du cas "un même fichier PDF mélange PLUSIEURS
  FOURNISSEURS ET plusieurs commandes, dont certaines déjà traitées"**
  (1re occurrence documentée plus haut, session précédente,
  `doc07178620260819170752.pdf`) : `doc07205620260824145119.pdf` (7 pages)
  contenait 3 pages VRAIMENT nouvelles (p0 : 109 Distribution, BL 39805,
  commande **M3.10.183**, LEG031456 qté 11 ; p1 : Coredime, BL CORB034089,
  commande **126.048** — écrite "BC126048" sur le papier, format Suivi réel
  trouvé par recherche large sur la référence HAGGE326EN plutôt que deviné
  ; p2 : 109 Distribution, BL 740135, commande **126.051**, 52302 qté 20 +
  70003 qté 10) ET 4 pages DOUBLONS de scan de BL déjà archivés plus tôt
  dans la même session (142.039, M3.14.364, M3.23.043 ×2 — reste-à-livrer
  et confirmation). Comme la 1re fois, `lire_bl()` ne sait détecter qu'UN
  SEUL fournisseur pour tout le fichier — traité manuellement (script
  scratchpad, pas dans le code du projet) : OCR du fichier entier une
  fois, `parse_bl_109()`/`parse_bl_coredime()` appelés page par page sur
  des tranches d'une seule page, rapprochement + écriture des 4 lignes
  sûres en un seul lot, puis extraction PAGE PAR PAGE via
  `_extraire_bl_vers()` vers `Traités/<commande>/` (pages neuves,
  archivage BC compris) ou vers `À vérifier/Doublons confirmés (à
  supprimer)/` (pages doublons, sans réécriture Suivi — contenu déjà
  correctement enregistré). Les 7 pages ont été intégralement traitées et
  rangées, sauvegarde Suivi horodatée avant écriture.
- Suite complète verte après le fix Code aricle (8/8 tests Coredime BL
  confirmés isolément avant la passe complète).

## Session suivante — détection de fournisseur PAR PAGE (fait), demande
explicite de l'acheteur : "il est extrêmement fastidieux de scanner page
par page, beaucoup plus simple d'envoyer en masse"

Le cas "un même fichier PDF mélange plusieurs fournisseurs" (rencontré 2
fois, traité à la main les deux fois — voir sections précédentes) est
désormais géré AUTOMATIQUEMENT par `moteur/rapprochement/lecture_bl.py`.

- **`lire_bl()` détecte désormais le fournisseur PAGE PAR PAGE**, pas
  seulement sur le texte entier du document. Si toutes les pages
  s'accordent (ou qu'une seule est reconnue) — le cas de très loin le plus
  courant — le comportement est STRICTEMENT celui d'avant : détection sur
  le texte ENTIER, un seul appel au parser sur toutes les pages (pour ne
  rien perdre d'un fournisseur qui ne se révèle que sur une page parmi
  plusieurs, ex. une page de garde). Seulement si PLUSIEURS fournisseurs
  DIFFÉRENTS ressortent, le fichier est découpé par groupes de pages
  (`_parser_groupe_fournisseur()`), chaque groupe traité indépendamment —
  un fournisseur peut réapparaître PLUS LOIN dans le fichier (pages non
  contiguës, cas réel constaté : RAVATE p0, COMINTER p1/p2/p7/p9, 109
  DISTRIBUTION p3/p6/p8...), le regroupement n'exige donc jamais la
  contiguïté.
- **Garde-fou critique : ne jamais fusionner à tort deux BL de commandes
  différentes sous un même fournisseur.** Certains fournisseurs
  (109 Distribution, Cominter Ouest, Electric Plus) répartissent déjà
  eux-mêmes plusieurs BL sur plusieurs pages via leur propre numéro de BL
  (`moteur.ocr.pages_par_identifiant`, `bl.pages` renseigné par le parser
  lui-même) — pour ceux-là, un groupe à plusieurs pages est passé en UN
  SEUL appel, le parser fait la répartition fine lui-même. D'autres
  (Coredime, Ravate, Stand 64, DEM...) ne savent PAS répartir un BL sur
  plusieurs pages : leur parser traite tout ce qu'on lui donne comme UN
  SEUL document. Si un groupe de plusieurs pages du MÊME fournisseur non
  contiguës (donc probablement 2 BL réellement différents, pas un même BL
  sur 2 pages) est passé à un tel parser, `_parser_groupe_fournisseur()`
  détecte que le résultat ne renseigne PAS `bl.pages` et REFAIT l'appel
  PAGE PAR PAGE plutôt que de garder le résultat fusionné — mieux vaut
  sous-découper (au pire une info incomplète, réexaminée à la main) que
  mélanger deux commandes réellement différentes sous un seul BL (c'est
  exactement le danger identifié sur le cas réel qui a motivé cette
  fonctionnalité, `doc07205620260824145119.pdf` — COREDIME y apparaissait
  sur 4 pages non contiguës, 3 commandes distinctes).
- **`bl.pages` est TOUJOURS renseigné** pour tout BL produit par le chemin
  multi-fournisseur (jamais `None`) — le laisser `None` ferait archiver
  TOUT le fichier source (voir `moteur.rapprochement.pipeline_bl`,
  `bl.pages if bl.pages is not None else list(range(pages_totales))`),
  emportant à tort les pages des AUTRES fournisseurs.
- **`lire_bl()` retourne désormais `(bons, raisons)`** — `raisons` est une
  LISTE (peut contenir plusieurs anomalies, une par page/groupe en échec),
  et peut être NON VIDE MÊME QUAND `bons` ne l'est pas (une partie du
  fichier résolue, une autre page en anomalie, ex. fournisseur reconnu
  mais sans parser BL, ou page illisible). Avant cette session, une seule
  raison (ou aucune) pour tout le fichier ; `analyser_dossier()` adapté en
  conséquence (boucle sur `raisons` au lieu d'un seul `if not bons`).
- **BUG RÉEL ÉVITÉ pendant la construction, trouvé par relecture attentive
  du code existant AVANT tout test** (pas signalé par l'acheteur — la
  fonctionnalité vient d'être construite, pas encore déposée en usage réel
  avec un cas qui l'aurait révélé) : `appliquer_et_archiver()`
  (`pipeline_bl.py`) traitait TOUTE présence dans `rapport.anomalies_lecture`
  comme un échec de lecture TOTAL du fichier (déplacement du fichier ENTIER
  vers "à vérifier", AVANT même que ses pages résolues aient pu être
  archivées individuellement — la 2e boucle trouvait alors le fichier déjà
  déplacé et ne faisait plus rien). Avec la détection par page, un fichier
  peut désormais avoir À LA FOIS une anomalie ET des BL résolus. Corrigé :
  `fichiers_en_echec_total` exclut désormais tout fichier qui a AU MOINS un
  `BonLivraison` (donc pas un échec total) — testé
  (`test_appliquer_et_archiver_anomalie_de_lecture_necoupe_pas_les_bl_resolus_du_meme_fichier`).
  **2e correctif lié** : le critère "fichier à un seul BL -> déplacer le
  fichier ENTIER tel quel" (`len(gs) == 1`) ne suffit plus — un fichier
  multi-fournisseur peut produire un SEUL `BonLivraison` dont `bl.pages` ne
  couvre qu'UNE SEULE page sur plusieurs (les autres appartenant à un autre
  fournisseur, ou illisibles). Le critère est désormais "un seul BL ET
  (`bl.pages` est `None`, OU couvre la TOTALITÉ des pages du fichier)" —
  sinon, comme pour un fichier multi-BL, découpage par
  `_traiter_bl_multiples_du_fichier()` (qui gère déjà correctement le
  cas `len(gs) == 1`, aucune logique nouvelle nécessaire là) — testé
  (`test_appliquer_et_archiver_un_seul_bl_mais_pages_partielles_est_decoupe_pas_deplace_en_bloc`).
- **Validé sur le VRAI fichier qui a motivé cette fonctionnalité**
  (`doc07205620260824145119.pdf`, déjà traité à la main la session
  précédente) : la nouvelle détection automatique reproduit EXACTEMENT le
  même résultat que le traitement manuel — 7 BL, 3 fournisseurs (109
  Distribution ×3 pages, Coredime ×4 pages non contiguës), chaque page
  correctement isolée avec sa propre commande, AUCUNE fusion à tort entre
  les 2 commandes Coredime distinctes portées par des pages non
  adjacentes. Seul écart (attendu, pas un bug de cette fonctionnalité) :
  le n° de commande de la page Coredime "BC126048" ressort vide
  (limitation PRÉEXISTANTE du motif de commande Coredime sur ce format
  collé sans séparateur reconnu — la session précédente l'avait résolu à
  la main via une recherche large sur la référence article, pas par le
  parser lui-même).
- **Tests** : `tests/test_lecture_bl.py` (nouveau — 3 tests, dont un
  construit en COMBINANT de vraies pages de fixtures déjà verrouillées
  ailleurs — `bl_coredime_1.pdf`/`bl_coredime_3.pdf`/`bl_dist109_1.pdf`/
  `bl_dist109_2.pdf` assemblées en un seul PDF à 4 pages via `fitz` —
  contenu 100% réel, seul l'assemblage dans un même fichier est
  synthétique, reproduit fidèlement le geste de l'acheteur qui scanne
  plusieurs BL papier à la suite) ; `tests/test_rapprochement_pipeline_bl.py`
  (2 nouveaux tests pour les 2 correctifs de routage ci-dessus). Suite
  complète : 235 passés (230 avant cette fonctionnalité), aucune
  régression.

## Session suivante — nouveau fournisseur BL YESSS, bug réel d'archivage
"/" trouvé en écrivant pour de vrai (fait)

Premier vrai lot de BL traité avec la détection par page en usage réel :
sur 4 fichiers déposés, 3 étaient des redépôts déjà rapprochés (confirmant
l'idempotence sur des re-scans réels), 1 était un **nouveau fournisseur
jamais vu, YESSS ÉLECTRIQUE** (marque du groupe CEF SAS — `BL M4.273
GENDARMERIE.pdf`, agence YESSS CAMBAIE) — confirmé avec l'acheteur avant
de construire le parser (comme pour le branchement référentiel : jamais
une évolution d'architecture ou un nouveau fournisseur construit sans
validation explicite).

- **Structure RÉELLEMENT inhabituelle** (`moteur/fournisseurs/yesss.py`,
  nouveau module, BL uniquement — aucun devis connu pour ce fournisseur) :
  chaque champ (Montant, Prix net, Désignation, Catalogue...) a son texte
  pivoté à 90° sur la page (confirmé en annotant le PDF rendu avec les
  boîtes OCR — mots hauts et étroits, signature d'un texte tourné). Chaque
  "colonne" du tableau d'origine (non tourné) devient une bande verticale
  étroite (~30-40px) : label empilé à gauche, valeur RÉELLE juste à
  droite, et une 3e bande encore plus à droite portant un "0.00"/vide —
  un emplacement de 2e ligne d'article TOUJOURS imprimé par le gabarit
  mais non rempli (une seule ligne d'article vue à ce jour). Chaque valeur
  est retrouvée par PROXIMITÉ (X ET Y) à son label plutôt que par un ordre
  de lecture haut/bas classique — une tolérance X étroite (~40px) est
  nécessaire : sans elle, une mention voisine sans rapport ("dispo sous
  48h", dans la bande vide adjacente) se trouve PLUS PROCHE en Y du label
  "Désignation" que la vraie désignation. Quantité déduite de Montant/Prix
  net (comme 109 Distribution/Cominter/Electric Plus/Ravate), jamais lue
  directement (la cellule Qté est elle aussi ambiguë, deux valeurs
  empilées "2"/"0" — la vraie et l'emplacement vide).
- **Date : 2 bugs réels corrigés en construisant le parser**, tous deux
  liés au même principe "chercher par proximité, jamais le premier trouvé
  dans un ordre arbitraire" : (1) le jour ("24") ressort comme un mot
  OCR SÉPARÉ du mois+année ("aout 2026") — un premier essai qui cherchait
  "l'unique mot à 1-2 chiffres de tout le document" échouait car d'autres
  mots (quantité, etc.) matchaient aussi ; corrigé en prenant le candidat
  le PLUS PROCHE (Y) du mot mois+année. (2) le pavé légal du bas de page
  cite une loi "du 25 janvier 1985" — un mois+année tout aussi valide
  regex-parlant, mais sans rapport ; en cherchant le premier match dans un
  ordre non trié, "janvier 1985" pouvait être pris à la place de la vraie
  date. Corrigé en ancrant la recherche du mois+année au label "Date"
  lui-même (même logique de proximité que le reste du gabarit), pas au
  premier match trouvé.
- **BUG RÉEL CRITIQUE trouvé en écriture réelle (pas en test)** : le n° de
  BL YESSS contient un "/" imprimé ("CAM/040759") — `_nom_archive_bl()`
  (`pipeline_bl.py`) l'interpolait tel quel dans le nom de fichier
  d'archive, et Windows interprète "/" comme un séparateur de DOSSIER :
  l'archivage complet échouait ("chemin d'accès introuvable"), alors que
  l'écriture dans le Suivi avait déjà réussi à ce stade (grâce au
  découplage écriture/archivage déjà en place — voir "Bug de robustesse"
  plus haut — rien perdu, juste un archivage à refaire une fois corrigé).
  Corrigé par une fonction de nettoyage PARTAGÉE (`_sans_caracteres_interdits()`,
  caractères Windows interdits `< > : " / \ | ? *`) — réutilisée aussi par
  `_nom_dossier_commande()` (qui avait déjà sa PROPRE logique de
  nettoyage, quasi identique mais dupliquée — consolidée en une seule
  fonction). **Portée générale** : protège tout futur fournisseur dont le
  n° de BL/commande contiendrait un caractère interdit, pas seulement
  YESSS. Testé
  (`test_archiver_bl_numero_bl_avec_slash_ne_casse_pas_le_chemin`).
- **`moteur/detecteur.py`** : motif `YESSS` ajouté (`\bYESSS\b`), plus
  `YESSS MAYOTTE` testé EN PREMIER (même principe que COMINTER MAYOTTE) —
  agence distincte qui existe dans la liste Fournisseurs du Suivi mais
  jamais rencontrée en pièce réelle, donc reconnue mais SANS parser BL
  dédié (tombe proprement dans l'anomalie "reconnu mais pas de parser" le
  jour où une pièce réelle se présentera). Nom canonique Suivi vérifié
  directement dans le classeur vivant ("YESSS", identique au nom détecteur
  — pas de remapping nécessaire, contrairement à GMR/Electric Plus) ;
  `moteur.panier.MAPPING_FOURNISSEURS` complété en conséquence (verrouillé
  par `tests/test_panier.py`, qui exige une entrée pour tout fournisseur
  du détecteur).
- **BL 123.101.pdf (109 Distribution), même lot** : qté/tarif déjà
  corrects dans le Suivi mais date enregistrée (24/08, date de traitement
  d'une session précédente) différente de la vraie date du BL (21/08) —
  cas "à confirmer" déjà couvert par le garde-fou de cohérence de date
  (voir plus haut) ; confirmé par l'acheteur, corrigé et archivé.
- **Recette réelle sur le vrai document YESSS** : commande M4.273, BL
  CAM/040759, référence 411651 (DX3-ID2P63AA30MATGA), qté 2, prix net
  39,07€, montant 78,13€ — tout exact du premier coup après le fix de
  date. Fixture réelle ajoutée (`tests/fixtures/bl_yesss_1.pdf`), testé
  (`tests/test_parsers_bl_yesss.py`, 2 tests). Suite complète : 238
  passés.

## Session suivante — lot mélangeant 5 fournisseurs, 4 bugs réels corrigés,
2 nouveaux fournisseurs (DEM, PROTECTHOMS) (fait)

Premier fichier réel à mélanger AUTANT de fournisseurs différents d'un
coup (RAVATE, Coredime, Cominter, Electric Plus, 109 Distribution — 5
pages, 5 fournisseurs) : la détection par page (voir section précédente)
les a tous séparés correctement dès le premier passage automatique.

- **BUG RÉEL CORRIGÉ — RAVATE, "AU" et la date collés par des tirets**
  (commande 135.049) : "AU-25/08/2026--135-049" — les deux `\s*` du motif
  de commande n'acceptaient QUE des espaces, jamais un tiret à cet endroit
  précis (distinct du tiret déjà toléré DANS la commande, "131-162"). La
  commande ressortait introuvable. `\s*` élargi en `[\s-]*` aux deux
  endroits — le tiret capturé reste normalisé en point par le code déjà
  en place.
- **BUG RÉEL CORRIGÉ — RAVATE, séparateur de date lu "."** :
  "25/08.2026" au lieu de "25/08/2026" — motif élargi pour accepter "/"
  OU "." aux deux séparateurs.
- **Limite RAVATE non corrigée (scan de moindre qualité, règle d'or)** :
  seules 2 des 4 lignes d'articles de ce document sont extraites (les 2
  autres ont des cellules chiffrées trop abîmées) — l'autocontrôle Total
  HT (1681,61€ affiché vs 1030,99€ extrait) le signale honnêtement.
  Fixture réelle ajoutée (`bl_ravate_9_commande_tirets_colles.pdf`),
  testé.
- **BUG RÉEL CORRIGÉ — ELECTRIC PLUS, en-tête de colonnes fusionné avec le
  1er article** (commande M4.269) : l'OCR a groupé "DESIGNATION QTE PRIX
  UNIT.HT..." sur la MÊME ligne visuelle que la référence+désignation du
  1er article ("PLA11525 EMBTMOULUREKEVA32MMX12MM DESIGNATION QTE..."),
  faisant échouer cette ligne entièrement et perdre la référence du 1er
  article (la ligne suivante récupérait à tort un bout de désignation
  comme référence). `_zone_tableau_electricplus()` détache désormais les
  cellules qui précèdent le mot d'en-tête et les reporte sur la ligne
  suivante.
- **BUG RÉEL CORRIGÉ — ELECTRIC PLUS, garde-fou contre une fausse
  référence** : une ligne chiffrée peut se retrouver sans aucune
  référence/désignation adjacente (regroupement Y défavorable, ou
  véritablement absente sur le document) — `cellules[0]` est alors la
  QUANTITÉ elle-même ("70,00MTR"), jamais une vraie référence.
  `_ligne_vers_article_electricplus()` refuse désormais de produire une
  ligne dans ce cas (plutôt que d'écrire "70,00MTR" comme référence) —
  l'écart de Total HT (86,40€ affiché vs 29,70€ extrait pour 2 lignes
  sûres) signale honnêtement qu'une ligne manque. Fixture réelle ajoutée
  (`bl_electricplus_10_entete_fusionne_1er_article.pdf`), testé.
- **Correction manuelle motivée — COREDIME, commande introuvable**
  (fichier `2026-08-19 - COREDIME - CORB033477.1 - BC inconnue.pdf`) :
  aucun label "BC"/"COMMANDE N°" reconnu sur ce document (label
  "Commanderef:" suivi de "108.26", tronqué — la vraie valeur "108.276"
  apparaît par ailleurs isolée juste avant "Référence LACOUTURE").
  Retrouvée avec certitude par recherche large sur les références de
  l'article dans le Suivi (5-6 correspondances convergentes sur la
  commande "108.276") — PAS un correctif de code (un seul exemple, format
  de label jamais revu ailleurs, règle d'or), corrigée à la main pour ce
  BL précis avant rapprochement.
- **Correction manuelle motivée — COREDIME, référence tronquée** : "600RAL"
  extrait sur le BL correspond à "600RAL7016" dans le Suivi (même
  désignation "RAL 7016", même quantité 8) — la cellule OCR a perdu le
  suffixe "7016". Corrigée à la main (même principe qu'au-dessus).
- **Limite COREDIME non corrigée** : 5 des 7 lignes de ce BL ne sont pas
  extraites — leur confirmation "<qté> x 1 unite" a perdu la quantité de
  tête à l'OCR (juste "x 1 unite" sans chiffre devant), contrairement aux
  2 lignes extraites qui, elles, ont gardé leur chiffre. Toutes ces 5
  lignes ont `qte_commandee=8` dans le Suivi — plausible qu'elles valent
  aussi 8, mais UN SEUL signal (pas de prix chez Coredime pour
  recouper) : pas assez pour deviner, laissées à la vérification de
  l'acheteur sur le papier.
- **Nouveau fournisseur BL : DEM** (déjà couvert côté devis,
  `moteur/fournisseurs/dem.py`, section GABARIT BL ajoutée à la suite du
  parser devis existant — **le parser devis n'a PAS été touché**, un
  premier essai l'avait accidentellement réécrit de mémoire au lieu de
  seulement ajouter du code, repéré et corrigé immédiatement via `git
  diff` avant tout test). Structure du tableau IDENTIQUE au devis (déjà
  documentée : prix AU CENT "/C", désignation sur la ligne suivante) —
  seule différence, c'est un scan OCR au lieu de texte PDF natif. 2 vrais
  BL vus (`bl_dem_1_deux_pages_reste_a_livrer.pdf`,
  `bl_dem_2_six_lignes.pdf`) :
  - **Chaque PAGE est un bon de livraison DEM indépendant** (pas de
    fusion inter-pages) : sur le 1er fixture, les 2 pages ont des n° de
    BL et des dates DIFFÉRENTS (706992 le 24/08, 706990 le 20/08) pour la
    MÊME commande M3.14.363 — `parse_bl_dem()` retourne donc une LISTE,
    comme 109 Distribution/Cominter/Electric Plus.
  - **1er cas réel de "Reste à livrer" chez DEM** : une ligne SANS prix
    ni montant imprimés (désignation collée sur la même ligne visuelle
    faute de place prise par les colonnes de prix vides) — jamais
    livrée, exclue. Contrairement à Coredime, pas besoin de mémoriser un
    drapeau par page : l'absence de prix est en elle-même le signal
    fiable ici. Preuve concrète que l'exclusion est correcte : cette
    MÊME référence/quantité réapparaît PRICÉE sur l'AUTRE page du même
    fichier (la livraison suivante qui la solde).
  - Recette : les 2 fixtures retombent exactement sur leurs Total HT
    respectifs (237,50€ / 87,50€ / 702,66€). Testé
    (`tests/test_parsers_bl_dem.py`, 3 tests) + le test devis existant
    (`test_parse_dem`) toujours vert.
- **Nouveau fournisseur BL : PROTECTHOMS** (`moteur/fournisseurs/
  protecthoms.py`, nouveau module, BL uniquement) — équipements de
  protection individuelle/amiante, PAS du matériel électrique comme tous
  les autres fournisseurs de ce projet, mais confirmé présent dans la
  liste Fournisseurs du Suivi (vérifié directement dans le classeur
  vivant) donc en périmètre. 1 vrai BL vu (commande M3.15.399). Structure
  simple et claire, tableau "Reference produit | Designation | Quantites
  | Reste à livrer" : chaque ligne visuelle est déjà un article complet.
  Référence produit repérée par sa FORME (1 chiffre + 2 lettres + 6
  chiffres, ex. "2VU043003") plutôt qu'une position de cellule ou un
  en-tête/pied de tableau. "Reste à livrer" (4e cellule, présente sur 1
  ligne du seul document vu) est **purement informatif ici, jamais
  soustrait** — contrairement au "Reste à livrer" Coredime/DEM (qui
  EXCLUT toute la ligne) : chez PROTECTHOMS, "Quantites" et "Reste à
  livrer" sont deux colonnes séparées, la ligne EST bien livrée à hauteur
  de "Quantites". Pas de prix du tout sur ce document (comme Coredime) —
  pas d'autocontrôle Total HT possible. Testé
  (`tests/test_parsers_bl_protecthoms.py`, 2 tests), 8/8 lignes exactes
  dès le premier essai.
- **Limite COMINTER non corrigée (`BL 131.165 RAL LAGOURGUE.pdf`, 0 ligne
  extraite)** : scan très dégradé, l'en-tête de colonnes ("Px unitaire
  Rem Px net") s'est retrouvé mélangé à la référence+désignation du seul
  article (même famille de problème que le bug Electric Plus corrigé
  cette session, mais combiné à une qualité de scan bien plus mauvaise) —
  un seul exemple, pas assez pour généraliser un correctif fiable ; laissé
  en l'état dans "à vérifier" pour vérification manuelle sur le papier
  (article probablement "Rail de montage 41X41 en C", qté 2, montant
  73,75€ d'après une lecture visuelle, non extrait automatiquement).

## Session suivante — 2e lot du jour, 3 bugs réels RAVATE corrigés, 2
commandes retrouvées à la main (fait)

- **BUG RÉEL CORRIGÉ — RAVATE, "/" séparateur de date lu "7"** (commande
  M3.18.223, 2 documents du même lot) : "AU16706/2026-M3-18.223" (un seul
  "/" corrompu) et "AU1670672026.M3.18.223" (les DEUX corrompus) — motif
  de commande élargi pour tolérer "/", "." OU "7" aux deux séparateurs de
  la date (sans risque de confusion : jour et mois ont une longueur fixe
  dans le motif). Le séparateur avant la commande elle-même tolère aussi
  "." en plus d'espace/tiret.
- **BUG RÉEL CORRIGÉ — RAVATE, en-tête "Reference" lu "Reterence"** (F
  confondu avec T) : l'ancre de début de tableau ne matchait plus DU TOUT,
  zone vide, 0 ligne extraite pour un Total HT de 204,86€ affiché. "F"
  rendu tolérant au "T".
- **BUG RÉEL CORRIGÉ — RAVATE, virgule décimale lue ":"** ("0:00" au lieu
  de "0,00") — même famille que la confusion "*" déjà tolérée, sans risque
  de confondre avec le cas COMBO existant (2 montants séparés par ":") :
  la partie après les 2 décimales ne peut alors jamais tenir dans les 3
  caractères de fin autorisés, le motif simple échoue proprement et le
  repli combo prend le relais. Fixtures réelles ajoutées
  (`bl_ravate_10_separateur_slash_lu_7.pdf`,
  `bl_ravate_11_entete_reterence_0_ligne.pdf`), testé (2 tests).
- **Limite RAVATE non corrigée** (`bl_ravate_11`, même lot) : la 2e ligne
  d'article de ce document a son Montant HT au point décimal déplacé
  ("4106." au lieu de "41,06") — un seul exemple, pas de règle inventée ;
  l'autocontrôle Total HT (204,86€ vs 163,80€ extrait) le signale
  honnêtement.
- **Correction manuelle motivée — COMINTER, commande APRÈS l'en-tête**
  (fichier multi-BL, 1er des 3 BL) : la recherche de commande est bornée
  entre l'en-tête "Numero/Date/Fin de Validité" et l'en-tête du tableau
  "Designation" (voir `parse_bl_cominter`) — sur ce document précis, la
  valeur "M3.10.166" (label "Référence") est imprimée APRÈS "Designation",
  hors de cette fenêtre pour la première fois observée. Un seul exemple,
  élargir la fenêtre de recherche risquerait de capturer un nombre de la
  1ère ligne d'article par erreur — corrigée à la main pour ce BL précis.
- **Correction manuelle motivée — 109 Distribution, "N° Réf. Client" en
  texte libre** (commande retrouvée : "161.008") : le champ contenait
  "kanoppe clim 161" (un nom de chantier, pas un code commande) au lieu du
  format habituel — retrouvée avec certitude par recherche large sur la
  référence de l'article dans le Suivi (résultat unique, "161" du nom de
  chantier correspond bien au préfixe de "161.008"). Pas un correctif de
  code (format jamais revu ailleurs, règle d'or).
- **Recette réelle sur les 8 fichiers (10 BL) du lot** : 13 lignes sûres
  écrites au total, 8 BL/pages archivés individuellement (dont les 3 BL
  d'un même fichier Cominter groupé). Restent en "à vérifier" : 2 pages
  RAVATE avec une référence introuvable dans le Suivi pour leur commande,
  et le Cominter "131.165" à 0 ligne déjà documenté ci-dessus.

## Session suivante — 3e lot du jour, 2 bugs réels YESSS corrigés sur son
2e vrai BL (fait)

- **BUG RÉEL CORRIGÉ — YESSS, commande imprimée "BC N°..."** (2e vrai BL
  vu, `BL M4.276.pdf`) : format DIFFÉRENT du 1er document ("N° commande
  M4.273", "commande" en toutes lettres) — ici "BC N°M4.276" en un seul
  mot OCR, sans espace entre "N°" et la valeur. Le label "N° commande"
  existe bien aussi sur ce document, mais comme label de colonne isolé
  (aucune valeur adjacente dans le même mot) : les deux formats sont
  désormais essayés, sans risque de faux positif sur ce label isolé (rien
  à capturer juste après "commande" dans son propre mot).
- **BUG RÉEL CORRIGÉ — YESSS, date imprimée en un seul mot bien formé** :
  "25 aout 2026" (jour+mois+année ensemble), alors que sur le 1er document
  le jour ressortait comme un mot OCR séparé du mois+année. Le motif
  mois+année capture désormais un jour optionnel directement dans le même
  mot, avant de retomber sur la recherche par proximité si absent
  (comportement du 1er document inchangé). Fixture réelle ajoutée
  (`bl_yesss_2_commande_bc_no_et_date_groupee.pdf`), testé.
- **Recette réelle sur les 3 fichiers du lot** : 2 lignes écrites, 2 BL
  archivés (RAVATE M4.275, YESSS M4.276). Reste en "à vérifier" : RAVATE
  126.049 (1 ligne à confirmer — repli référence proche "Z" vs "2", 1
  ligne inconnue — référence "T4" absente du Suivi pour cette commande,
  potentiellement un article réellement différent).

## Session suivante — recette d'un lot de 11 fichiers/12 BL, incident de
sécurité sur le Suivi (écriture pendant qu'il était ouvert) et
récupération complète, puis 4 cas "à vérifier" tranchés (fait)

- **INCIDENT DE SÉCURITÉ, leçon retenue pour toute session future sur ce
  projet** : l'acheteur avait explicitement prévenu "attention le suivi
  est ouvert, demande-moi quand tu veux y accèder" — une écriture a quand
  même été lancée sans lui redemander confirmation explicite au moment
  précis d'écrire, en se fiant au fait que le verrou technique
  (`ClasseurVerrouille`, basé sur la présence d'un fichier `~$...` créé
  par Excel) ne remonterait probablement pas d'erreur. Cette fois-là, le
  verrou n'a PAS bloqué l'écriture (cause exacte non identifiée avec
  certitude — partage réseau, mode d'ouverture Excel différent, fichier
  `~$` pas encore créé ou déjà disparu au moment du test) alors que le
  fichier était réellement ouvert côté acheteur. Sa sauvegarde de sa
  propre session Excel (qui ne contenait pas ces écritures) a ensuite
  écrasé le fichier, faisant apparemment perdre 10 lignes fraîchement
  écrites — "tu as ecris sur le fichier alors qu'il était ouvert de mon
  côté, c'est vraiment pas bien". **Récupération complète faite sans
  jamais toucher à l'état Excel de l'acheteur** : les 9 lignes
  récupérables ont été re-dérivées à partir des BL déjà archivés
  (relecture OCR + rapprochement + réécriture), vérifiées identiques aux
  valeurs d'origine — rien perdu au final, mais un incident évitable et
  un moment de confiance entamé. **Règle désormais appliquée sur ce
  projet, retenue en mémoire (feedback)** : l'absence d'exception
  `ClasseurVerrouille` n'est JAMAIS une confirmation suffisante que le
  Suivi est fermé — dès qu'un accès a été signalé ouvert (ou pourrait
  l'être), ne plus écrire tant que l'acheteur n'a pas répondu
  explicitement "c'est fermé" à une question posée dans le chat, même si
  le code ne remonte aucune erreur. **Confirmé utile de nouveau plus tard
  dans cette même session** (voir M4.277 ci-dessous) : le verrou technique
  a cette fois bien bloqué une écriture alors que le Suivi était
  effectivement rouvert — reste un filet de sécurité réel, mais jamais
  suffisant seul, toujours la parole de l'acheteur qui prime.
- **M3.27.002, ISHOP.pdf/ISHOP2.pdf confirmés comme 2 BL réellement
  différents par l'acheteur** ("On voit clairement qu'il s'agit de deux
  BL différents !") — le désamorçage anti-doublon avait à raison empêché
  d'écrire les deux en même temps (même ligne Suivi visée, référence
  251745). Vérification a posteriori : BL 741681 (8 unités) + BL 741706/
  ISHOP2 (6 unités) = 14, exactement la quantité commandée — confirmation
  numérique forte qu'il s'agit de deux livraisons partielles réelles, pas
  d'un doublon de scan. Écrit (14 livrées) et archivé dans
  `Traités/M3.27.002/`.
- **M3.15.399 (PROTECTHOMS), 2e livraison réelle confirmée par
  l'acheteur, mais PAS uniforme sur les 5 lignes du 2e BL** — leçon à
  retenir pour tout futur cas "2e BL du même fournisseur/commande" :
  comparer chaque ligne individuellement à la quantité commandée, ne
  jamais supposer que tout le document représente une livraison nouvelle
  sous prétexte qu'au moins une ligne l'est (ni l'inverse). BL097191
  (22/08, déjà traité) vs BL097312 (26/08, nouveau) : sur les 5 articles
  listés sur le 2e BL, un seul (**2VU043004**, combinaison XL) est une
  vraie nouvelle livraison (20 de plus, 30+20=50=exactement la quantité
  commandée) — écrit. Les 4 autres (2VU043003, 2VU061300, 5CO010100,
  1MA010701) réapparaissent avec EXACTEMENT la même quantité que la 1ère
  livraison ; 3 des 4 étaient déjà à 100% de leur quantité commandée
  (150/150, 30/30, 300/300) — une redélivraison identique d'un article
  déjà soldé n'ayant pas de sens commercial, ces 4 lignes ont été
  laissées SANS écriture (le 2e BL les reliste simplement, sans nouvelle
  livraison réelle pour elles). Décision communiquée à l'acheteur, pas
  encore commentée par elle au moment de la rédaction de cette note — à
  corriger si son avis diffère.
- **M3.23.046 (DEM), confirmé comme un vrai doublon de scan** (tranché
  par les données elles-mêmes, sans avoir besoin de redemander à
  l'acheteur) : le "nouveau" BL (n°741449 du 21/08, contre 706994 du
  24/08 déjà archivé) montre des quantités IDENTIQUES sur 5 des 6 lignes
  (la 6e, LEG031490, absente de l'OCR sur ce 2e scan) — et les 6 articles
  de la commande sont déjà TOUS à 100% de leur quantité commandée dans le
  Suivi. Contrairement au cas M3.15.399 ci-dessus (où 1 ligne sur 5
  montrait un delta réel qui complétait exactement la commande), ici
  AUCUNE ligne ne montre de delta — signal net de duplication plutôt que
  de 2e livraison. Déplacé vers `À vérifier/Doublons confirmés (à
  supprimer)/` (jamais supprimé directement par la session).
- **M4.277 (COMINTER), commande retrouvée depuis un passage précédent de
  cette session** : absente du Suivi la première fois, elle y figure
  maintenant (probablement saisie entre-temps par l'acheteur) —
  référence Suivi "A9F87463" correspond à la référence BL "MEA9F87463"
  (préfixe "ME" en trop, même famille que les cas déjà connus type
  MEA9Y13625). 1ère tentative bloquée par le verrou Excel (Suivi rouvert
  par l'acheteur en cours de session, voir l'incident ci-dessus) ; écrite
  et archivée dans `Traités/M4.277/` après confirmation explicite de
  l'acheteur que le classeur était refermé.
- **M3.10.186 (COREDIME), 0 ligne à l'OCR par défaut — PAS une limite du
  document, une résolution de rendu insuffisante.** L'acheteur a
  directement contesté la conclusion "limite OCR connue" en partageant un
  extrait du PDF où les quantités ("1" et "4") sont parfaitement lisibles
  — à raison : à l'OCR par défaut (200 DPI), les deux chiffres de
  quantité ET le début de leur confirmation ("1 X 1 unite"/"4 X 1 unite")
  ont totalement disparu (ne restait que "unite" seul), pas seulement
  tronqués comme dans les cas déjà documentés ailleurs. En relançant l'OCR
  de cette seule page à 350 DPI (`mots_document(..., dpi=350)`), les deux
  lignes ressortent immédiatement complètes et exactes
  ("1 x 1 unite"/"4 X 1 unite"). **Pas généralisé au reste du pipeline**
  (un seul exemple à ce jour, règle d'or — passer tout le rapprochement
  BL à 350 DPI par défaut ralentirait chaque lecture d'environ 3× la
  surface de pixels, pour un gain qui n'a été nécessaire qu'une fois) :
  traité comme une correction ponctuelle sur ce document précis. Écrit
  (SCHA9P22602 qté=1, LEG033327/033327 qté=4, aucun prix — comme toujours
  chez ce fournisseur) et archivé. **Leçon générale à retenir** : avant de
  conclure à une "limite OCR connue" sur un nouveau document, essayer une
  résolution plus élevée sur CE document précis avant d'abandonner —
  certains cas qui semblent être une vraie limite structurelle ne sont en
  réalité qu'un problème de résolution de rendu, résoluble sans toucher
  au code.

## Session suivante — incident critique "quantités non entières", audit
complet du Suivi et de Traités/, 2 bugs réels corrigés (Cominter, RAVATE)
(fait)

**Signalement critique de l'acheteuse** : appelée par le responsable
d'affaires du GYSM pour savoir si la commande 143.194 avait été livrée,
elle constate en ouvrant le Suivi des quantités livrées ABERRANTES —
"1,32 cutter livrés sur 4 commandés", "1,3015 boîte d'embouts sur 2",
tarifs "totalement fantaisistes" — et signale qu'en cherchant le BL dans
"Traités/143.194/" elle ne trouve que le BdC, pas le BL. Règle donnée,
générale et sans exception cochée par elle : **"à part s'il s'agit de
main d'œuvre, d'une prestation ou d'une location, il ne peut y avoir que
des entiers en quantité."** Elle demande explicitement une remise à plat
de la procédure de classement ("si commandes et BL ne sont pas attachés
dans le dossier de rapprochement, alors il ne faut pas écrire") et un
audit de bout en bout de ce qui est dans "Traités". Ces quantités non
entières traînaient dans le Suivi depuis un peu plus tôt dans CETTE MÊME
session (récupération après l'incident du verrou Excel, voir plus haut) —
un vrai signal de bon sens (qté non entière = alerte immédiate) qui
aurait dû être vérifié avant d'écrire et ne l'a pas été.

- **Audit complet du Suivi (~6 370 lignes, toutes années/sessions
  confondues), sur décision explicite de l'acheteuse** : scan de toute la
  colonne "Qté livrée" à la recherche de valeurs non entières. **11
  anomalies trouvées, dont 2 légitimes** (LECHER0618, une LOCATION
  d'échafaudage, qté=2,5 — une durée peut être fractionnaire ; MOA2, MAIN
  D'ŒUVRE, qté=0,5 — explicitement les 2 exceptions de la règle) — **9
  vraies anomalies** réparties sur 5 commandes (M2.17.006 ×4, M3.23.042
  ×1, M4.272 ×1 déjà présente + 4 manquantes découvertes en creusant,
  143.194 ×2, M3.27.002 ×1). Aucune colonne "unité" n'existe dans le
  Suivi — seule la désignation permet de juger au cas par cas si une
  ligne est légitimement non entière.
- **BUG RÉEL CORRIGÉ (Cominter, code) — la capacité de coupure du
  disjoncteur DNX³ ("4.5KA"/"4.5 KA", normalement collée en fin de
  désignation) se retrouve parfois dans SA PROPRE cellule OCR**, sur 3
  documents réels distincts (M2.17.006/OBL108540, M3.23.042/OBL108537,
  M4.272/OBL108653). Comme elle commence aussi par un chiffre, l'ancien
  test `re.match(r"^\d", ...)` (dans `_ligne_bl_vers_article_cominter()`,
  boucle de détection de fin de désignation) la prenait à tort pour la
  cellule Qté(+Unité), décalant tout le reste d'une cellule — d'où les
  "4,5" disjoncteurs livrés (au lieu de 10/20/7/5/4/1, la vraie quantité,
  toujours à la cellule suivante). **DEUX correctifs cassés avant le
  bon, gardés en commentaire dans le code pour ne pas y retomber** :
  1. Exiger un `fullmatch` sur exactement 2 décimales — cassait le format
     M4.272, où Qté+Unité sont dans la MÊME cellule ("7,00 Unite", du
     texte suit les 2 décimales, un fullmatch échoue).
  2. Exiger une VIRGULE (pas un point) comme séparateur — cassait une
     vraie quantité écrite avec un point ("3.00 Unite",
     bl_cominter_3.pdf, test déjà existant) : virgule/point ne
     distinguent PAS de façon fiable une vraie quantité de "4.5KA" (les
     deux séparateurs existent des deux côtés selon les documents).
  Signal qui tient sur TOUS les cas réels observés (2 fournisseurs
  confondus, virgule ET point) : une vraie cellule Qté a TOUJOURS
  exactement 2 chiffres après le séparateur ("10,00", "3.00", "542,00"),
  jamais un seul comme "4.5KA". `re.match` (préfixe, pas `fullmatch`)
  pour continuer d'accepter tout ce qui suit, collé ou espacé. Fixture
  réelle ajoutée (`tests/fixtures/bl_cominter_8_kA_dans_sa_propre_cellule.pdf`,
  `test_parse_bl_cominter_8_capacite_de_coupure_dans_sa_propre_cellule`),
  suite complète Cominter (10 tests) et suite complète du projet (262
  tests) vertes après ce correctif.
- **M2.17.006 et M3.23.042 (déjà archivés)** : simplement re-parsés avec
  le code corrigé — les valeurs correctes (10/10/20/10 et 10) tombent
  directement, sans aucune reconstruction manuelle nécessaire. Écrites en
  correction directe.
- **M4.272 (BL resté dans "à vérifier", jamais écrit ni archivé
  auparavant)** : re-parsé avec le code corrigé, les 11 lignes tombent
  TOUTES exactement sur leur quantité commandée (confirmation que la
  commande est entièrement soldée). 2 références corrompues par le même
  phénomène "4.5KA" mais différemment cette fois (touchant la référence
  elle-même, pas seulement la quantité) : "L4067734" → 406773 (un "4" de
  trop, probablement un fragment de "4.5KA" glissé dans la référence) et
  "L41165C" → 411650 (dernier chiffre "0" lu "C") — reconnues par
  élimination (qté déduite matchant exactement la seule ligne Suivi
  restante de cette commande) et corrigées à la main (un seul exemple de
  ce sous-cas précis, pas de règle générale codée). Écrit et archivé.
- **143.194 (BL resté dans "à vérifier" avec 1 ligne A_CONFIRMER
  ancienne, jamais archivé)** : reconstruit entièrement à la main à
  partir de l'OCR brut (cellules très éclatées sur ce document RAVATE) —
  1004683 (qté 4, prix 6,70€), DT7172 (qté 2, prix 5,59€), DT7386T/le cas
  "DT73B6T" déjà documenté (qté 2, prix 5,98€), DT7520 (déjà correct,
  qté 4). **Vérifié par DEUX signaux convergents** : chacune des 4
  quantités correspond EXACTEMENT à sa quantité commandée, ET la somme
  des montants reconstruits (26,80+11,18+11,96+18,44 = 68,38€) retombe
  EXACTEMENT sur le Total HT imprimé — confirmation qu'aucune des 4
  lignes n'était une vraie anomalie, seulement des quantités mal
  extraites par l'ancien code (division bruitée / cellule mal appariée).
  Écrit et archivé.
- **BUG RÉEL CORRIGÉ (RAVATE, code) — préfixe de nom de chantier collé
  devant la commande sans espace**, découvert en archivant 143.194 : le
  document affiche "AU 24/0B/2026GYSM-143.194" — le motif de commande ne
  tolérait qu'UNE SEULE lettre de préfixe (le "M"/"i"/"o" habituel FAISANT
  PARTIE de la commande elle-même), pas tout un mot comme "GYSM".
  `numero_commande` ressortait vide et le BL s'est d'abord archivé à tort
  dans "Commande inconnue" au lieu de "143.194" (repéré et corrigé dans
  la foulée, avant que l'acheteuse ne le voie). Motif élargi pour
  tolérer un mot de lettres majuscules suivi d'un tiret, optionnel, AVANT
  la commande elle-même — ne change rien aux cas déjà couverts (un "M"
  isolé ne peut jamais matcher "LETTRES-", faute du tiret qui suit
  immédiatement dans la commande réelle). Fixture réelle ajoutée
  (`tests/fixtures/bl_ravate_12_prefixe_chantier_colle_commande.jpg`,
  `test_parse_bl_ravate_12_prefixe_chantier_colle_commande`), suite
  RAVATE (27 tests) et suite complète (262 tests) vertes.
- **260630 (M3.27.002, 109 Distribution, déjà archivé)** : qté 1,0001 →
  1,0 — bruit d'arrondi pur (Total HT affiché "39,56" est le P.U.Net
  "39,55502" arrondi à 2 décimales pour un seul article, pas une vraie
  division qté×prix) sur un fournisseur dont le mécanisme (Total/PxNet
  systématique, jamais de cellule Qté imprimée préférée) reste
  correct pour la quasi-totalité des documents déjà traités — traité en
  correction manuelle ponctuelle (un seul exemple de ce bruit résiduel à
  ce jour), pas un correctif de code sur `dist109.py`.
- **Audit complet de `a_traiter/BL/Traités/`, sur décision explicite de
  l'acheteuse (136 dossiers commande)** : seulement 2 écarts trouvés.
  143.194 (BdC présent, BL absent — en cours de résolution ci-dessus,
  désormais réglé) et M3.15.399 (Protecthoms, BdC absent mais les 2 vrais
  BL bien présents — recherché explicitement dans l'archive externe des
  BC, réellement introuvable là-bas, pas un bug de la recherche ; gap
  mineur, signalé mais non bloquant). Aucun dossier vide. Tous les autres
  dossiers ont bien leur BdC ET leur BL.
- **Récapitulatif des écritures de cette session** : 15 lignes
  corrigées/écrites au total sur 5 commandes, suite complète du projet
  (262 tests) verte après les 2 correctifs de code, plus aucune quantité
  non entière dans tout le Suivi hors les 2 cas légitimes (location,
  main d'œuvre).

**Leçon générale retenue (feedback), à appliquer systématiquement dans
toute session future** : une quantité livrée non entière (hors main
d'œuvre/prestation/location) est TOUJOURS un signal d'alerte — jamais
l'écrire sans d'abord vérifier qu'elle correspond à un signal fort
(quantité commandée exacte, ou reconstruction qui retombe sur le Total HT
affiché). Ce contrôle de bon sens, simple et rapide, aurait évité tout cet
incident.

## Session suivante — lot de 8 BL du jour, règle "quantité entière" appliquée
systématiquement avant écriture, 3 cas de mauvaise cellule de quantité (fait)

Premier lot traité après la leçon ci-dessus : chaque ligne "sûre" par
l'automatique a été recroisée avec la quantité commandée avant écriture — a
immédiatement débusqué 2 lignes RAVATE avec des quantités fractionnaires
(0,7336 et 1,1962) que le rapprochement automatique aurait classées "sûres"
sans lever aucune alerte (la quantité obtenue restait sous la quantité
commandée, donc pas de garde-fou sur-livraison déclenché).

- **RAVATE (139.118, 161.012) — la cellule Px Brut, si illisible par l'OCR,
  décale la fenêtre des 4 dernières cellules et fait confondre Remises avec
  Px Net.** `_ligne_chiffree_bl_ravate()` bascule sur un repli à 3 cellules
  dès qu'UNE des 4 dernières ne parse pas comme un montant — ce repli
  suppose TOUJOURS que c'est Px Brut qui a disparu (absorbé par une cellule
  qté+unité collée, cas déjà documenté, M3.10.182) et garde donc les 3
  dernières comme [Px Brut, Px Net, Montant]. Mais ici c'est Px Brut
  LUI-MÊME qui échoue à parser ("1666" au lieu de "16,66", séparateur
  perdu ; "8T6" au lieu de "8,76", lettre substituée) — les 3 dernières
  cellules sont alors réellement [Px Net, Remises, Montant], pas
  [Px Brut, Px Net, Montant] : Remises se retrouve utilisé comme Px Net
  (192,20€ au lieu de 7,05€ ; 125,40€ au lieu de 5,00€), donnant des
  quantités livrées à 0,7336/1,1962 au lieu de 20/30 (la vraie quantité,
  imprimée en clair juste avant ce bloc de 4 cellules). **Pas de correctif
  de code** (2 documents du même lot, cause chaque fois différente —
  séparateur perdu vs lettre substituée — pas assez homogène pour une règle
  fiable sans risquer de casser le repli existant, qui, lui, reste
  nécessaire pour son cas d'origine ; règle d'or) : corrigé à la main sur
  ces 2 lignes, vérifié par la quantité commandée (20 et 30, exactement) ET
  par le Total HT imprimé (141,00€ et 150,00€, chacun retombant pile avec
  qté×Px Net).
- **COREDIME (123.108) — "Gaine ICT" vendue en couronnes de 100m, un code
  de désignation numérique se glisse juste avant la vraie quantité.** Cellules
  réelles : `['LEG06620', 'ICTA3422', '20', 'ATF', 'STANDARD', '100M', '400', 'M*']`
  — le "20" (code de taille, Ø20mm, faisant partie de la désignation) et le
  "400" (la VRAIE quantité, en mètres, juste avant l'unité "M*") sont tous
  les deux des nombres isolés dans des cellules séparées ; le code actuel a
  pris le premier ("20") au lieu du dernier avant l'unité ("400"). Qté
  livrée réelle = 400 (LEG06620) et 300 (LEG06625) — vérifié : correspond
  EXACTEMENT à la quantité commandée dans les deux cas (contre 20/25 extraits
  à tort). Un seul document à ce jour, traité en correction manuelle.
- **Nouveau préfixe fabricant confirmé par l'acheteuse : "SIB"** (référence
  BL "SIBP02120" = référence Suivi "P02120", "ajouté au code dans la base
  article SONEPAR") — à ajouter à la liste des préfixes déjà connus
  (LEG/EBE/PW/BT) si un 2e cas se présente avec un vrai gain à en tirer via
  `deduire_prefixes()`.
- **ELECTRIC PLUS (139.117) — décalage cascadé entre référence et données
  chiffrées sur tout le document**, plus grave que les cas déjà documentés
  (désignation qui déborde, "BLANC" pris pour une référence) : sur ce
  document précis, la donnée chiffrée qui SUIT une référence dans le flux
  OCR n'est PAS toujours la sienne — reconstruit entièrement par
  correspondance avec la quantité commandée de chaque référence (5
  correspondances trouvées, chacune vérifiée par une égalité EXACTE
  qté×prix=montant) : 077040L=2 (5,11€), 077111L=100 (2,70€), 077011L=6
  (3,39€), 076565=25 (7,86€), 080251=50 (0,82€). **077001L (qte_cmd=2)
  n'a aucune donnée correspondante trouvée sur ce document** — pas livrée
  cette fois, pas forcée. **~10,56€ d'écart résiduel sur le Total HT
  (548,62€ affiché contre 538,06€ reconstitué) reste inexpliqué** — laissé
  tel quel plutôt que deviné une 6e ligne fantôme (l'autocontrôle existant
  le signale honnêtement).
- **`BL E26.009.pdf` : PAS un bon de livraison fournisseur** — c'est un BON
  DE COMMANDE émis par ELECTRICITE SERVICES REUNION elle-même (gants
  Maxiflex, gilets WTP — EPI), déposé par erreur dans `a_traiter/BL/` au
  lieu d'être un BL d'un fournisseur. Laissé tel quel, signalé à l'acheteur
  plutôt que forcé dans le pipeline de rapprochement (aucun fournisseur à y
  reconnaître, ce n'est structurellement pas le même type de document).
- **Recette sur les 7 vrais BL du lot** : 12 lignes écrites, toutes
  entières, toutes vérifiées soit par correspondance exacte avec la
  quantité commandée soit par reconstitution du Total HT — 7 BL archivés
  avec leur BC.

## Session suivante — lot de 16 BL du jour, ligne RAVATE silencieusement
perdue trouvée par recoupement Total HT, cluster 123.107 (10 lignes)
reconstruit à la main, typo OCR "COMMANBDE" (fait)

- **RAVATE (M3.23.051), même famille de bug que la session précédente
  (Px Brut illisible décale la fenêtre)** : la ligne 404926 (qté_cmd=20)
  ressortait à 11,6583 — cellule qté+unité collée (":UN:20,00:") non
  reconnue par `_ligne_chiffree_bl_ravate()`, retombant sur le repli à 3
  cellules qui a pris Remises pour Px Net. Corrigé à la main : qté=20
  (imprimée en clair), prix_net=1,99€. **Plus grave, trouvé en recoupant
  le Total HT (227,75€) plutôt que ligne par ligne** : une 2e ligne du
  même BL (069831L, qté_cmd=20) n'apparaissait dans AUCUNE catégorie du
  rapport (ni SUR, ni A_CONFIRMER, ni INCONNU) — silencieusement perdue
  car ses cellules chiffrées ("80.6", ":29.9") ne parsaient pas du tout.
  Reconstruite par soustraction (Total HT − les 3 autres lignes connues =
  132,60€, qté=20 → prix_net=6,63€) — la somme des 4 montants
  (39,80+156,00+195,00+58,50... voir détail complet ci-dessous) retombe
  exactement sur 227,75€. **Leçon générale, au-delà de ce cas précis** :
  ne pas se contenter de vérifier que les lignes RENVOYÉES sont
  cohérentes — vérifier aussi que leur NOMBRE correspond aux lignes
  visibles sur le document ("X lignes" imprimé en pied de tableau chez
  RAVATE), sans quoi une ligne entière peut disparaître sans qu'aucune
  anomalie ne soit levée.
- **RAVATE (123.107, 2 pages), reconstruction complète de 10 lignes de
  câbles en couronnes de 100m** : toutes les références ressortaient
  INCONNU car le préfixe "H07" est systématiquement absent de l'OCR
  ("VU1.5VJ" au lieu de "H07VU1.5VJ" — Suivi attend la forme complète).
  La page de continuation ("123.107 SUITE.jpg", pied de tableau avec le
  Total HT réel de 1627,00€) n'a donné 0 ligne — seule la 1ère page
  contient le tableau. Reconstruit entièrement à la main à partir du
  nombre de couronnes imprimé par ligne (×100m, prix par couronne divisé
  par 100) : 9 des 10 lignes ont un nombre de couronnes directement
  lisible et un prix par couronne strictement identique au sein de chaque
  famille de section (19,50€/couronne pour tout le 1,5mm², 29,50€ pour
  tout le 2,5mm²) — ce motif a servi à déduire la 10e ligne (VU2.5R,
  qté=1100, aucune cellule de qté lisible) par élimination. Vérifié :
  somme des 10 montants = 1627,00€ = Total HT exact de la page 2.
- **COREDIME (M3.23.049) : l'intitulé même de l'ancre était corrompu par
  l'OCR** — "Rfrence COMMANBDE No M3.23.049" (un "B" inséré dans
  "COMMANDE") a fait échouer le motif de recherche de commande, jamais un
  problème de cellule de données cette fois. Commande retrouvée par
  lecture directe du texte brut (aucune ambiguïté, le nom du fichier la
  confirmait aussi) ; référence SIBP01120 = P01120 (même famille de
  préfixe SIB que la session précédente), qté=80 confirmée par "80 X 1
  unite" imprimé, correspond exactement à la quantité commandée.
- **Fichiers hors périmètre confirmés** : `BL E26.009.pdf` (toujours le
  même bon de commande interne, pas déplacé faute de consigne) et un
  nouveau document RAVATE PRO (`doc...164035.pdf`) — fournisseur reconnu
  par le détecteur mais aucun parser BL n'existe encore pour cette
  entité (distincte de RAVATE Elec, voir tableau fournisseurs plus haut) ;
  laissé de côté, pas forcé.
- **2 fichiers archivés au mauvais endroit ("Commande inconnue") malgré une
  écriture manuelle réussie** : `123.107 SUITE.jpg` (0 ligne/pas de
  commande propre à cette page de continuation) et `M3.23.049.pdf` (typo
  OCR ci-dessus, non résolue par un 2e passage automatique) — corrigés à
  la main juste après l'archivage (dossier `123.107/` et nouveau dossier
  `M3.23.049/` créé), aucune donnée perdue. **Leçon opérationnelle** :
  une correction manuelle de commande faite pour l'ÉCRITURE Suivi n'est
  PAS automatiquement répercutée sur l'ARCHIVAGE si ce dernier repose sur
  un nouveau passage de `lire_bl()` — toujours vérifier le dossier de
  destination après un archivage impliquant un cas corrigé à la main.
- **Script d'archivage interrompu en cours de lot (13 fichiers), cause
  non identifiée avec certitude** (aucun message d'erreur informatif au
  moment du plantage, juste une exception `shutil`/`CopyFile2`) — les 4
  premiers fichiers archivés avant l'arrêt, relancé sans problème sur les
  11 restants juste après. Cohérent avec les pannes ponctuelles déjà
  documentées (fichier verrouillé transitoirement par un autre
  processus) plutôt qu'un vrai bug — à surveiller si ça se reproduit.
- **Recette finale** : 58 lignes du Suivi écrites au total (46 sûres
  automatiques + 12 reconstruites/corrigées à la main), toutes entières,
  14 BL/pages archivés avec leur BC quand trouvé.

## Rapprochement factures — cadrage F1 (préparation, aucun parser écrit)

Suite de [Rapprochement AI](#rapprochement-ai-moteurrapprochement--branche-en-cours) :
le rapprochement des BL a résorbé une bonne partie des lignes
"🚚 Reçue (Attente Facture)" ; reste l'étape hebdomadaire des FACTURES
fournisseurs, question restée ouverte depuis R1 ("comment le rapprochement
factures se traduit-il sur la feuille, faute de colonne dédiée ?"). Décidée
cette session par l'acheteur : DE NOUVELLES COLONNES SAISIES dédiées. Cette
session prépare le terrain en 3 volets, sans écrire de parser — le Volet 1
est une PROPOSITION à valider, le Volet 2 est un cadrage sur données
réelles (79 vraies factures 109 Distribution déjà déposées), le Volet 3 est
un jeu de décisions d'architecture proposées.

### Volet 1 — Colonnes facture (colonnes et emplacement validés par l'acheteur)

**Colonnes à ajouter dans le tableau structuré "Commandes"** (jamais par
l'outil — ajouter des colonnes à un classeur de 16 tableaux structurés
n'est pas dans le périmètre du patch XML chirurgical, voir
`moteur/rapprochement/ecriture.py` — c'est l'acheteur qui les crée dans
Excel, "Insérer > Colonne de tableau" DEPUIS l'intérieur du tableau pour
que les 16 tableaux structurés et leurs formules s'étendent automatiquement
aux nouvelles colonnes) :

| Colonne | Type | Saisie ou calculée |
|---|---|---|
| N° facture | texte | **saisie** (Rapprochement AI écrira dedans) |
| Date facture | date | **saisie** |
| Qté facturée | nombre | **saisie** |
| PU facturé | nombre | **saisie** |
| Montant facturé HT | nombre | calculée = `Qté facturée × PU facturé` |
| Écart facture | nombre | calculée = `Montant facturé HT − Facturé BL` |

**Pourquoi Qté facturée ET PU facturé séparés** (plutôt qu'un seul Montant
facturé HT saisi directement) : une facture Coredime/DEM peut être
PARTIELLE par rapport au BL (la même situation que "Reste à livrer" côté
BL, déjà documentée) — avoir la quantité à part permet de voir d'un coup
d'œil si la facture couvre bien toute la quantité livrée. Et avoir le PU à
part rend un écart de PRIX visible sans avoir à diviser à la main — le cas
réel déjà rencontré côté BL (109/GOUJON, Cominter/Planet Wattohm...) où le
prix facturé diffère légèrement du prix du BL doit sauter aux yeux, pas
être caché dans un montant global.

**Emplacement proposé** : juste après "Tarif BL" (colonne 17 actuelle),
pour rester groupées avec les 4 colonnes de saisie déjà existantes
(Date de livraison / Qté livrée / Tarif BL / Note) — "Montant facturé HT"
et "Écart facture" peuvent suivre juste derrière, ou "Écart facture"
peut aller à côté de "Facturé BL" (colonne 48 actuelle) puisqu'elle
compare directement les deux. Comme les formules du classeur utilisent
des références structurées (`Commandes[[#This Row],[Nom colonne]]`, pas
des lettres de colonne), la position exacte n'a AUCUN impact technique —
c'est un pur choix d'ergonomie de saisie pour l'acheteur, elle tranche.

**Formule de "Facturé BL" actuelle (rappel, déjà en place)**, sur laquelle
"Écart facture" s'appuie :
```
=IFERROR(IF(Commandes[[#This Row],[Tarif BL]]<>0,
             Commandes[[#This Row],[Tarif BL]]*Commandes[[#This Row],[Qté livrée]],
             Commandes[[#This Row],[Tarif convenu]]*Commandes[[#This Row],[Qté livrée]]),"")
```
confirme le bandeau de `ecriture.py` : c'est un MONTANT calculé
(Tarif BL × Qté livrée, ou Tarif convenu × Qté livrée si pas de BL), pas un
indicateur de facture reçue.

**Formules calculées proposées :**
```
Montant facturé HT :
=IFERROR(Commandes[[#This Row],[Qté facturée]]*Commandes[[#This Row],[PU facturé]],"")

Écart facture :
=IFERROR(Commandes[[#This Row],[Montant facturé HT]]-Commandes[[#This Row],[Facturé BL]],"")
```

**Adaptation de "Statut commande"** — formule ACTUELLE (relue sur une
COPIE du classeur vivant, `openpyxl` en lecture seule, jamais le fichier
ouvert par l'acheteur) :
```
=IF(Commandes[[#This Row],[Note]]="Commande annulée","⚪ Annulée",
IF(Commandes[[#This Row],[Note]]="Rupture fournisseur","⛔ Rupture fournisseur",
IF(N(Commandes[[#This Row],[Qté livrée]])=0,"🔵 En attente livraison",
IF(OR(AND(IF(N(Commandes[[#This Row],[Tarif convenu]])>0,N(Commandes[[#This Row],[Tarif convenu]]),N(Commandes[[#This Row],[Tarif Affaire]]))>0,
          N(Commandes[[#This Row],[Tarif BL]])>IF(N(Commandes[[#This Row],[Tarif convenu]])>0,N(Commandes[[#This Row],[Tarif convenu]]),N(Commandes[[#This Row],[Tarif Affaire]]))+0.01),
       AND(IF(N(Commandes[[#This Row],[Tarif convenu]])>0,N(Commandes[[#This Row],[Tarif convenu]]),N(Commandes[[#This Row],[Tarif Affaire]]))=0,
           N(Commandes[[#This Row],[Tarif raisonnable]])>0,
           N(Commandes[[#This Row],[Tarif BL]])>N(Commandes[[#This Row],[Tarif raisonnable]])*1.15)),
   "⚠️ Surfacturation",
IF(OR(N(Commandes[[#This Row],[Qté livrée]])>=N(Commandes[[#This Row],[Qté commandée]]),Commandes[[#This Row],[Note]]="Reliquat soldé"),
   IF(OR(IF(N(Commandes[[#This Row],[Tarif convenu]])>0,N(Commandes[[#This Row],[Tarif convenu]]),N(Commandes[[#This Row],[Tarif Affaire]]))>0,
         AND(N(Commandes[[#This Row],[Tarif BL]])>0,N(Commandes[[#This Row],[Tarif raisonnable]])>0,
             N(Commandes[[#This Row],[Tarif BL]])>=N(Commandes[[#This Row],[Tarif raisonnable]])*0.85,
             N(Commandes[[#This Row],[Tarif BL]])<=N(Commandes[[#This Row],[Tarif raisonnable]])*1.15)),
      "✅ Soldée — Tout OK",
      IF(N(Commandes[[#This Row],[Tarif BL]])>0,"⬜ Soldée (Sans référence / à vérifier)","🚚 Reçue (Attente Facture)")),
   "🟠 Partielle (Reliquat)")))))
```
Elle ne regarde AUCUNE donnée de facturation aujourd'hui (juste Qté livrée
vs Qté commandée, et Tarif BL vs Tarif convenu/Affaire/raisonnable) — c'est
exactement le trou que l'acheteur veut combler.

**Écart avec l'énoncé initial de la consigne ("le seul cas ✅ Soldée
existant aujourd'hui, 2 lignes") — TOUJOURS PAS élucidé, en attente de
réponse de l'acheteur** : 1ère réponse "je ne vois pas quelles sont ces
deux lignes, oublions ce point pour l'instant", puis "en fait il faudra y
revenir, pose moi des questions complémentaires" (questions posées :
source du souvenir — Suivi actuel / fichier de refonte v4 / autre ;
formule ou texte tapé à la main ; texte exact ; repère chantier/commande),
puis un message coupé ("il y a eu une erreur de conception :") suivi de
"non stop". **Devenu sans objet pour "✅ Soldée (Facturée)"** : ce statut
n'est finalement plus construit du tout (voir juste en dessous, décision
explicite de ne pas toucher "Statut commande") — mais la question reste
ouverte en soi si l'acheteur souhaite y revenir un jour.

**DÉCISION EXPLICITE DE L'ACHETEUR (correctif après la 1ère proposition) :
NE PAS toucher à "Statut commande"** — "cette formule vient d'être
refondue autour du contrôle de prix et le Dashboard s'appuie sur ses
libellés exacts (...), y empiler la dimension facture la rendrait
illisible et casserait ces agrégats." **Vérifié sur le classeur vivant
(nouvelle copie, comparaison caractère par caractère) : la formule
"Statut commande" et les comptages exacts par statut (4 507 / 1 571 / 115
/ 113 / 55 / 27 / 23 / 8, identiques à la 1ère lecture de cette session,
voir plus haut) n'ont PAS changé entre la 1ère lecture et cette
vérification** — donc pas de "refonte survenue après la dernière mise à
jour de CLAUDE.md" à ce stade précis ; en revanche, la vérification a
confirmé un point réellement utile et jusque-là non regardé : **le
classeur contient bien 3 feuilles supplémentaires, `Dashboard`, `Analyses`
et `Calculs`** (`wb.sheetnames` sur une copie fraîche), qui référencent la
table structurée `Commandes` — absentes du périmètre de ma 1ère lecture
(bornée à la feuille `Commandes` par `FEUILLE_COMMANDES` dans
`ecriture.py`). **Conséquence actionnable, à ne pas oublier avant toute
écriture réelle sur les futures colonnes facture** : rejouer
`tests/test_rapprochement_ecriture.py::test_ecriture_chirurgicale_sur_le_vrai_suivi_commandes`
en vérifiant explicitement que le contenu de CES 3 feuilles (pas
seulement `Commandes`) reste identique octet pour octet après un patch —
le calcChain/les dépendances inter-feuilles sont plus riches qu'avec la
seule feuille `Commandes` déjà testée jusqu'ici.

**`trouver_fichier_suivi_vivant()` (`moteur/rapprochement/pipeline_bl.py`)
revérifié** : un seul fichier correspond au motif `*Suivi commandes*.xlsx`
dans le dossier "Commandes courantes" aujourd'hui (`1.3.0.1. Suivi
commandes - 2026.xlsx`) — aucun "Suivi nouveau..." ni "copie ...xlsx" trouvé
à côté à ce jour, donc rien à corriger dans l'exclusion actuelle ; à
revérifier si/quand un fichier de refonte réapparaît réellement dans ce
dossier.

**"Statut facture" — nouvelle colonne CALCULÉE, séparée, proposée à la
place de la modification de "Statut commande"** : même style de formule
(IF en cascade sur références structurées), mais totalement indépendante
— aucun risque pour "Statut commande" ni pour le Dashboard qui en dépend.
Reprend la tolérance déjà tranchée par l'acheteur (Volet 3 : aucune
tolérance, Écart facture doit valoir exactement 0) :
```
=IF(N(Commandes[[#This Row],[Qté livrée]])=0,"",
IF(Commandes[[#This Row],[N° facture]]="","🧾 En attente facture",
IF(N(Commandes[[#This Row],[Écart facture]])=0,"✅ Facturée","⚠️ Écart facture")))
```
(vide tant que rien n'est livré ; "🧾 En attente facture" si livré mais
sans N° facture ; "✅ Facturée" si écart nul ; "⚠️ Écart facture" sinon —
libellés/émoji à valider par l'acheteur, choisis pour rester visuellement
distincts des libellés déjà utilisés par "Statut commande", notamment
"🚚 Reçue (Attente Facture)" qui ne doit pas être confondu avec "🧾 En
attente facture"). Formule courte, ne nécessite PAS `LET()` — la question
de la disponibilité de `LET()` sur "Excel 2026 en local" (poste de
l'acheteur, version non confirmée) devient donc sans objet pour cette
colonne. À coller dans le tableau structuré une fois la colonne créée,
testée d'abord sur une ligne avant application à toute la colonne (même
prudence que prévu, moins critique désormais puisque "Statut commande"
n'est plus touché).

**Si l'acheteur souhaite EN PLUS voir la facturation dans le Dashboard** :
évolution séparée de la feuille `Dashboard` (nouveaux indicateurs à partir
de "Statut facture"/"Écart facture"), jamais une modification de "Statut
commande" — hors périmètre de cette session, à cadrer plus tard si
demandé.

**Anomalie repérée en passant, sans rapport avec ce cadrage** : la ligne 2
du tableau (première ligne de données, commande réelle "i26.001" COMINTER,
"Immobilisations") a des formules DIFFÉRENTES et visiblement incohérentes
pour les colonnes Reliquat/RAL/Soldé/Reste à facturer (chacune semble
référencer les mauvaises colonnes, comme si la formule de "Statut
commande" avait été copiée-collée avec un décalage) — alors que TOUTES les
autres lignes (vérifié lignes 3 et 4) ont des formules simples et
cohérentes entre elles. Cette ligne 2 n'a pas encore de livraison (Qté
livrée vide), donc l'impact visible est probablement nul aujourd'hui.
Signalé à l'acheteur, qui répond "je regarderai" — hors périmètre de
cette session, pas touché.

**Une fois les colonnes créées par l'acheteur** (prochaine session) :
vérifier cellule par cellule sur une COPIE fraîche du vivant qu'elles sont
bien vides de formule (sauf Montant facturé HT / Écart facture),
étendre `COLONNES_MODIFIABLES` dans `moteur/rapprochement/ecriture.py`
à `("Date de livraison", "Qté livrée", "Tarif BL", "Note", "N° facture",
"Date facture", "Qté facturée", "PU facturé")`, et étendre
`tests/test_rapprochement_ecriture.py` (patch réel sur une copie,
relecture, vérification qu'aucune autre partie du zip n'a changé — même
schéma que le test déjà existant sur les 4 colonnes actuelles). **Noter la
date à partir de laquelle ces colonnes existent** : toute ligne dont le
statut est "🚚 Reçue (Attente Facture)" AVANT cette date n'a simplement
jamais eu l'occasion d'être facturée dans l'outil — pas un signal
d'anomalie rétroactif.

### Volet 2 — Cadrage du flux factures, 109 Distribution

**Pas besoin de redemander des pièces** : l'acheteur a déjà déposé
**79 vraies factures 109 Distribution** dans `a_traiter/Factures/`
(`Facture_360310.pdf` … `Facture_365699.pdf`, période du 15/07 au
17/08/2026) — largement de quoi cadrer le format sans rien inventer.

**Format** : PDF **texte natif** (jamais de scan, confirmé sur les 79 —
aucun OCR nécessaire pour ce fournisseur, contrairement à ses BL). 3 pages
la plupart du temps ; en pratique seule la page 0 porte du contenu, SAUF
que **58 factures sur 79 (73 %)** ont EN PLUS, sur une page suivante, une
copie de NOTRE PROPRE bon de commande (reconnaissable : "BON DE COMMANDE",
"Electricité Services Réunion", "William AIMAR", "achats@espace-soleil.re"
— c'est le PDF généré par notre propre flux, pas un document de 109) —
**expliqué a posteriori (session F4/Stand64+EDOI, voir plus bas) : exigence
du service comptable de l'acheteur, pas un ajout de 109 Distribution
elle-même** — une facture doit être transmise accompagnée de son BL et de
son bon de commande, d'où ce regroupement systématique.

**Champs présents, tous en texte natif PyMuPDF** (l'ordre de lecture
`get_text()` NE SUIT PAS l'ordre visuel — même défi que le gabarit DEVIS de
ce même fournisseur, `moteur/fournisseurs/dist109.py`, un futur parser
facture devra probablement s'appuyer sur `_gabarit.scan_ancre`/
`scan_regex` plutôt que sur le flux brut) :
- **N° facture** : "Facture client n° 360 310 du 15/07/2026" (numéro avec
  espace interne à nettoyer, + date d'émission).
- **Date d'échéance** : "Date d'échéance :29/08/2026" — toujours 45 jours
  net après la date facture sur les 79 pièces ("Clt - 45 jours").
- **N°Réf.Client** (notre référence transmise à 109) : présente sur les 79,
  mais SEULEMENT **68/79 (86 %)** au format exact du Suivi ("123.096" ou
  "M3.14.339"). Le reste (**11/79, 14 %**) n'est PAS directement
  exploitable : 4 au format interne 109 ("BC 241659", 6 chiffres, aucun
  rapport avec notre numérotation), 7 en texte libre/dégradé ("bc 241660"
  minuscule, "cmd n° 240610", une adresse mail collée à un fragment de
  commande, ou une concaténation illisible de PLUSIEURS commandes —
  "ORANE; 132.008; 132.00" tronqué, cas réel détaillé plus bas).
- **Bon de livraison n°XXXXXX du DATE traité par NOM** : cité au moins une
  fois par facture, JUSTE avant le bloc de lignes qu'il couvre. **9/79
  (11 %) factures citent PLUSIEURS BL**, chacun avec son propre bloc de
  lignes ET son propre "Sous total" — confirmé sur 2 cas réels examinés en
  détail :
  - `Facture_360366.pdf` (129.034, 3 BL) : les 3 BL couvrent la MÊME
    commande (livraisons successives du 07 et 08/07/2026) — un vrai relevé
    mono-commande, multi-BL.
  - `Facture_365533.pdf` (réf. client "ORANE; 132.008; 132.00…", 3 BL) :
    un BL du 28/04/2026 et deux BL du 11/08/2026 — dates trop éloignées et
    mentions de chantier différentes ("EUGENIA" apparaissant près de 2 des
    3 sous-totaux, "ATTENTE DATE LIVRAISON") pour être la même commande :
    **confirme qu'une facture peut couvrir plusieurs commandes
    différentes**, pas seulement plusieurs BL d'une même commande.
  **Conséquence pour un futur parser** : rapprocher facture → commande
  PAR BLOC DE BL (en utilisant le n° de BL, déjà connu si ce BL a été
  rapproché précédemment côté `moteur/rapprochement/`) est plus robuste
  que de se fier au seul "N°Réf.Client" globale de la facture, qui peut ne
  s'appliquer qu'à UNE partie des lignes sur une facture multi-commandes.
- **Lignes d'article** : Référence / Description / Qté / P.U.Net / Eco-part
  (souvent vide, mais RENSEIGNÉE et non-nulle sur au moins 1 facture
  réelle, `Facture_365533.pdf`, 4,41€) / Rem% / Code TVA / Total — dans cet
  ordre visuel de colonnes, mais le flux `get_text()` les restitue dans le
  désordre (Total, Description, Code TVA, P.U.Net, Qté, Référence).
  Immense majorité des lignes en code TVA "C1" (exonération TVA — chantiers
  défiscalisés), mais du "C0" et du vrai "C4" (8,5 %) vus aussi sur les 79
  pièces — donc TVA à gérer réellement, pas juste 0 partout.
- **Totaux** : "Sous total" par bloc de BL, tableau de ventilation TVA
  (Code / Taux / Base HT / Montant TVA), puis Total HT / TVA / Net à payer
  + Total Eco-part HT / Total TVA Eco-part en pied de page.
- **Aucun "avoir" parmi les 79 pièces déposées** (recherché "AVOIR" dans le
  texte entier, aucune occurrence) — l'acheteur signale que sur ce
  fournisseur, une facture d'avoir se reconnaît normalement par la mention
  **"FACTURE D'AVOIR"** dans l'en-tête (à la place de "Facture client
  n°...") — noté pour un futur parser, pas encore vu sur pièce réelle.

**Tableau de flux (109 Distribution uniquement — les autres fournisseurs
sont seulement inventoriés ci-dessous, pas cadrés)** :

| Aspect | 109 Distribution |
|---|---|
| Format PDF | Texte natif, jamais de scan (79/79) |
| Nb pages | 3 en général ; souvent (73 %) + copie de notre propre BdC |
| N° commande exploitable directement | 86 % (68/79) ; sinon repli nécessaire par n° de BL |
| Mono ou multi-BL par facture | 89 % mono-BL ; 11 % multi-BL (jusqu'à 4 vus), parfois multi-COMMANDES |
| Prix par ligne | Toujours présent (contrairement au BL de ce même fournisseur qui, lui, affiche aussi les prix — cohérent) |
| TVA | Majoritairement C1 (exonéré), mais C0/C4 réels rencontrés |
| Avoir | Aucun exemple réel à ce jour ; repérable par "FACTURE D'AVOIR" en en-tête (indication de l'acheteur) |

**Autres fournisseurs — inventaire, pas de cadrage** : `a_traiter/Factures/`
ne contient AUCUN autre fournisseur à ce jour (79/79 = 109 Distribution).
Les fournisseurs déjà couverts côté BL (Coredime, Cominter Ouest, Electric
Plus/GMR, Ravate, Stand 64, DEM, YESSS, Protecthoms) restent donc à
cadrer plus tard, dès que l'acheteur y déposera de vraies factures — même
règle d'or que partout ailleurs dans ce projet (jamais de gabarit sans
pièce réelle).

### Volet 3 — Décisions d'architecture

**Dossiers** (validé par l'acheteur) :
- `a_traiter/Factures/` existe déjà (79 PDF à plat aujourd'hui).
- `a_traiter/Factures/À vérifier/` à créer — même rôle que son pendant côté
  BL (`a_traiter/BL/À vérifier/`) : toute facture lue mais avec au moins
  une ligne "inconnue"/"à confirmer" non tranchée y est déplacée telle
  quelle, jamais mélangée avec les factures pas encore traitées.
- **Archivage définitif dans `a_traiter/BL/Traités/<n° commande>/`**
  (dossiers DÉJÀ créés par le flux BL, avec copie du BdC dedans — vérifié
  sur `108.271/` : `2026-07-15 - COMINTER - OBL107471 - BC 108.271.pdf` +
  `BC - Lacouture - BC 108.271 - COMINTER.pdf`) — PAS un nouveau dossier
  "Factures" séparé, pour garder le flux commande→BL→facture consultable
  au même endroit, comme demandé par l'acheteur pour l'archivage BL.
- **Facture multi-commandes** (cas réel confirmé, `Facture_365533.pdf`) :
  contrairement à un BL Cominter multi-BL (qu'on peut DÉCOUPER par page
  avec `fitz`), une facture 109 est un PDF à une seule page utile — donc
  pas de découpage possible. Proposition : **copier le PDF entier dans
  CHAQUE dossier de commande concerné**, avec un nom qui reflète la
  commande de CE dossier (donc le même fichier physique porte un nom
  différent selon l'endroit où il atterrit).
- **Nommage proposé** (calé sur celui des BL) :
  `<date facture> - 109 DISTRIBUTION - Facture <n° facture> - BC <n° commande>.pdf`
  (ex. `2026-07-15 - 109 DISTRIBUTION - Facture 360310 - BC 241659.pdf`
  pour un cas où le n° réel n'est pas exploitable).

**Tolérance d'écart facture vs BL/tarif convenu — décision de l'acheteur :
AUCUNE tolérance.** Proposition initiale (2 % + plancher 0,50 €) refusée :
"il y a énormément d'articles à très faible valeur, pas de tolérance,
prix BL = prix facture". Le PU facturé doit être EXACTEMENT égal au tarif
déjà utilisé pour "Facturé BL" (Tarif BL, ou Tarif convenu en repli) —
Écart facture doit valoir 0, sans quoi la ligne reste "à confirmer"/statut
d'alerte, jamais "✅ Soldée (Facturée)" automatiquement. Cohérent avec le
volume réel d'articles à faible valeur unitaire observé dans le Suivi (où
un pourcentage ou un plancher masquerait facilement une vraie erreur de
prix sur une petite ligne).

**Facture arrivant AVANT que son BL ait été rapproché** (Qté livrée
vide) : proposition — **"à confirmer" par défaut**, sauf pour les
fournisseurs déjà listés comme "facture = BL" dans CLAUDE.md. Aujourd'hui,
un SEUL fournisseur est dans ce cas : **Electric Plus/GMR** ("GMR n'envoie
PAS de bon de livraison séparé... le rapprochement se fait directement à
partir de ses FACTURES", déjà documenté plus haut). 109 Distribution N'EN
FAIT PAS PARTIE — il envoie des BL séparés (confirmé cette session), donc
une facture 109 arrivée avant son BL doit toujours attendre confirmation,
jamais valoir livraison automatique.

**Avoirs** : jamais automatiques, toujours "à confirmer", rattachés à la
facture ou au BC d'origine cités dessus — cohérent avec le fait qu'aucun
avoir réel n'a encore été vu (règle d'or : rien à coder tant qu'aucune
pièce réelle n'est disponible). Repère à surveiller signalé par l'acheteur :
mention **"FACTURE D'AVOIR"** en en-tête.

## Rapprochement factures — session F2 (implémentation, 109 Distribution)

Suite directe de F1 ci-dessus. Livre le rapprochement des factures 109
Distribution de bout en bout, en réutilisant au maximum le flux BL déjà
éprouvé (mêmes principes : lecture seule d'abord, écriture confirmée
ensuite, référentiel articles partagé, garde-fous de conflit/idempotence).

**DÉCOUVERTE BLOQUANTE, à traiter EN PREMIER avant toute écriture réelle** :
la consigne de cette session affirmait que les colonnes facture "existent
dans le Suivi et sont dans COLONNES_MODIFIABLES" — **FAUX, vérifié
directement sur le vrai classeur vivant en tout début de session** :
`moteur/rapprochement/ecriture.py::COLONNES_MODIFIABLES` ne contenait
encore QUE les 4 colonnes BL, et les en-têtes réels du Suivi (`1.3.0.1.
Suivi commandes - 2026.xlsx`, feuille "Commandes", lues colonne par
colonne) ne comportent AUCUNE des 6 colonnes proposées en Volet 1 (N°
facture / Date facture / Qté facturée / PU facturé / Montant facturé HT /
Écart facture) — Volet 1 restait une PROPOSITION validée dans son
principe, jamais appliquée dans Excel. **Conséquence concrète pour cette
session** : le pipeline de LECTURE/RAPPROCHEMENT a été construit et validé
de bout en bout sur les 79 vraies factures + le vrai Suivi (voir plus bas),
mais AUCUNE écriture réelle n'a pu être démontrée sur le classeur vivant —
le mécanisme d'écriture est prouvé sur un classeur SYNTHÉTIQUE qui, lui, a
les colonnes (voir "Tests"). `COLONNES_MODIFIABLES` a quand même été étendu
dès maintenant (les 4 colonnes de saisie facture) : `ecriture.appliquer()`
refuse déjà proprement, avec un message clair
("colonne introuvable dans les en-têtes"), tant que l'acheteur n'a pas créé
ces colonnes dans Excel — rien à changer côté code une fois qu'elle
l'aura fait. **Prochaine étape avant toute recette réelle d'écriture** :
demander à l'acheteur de créer les 4 colonnes de saisie (+ les 2 colonnes
formule, voir Volet 1) dans le tableau structuré "Commandes" de son
classeur.

### Architecture (réutilise BL, ne duplique pas)

- `moteur/rapprochement/modele_facture.py` — `Facture`/`LigneFacture`,
  même esprit que `modele_bl.py`. `LigneFacture.numero_commande` reste vide
  au parsing (le parser n'a pas accès au Suivi) — renseigné par
  `matching_facture`/`pipeline_facture`.
- `moteur/fournisseurs/dist109.py::parse_facture_109` — section "GABARIT
  FACTURE" ajoutée à la suite du devis/BL existants. **Facture en texte PDF
  NATIF, jamais de scan chez ce fournisseur** (confirmé sur les 79 vraies
  pièces, contrairement à ses BL) — pas d'OCR nécessaire, `moteur.lecture_pdf.
  lire_pdf` suffit. Bloc ancré sur le code TVA (C0 à C9, même famille de
  motif que le devis/BL du même fournisseur), mais dans un ORDRE DE CHAMPS
  différent : Total(-2), Description(-1), TVA(0, ancre), P.U.Net(+1),
  Qté(+2), Référence(+3) — l'Eco-part s'intercale entre Description et TVA
  UNIQUEMENT quand elle est non nulle (décale tout d'un cran, détecté en
  regardant si la cellule juste avant le TVA ressemble à un prix).
  **Validé sur les 79 vraies factures du lot de cadrage F1 : 79/79 exactes
  sur le Total HT affiché** (aucun écart, y compris la seule facture à
  Eco-part non nul).
- **BUG RÉEL ÉVITÉ avant tout code livré** (trouvé en prototypant, jamais
  passé en production) : un premier essai scannait le texte ENTIER de la
  facture pour les codes TVA, sans borner — le pied de page (tableau de
  ventilation TVA, "Code/Taux/Base HT/Montant TVA") contient LUI AUSSI des
  codes TVA isolés qui, sur la SEULE facture à 2 taux de TVA différents du
  lot (`Facture_365533.pdf`), formaient un faux article ("référence 8,50,
  qté 718,41") — resté invisible sur l'autocontrôle Total HT SEULEMENT
  parce que son "Total" valait 0,00 par pure coïncidence sur cette pièce
  précise (pas une garantie générale sur un futur PDF). Corrigé
  DÉFINITIVEMENT par une zone bornée PAR BLOC DE BL (entre le marqueur
  "Bon de livraison n°X du date" et le "Sous total" qui le suit) — élimine
  le risque structurellement ET donne en prime l'association ligne -> n°
  de BL sans heuristique de proximité. Fixtures réelles ajoutées
  (`tests/fixtures/facture_dist109_1..4*.pdf`), testé
  (`tests/test_parsers_facture_dist109.py`, 4 tests, dont un qui verrouille
  explicitement l'absence de cette ligne fantôme).
- `moteur/rapprochement/parsers_facture.py` / `lecture_facture.py` — même
  principe que leurs équivalents BL (registre auto-découvert, tolérance
  aux pannes). Pas de détection PAR PAGE (contrairement à `lecture_bl.py`) :
  aucune facture réelle ne mélange plusieurs fournisseurs à ce jour — à
  ajouter sur le même modèle si un cas réel se présente.
- `moteur/rapprochement/matching_facture.py` — **réutilise TEL QUEL** les
  primitives de correspondance de référence déjà éprouvées côté BL
  (`matching._memes_references`/`_repli_reference_proche`/
  `_repli_referentiel`, importées directement — fonctionnent par duck
  typing, `LigneFacture` partage les champs `reference_fournisseur`/
  `designation` de `LigneBL`) plutôt que d'en dupliquer une 2e version.
  Seule la SÉMANTIQUE DE COMPARAISON change (`_comparer_facture`) : une
  facture n'est jamais "cumulée" comme une livraison — elle est confrontée
  à ce qui est DÉJÀ enregistré (Qté livrée, Tarif BL sinon Tarif convenu,
  et le N° de facture déjà présent le cas échéant pour l'idempotence).
  **Aucune tolérance sur l'écart de prix** (décision explicite de
  l'acheteur, voir Volet 3 : "prix BL = prix facture"). `LigneSuiviFacture`
  lit aussi les 4 colonnes facture SI PRÉSENTES (`colonnes_facture_
  disponibles()`) — sinon elles restent `None` sur chaque ligne, ce qui
  laisse le diagnostic en LECTURE SEULE utilisable même sans elles (voir
  découverte bloquante ci-dessus).
- `moteur/rapprochement/pipeline_facture.py` — orchestration
  (`rapprocher_dossier_factures()` lecture seule, `appliquer_et_archiver_
  factures()` écriture + archivage), réutilise directement de
  `pipeline_bl.py` : `trouver_fichier_suivi_vivant`, `trouver_dossier_
  commandes`/`trouver_bon_de_commande`, `_dossier_pour_commande`,
  `_copier_bon_de_commande_si_absent`, `_sans_caracteres_interdits`,
  `_parser_date_bl` (générique malgré son nom, réutilisable tel quel pour
  parser `Facture.date_facture`), `deplacer_vers_a_verifier`. Référentiel
  articles partagé (même `moteur/articles.db`), mais sa PROPRE file
  d'attente `referentiel/A_confirmer_Facture.xlsx` (ajouté au `.gitignore`,
  comme `A_confirmer_BL.xlsx` — un oubli du `.gitignore` corrigé au
  passage, les deux n'y étaient pas alors qu'`A_confirmer.xlsx`, lui,
  l'était déjà).

### Résolution de commande par bloc de BL (`_resoudre_commandes_facture`)

Une facture peut citer PLUSIEURS BL (jusqu'à 4 vus sur le lot de cadrage),
chacun potentiellement rattaché à une commande différente (1 seul cas réel
confirmé sur 79, `Facture_365533.pdf`). Algorithme, validé contre le vrai
Suivi sur les 9 factures multi-BL du lot de cadrage :

1. **N°Réf.Client fait autorité dès qu'il donne UN candidat au format
   Suivi** (regex réutilisée de `dist109.MOTIF_COMMANDE_BL`) — appliqué à
   TOUS les blocs de BL de la facture (89% des factures multi-BL réelles
   ne couvrent qu'UNE SEULE commande, livrée en plusieurs fois — validé :
   8 des 9 factures multi-BL du lot retombent proprement sur ce cas, 0
   divergence trouvée en cross-validant par déduction de contenu).
2. **En repli** (en-tête vide ou format interne 109 non exploitable, ex.
   "BC 241659", "cmd n° 240610" — 11/79 factures du lot dans ce cas) : la
   commande de CE bloc est déduite de son propre contenu, en réutilisant
   TEL QUEL `matching.deduire_commande_par_contenu()` (même mécanisme que
   pour un n° de commande illisible côté BL — au moins 2 lignes
   concordantes, score sans ambiguïté). Jamais utilisée pour un
   rapprochement "sûr" automatique — toute correspondance obtenue via une
   commande déduite bascule "à confirmer", raison explicite donnée (même
   principe que `pipeline_bl.py`).
3. Si ni l'un ni l'autre : le bloc part dans `rapport.anomalies_facture`
   ("n° de commande introuvable") — bac dédié, entrée naturelle d'une
   future session F4 (aide à la création manuelle de la commande côté
   achat) — jamais un choix au hasard.

Sur le SEUL cas réel multi-commande du lot (`Facture_365533.pdf`, 2
commandes réellement différentes derrière "ORANE; 132.008; 132.00"
tronqué) : l'en-tête ne donne aucun candidat exploitable, et la déduction
par contenu échoue aussi (chaque bloc de BL n'a que 1-2 lignes, sous le
seuil de confiance de 2 lignes concordantes) — les 3 blocs partent
honnêtement en "commande introuvable", pour décision humaine. Comportement
jugé correct : mieux vaut ne rien écrire qu'un mauvais rapprochement sur le
seul cas réellement ambigu du lot.

### Archivage — copie, jamais découpage

Contrairement à un BL Cominter Ouest/109 Distribution multi-BL (qu'on peut
DÉCOUPER par page avec `fitz`, voir `pipeline_bl._extraire_bl_vers`), une
facture 109 est un PDF à une seule page utile — pas de découpage possible.
Une facture multi-commande entièrement résolue est donc **COPIÉE** (jamais
déplacée) dans `a_traiter/BL/Traités/<n° de commande>/` **pour CHAQUE
commande concernée** (`pipeline_facture.archiver_facture`), le fichier
source n'étant supprimé qu'une fois toutes les copies faites avec succès —
même arbre que les BL et leur BdC, pour le flux commande→BL→facture
consultable au même endroit demandé par l'acheteur en session BL. Une
facture avec au moins un bloc non résolu part ENTIÈRE (jamais partiellement
copiée) vers `a_traiter/Factures/À vérifier/`.

### Résultat de la recette en LECTURE SEULE sur les 79 vraies factures +
le vrai Suivi commandes (aucune écriture, colonnes facture absentes)

`147 sûres · 29 à confirmer · 0 déjà à jour (attendu, aucune facture
n'a jamais pu être enregistrée faute des colonnes) · 9 inconnues · 13
blocs non rapprochés (commande introuvable)`. Quelques observations
réelles utiles pour la suite :

- **Faux positif de précision trouvé, à garder en tête (pas corrigé — la
  décision "aucune tolérance" est celle de l'acheteur)** : `Facture_365387.pdf`
  ressort "à confirmer" pour un écart de 0,005€ (PU facturé 39,56€ vs
  Tarif BL stocké 39,555€, 3 décimales — probablement une moyenne calculée
  à l'écriture d'un BL, jamais arrondie à 2 décimales comme une facture
  l'imprime toujours). Un cas parmi 176 lignes rapprochées, isolé — pas
  assez pour remettre en cause "aucune tolérance" sans en reparler à
  l'acheteur, mais le genre de cas qu'elle voudra probablement voir une
  fois les colonnes créées.
- Plusieurs "à confirmer" sont des **facturations partielles réelles**
  (Qté facturée < Qté livrée déjà enregistrée, ex. `Facture_362846.pdf` —
  cohérent avec le fait que Coredime/DEM facturent aussi parfois en
  plusieurs fois, rien d'anormal en soi, juste à vérifier au cas par cas).
- Le repli référence proche (réutilisé de `matching.py`) a immédiatement
  servi sur données réelles (`Facture_363011.pdf` "BTAER4X25+2X1." vs Suivi
  "BTAER4X25+2X1.5", `Facture_365398.pdf` "H07VK16BL" vs Suivi "H07VK16B") —
  aucune règle nouvelle nécessaire, la réutilisation fonctionne comme prévu.
- **Compteur de résorption** (`compter_lignes_a_facturer`, l'indicateur du
  chantier F2) sur le vrai Suivi, 109 DISTRIBUTION : `1000 lignes livrées,
  1000 encore "à facturer"` — 100% par construction aujourd'hui puisque
  aucune ligne n'a jamais pu recevoir de N° facture (colonnes absentes) ;
  redeviendra un vrai indicateur utile dès qu'une 1ère vraie écriture aura
  eu lieu.

### Tests

`tests/test_parsers_facture_dist109.py` (4, sur vraies fixtures — simple/
BC interne/multi-BL même commande/multi-BL multi-commande+eco-part),
`tests/test_lecture_facture.py` (5, tolérance aux pannes), `tests/
test_rapprochement_matching_facture.py` (12, logique de comparaison sur
objets synthétiques), `tests/test_rapprochement_pipeline_facture.py` (15,
résolution de commande/archivage multi-copie/résorption/écriture de bout
en bout), plus 2 tests ajoutés à `tests/test_rapprochement_ecriture.py`
(colonnes facture dans `COLONNES_MODIFIABLES`, écriture réelle sur un
classeur synthétique). **Les tests d'écriture réelle tournent sur un
classeur SYNTHÉTIQUE qui A les colonnes facture** (le vrai Suivi ne les a
pas encore, voir découverte bloquante ci-dessus) — c'est la preuve que le
mécanisme fonctionne, en attendant la vraie recette d'écriture sur le
classeur vivant une fois les colonnes créées par l'acheteur.

### Reste à faire (session F3)

1. ~~Créer les colonnes de saisie dans Excel~~ **FAIT autrement** — voir
   section suivante : plus besoin que l'acheteur crée quoi que ce soit à la
   main dans Excel, ni ne colle de formule ; l'outil a créé les 5 colonnes
   lui-même, directement dans le tableau structuré, par patch XML. Le
   "patch XML chirurgical hors périmètre pour ajouter des colonnes à un
   tableau structuré" évoqué ici s'est révélé faux — voir
   `ajouter_entetes_saisie()` dans `moteur/rapprochement/ecriture.py`.
2. Reparler du cas de précision ci-dessus avec l'acheteur (0,005€ d'écart
   dû à un Tarif BL stocké à 3 décimales) avant de conclure que "aucune
   tolérance" produit trop de faux positifs — un seul exemple à ce jour.
3. Vraie recette d'écriture sur le classeur vivant, comme pour chaque
   fournisseur BL avant lui (dépôt réel, clic "Écrire", vérification ligne
   par ligne) — les colonnes existent maintenant pour de vrai dans le
   vivant (voir section suivante), reste à y faire transiter de vraies
   valeurs de facture.
4. Étendre aux autres fournisseurs dès que de vraies factures seront
   déposées pour eux (`a_traiter/Factures/` ne contient que du 109
   Distribution à ce jour, voir Volet 2).
5. AVOIR : toujours aucun exemple réel — `type_document == "AVOIR"` déjà
   détecté et mis de côté sans y toucher (voir `MOTIF_AVOIR_FACTURE`),
   rien de plus à coder tant qu'aucune pièce réelle n'existe.
6. Décider comment "Montant facturé HT" est rempli (voir section suivante :
   c'est une simple colonne de saisie pour l'instant, pas un calcul) — et
   construire "Écart facture"/"Statut facture", hors périmètre de la
   session qui a créé les colonnes.

## Rapprochement factures — colonnes créées dans le Suivi vivant, refonte
du classeur découverte, "Statut commande" jamais touché (fait)

Suite directe de F2. Objectif de cette session, cadré strictement par
l'acheteur ("un seul plan, tu ne changes pas d'approche en cours de
route") : préparer le terrain pour F3 en créant les 5 colonnes de saisie
facture dans le VRAI classeur vivant — **sans lui demander une seule
manipulation dans Excel**, contrairement au plan F2/Volet 1 (jamais
appliqué) qui prévoyait qu'elle les crée elle-même et y colle des
formules.

- **Le classeur vivant a été refondu depuis la dernière fois qu'il a été
  inspecté en détail** (dernier état documenté : ~5 900 lignes, feuille
  "Stats") : re-vérifié en tout début de session, `trouver_fichier_suivi_
  vivant()` pointe bien sur `1.3.0.1. Suivi commandes - 2026.xlsx` dans
  "1.3.0.1. Commandes courantes" (seul candidat, pas d'ambiguïté avec un
  "Suivi nouveau..." ou une copie manuelle à ce jour) — feuilles
  `Commandes, Dernières commandes, Tarifs Affaire - Chantiers, Listes
  Paramètres, Base Articles, Tarifs Affaire - Maintenance, Dashboard,
  Analyses, Calculs` (9 feuilles ; l'ancien export périmé à la racine du
  dépôt, lui, a encore une feuille "Stats" et pas de Dashboard/Analyses/
  Calculs — ne JAMAIS le confondre avec le vivant, voir
  `trouver_fichier_suivi_vivant` vs `moteur.panier.trouver_fichier_suivi`).
  50 colonnes (A..AX) au départ, table structurée "Commandes" couvrant
  exactement `A1:AX6420` (16 Excel Tables toujours présentes). Le nombre
  de lignes a grossi PENDANT la session (6420 -> 6430, l'acheteur a
  continué de saisir des commandes normalement) — preuve que le classeur
  reste en usage actif tout du long, pas figé pour l'occasion.
- **DÉCOUVERTE, non traitée (hors périmètre strict de cette session,
  notée pour la future session "formules"/"Statut facture")** : la
  formule RÉELLEMENT appliquée cellule par cellule à "Statut commande"
  (ex. U100) est orientée PRIX (compare Tarif BL à Tarif convenu/Tarif
  Affaire + 0,01, ou à Tarif raisonnable ±15% -> "⚠️ Surfacturation" ;
  une fois soldée, "✅ Soldée — Tout OK" si le tarif est cohérent, sinon
  "⬜ Soldée (Sans référence / à vérifier)" si un Tarif BL existe, sinon
  seulement "🚚 Reçue (Attente Facture)") — **différente du
  `calculatedColumnFormula` stocké dans la définition du tableau**
  (`xl/tables/table1.xml`, gabarit utilisé par Excel pour une nouvelle
  ligne), qui porte encore une version plus ancienne ("⚠️ Problème Tarif",
  "✅ Soldée (Facturée)" testé via `Facturé BL]=1` comme si c'était un
  booléen — alors que "Facturé BL" est un MONTANT calculé, pas un
  indicateur "facturé oui/non", incohérence probablement héritée d'avant
  la refonte prix). Concrètement : "🚚 Reçue (Attente Facture)" ne
  signale donc PAS "en attente de facturation" au sens large — seulement
  "livré en totalité mais aucun Tarif BL renseigné du tout", un cas bien
  plus étroit — confirme, s'il en fallait, pourquoi une vraie dimension
  facture séparée (les colonnes ci-dessous, puis un "Statut facture" à
  construire) est nécessaire. **Ni cette dérive de formule ni "Statut
  commande" en général n'ont été touchés** — strictement hors périmètre,
  jamais modifié par ce module (voir bandeau de `ecriture.py`).
- **1er essai (colonnes HORS tableau) construit, testé, montré à
  l'acheteur sur une copie, puis EXPLICITEMENT REJETÉ par elle avant toute
  écriture réelle** : "les 5 colonnes ont été ajoutées mais ne font pas
  partie du tableau, je comprends pas l'intérêt" — puis, après explication,
  correction ferme et détaillée : **"CORRECTION DÉFINITIVE DU PLAN — les 5
  colonnes vont DANS le tableau structuré Commandes, pas à côté. Raison :
  un tri du tableau ne déplace pas les colonnes hors tableau, les données
  facture se retrouveraient sur les mauvaises lignes."** Leçon générale :
  même un plan validé étape par étape doit rester ouvert à une correction
  de l'acheteur en cours de route si elle en voit la nécessité — la
  rigueur du plan initial protège contre les dérives de LA SESSION, pas
  contre un vrai changement de décision de sa part.
- **`ajouter_entetes_saisie(fichier, noms, dossier_backups, feuille=...)`**
  (nouvelle fonction, `moteur/rapprochement/ecriture.py`) : ajoute des
  colonnes DANS le tableau structuré, à sa suite — patch chirurgical de
  DEUX parties du zip seulement (même esprit que `appliquer()`, jamais de
  `openpyxl.save()`, voir bandeau du fichier) :
  1. La feuille (`xl/worksheets/sheetN.xml`) : cellule d'en-tête inlineStr
     par nom, en ligne 1, insérée triée par colonne (réutilise
     `_remplacer_dans_ligne`/`_ecrire_cellule`, déjà éprouvés côté
     `appliquer()`) + `<dimension>` étendue.
  2. La définition du tableau (`xl/tables/tableN.xml`, retrouvée par son
     attribut `name=`, pas besoin de suivre les relations de la feuille) :
     `ref` du `<table>` ET du `<autoFilter>` étendus (même colonne de fin,
     jamais les lignes), `<tableColumns count="...">` incrémenté, un
     `<tableColumn id="…" name="…"/>` par nom ajouté en toute fin (l'ORDRE
     dans le XML = l'ordre affiché par Excel, vérifié sur le vrai
     classeur — l'attribut `id=` n'est PAS séquentiel par position, choisi
     comme max(id existants)+1, +2...). **Rigoureusement AUCUN
     `calculatedColumnFormula`, AUCUNE `totalsRow`** — seulement `id` et
     `name`, comme demandé explicitement par l'acheteur.
  Garde-fous identiques à `appliquer()` : verrou Excel, sauvegarde
  horodatée AVANT écriture, refus (`ValueError`) si un nom existe déjà
  (en-tête de la feuille OU colonne du tableau — deux vérifications
  indépendantes, l'une ne présume jamais de l'autre) ou si une cellule
  cible n'est pas vide (même une cellule stylée sans valeur — relu depuis
  le XML brut, pas seulement via `lire_entetes()`).
- **Validé de bout en bout sur une VRAIE copie du classeur vivant avant
  toute écriture réelle** (`tests/test_rapprochement_ecriture.py`,
  19 tests dont 3 tournent sur le vivant, `skipif` absent du poste) :
  reconstruction INDÉPENDANTE (remplacement de texte brut, pas la même
  logique regex que l'implémentation) du contenu XML attendu du tableau
  après patch, comparaison zip membre par membre (exactement DEUX parties
  changent : la feuille et la table, tout le reste octet pour octet
  identique), formules U100/AV100 relues avant/après et comparées
  strictement égales. Le test sur le vivant lit son état "avant"
  dynamiquement (ref/nombre de colonnes actuels, jamais figés en dur) et
  utilise deux noms de colonne jamais utilisés en pratique — pour rester
  valide indéfiniment, y compris après cette session, quel que soit l'état
  futur du classeur.
- **Écriture réelle faite le 2026-09-01, sur confirmation explicite de
  l'acheteur que le classeur était fermé** (redemandé explicitement juste
  avant d'écrire, même après une 1ère confirmation donnée plus tôt dans la
  session — l'absence d'erreur `ClasseurVerrouille` ne suffit jamais seule
  comme preuve que le fichier est fermé, toujours reconfirmer avec
  l'acheteur au moment précis d'écrire) : les 5 colonnes
  ("N° facture", "Date facture", "Qté facturée", "PU facturé", "Montant
  facturé HT") existent maintenant pour de vrai dans le Suivi commandes
  vivant, colonnes 51 à 55 (AY à BC au moment de l'écriture — décalera
  naturellement si des colonnes sont insérées avant elles un jour). Table
  "Commandes" étendue à `A1:BC6430` (55 colonnes). Sauvegarde horodatée
  créée dans `backups/` avant l'écriture (rotation 30 jours, comme
  toujours). Relecture du vivant après coup : en-têtes aux bonnes
  colonnes, formules U100/AV100 identiques, aucun autre membre du zip
  modifié. **Acheteur elle-même a ouvert deux copies de test dans Excel
  avant l'écriture réelle** (1ère version hors tableau, puis 2e version
  dans le tableau après correction) : aucun message de réparation à
  chaque fois, style du tableau + flèches de filtre présents sur la 2e
  version, Dashboard non cassé.
- **`COLONNES_MODIFIABLES` étendu à 9 entrées** (4 colonnes BL + 5
  colonnes facture, "Montant facturé HT" comprise) — nouveau
  `ENTETES_FACTURE` (tuple des 5 noms dans l'ordre exact où ils ont été
  créés) exporté par `ecriture.py` pour que F3 les réutilise tel quel,
  sans redéfinir la liste ailleurs.
- **"Montant facturé HT" est UNE SIMPLE COLONNE DE SAISIE pour
  l'instant, pas un calcul** — contrairement à la proposition F2/Volet 1
  (où elle devait être une formule Excel `Qté facturée × PU facturé`) :
  aucune formule n'a été ajoutée nulle part (ni dans le tableau, ni dans
  les cellules) pour aucune des 5 colonnes. Comment elle sera
  effectivement remplie (valeur calculée et écrite directement par le
  futur moteur de rapprochement factures, ou vraie formule Excel ajoutée
  plus tard par le même mécanisme de patch XML) reste une décision NON
  prise, à trancher dans la session dédiée aux formules — avec "Écart
  facture" et "Statut facture" (nouvelle colonne calculée, indépendante de
  "Statut commande", toujours juste proposée dans le cadrage F1, jamais
  construite). **"Statut commande" n'a pas été et ne sera jamais modifié
  par ce module** — la formule ni le `calculatedColumnFormula` du tableau
  n'ont été touchés, y compris la dérive entre les deux découverte cette
  session (voir plus haut).

## Rapprochement factures — 1ère écriture réelle (F2 suite, 2026-09-01)

Suite directe de la session "colonnes créées dans le Suivi vivant" —
`pipeline_facture.py`/`matching_facture.py` adaptés aux 5 colonnes réelles,
puis 1ère vraie recette d'écriture (109 Distribution) sur le classeur
vivant.

- **`ecritures_pour_facture()` réutilise `ecriture.ENTETES_FACTURE`**
  (jamais une 2e liste de noms redéfinie) et écrit désormais les 5 colonnes,
  "Montant facturé HT" compris — colonne de SAISIE, pas une formule Excel
  (voir session précédente). Reprend le montant IMPRIMÉ sur la facture
  (`LigneFacture.montant_ht`, toujours renseigné chez 109 Distribution) ;
  un futur fournisseur qui n'imprimerait pas ce montant ligne à ligne le
  ferait recalculer (Qté × PU) avec un signalement explicite dans
  `resume["montants_recalcules"]` et le rapport écrit — **jamais
  silencieusement**, aucun cas réel à ce jour (109 Distribution imprime
  toujours son montant de ligne).
- **Correctif : un fournisseur reconnu mais sans parser facture reste EN
  PLACE dans `a_traiter/Factures/`**, jamais déplacé vers `À vérifier/`
  (`_est_anomalie_sans_parser`, `resume["factures_sans_parser"]`) — ce
  n'est pas une décision humaine en attente, juste un fournisseur pas
  encore couvert. Sans objet sur ce lot (les 79 fichiers du dossier sont
  tous 109 Distribution, vérifié via le détecteur), mais nécessaire pour
  la suite (Coredime, Cominter... dès que de vraies factures seront
  déposées pour eux).
- **BUG D'EXÉCUTION (pas un bug du moteur) — script `py -3 <fichier>.py`
  depuis un chemin hors du dépôt ne trouve pas le paquet `moteur`** :
  contrairement à `py -3 -c "..."` (qui ajoute le répertoire courant à
  `sys.path`), `py -3 <script>` ajoute le répertoire DU SCRIPT — un script
  d'exécution ponctuel dans le scratchpad de session doit donc faire
  `sys.path.insert(0, str(PROJET))` explicitement. Repéré avant toute
  écriture (échec à l'import), rien perdu — mais à retenir pour toute
  future session qui écrirait un script d'exécution hors du dépôt plutôt
  que d'utiliser `py -3 -c`.
- **LEÇON RÉELLE — le Suivi peut changer ENTRE la simulation (lecture
  seule) montrée à l'acheteur et l'écriture confirmée, même à quelques
  minutes d'intervalle, si elle continue à y travailler.** Constaté en
  conditions réelles cette session : la simulation montrée pour
  confirmation donnait 146 sûres ; une 1ère relecture juste avant
  d'écrire en donnait 147 (`Facture_365379.pdf`, réf. 59210, commande
  M3.14.350, passée d'inconnue à sûre — la ligne Suivi avait dû être
  éditée entre-temps) ; puis, l'exécution réelle (après un aller-retour de
  diagnostic) l'a retrouvée à nouveau à 146, la ligne 59210 étant repassée
  "à vérifier" (`Aucune ligne du Suivi ne correspond...`) — signe que
  l'acheteur retouchait activement cette ligne pendant la session (essai/
  annulation probable). **Le nombre réellement écrit (146) correspond
  exactement à ce qui avait été confirmé** — coïncidence de calendrier,
  pas une garantie générale. Leçon retenue, à appliquer systématiquement :
  ne JAMAIS réutiliser un rapport de simulation périmé pour construire les
  `Ecriture` à écrire — toujours relire fraîchement (`rapprocher_dossier_
  factures()`) juste avant `appliquer_et_archiver_factures()`, et
  comparer les comptes (à_confirmer, anomalies_facture inchangés ; sûres
  jamais en dessous du nombre confirmé) avant d'écrire quoi que ce soit —
  exactement le garde-fou déjà appliqué cette session (voir script de
  recette, conservé hors dépôt).
- **LIMITE CONFIRMÉE SUR DONNÉES RÉELLES (déjà repérée en test la session
  précédente, désormais vérifiée sur la vraie écriture) : "Date facture"
  se relit comme un NOMBRE DE SÉRIE Excel brut (ex. 46218.0), pas comme
  une date, tant que la colonne n'a pas encore de format de date
  appliqué** — une cellule fraîchement créée (colonne née le 2026-09-01,
  jamais écrite avant sur AUCUNE ligne) n'hérite d'aucun style, donc
  aucun format d'affichage. La VALEUR écrite est correcte (vérifié :
  46218 = 15/07/2026, cohérent avec la date de la facture 360311) — seul
  l'AFFICHAGE dans Excel restera "46218" tant que l'acheteur n'aura pas
  appliqué un format de date à cette colonne (une seule fois, dans Excel —
  hors périmètre de l'outil, qui n'a pas vocation à poser des formats de
  cellule). À signaler à l'acheteur à la prochaine occasion.
- **Recette réelle du 2026-09-01 : 146 lignes écrites** (exactement le
  nombre confirmé par l'acheteur après la simulation), sauvegarde
  horodatée créée avant écriture (`backups/1.3.0.1. Suivi commandes -
  2026_20260901_153637.xlsx`). **46 factures entièrement résolues**
  archivées individuellement dans `a_traiter/BL/Traités/<n° commande>/`
  (aux côtés du BdC et des BL de la même commande), **33 factures avec au
  moins une ligne non résolue** déplacées vers `a_traiter/Factures/À
  vérifier/` (fichier entier, jamais découpé — une facture 109 Distribution
  tient sur une seule page utile). 0 échec d'archivage, 0 montant
  recalculé (tous les montants de ligne de ce lot étaient imprimés).
  **Compteur de résorption après écriture, 109 DISTRIBUTION : 854 lignes
  livrées encore sans facture sur 1 000 livrées au total (146 désormais
  facturées)** — premier vrai chiffre non trivial pour cet indicateur,
  jusque-là toujours 100% "à facturer" faute d'écriture. Rapport détaillé :
  `rapports/rapprochement_facture_20260901_154355.txt`.
- **Pattern réel à approfondir en F4 — 13 blocs "commande introuvable" sur
  9 factures, TOUS avec N°Réf.Client LITTÉRALEMENT VIDE** (pas un format
  interne 109 non reconnu comme "BC 241659" — ici rien du tout n'est
  imprimé). Différent des cas déjà documentés (code interne 109, texte
  libre partiellement exploitable) : une absence totale et cohérente sur
  ces 9 factures précises pourrait signaler des commandes passées sans
  transmettre notre référence à 109 (retrait comptoir, commande
  téléphonique, carnet manuel d'un chargé de travaux — même famille que
  les "carnets manuels des gars" déjà documentés côté BL, voir plus haut)
  — **hypothèse non confirmée avec l'acheteur à ce jour**, juste une
  observation factuelle sur les 79 pièces de ce lot (règle d'or : rien
  inventé au-delà de ce qui est réellement observé). À vérifier avec elle
  avant de construire quoi que ce soit dessus.
- **Ouvert pour la suite** : le cas de précision 365387/260630 (écart
  0,5 centime, Tarif BL à 3 décimales) réapparaît identique à cette
  recette — toujours pas retranché avec l'acheteur si "aucune tolérance"
  doit rester strict à ce niveau. "Écart facture"/"Statut facture" restent
  à construire (session dédiée aux formules). Étendre à d'autres
  fournisseurs dès que de vraies factures seront déposées pour eux.

## Rapprochement factures — F4 : Coredime, 2e fournisseur (fait, 2026-09-01)

Suite directe de F2 (109 Distribution). Cadrage donné par l'acheteur avant
tout code : attaquer les fournisseurs dans l'ordre du poids réel du
backlog (Coredime d'abord, puis Cominter), ne jamais mélanger le flux BL
et Facture pour Electric Plus/GMR (déjà couvert côté BL), rester prudent
sur les scans Ravate (toujours "à confirmer"), et une exception ENCADRÉE
proposée pour Tarif BL depuis la facture — validée en une phrase ("ok
100% !") : whitelist explicite (Coredime pour commencer), uniquement si
Tarif BL est vide, uniquement pour une ligne réellement écrite, tracée
dans le rapport (`FOURNISSEURS_TARIF_BL_DEPUIS_FACTURE`,
`moteur/rapprochement/pipeline_facture.py`).

**Pièces réelles** : 70 vraies factures Coredime de juillet-août 2026,
récupérées depuis un mail de Prisca LEBLÉ (comptable) contenant un .msg
imbriqué par facture — extraites une par une via `win32com.client`
(`OpenSharedItem`/`Attachments`/`SaveAsFile`, `sous_item.Close(0)` dans un
`finally` pour libérer le handle avant `unlink()` — sinon `PermissionError`
sur une extraction relancée).

**Parser facture Coredime** (`moteur/fournisseurs/coredime.py`, section
GABARIT FACTURE, 8 fixtures réelles, `tests/test_parsers_facture_coredime.py`) :
- **Bornage de page = le VRAI repère structurel, jamais un repère de
  contenu.** Un premier essai bornait chaque bloc "BON D'EXPEDITION" sur
  la mention "----- IMPORTANT -----" — texte qui peut apparaître n'importe
  où selon le document (pas une vraie limite de page) — perdant des lignes
  imprimées APRÈS ce repère dans le flux PyMuPDF scramblé. Corrigé en
  bornant sur le vrai repère de page, un bloc de métadonnées répété
  `##ESIGUID;...` — présent sur CHAQUE début de page, fiable sur les 70
  pièces.
- **Référence avec tiret** (`WAG221-425`) et **référence purement
  numérique** (`227060133`, un vrai article, à distinguer de `9993`
  "LIVRAISON AVION" — un frais de port qui ressort simplement "inconnu" au
  rapprochement, jamais un mauvais rattachement) : classe de caractères de
  `MOTIF_LIGNE_FACTURE_COREDIME` élargie aux deux formats.
- **Remise double ("35,00+26,00%") imprimée sur une ligne totalement
  disjointe** de sa référence/désignation/quantité dans le flux PyMuPDF —
  `_lignes_remise_double_coredime()` : appariement 1:1 UNIQUEMENT quand il
  n'y a qu'UNE ligne incomplète et UNE ligne "Remise" dans le bloc (aucune
  ambiguïté) ; dès que plusieurs lignes de chaque sorte coexistent dans le
  même bloc, l'appariement devient ambigu et ces lignes restent NON
  extraites — extraction partielle honnête, l'écart avec le Total HT
  affiché le signale (`test_parse_facture_coredime_8...`, 12 lignes sur 19
  réelles, jamais deviné).
- **AVOIR détecté via la ligne de métadonnées** (`...;Avoir;...` dans le
  bloc `##ESIGUID`/`#####DEMAT-FJ`) — `type_document="AVOIR"`, retourne
  immédiatement SANS extraire aucune ligne (jamais rapproché
  automatiquement, format numérique de toute façon différent).
- **Total HT recalculé et comparé à chaque parse** (`_total_ht_facture_coredime`,
  ancré sur la ligne nue `COR [FA]<num>`) — tout écart imprime un `!!` en
  clair, jamais silencieux (5 des 70 pièces réelles ont un écart résiduel
  connu, dû à ce garde-fou remise-double volontairement conservateur).

**Deux changements de règle sur `_comparer_facture()`
(`moteur/rapprochement/matching_facture.py`), tous deux des décisions
explicites de l'acheteur données EN COURS de session, chacune corrigeant
la précédente — leçon générale : un plan déjà validé reste ouvert à une
vraie décision de l'acheteur, y compris quelques secondes après l'avoir
donnée :**

1. **Une facture reçue AVANT que son BL soit rapproché (Qté livrée encore
   à 0) n'est plus un motif de blocage.** Mot de l'acheteur : *"il faut
   l'écrire ! Ce ne sont pas des factures non parvenues puisqu'on les a
   reçues !"*, puis, en creusant pourquoi tant de lignes Coredime
   traînaient en "à confirmer" : *"en plus, il y a les BL manquants
   là-dedans ! Ils sont signés"* — la livraison a réellement eu lieu (BL
   papier signé), seul son rapprochement dans l'outil n'a pas encore eu
   lieu. Le contrôle de cohérence de quantité (facturée vs livrée) est
   simplement IGNORÉ dans ce cas (rien à comparer, pas un signe
   d'anomalie) — condition `ligne_suivi.qte_livree > 0 and abs(...) > 0.001`
   au lieu d'un retour anticipé bloquant.
2. **Le PRIX facturé n'est PLUS DU TOUT un motif de blocage non plus** —
   décision affinée en une seule respiration dans le chat : d'abord une
   tolérance de 0,01€ envisagée (pour absorber le bruit d'arrondi type
   `GEWDX40020`/`GEWDX27720`, Tarif BL/convenu stocké à 3-4 décimales,
   jamais arrondi comme une facture l'imprime toujours à 2), PUIS,
   quelques secondes plus tard : *"pardon changeons ça, 0,01€ de
   tolérance"* suivi immédiatement de *"en fait il faut écrire tout ce qui
   apparaît sur les factures rapprochables à des commandes, quel que soit
   le prix"* — la tolérance à 0,01€ a été codée puis testée, puis
   entièrement retirée avant même d'être exploitée en écriture réelle.
   Toute la section prix de `_comparer_facture()` (comparaison PU
   facturé/Tarif BL/Tarif convenu) a été supprimée : le PU facturé est
   désormais écrit tel quel, l'écart avec Tarif BL/Tarif convenu reste
   visible directement dans le Suivi (colonnes côte à côte) sans plus
   jamais bloquer l'écriture. Revient entièrement sur le "aucune
   tolérance" du cadrage F1 (Volet 3). **Seul ce qui concerne l'IDENTITÉ
   du rapprochement reste un motif de blocage** : un autre n° de facture
   déjà présent sur la ligne (doublon/litige), une quantité facturée
   incohérente avec une quantité RÉELLEMENT déjà livrée, une commande
   déduite (jamais "sûr" automatiquement, inchangé), ou une référence
   ambiguë/introuvable (inchangé, matching de référence non touché).
   Testé de bout en bout (`tests/test_rapprochement_matching_facture.py`,
   tous les tests de prix réécrits pour refléter "SUR quel que soit le
   prix" plutôt que "à confirmer si écart").
   **Portée** : ce changement s'applique à TOUS les fournisseurs (pas
   propre à Coredime, contrairement à l'exception Tarif BL) — a
   immédiatement débloqué des lignes 109 Distribution qui traînaient dans
   `À vérifier/` depuis la session F2 (ex. `F2R160T`, `59210`), sans
   action spécifique nécessaire.

**BUG RÉEL TROUVÉ ET CORRIGÉ EN CONDITIONS RÉELLES — même piège que côté
BL, mais jamais rencontré côté Facture avant cette session** :
`appliquer_et_archiver_factures()` suppose que `dossier_a_traiter` est
TOUJOURS `a_traiter/Factures/` lui-même, jamais un de ses sous-dossiers.
Pour retraiter directement les 84 factures déjà déplacées dans
`a_traiter/Factures/À vérifier/` (après le correctif "prix jamais
bloquant"), un appel a pointé `dossier_a_traiter` DIRECTEMENT sur ce
sous-dossier — créant un `À vérifier/À vérifier/` imbriqué pour les
factures encore non résolues au lieu de les y laisser à plat. **Aucune
perte de données** (repéré en quelques minutes via 3 vérifications
PowerShell indépendantes — Get-ChildItem, Test-Path, contenu détaillé —
après que deux lectures consécutives, Python puis PowerShell, aient
d'abord semblé montrer un dossier VIDE, un faux symptôme dû au délai de
cohérence du partage réseau juste après l'opération de masse) : les 78
fichiers ont été remontés à la main (`Move-Item`) dans
`a_traiter/Factures/À vérifier/`, le dossier imbriqué vide supprimé, puis
RE-remontés une 2e fois jusqu'à `a_traiter/Factures/` (racine) pour
reproduire le pattern normal d'usage (l'acheteur redépose dans le dossier
principal pour un nouveau passage, jamais directement dans "À vérifier"),
avant de relancer proprement. Avertissement ajouté dans le docstring de
`appliquer_et_archiver_factures()`, symétrique à celui déjà en place côté
`pipeline_bl.appliquer_et_archiver` — **leçon générale pour toute session
future qui voudrait retraiter un lot déjà dans "À vérifier"** : toujours
remonter les fichiers dans le dossier principal (`a_traiter/BL/` ou
`a_traiter/Factures/`) avant un appel à la fonction d'écriture,
`rapprocher_dossier(_factures)()` (lecture seule) pouvant lui, seul, être
pointé sans risque directement sur "À vérifier" pour diagnostic.

**Clarification de l'acheteur sur le paquet de commandes "24X.XXX"/"242.5XX"
introuvables dans le Suivi** (repéré comme un motif net : 18 des 20 blocs
"commande introuvable" du 1er passage de recette concentrés dans cette
même plage numérique) : *"les factures sous le format '24XXXX' (X =
chiffre) sont des BdC manuels, F4 s'en occupera"* — catégorie déjà connue
(carnets manuels des "gars", voir sessions BL précédentes), différée à une
prochaine étape de cette même branche F4, pas une anomalie de
rapprochement à corriger maintenant. Le seul fichier "fournisseur non
reconnu" du lot (`BC241723.pdf`) colle à cette même famille de numéros —
vraisemblablement le BdC papier scanné d'une de ces commandes manuelles,
pas une facture, jamais forcé dans le pipeline.

**Recette réelle — 3 écritures successives sur le vrai Suivi vivant, même
soir** (chaque écriture précédée d'une relecture fraîche du Suivi et d'une
reconfirmation explicite "Suivi fermé ?" — jamais réutilisé un rapport de
simulation périmé, leçon retenue de la session F2) :

| Écriture | Lignes | Montant | Factures archivées | Tarif BL complété |
|---|---|---|---|---|
| 1 (sûres initiales) | 55 | 19 197,77 € | 19 | 15 |
| 2 (facture avant BL débloqué) | 21 | 2 692,46 € | 6 | 17 |
| 3 (prix jamais bloquant débloqué) | 28 | 3 443,39 € | 15 | 22 |
| **Total session** | **104** | **25 333,62 €** | **40** | **54** |

Résorption finale : **109 DISTRIBUTION** 852 lignes livrées encore sans
facture sur 1 000 (148 déjà facturées, contre 146 en fin de session F2) ;
**COREDIME** 800/864 (64 déjà facturées, contre 0 en début de session).
Reste 63 factures dans `a_traiter/Factures/À vérifier/` — majoritairement
le paquet "BdC manuels" ci-dessus, quelques vraies facturations partielles
(Qté facturée < Qté déjà livrée, légitime), la facture 6108788 (commande
113.071 déduite, jamais écrite automatiquement par principe, 9 lignes), et
5 vrais trous d'extraction/documents hors périmètre (`6107308.pdf`,
`6107594.pdf`, `6108474.pdf` — Total HT affiché mais 0 ligne extraite,
règle d'or : un seul exemple par cas, pas de correctif tenté ; `6108972.pdf`
l'AVOIR, mis de côté comme prévu ; `BC241723.pdf`, hors périmètre).

**Reste à faire (F4 suite)** :
1. Cominter (2e plus gros fournisseur par backlog réel, selon le
   classement donné par l'acheteur) — pièces BL déjà couvertes côté
   rapprochement BL, reste à cadrer/construire son parser FACTURE.
2. Proposer l'extension Hermes (lecture seule, extraction automatique des
   PDF de factures fournisseurs depuis ral@/achats@ depuis janvier 2026
   vers `a_traiter/Factures/`, anti-doublon par EntryID + n° de facture déjà
   présent dans le Suivi) — prévu après le 2e fournisseur, pas encore fait.
3. Regénérer l'état FNP d'août une fois qu'un fournisseur majeur est
   significativement résorbé, pour montrer la baisse du total à la DAF.
4. Décider avec l'acheteur du sort des 63 factures encore en "À
   vérifier" (notamment la grosse 6108788).

## Rapprochement factures — F4 : Cominter, 3e fournisseur (fait, 2026-09-02)

Suite directe de F4/Coredime : *"on passe à Cominter, je t'ai mis les
factures sous le même format que Coredime tout à l'heure dans le dossier."*
5 .msg de Prisca LEBLÉ (comptable) — 4 pour Cominter Réunion, 1 déposé en
cours de session pour Cominter Mayotte (*"normalement elle n'avait transmis
que Cominter [Réunion]"*) — extraits avec la même mécanique win32com que
Coredime (139 PDF au total). **Structure nettement plus hétérogène que
Coredime** : au moins 2 entités légalement distinctes détectées sous
"COMINTER" (Comptoir Ouest/Saint-Paul, Sainte-Clotilde, Saint-Pierre — SIRET
et adresses différents mais même gabarit de document) et une 3e vraiment à
part, "COMINTER MAYOTTE" (déjà séparée côté devis/BL, `cominter_mayotte.py`)
— **2 parsers facture distincts écrits**, un par entité.

**~40 des 139 fichiers sont des ANNEXES scannées** (`xxxxxx_BL_OBLxxxxxx.pdf`,
`xxxxxx_MAN.pdf`, `Annexe_1/2_MFACxxxxx.pdf`, `PJ_1/2_MFACxxxxx.pdf`) — 0
caractère de texte natif, hors périmètre du parser facture (le rapprochement
BL de ce fournisseur existe déjà séparément). **Piège réel identifié avant
tout code** : certains `_MAN.pdf`/`OFC194324.pdf` ont en fait du texte natif
mais ne contiennent QUE notre propre "DETAIL DE LA COMMANDE"/"BON DE
COMMANDE" (aucun champ Cominter) — détectés "COMINTER" (le mot apparaît
dans "DESTINATAIRE COMINTER") mais 0 ligne honnêtement extraite, aucun
traitement spécial nécessaire (même famille que les BC 109 Distribution
déjà documentés).

### Parser Cominter Réunion (`moteur/fournisseurs/cominter.py`, section
GABARIT FACTURE, 4 fixtures réelles)

Structure de ligne d'article structurellement identique au BL scanné du
même fournisseur (déjà éprouvé, voir GABARIT BL) mais en texte NATIF
(jamais de scan pour la facture, contrairement au BL) et un ordre de champs
scramblé différent : Qté, Px unitaire, Remise% optionnelle, Montant net,
Code TVA, Désignation, Cdt/Unité optionnel, Référence. Ancrage sur le
MONTANT (seule cellule fiable sur CHAQUE ligne, avec ou sans remise/Cdt) —
la désignation+référence de chaque article est bornée par les deux montants
consécutifs, jamais par un Cdt (absent sur les lignes d'éco-participation).
Zone bornée entre "Signature" et "Article 7. PROCEDURE..."/"NET A PAYER".

**4 bugs réels trouvés en confrontant 5 factures riches** (dont
`OFC193413.pdf`/commande M2.22.082 — la MÊME commande déjà documentée dans
l'incident "quantités non entières" d'une session précédente, "L406773"/
"4.5 KA" inclus, confirmant la parenté structurelle avec le BL) :
1. **Note de pied de tableau après la référence, sur la DERNIÈRE ligne
   d'articles du document** (`OFC194316.pdf`, "Livraison chantier
   Anzemberg... CAMBAIE CG CG" après la référence CAETM4288) : la référence
   n'est PAS toujours `zone[-1]` — recherchée en partant de la FIN de la
   zone jusqu'à trouver un token qui ressemble vraiment à une référence
   (`MOTIF_REF_ARTICLE_BL_COMINTER`, déjà éprouvé côté BL), tout ce qui suit
   étant simplement ignoré.
2. **Code TVA collé au montant sur la MÊME ligne** (`NFA018127.pdf`,
   "17,33 € 1" au lieu de deux lignes séparées) — sans ce repli, aucune
   ligne du document ne matchait le motif montant (ancré `$`) et toute la
   facture ressortait à 0 ligne.
3. **AUCUN repère "Signature"** avant le bloc [date, n° de BL, n° de
   commande] sur certaines pièces (`NF155008.pdf`, agence Saint-Pierre) —
   repli sur la 1ère ligne qui ressemble à un n° de BL
   (`MOTIF_BL_FACTURE_COMINTER`, sans risque de faux positif : 2-5 lettres
   + 5-7 chiffres, rien d'autre sur la ligne, ne matche aucune ligne du
   bandeau légal/adresse).
4. **Référence AVANT la désignation** (même `NF155008.pdf`, "L69731L /
   Prise 2P+T saillie Plexo gris / unite" — Cdt en minuscule aussi,
   comparaison désormais casse-insensible) au lieu d'après comme les 3
   autres fixtures : quand la référence trouvée est en PREMIÈRE position de
   la zone, la désignation est cherchée APRÈS elle (bornée par le prochain
   Cdt trouvé, ou à défaut à UNE seule ligne — pour ne jamais avaler une
   note de pied de document type "BC N°24 1581 DU 07/07/26" qui peut
   traîner juste après sur la dernière ligne d'articles).

N° de commande : le "BC N°..." de l'en-tête (avant "Signature") peut être
TRONQUÉ par rapport à sa réimpression dans le bloc Signature (cas réel
`NFA018127.pdf` : en-tête "BC:3240" vs Signature "BC N°24 3240", complet) —
le bloc Signature fait autorité quand il en fournit un, repli sur l'en-tête
sinon. Séparateur ESPACE toléré en plus de point/tiret
(`MOTIF_COMMANDE_FACTURE_COMINTER`, même famille que Cominter Mayotte).

### Parser Cominter Mayotte (`moteur/fournisseurs/cominter_mayotte.py`,
section GABARIT FACTURE, 3 fixtures réelles initiales + 7 supplémentaires
confirmées lors d'un 2e dépôt)

Structure proche mais DEUX différences réelles avec Cominter Réunion,
justifiant un module séparé (déjà la convention pour ce fournisseur côté
devis/BL) plutôt qu'une branche conditionnelle dans le fichier commun :
- **Aucun code TVA après le montant** — la référence suit DIRECTEMENT
  (jamais vu l'inverse sur les 10 pièces confrontées).
- **Aucun repère "Signature"** du tout — le repli "1ère ligne qui ressemble
  à un n° de BL" déjà écrit pour Cominter Réunion sert ici de repère
  PRINCIPAL, pas seulement de secours.
- **N° de commande** : étiquette EXPLICITE "- N° de Commande : ..." plus
  loin dans le document, préférée à l'en-tête (toujours vue, jamais
  tronquée contrairement à Réunion) — séparateur ESPACE normalisé par
  `re.sub(r"\s+", ".", ...)` (pas un simple `replace(" ", ".")`, sinon un
  double espace donne un double point, cas réel "24  3109").
- **Note parasite possible entre le Cdt d'un article et la Qté du
  suivant** (cas réel "VARIANTE DISPO GTL", `MFAC15576.pdf`, 13 lignes) —
  sans risque, la désignation reste bornée au Cdt trouvé, la note ne
  rentre jamais dans la zone utile.
Validé sur `MFAC15576.pdf` (13 lignes, total recalculé 6 516,33€ = Total
HT affiché exactement) et confirmé sur 7 factures fraîches supplémentaires
(2e .msg) sans nouveau bug trouvé.

### Recette réelle et écriture (2026-09-02)

Lecture seule sur les 139 fichiers (103 factures réellement lues, 50
"fournisseur non reconnu" = annexes scannées, 7 "aucune ligne extraite" =
BC ré-attachés) : **228 lignes sûres, 38 035,84 €**, 28 à confirmer
(majoritairement de vraies facturations partielles, Qté facturée < Qté
déjà livrée — légitime), 37 inconnues (surtout des éco-taxes ambiguës,
plusieurs lignes Suivi partageant la même référence générique), 31 blocs
"commande introuvable" — **la grande majorité au format "24.XXXX" côté
Mayotte** (même famille de BdC manuels déjà actée pour Coredime, pas un
bug), quelques-uns côté Réunion avec des formats courts inhabituels
("1.030", "1.329", "2.091") à confirmer avec l'acheteur si l'occasion se
présente.

Suite complète (361 tests) vérifiée verte avant écriture. **Écriture
réelle confirmée** : 228 lignes écrites, sauvegarde horodatée avant coup,
39 factures archivées avec leur BC dans `Traités/<commande>/` (37
Cominter + 2 Cominter Mayotte), 114 déplacées vers `À vérifier/` (0 échec
d'archivage, 0 fichier sans parser). Contrôle anti-imbrication explicite
dans le script d'écriture (leçon retenue de Coredime) : confirmé absent.
Résorption : **COMINTER** 1202 lignes livrées encore sans facture sur 1389
(187 déjà facturées, contre 0 en début de session) ; **COMINTER MAYOTTE**
0/0 (aucune ligne Mayotte n'a encore de Qté livrée non nulle dans le
Suivi — cohérent avec le motif "24.XXXX/commande introuvable" constaté :
le BL n'est pas non plus encore rapproché pour ces commandes).

**Reste à faire (F4 suite)** :
1. Les 114 factures en `À vérifier/` — surtout le paquet "24.XXXX"
   Mayotte à trancher avec l'acheteur, quelques formats courts Réunion à
   éclaircir.
2. Extension Hermes (proposée, pas encore construite) et regénération de
   l'état FNP d'août, toujours en attente d'un fournisseur "significativement
   résorbé" — Coredime ET Cominter le sont maintenant, bon moment pour
   reprendre ces deux points.
3. Fournisseurs BL déjà couverts mais sans facture testée : Electric
   Plus/GMR (facture=BL, déjà documenté), Ravate, DEM, YESSS,
   Protecthoms — dès que de vraies factures seront déposées pour eux.

## Rapprochement factures — F4 : Stand 64, 4e fournisseur + nouveau lot
Coredime + EDOI amorcé (fait, 2026-09-02)

Suite directe de F4/Cominter. L'acheteur dépose 3 lots en une fois :
factures Stand 64 (nouveau fournisseur, ~25 pièces, déposées directement en
PDF — pas de .msg à extraire cette fois), de nouvelles factures Coredime
(19 pièces, format `Nouvelles factures disponibles` automatique du portail
Sonepar — 1 seul PDF en pièce jointe DIRECTE par .msg, pas d'e-mail
imbriqué comme chez Prisca LEBLÉ, extraction bien plus simple), et EDOI
(2 pièces, même mécanisme Sonepar) — *"la filiale soeur SONEPAR de
Coredime à Mayotte [...] les deux principaux cas de BL non chiffrés"*.

### Parser facture Stand 64 (`moteur/fournisseurs/stand64.py`, section
GABARIT FACTURE, 4 fixtures réelles)

Texte NATIF (jamais de scan pour la facture, contrairement au BL du même
fournisseur qui, lui, en a besoin malgré l'apparence nette). Structure par
article : **le Total HT est imprimé AVANT la désignation** (pas après,
contrairement à la plupart des autres fournisseurs) — vérifié par
cohérence arithmétique exacte sur une pièce à 6 articles (la somme des 6
Total HT retombe pile sur le Total HT affiché). Puis Désignation (1-2
lignes), Référence, code TVA (Cx), puis Qté/P.U Net/P.U/Rem% dans cet
ordre.

**3 bugs réels trouvés et corrigés en confrontant plusieurs factures
riches** :
1. **Un mot de désignation tout en MAJUSCULES sans chiffre ni tiret**
   ("CHAINETTE", fin d'une désignation qui déborde sur 2 lignes) était pris
   à tort pour une référence par un motif trop permissif — TOUTES les
   vraies références contiennent au moins UN TIRET (même
   "WESTI-COMET-KITLUM-N", sans aucun chiffre) : signal fiable retenu à la
   place de "au moins un chiffre" (qui aurait exclu cette référence
   légitime). Référence contenant un "+" ("ELIOT-ES52-2678-BLC+BLC", déjà
   connue côté BL du même fournisseur) : toléré dans les segments.
2. **BUG DE FOND, le plus important : l'Eco-part est présente ou non LIGNE
   PAR LIGNE au sein d'un MÊME document** (cas réel `Facture_34184.pdf` :
   2 lignes sur 3 avec Eco-part renseignée — une 5e valeur numérique après
   le code TVA en plus des 4 habituelles —, 1 sans) — un compte de valeurs
   numériques fixe après le code TVA décalait alors une ligne sur deux
   (perdait la ligne suivante, polluait sa désignation avec des chiffres
   errants). Corrigé par un changement d'ANCRAGE : au lieu de compter les
   valeurs numériques en avançant depuis le code TVA, `_lignes_facture_stand64()`
   trouve d'abord TOUTES les références de la zone (motif fiable dans les
   deux sens), puis pour chacune : Qté et P.U Net sont à position FIXE
   juste après le code TVA (fiable quel que soit le nombre de valeurs
   restantes, jamais besoin de les compter) ; Total HT et désignation sont
   retrouvés en REMONTANT depuis la référence (tout ce qui n'est pas
   numérique juste avant elle est désignation, le premier nombre rencontré
   au-delà est le Total HT) — ne dépend jamais du nombre de valeurs
   consommées par la ligne précédente. Un caractère isolé ("+") parfois
   imprimé entre l'Eco-part d'une ligne et le Total HT de la suivante est
   ignoré naturellement par cette approche.
3. **N° de commande** : l'en-tête ("BC N°...") peut avoir un point réel
   décalé par rapport à notre propre "BON DE COMMANDE" reproduit en pied de
   certaines factures (cas réel : en-tête "BC N°M2.220.78" vs BON DE
   COMMANDE "M2.22.078", mêmes chiffres "22078") — celui-ci est préféré
   quand disponible, pas de correctif général tenté sur l'en-tête (un seul
   exemple, règle d'or).

Balayage complet sur les 23 vraies factures Stand 64 disponibles : **23/23
exactes** (Total HT recalculé = Total HT affiché à chaque fois) après ces
correctifs.

### Nouveau lot Coredime (parser déjà existant, aucun code nécessaire)

19 nouvelles factures extraites depuis les .msg de notification portail
Sonepar (structure `Cher(e) client(e), veuillez trouver en pièce jointe...`,
1 PDF direct par .msg — bien plus simple à extraire que le montage
imbriqué de Prisca LEBLÉ). Traitées par le parser déjà construit en F4
Coredime, sans aucune modification. 2 anomalies de lecture connues (même
famille que les gaps déjà documentés : Total HT affiché mais 0 ligne
extraite, limite acceptée du garde-fou remise-double conservateur).

### EDOI — investigué, PAS ENCORE construit

Les 2 pièces disponibles (format `6005921.pdf`/`6005922.pdf`, mêmes .msg de
notification Sonepar qu'un nouveau Coredime) révèlent une structure
prometteuse mais pas encore assez éprouvée pour coder un parser fiable :
- **Mêmes repères de page que Coredime** (`##ESIGUID;...`, `##NM#FR;...`,
  `#####DEMAT-FJ;...;Facture;<date>;<montant>;<numero>;...;EDOI;...;`) —
  cohérent avec le fait que Coredime et EDOI sont deux filiales Sonepar
  partageant probablement la même plateforme de facturation. Une bonne
  base de départ pour construire le parser EDOI en réutilisant l'ossature
  déjà éprouvée côté Coredime plutôt qu'en repartant de zéro.
- **Un même PDF contient à la fois la FACTURE (page 1) et le BON DE
  LIVRAISON correspondant** (page suivante, même numéro de document
  "EDO Bxxxxxx", "BON D'EXPEDITION N°Bxxxxxx"). **Expliqué par
  l'acheteur, PAS une bizarrerie EDOI** : "c'est une exigence de notre
  service comptable, les factures doivent venir avec leur BL et leur
  commande" — même règle que le regroupement facture+BdC déjà observé chez
  109 Distribution (73 % des factures, voir cadrage F1 plus haut, "piste à
  creuser" alors non résolue). **Portée générale, à garder en tête pour
  tout futur fournisseur** : ce n'est pas une structure de document propre
  à un fournisseur donné, donc pas surprenant de la revoir ailleurs — ne
  plus la signaler comme une anomalie de format, juste borner la zone
  BL/BC pour ne pas la confondre avec la zone facture (déjà fait ici et
  chez 109 Distribution).
- Énormément de texte de CGV (conditions générales de vente, ~500 lignes)
  entoure le contenu utile — la zone d'articles réelle devra être bornée
  aussi précisément que chez Coredime, voire plus.
- **Seulement 2 pièces disponibles à ce jour, chacune avec 1 seule ligne
  d'article** — pas assez pour valider un motif de ligne multi-article
  fiable (règle d'or : jamais de règle de parsing inventée sans assez
  d'exemples réels). À reprendre dans une session dédiée dès que davantage
  de factures EDOI seront disponibles. En attendant, les 2 pièces restent
  `a_traiter/Factures/` (fournisseur EDOI déjà reconnu par le détecteur —
  utilisé côté devis — mais sans parser facture : `factures_sans_parser`,
  jamais déplacées vers "À vérifier/", comportement déjà tolérant établi).

### Recette réelle et écriture (2026-09-02)

Lecture seule sur les 47 fichiers du dossier (41 factures lues : 19
Coredime + 22 Stand 64 ; 2 EDOI mises de côté proprement ; 4 fichiers
`FA <numéro>.pdf` restés MYSTÉRIEUX — scans purs, 0 caractère, jamais
identifiés avec certitude, à clarifier avec l'acheteur) : **61 lignes
sûres, 27 073,72 €**, 5 à confirmer (facturations partielles légitimes +
1 repli référence-proche), 14 inconnues (surtout des éco-taxes ambiguës),
7 blocs "commande introuvable" (dont un format Stand 64 sans aucun
séparateur, "240618", à éclaircir).

Suite complète vérifiée verte avant écriture. **Écriture réelle
confirmée** : 61 lignes écrites, sauvegarde horodatée, 21 factures
archivées avec leur BC (7 Coredime + 14 Stand 64), 24 déplacées vers
`À vérifier/` (0 échec d'archivage). Résorption : **COREDIME** 779/864
lignes livrées encore sans facture (85 déjà facturées) ; **STAND 64**
73/102 (29 déjà facturées — premier vrai chiffre pour ce fournisseur tout
neuf, contre 0 en début de session).

**Reste à faire** :
1. ~~Identifier les 4 fichiers `FA <numéro>.pdf`~~ **FAIT** — confirmés par
   l'acheteur comme des factures **Stand 64** (voir section suivante,
   scans purs, pas encore de `parse_facture_ocr` pour ce fournisseur).
2. Construire le parser facture EDOI dès que davantage de pièces réelles
   seront disponibles (voir ci-dessus).
3. Les 24 factures en `À vérifier/` de cette session, plus le reliquat des
   sessions précédentes (Coredime/Cominter).
4. Toujours en attente : extension Hermes, regénération FNP d'août.

## Rapprochement factures — F4 : Electric Plus/GMR, 5e fournisseur, repli
OCR générique du flux Facture (fait, 2026-09-02)

Suite directe de la session Stand 64 ci-dessus, sur demande explicite de
l'acheteur : *"je t'ai maintenant ajouté les factures Electric Plus (GMR),
la particularité c'est que eux n'ont pas de BL à proprement parler, ce sont
directement des factures qui font office de BL. On suit le même processus
tout de même, on enchaîne."*

### Repli OCR générique dans le flux Facture (nouveau, architecture)

Jusqu'ici, **tous** les parsers facture (109 Distribution, Coredime,
Cominter, Cominter Mayotte, Stand 64) lisaient du texte PDF NATIF
(`moteur.lecture_pdf.lire_pdf`) — `moteur/rapprochement/lecture_facture.py`
appelait `parser_facture(fournisseur, texte)` sans alternative. Electric
Plus/GMR est le premier fournisseur FACTURE dont les pièces sont des SCANS
purs (comme son BL, déjà couvert) : `lire_pdf()` y renvoie une chaîne
vide.

Plutôt que de câbler un cas spécial "Electric Plus", **repli OCR
générique** ajouté à `lire_facture()` : si `lire_pdf()` ne renvoie AUCUN
texte, la lecture bascule sur l'OCR (`moteur.ocr.mots_document`, même
détection de fournisseur qu'avant mais sur le texte OCR) et appelle un
NOUVEAU registre `parser_facture_ocr(fournisseur, mots_par_page)` (
`moteur/rapprochement/parsers_facture.py`, découvre un attribut
`parse_facture_ocr` optionnel sur chaque module fournisseur, symétrique à
`parse_facture`). **Profite à tout futur fournisseur facture scanné**, pas
seulement Electric Plus — par exemple les 4 factures Stand 64 "FA
<numéro>.pdf" identifiées ci-dessus (scans purs) seront automatiquement
prises en charge dès qu'un `parse_facture_ocr` existera pour ce
fournisseur, sans retoucher `lecture_facture.py`.

### Parser (`moteur/fournisseurs/electricplus.py::parse_facture_electricplus_ocr`)

Réutilise au maximum l'extraction déjà éprouvée côté BL (même document,
même OCR — GMR n'a qu'UN SEUL document, qui sert aux deux flux) :
refactor pur de `_ligne_vers_article_electricplus`/
`_parse_une_facture_electricplus` pour extraire une fonction commune
(`_champs_ligne_electricplus`, `_entete_et_lignes_electricplus`) réutilisée
par les DEUX constructeurs (`_ligne_vers_article_electricplus` pour
`LigneBL`, `_ligne_vers_ligne_facture_electricplus`/
`_construire_facture_electricplus` pour `Facture`/`LigneFacture`) — vérifié
sans aucune régression sur les 14 tests BL/devis déjà existants (identique
avant/après refactor).

**Découverte majeure, nouvelle pratique de l'acheteur (pas une bizarrerie
GMR)** : suite à l'exigence de son service comptable ("les factures
doivent venir avec leur BL et leur commande", déjà documentée plus haut
pour 109 Distribution — même règle, 2e fournisseur qui la révèle), chaque
facture GMR de ce lot est accompagnée de NOTRE PROPRE "BON DE COMMANDE"
généré par ce projet, et PARFOIS du DEVIS d'origine — plusieurs types de
documents empilés dans un même fichier PDF.

**2 bugs réels trouvés et corrigés en confrontant 21 vraies factures** :

1. **Désignation étalée sur PLUSIEURS lignes visuelles OCR** (3 des 4
   articles d'une même facture réelle, `4205720`) : référence + début de
   désignation sur une ligne, quantité/prix SANS suffixe PF/PR sur la
   ligne suivante, puis un COMPLÉMENT de désignation (taille, ex.
   "2,4x180") sur une 3e ligne ENCORE APRÈS les prix. Chaque ligne prise
   isolément échouait totalement (ni ancre PF, ni assez de cellules pour
   le repli positionnel) — ces 3 articles disparaissaient purement et
   simplement. Nouvelle fonction `_regrouper_articles_electricplus()` :
   une ligne dont la 1re cellule ressemble à une vraie référence
   (`MOTIF_DEBUT_REFERENCE_ELECTRICPLUS` — lettres majuscules suivies d'au
   moins un chiffre, jamais vrai pour un nombre nu ni pour un mot pur type
   "BLIST") démarre un nouvel article ; toute ligne suivante est absorbée,
   CHAQUE cellule reclassée individuellement en désignation ou en nombre
   (`_cellule_ressemble_a_nombre_electricplus`) plutôt que simplement
   concaténée dans l'ordre de lecture — indispensable car la
   désignation-complément arrive APRÈS les nombres dans l'ordre de
   lecture : une simple concaténation aurait décalé le repli positionnel
   "4 dernières cellules" et pris la désignation-complément pour le
   Montant.
   **Piège trouvé en l'appliquant d'abord au code PARTAGÉ** : un 1er essai
   branchait ce regroupement directement dans `_entete_et_lignes_
   electricplus()` (partagée BL/Facture) — a fait régresser 8 des 14 tests
   BL déjà verrouillés (une ligne de bruit, auparavant isolée et donc
   silencieusement ignorée sur les fixtures BL existantes, se retrouvait
   absorbée à tort dans la désignation d'un article réel). Corrigé en
   scopant le regroupement au SEUL chemin Facture
   (`_construire_facture_electricplus`), jamais le chemin BL — aucun cas
   réel à ce jour ne montre ce besoin côté BL. **Leçon générale** : un
   refactor "extraire une fonction commune" ne veut pas dire "partager
   TOUS les traitements futurs" — un correctif motivé par un seul flux
   doit rester scopé à ce flux tant qu'un besoin réel côté de l'autre flux
   n'est pas démontré.
2. **Pages DEVIS/BON DE COMMANDE risquant de fausser le découpage par
   n° de facture** (`pages_par_identifiant`, ancré sur un nombre nu de 6-7
   chiffres) : une page DEVIS porte SA PROPRE numérotation (ex. "4104132"
   pour la facture "4205769" du même article — cas réel) — sans filtrage,
   elle démarre un groupe à part entière, produisant une 2e "Facture"
   fantôme avec la ligne d'article DUPLIQUÉE (silencieux, potentiellement
   grave). Une page de NOTRE PROPRE bon de commande peut, elle, porter un
   nombre de 6-7 chiffres SANS RAPPORT (ex. une date collée "120720") qui
   déclenche le même risque, mais avec un symptôme différent et plus
   sournois : une facture par ailleurs PARFAITEMENT exacte (Total HT
   extrait = affiché au centime près) se voyait quand même accompagnée
   d'une 2e "Facture" fantôme à 0 ligne — suffisant pour faire basculer
   TOUT LE FICHIER vers "à vérifier" malgré une extraction sans le moindre
   défaut. Nouvelle fonction `_est_page_hors_perimetre_electricplus()` :
   exclut toute page mentionnant "DEVIS" ou "BON DE COMMANDE" (repéré
   fiable sur les 3 vraies pages BC confrontées, malgré des libellés de
   détail différents ensuite) AVANT tout regroupement par identifiant —
   appliquée UNIQUEMENT dans `parse_facture_electricplus_ocr` (jamais côté
   BL, où aucune fixture n'a ce problème à ce jour).

### Recette réelle sur 21 vraies factures (lecture seule + confrontation
individuelle) puis écriture réelle confirmée

**20/21 fichiers parfaitement exacts** (Total HT recalculé = Total HT
affiché, commande résolue via le libellé "CDE" quand présent) — y compris
plusieurs factures multi-articles réelles qui valident le correctif n°1 à
l'échelle (11, 13, 15 lignes sur 3 factures différentes, chacune exacte au
centime). 6 fichiers regroupent PLUSIEURS factures GMR (jusqu'à 3, comme
déjà vu côté BL) — chacune correctement séparée par
`pages_par_identifiant`, aucune fusion à tort entre 2 commandes
différentes.

**1 fichier à part, 3e bug réel corrigé** (`ESPACE SOLEIL-1206686+1206681+
1206661+BC-16.07.2026.pdf`, 3 factures bundlées) : les 3 groupes
ressortaient d'abord à 0 ligne / numéro de facture vide. Creusé en détail :
sur CES 3 pages précises (et seulement celles-ci, parmi les 21 fichiers),
**l'ordre des LIGNES ET DES CELLULES est inversé de bout en bout** —
l'en-tête ressort "MONTANT HT | P.U.NET HT | PRIX UNIT.HT | QTE |
DESIGNATION | REFERENCES" (la référence en DERNIER) au lieu de l'ordre
habituel, le pied de tableau ("TOTAL HT"/"CODES TVA") apparaît AVANT les
lignes d'articles plutôt qu'après, et le bloc [date, date échéance, n° de
facture] suit l'ordre inverse. Cohérent avec un scan de CE lot précis
effectué à l'envers (les 3 factures montrent EXACTEMENT la même
inversion) — pas 3 erreurs OCR indépendantes.

Corrigé par 2 replis, actifs uniquement quand le motif normal échoue
(comportement des 20 autres fichiers strictement inchangé) :
- `_zone_tableau_electricplus()` détecte l'inversion en comparant la
  position de la cellule "REFERENCES" dans la ligne d'en-tête (2e moitié
  plutôt qu'en tête) : si détectée, cherche le pied de tableau AVANT
  l'en-tête plutôt qu'après (structurellement impossible sur une page
  normale, donc un signal sûr) pour délimiter la zone, PUIS réinverse
  CHAQUE ligne de la zone — tout le reste du code (ancre PF, repli
  positionnel, regroupement multi-lignes) suppose ensuite l'ordre normal
  sans dupliquer aucune logique.
- `MOTIF_FACTURE_DATE_ELECTRICPLUS_MIROIR` : repli sur l'ordre inversé
  [n° facture, date échéance, date] quand l'ordre normal ne matche pas.

**2 des 3 factures retombent exactement sur leur Total HT après ce
correctif** (460,00€/5 lignes et 5 600,76€/12 lignes — cette dernière
valide le correctif à belle échelle, mélangeant câbles/bobines/embouts de
références très différentes). **La 3e (1206681, 1 seule ligne à 8,90€)
reste à 0 ligne** : un fragment de cellule illisible ("680", ni un nombre
ni une référence reconnaissable une fois l'ordre restauré) fait échouer à
la fois l'ancre PF (une cellule de trop entre l'ancre et la fin de ligne)
et le repli positionnel (repli sur `to_float()` d'une cellule non
numérique → 0 → ligne rejetée). Laissé tel quel — honnêtement signalé
(aucune ligne extraite, la facture partira en "à vérifier"), jamais
deviné : la 3e occurrence d'une même famille de bug sur un SEUL fichier
suffisait pour ce correctif-ci, pas de raison d'aller plus loin pour 8,90€
de résiduel sur une seule ligne (règle d'or).

### Recette finale : 20/21 fichiers parfaitement exacts, 2/3 factures du
21e récupérées après le correctif miroir — soit **22 factures exactes sur
23 réellement rapprochables**, la seule restante (1206681) honnêtement
signalée plutôt que devinée.

### Tests

`tests/test_parsers_facture_electricplus.py` (4 tests, sur 4 vraies pièces
copiées dans `tests/fixtures/facture_electricplus_1..4_*.pdf`) : le cas
désignation multi-ligne (4 articles, 1 seul extrait avant le correctif),
le cas devis+BC agrafés (vérifie qu'une SEULE facture ressort malgré 3
pages de nature différente), le cas commande via libellé CDE + suffixe PR,
et le cas 3-pages-miroir (2 factures exactes, la 3e à 0 ligne verrouillée
telle quelle). Suite BL/devis existante (14 tests) revérifiée verte deux
fois (après le refactor partagé, puis après le correctif miroir) — aucune
régression sur les 17 tests au total.

### Écriture réelle confirmée (2026-09-02, Suivi fermé, confirmé
explicitement par l'acheteur juste avant chaque étape — lecture PUIS
écriture, elle a rouvert le fichier entre les deux et l'a refermé avant
que j'écrive)

Relecture fraîche juste avant écriture (jamais un rapport périmé — leçon
déjà retenue en F2/F4) : **69 lignes sûres, 19 752,05 €**, 1 à confirmer
(repli référence-proche, "CAB013010C100" vs Suivi "CAB013010T1000"), 40
inconnues, 3 blocs "commande introuvable" (N°Réf.Client vide + déduction
par contenu non concluante).

**Motif net parmi les inconnues, à surveiller** : une bonne partie
(~30/40) sont des références "CAB000XXXC100" (bobines de câble HO7V-U
1,5mm² par couleur, ex. "CAB000170C100"=violet, "CAB000110C100"=rouge...)
— ressortent SYSTÉMATIQUEMENT "aucune ligne du Suivi ne correspond" sur
PLUSIEURS commandes différentes et sans rapport (1206671, 1206684,
1206685, 1206686, 1206661). Pas assez d'un seul cas pour un correctif
(règle d'or), mais la RÉCURRENCE sur autant de commandes différentes
suggère que le Suivi enregistre peut-être ces bobines sous une référence
différente de façon systématique plutôt que ligne par ligne — à vérifier
avec l'acheteur si ce motif revient.

**Écriture réelle confirmée** : 69 lignes écrites, sauvegarde horodatée,
**13 factures archivées** avec leur BC dans `Traités/<commande>/`, **9
déplacées vers `À vérifier/`** (0 échec d'archivage). EDOI (2 fichiers)
laissé en place, toujours sans parser. Contrôle anti-imbrication confirmé
absent. Résorption : **ELECTRIC PLUS** 311 lignes livrées encore sans
facture sur 377 (66 déjà facturées — premier vrai chiffre pour ce
fournisseur côté facture, distinct de sa résorption côté BL).

## Rapprochement factures — Coredime au complet, 1er semestre 2026 +
avoirs (fait, 2026-09-02)

Sur demande explicite de l'acheteur : *"je t'ai déposé les dernières
factures de 109 ainsi que les factures de Coredime pour les six premiers
mois de l'année, avoirs compris, sous format zip. On va traiter un
fournisseur au complet comme ça !"* — le plus gros lot jamais soumis à ce
projet : **328 factures Coredime** (6 zips) + **29 factures 109
Distribution** récentes (déposées à plat, non zippées) + les 2 EDOI déjà
en attente.

**Incident réseau pendant l'extraction** (même famille que la panne déjà
documentée en session BL) : le lecteur X: est devenu temporairement
inaccessible en pleine extraction des zips (`unzip` a levé des erreurs
"bad zipfile offset" en cascade). Diagnostiqué comme une vraie coupure
(même `echo` en Bash et `Get-ChildItem` en PowerShell échouaient) plutôt
qu'une corruption : `unzip -t` sur le zip concerné, une fois la connexion
rétablie (confirmée par l'acheteur), a montré "No errors detected" — zéro
donnée perdue. **Leçon retenue, suggérée par l'acheteur** : extraire et
traiter les zips UN PAR UN plutôt qu'en boucle, pour limiter la casse et
faciliter le diagnostic en cas de nouvelle coupure — adopté pour la suite
de cette session.

### 3 bugs réels corrigés dans le parser facture Coredime
(`moteur/fournisseurs/coredime.py`), trouvés en confrontant les 328
pièces réelles à l'autocontrôle Total HT

Avant tout correctif : **48 factures sur 328 avec un écart de Total HT,
11 828,13 € d'écart cumulé** — largement au-delà des petits résidus déjà
connus (garde-fou remise-double, voir plus bas). Sur les pires cas,
jusqu'à 85 % du montant d'une facture manquait silencieusement.

1. **Troncature prématurée sur les factures multi-folios** (cause de
   loin la plus lourde : la quasi-totalité des 11 828 € ci-dessus). Le
   repère de fin de zone facture supposait "2e occurrence du marqueur de
   page ##ESIGUID" = fin de la facture, un raisonnement valide seulement
   pour une facture à 1 SEULE page de contenu + 1 annexe. Une facture à
   beaucoup de lignes s'étale en réalité sur PLUSIEURS folios (pages) DE
   CONTENU, chacun avec son propre "##ESIGUID" répété en en-tête — le 2e
   folio (pas une annexe) était donc pris à tort pour la fin de la
   facture, tronquant tout le reste. Cas réel le plus parlant
   (`6105181.pdf`) : 24 lignes réelles sur 2 folios, seules 11 extraites,
   2 984,98 € manquants sur 8 438,47 €.
2. **Annexe "BON DE LIVRAISON" de Coredime lui-même, DOUBLANT le total**
   (cas réel `6200396.pdf`, 240,00 € extraits au lieu de 120,00 €) :
   Coredime peut annexer son propre bon de livraison à la suite de la
   facture, avec un tableau d'articles dans EXACTEMENT le même format
   qu'une vraie ligne de facture — sans repère pour l'exclure (la zone
   s'étendait "jusqu'à la fin du texte" faute de bon de commande sur cette
   pièce précise), ce 2e tableau était absorbé intégralement.
3. **Notre propre numéro de téléphone confondu avec une ligne de facture
   incomplète** (cas réel `6401314.pdf`, 0 ligne extraite pour une
   facture d'1 seule ligne, 29,16 €) : l'annexe "BON DE COMMANDE" contient
   un bloc signature ("DATE/ACHETEUR/VISA", notre téléphone "0693 86 68
   03") resté dans la zone scannée, qui matchait accidentellement le motif
   de "ligne incomplète" (mécanisme de récupération remise-double) comme
   un 2e faux candidat à côté du vrai — désamorçant l'appariement 1:1
   pourtant sans ambiguïté sur cette pièce.

**Leçon commune aux bugs 2 et 3, trouvée en deux temps** : le premier
réflexe pour exclure une annexe (chercher son TITRE — "BON DE COMMANDE",
"B O N  D E  L I V R A I S O N" imprimé par Coredime en lettres espacées
comme "F A C T U R E") échoue dans les deux cas, car ces titres s'impriment
en PIED de leur propre bloc — donc APRÈS le contenu qu'il fallait
justement exclure. Repères EN TÊTE de bloc retenus à la place :
"COR B<num>" (référence isolée de l'annexe BL, même famille que
"COR F<num>" déjà utilisé pour les totaux) et "DESTINATAIRE" (toute
première ligne de notre propre BC).

**Résultat final** : 40 factures avec un écart résiduel (contre 48),
**706,75 € d'écart cumulé (contre 11 828,13 €, -94 %)** — l'immense
majorité des résidus restants sont désormais de l'ordre du centime à
quelques euros, cohérents avec la limite déjà connue et acceptée
(remise multiple ambiguë, voir "Points fragiles"). 4 tests ajoutés
(`tests/test_parsers_facture_coredime.py`, 11 au total), 3 nouvelles
fixtures réelles (`facture_coredime_9..11_*.pdf`) verrouillant chacun des
3 bugs ci-dessus.

### Recette réelle et écriture (2026-09-02, Suivi libéré par l'acheteur
pour "les prochaines heures")

Lecture seule sur les 358 fichiers (356 factures lues sur 358 — 2 échecs
de lecture nets : `6100597_1f784aa1-....pdf`, un doublon de scan à 0
ligne, et `6300524.pdf` — 4 AVOIRS reconnus et mis de côté, 1ère
confirmation réelle du mécanisme AVOIR sur données Coredime réelles) :
**434 lignes sûres, 73 818,66 €** (COREDIME 382 lignes/67 347,85 € +
109 DISTRIBUTION 52 lignes/6 470,81 €), 232 à confirmer, 154 inconnues —
en très large partie des commandes 2025 introuvables dans le Suivi (qui
ne couvre que l'année civile en cours), conformément à l'avertissement de
l'acheteur avant même de déposer le lot.

**Écriture réelle confirmée** (relecture fraîche juste avant, comptes
identiques à la lecture seule — rien n'avait changé entre-temps) :
**434 lignes écrites**, sauvegarde horodatée, **105 factures archivées**
avec leur BC dans `Traités/<commande>/`, **251 déplacées vers
`À vérifier/`** (0 échec d'archivage), **73 "Tarif BL" renseignés depuis
la facture** (liste blanche Coredime, Tarif BL vide comblé). EDOI (2
fichiers) laissé en place, toujours sans parser.

**Résorption, avant/après ce lot — 1ère vraie mise à l'échelle de
Coredime** :

| Fournisseur | Avant | Après |
|---|---|---|
| COREDIME | 0 déjà facturées | **442/865 déjà facturées** (423 restantes) |
| 109 DISTRIBUTION | 148 déjà facturées | **191/1001 déjà facturées** (810 restantes) |

**Reste à faire** : les 251 factures en `À vérifier/` (dominées par le
motif "commande introuvable" pré-2026, à trier avec l'acheteur — pas
grand-chose d'actionnable dessus tant que le Suivi ne couvre pas les
années antérieures) ; EDOI toujours sans parser ; extension Hermes et
FNP d'août toujours en attente.

## Rapprochement factures — session S0 (2026-09-03) : matching v1.1
(agrégation multi-BL, frais connus, causes codées) + INCIDENT Suivi majeur

Session pilotée par un plan écrit à l'avance par l'acheteur (règles 0.4,
"un seul plan, exécuté dans l'ordre"). Étapes 0-3 : 5 corrections dans
`moteur/rapprochement/matching_facture.py`/`pipeline_facture.py`, un lot
ciblé de factures réécrites, puis un INCIDENT sérieux sur le Suivi vivant
en plein milieu — traité en détail plus bas, à lire avant toute session
future qui toucherait à l'écriture facture.

**Étape 0 — 8 fixtures créées, 1 correction de trajectoire mineure** : le
plan indiquait `a_traiter/Factures/S0/` comme emplacement des 8 PDF réels à
transformer en fixtures — ce dossier n'existait pas, les 8 factures
étaient en réalité déjà dans `a_traiter/Factures/À vérifier/` sous leurs
noms d'origine (retrouvées avec certitude par les numéros de facture déjà
donnés dans les noms de fixtures cibles, aucune ambiguïté) — copiées
(jamais déplacées) vers `tests/fixtures/` :
`facture_109_362840_multi_bl_meme_ref.pdf`, `facture_109_362763_partielle.pdf`,
`facture_109_362777_bdc_manuel.pdf`, `facture_coredime_6108234_suffixe.pdf`,
`facture_coredime_6107800_ecotaxe.pdf`, `facture_coredime_6100226_bdc_manuel.pdf`,
`facture_coredime_6108846_remise_double_x3.pdf`, `facture_stand64_33330_scan.pdf`
(cette dernière hors périmètre de cette session — voir "Pas de Stand 64" du
plan — c'est le 4e fichier "FA \<numéro\>.pdf" scanné déjà identifié comme
Stand 64 sans `parse_facture_ocr`, simplement mis en réserve).

**Étape 1 — 5 corrections, toutes vérifiées sur PDF réels** :

- **(a) Agrégation multi-BL même référence** (`moteur.rapprochement.
  matching_facture.agreger_lignes_meme_reference`, nouvelle fonction) : cas
  réel Facture_362840.pdf (109, commande 123.089) — P03200 et F2U15RVVOO
  sont chacun répartis sur 2 "Bon de livraison" DIFFÉRENTS de la MÊME
  facture (livraison fractionnée). Avant ce correctif, `pipeline_facture.
  rapprocher_dossier_factures` groupait les lignes PAR BL BRUT (pas par
  commande résolue) avant d'appeler `apparier_facture` — chaque bloc,
  comparé isolément à la Qté livrée TOTALE du Suivi, ressortait "à
  confirmer" à tort (100 facturés sur ce bloc vs 200 déjà livrés en tout).
  Deux correctifs liés : `agreger_lignes_meme_reference()` (interne à
  `apparier_facture`, somme les quantités/montants des lignes qui
  partagent la même référence normalisée, PU identique exigé — sinon
  aucune agrégation, chaque ligne reste séparée et ressort "à confirmer",
  cause=PRIX_DIFF_MEME_REF, jamais un prix deviné) ET le regroupement dans
  `rapprocher_dossier_factures` changé de "par BL brut" à "par commande
  RÉSOLUE" (plusieurs BL d'une même facture qui pointent vers la même
  commande sont désormais fusionnés avant l'appel à `apparier_facture`).
  Validé de bout en bout sur la vraie pièce : 11 lignes brutes → 9
  correspondances (agrégation réussie = MOINS de lignes, jamais autant),
  toutes sûres.
- **(b) Repli "premier token" — référence Suivi à suffixe libre**
  (`_repli_premier_token`, nouveau, entre le repli référence-proche et le
  repli référentiel dans `apparier_facture`) : cas réel
  Facture_6108234.pdf (Coredime, commande M3.14.342) — Suivi
  `"SIXGPCP35 PVC"` (texte descriptif ajouté à la saisie) vs facture
  `"SIXGPCP35"` — comparaison sur le premier terme (avant le 1er espace)
  du Suivi, UN SEUL candidat exigé, toujours "à confirmer".
- **(c) Frais connus, jamais bloquants** (`referentiel/frais_fournisseurs.csv`,
  nouveau fichier + `charger_frais_fournisseurs()`/nouveau `StatutFacture.
  FRAIS` dans `apparier_facture`) : 2 entrées réelles évidencées —
  `COREDIME;ECO-23` (voir correctif d), `COREDIME;9993` (LIVRAISON AVION,
  déjà documenté session F4). **PAS d'entrée 109 DISTRIBUTION** :
  volontaire, son éco-part est une colonne supplémentaire au sein d'une
  ligne normale (`moteur.fournisseurs.dist109._bloc_ligne_facture`), jamais
  une ligne séparée — rien à ajouter tant qu'aucune facture réelle ne
  montre un frais 109 sous forme de ligne à part. Une référence ABSENTE du
  fichier suit le comportement normal quel que soit son montant (jamais de
  seuil "petit montant = ignoré" inventé, décision explicite de cadrage).
  Validé sur données réelles au moment de l'écriture (voir plus bas) : 5
  vraies lignes ECO-23 correctement isolées, jamais bloquantes.
- **(d) BUG RÉEL CORRIGÉ — Coredime, montant éco-taxe faux (0€ au lieu du
  vrai montant)** (`MOTIF_LIGNE_ECOTAXE_COREDIME`, nouveau, dans
  `moteur/fournisseurs/coredime.py`) : ligne réelle
  `" ECO-23    ECOTAXE    *** 10   UN   0,08   0,08   0,80 1"` — DEUX prix
  consécutifs (brut et net, toujours identiques sur les 2 pièces vues,
  jamais de remise observée sur une éco-taxe) là où une ligne normale n'en
  affiche qu'UN SEUL. `MOTIF_LIGNE_FACTURE_COREDIME` matchait quand même
  (la désignation non-greedy absorbe le `"***"`) mais son groupe montant
  (`[\d\s,]+?`, tolère des espaces internes) fusionnait les deux derniers
  nombres ("0,08       0,80") en une chaîne que `_f()` ne convertit pas
  correctement → montant=0,0 au lieu de 0,80. **Silencieux** : le
  garde-fou qté×prix vs montant (`abs(0 - 0.8) > max(0.05*0, 1.0)` =
  `0.8 > 1.0` = Faux) ne se déclenche pas sur un si petit écart absolu.
  Repère fiable, jamais vu sur un article normal : le `"***"` imprimé à la
  place d'une vraie quantité de vente précède TOUJOURS cette structure à 2
  prix. Traité en amont du passage normal (les lignes éco-taxe sont
  isolées et retirées du texte soumis à `MOTIF_LIGNE_FACTURE_COREDIME`
  AVANT qu'il ne s'exécute — jamais une double extraction). Validé sur les
  2 pièces réelles : 6107800.pdf (0,80€, total 55,80€ exact) et
  6100226.pdf (0,40€, total 37,81€ exact) — les DEUX reconciliaient
  exactement leur Total HT après ce correctif, alors qu'AVANT ils
  montraient un écart (55,00€/37,41€ extraits).
- **(e) Cause codée par ligne/anomalie** (`CauseFacture`, nouvel enum à 15
  valeurs dans `matching_facture.py` ; `classifier_cause_anomalie()` dans
  `pipeline_facture.py`) : assignée DIRECTEMENT pour tout ce qui transite
  par une `CorrespondanceFacture` (FRAIS, QTE_PARTIELLE/QTE_SUPERIEURE,
  DOUBLON_FACTURE, CLE_PARTIELLE, REF_INCONNUE, PRIX_DIFF_MEME_REF) ;
  dérivée en BEST-EFFORT du texte des anomalies "fichier entier"
  (anomalies_lecture/anomalies_facture, restées des 2-tuples EXPRÈS —
  les changer de forme aurait cassé `gui_rapprochement_facture.py`, qui
  fait `for facture, raison in rapport.anomalies_facture`). Compteur par
  cause ajouté en fin de rapport texte (`appliquer_et_archiver_factures`).
  **`MOTIF_BDC_MANUEL_24X`** (`est_bdc_manuel_24x()`) : détecte un bon
  manuel "BC/BCN 24XXXX" (carnet papier) à partir d'un nouveau champ
  `Facture.numeros_commande_bruts` (candidat N°Réf.Client/Réf.: BRUT,
  renseigné même quand le parser ne peut PAS le convertir en commande
  exploitable — nouveau côté 109 ET Coredime, ce dernier via
  `_ref_brute_coredime()`) — confirmé sur 3 cas réels : "BC 241766"
  (362777.pdf), "BC 241659" (fixture déjà existante), "BCN 241461"
  (6100226.pdf, label "Réf.:" chez Coredime).
  **BUG RÉEL ÉVITÉ (trouvé en construisant ce correctif, potentiellement
  grave) — `_verifier_total_ht_facture()`** (nouveau, `pipeline_facture.py`) :
  une facture dont AUCUNE ligne n'est extraite (donc rien "d'inconnu" non
  plus) pouvait être considérée "entièrement résolue" par `_est_resolu_facture`
  (0 sur 0 lignes à traiter) et ARCHIVÉE EN SILENCE, perdant toute trace du
  montant manquant — cas réel qui l'a révélé :
  `facture_coredime_6108846_remise_double_x3.pdf`, 0 ligne extraite pour un
  Total HT affiché de 196,92€ (limite déjà connue et acceptée du garde-fou
  remise-double, voir "Points fragiles"). Réconciliation Total HT (même
  seuil 0,02€ que le contrôle console déjà existant côté parser) désormais
  SURFACÉE dans `anomalies_facture` (cause TOTAL_ECART) — ce qui empêche
  l'archivage automatique tant que l'écart n'a pas été vérifié à la main.

**Suite verrouillée par pytest** : 405 passés (372 avant cette session, +33
nouveaux tests couvrant les 5 corrections, dont plusieurs sur les VRAIES
pièces ci-dessus).

**Étape 2 — périmètre du lot RÉDUIT à la demande de l'acheteur, après un
écart chiffré important avec le plan** : le plan attendait 63 factures
Coredime et 114 Cominter (+109 Mayotte) "à vérifier" à remonter à la
racine ; le décompte réel (lecture seule, `rapprocher_dossier_factures`
pointé directement sur `À vérifier/`) donnait 289 Coredime et 40+17
Cominter/Mayotte. Écart expliqué par 234 blocs `commande_absente`
(commandes pré-2026, hors périmètre du Suivi, déjà documenté comme non
résoluble — voir session Coredime H1 2026 ci-dessus) — le plan comptait
vraisemblablement un sous-ensemble "réellement actionnable" déjà filtré
par l'acheteur, pas le total brut du dossier. Face à cet écart, l'acheteur
a choisi (question posée explicitement) de ne remonter QUE les **11
fichiers ayant gagné au moins une ligne sûre** grâce aux corrections du
jour (2 × 109, 6 × Coredime, 3 × Cominter) plutôt que les 379 fichiers des
3 fournisseurs au total — évite de rescanner ~370 fichiers structurellement
non résolubles aujourd'hui.

**Étape 3 — écriture réelle, PUIS incident majeur, PUIS ré-écriture
complète vérifiée** :

- 1re écriture (13h35) : 14 lignes sûres écrites (109×5, COREDIME×6,
  COMINTER×3 sur les 11 fichiers), 4 factures archivées avec leur BC
  (`6100806.pdf`→M3.18.171, `6109361.pdf`→M3.23.043, `Facture_362840.pdf`
  →123.089, `Facture_362846.pdf`→133.022), 7 vers `À vérifier/` (cause
  précisée : 2 QTE_PARTIELLE, 1 QTE_SUPERIEURE, 1 CLE_PARTIELLE, 3
  REF_INCONNUE). Résorption : 109 -5, COREDIME -6, COMINTER -3.
  Vérifiée par relecture directe de 4 cellules — correcte.

**INCIDENT MAJEUR — le Suivi vivant "retourné à hier" (~13h49), toutes les
écritures du jour perdues (les miennes ET une partie de celles de
l'acheteur), cause non identifiée avec certitude.** Chronologie
reconstituée : mon écriture de 13h35 vérifiée correcte immédiatement après
(cellule par cellule) ; le fichier n'a plus bougé jusqu'à 13h49:41, où sa
taille ET sa date de modification changent sans qu'aucune action de ma
part ne l'explique (je n'avais plus rien écrit depuis 13h35) — une
RELECTURE à ce moment-là montre les 14 lignes de N° facture redevenues
vides. L'acheteur, en ouvrant le fichier peu après, le trouve "retourné à
hier" (toutes les commandes du jour disparues, pas seulement mes 14
lignes) — signature typique d'un fichier daté de la veille sauvegardé
PAR-DESSUS le fichier du jour (même famille que l'incident de verrou déjà
documenté en session précédente, mais ici SANS que le verrou Excel n'ait
techniquement bloqué quoi que ce soit — cause exacte non identifiée : pas
d'explication certaine trouvée cette fois, contrairement à l'incident
précédent). **Heureusement, l'acheteur avait fait SA PROPRE sauvegarde
manuelle à 12h20** (avant de me donner accès au fichier ce jour-là),
confirmée par l'acheteur comme ne contenant AUCUNE perte de son propre
travail. Le fichier vivant a été remplacé par cette sauvegarde ; le
fichier "cassé" (retourné à hier) a été archivé dans
`1.3.0.1. Commandes courantes\Archives\1.3.0.1. Suivi commandes - 2026 -
020926.xlsx` (nom reflétant le contenu constaté, daté du 02/09).

**Récupération faite en 2 temps, aucune donnée reperdue** :
1. Les 4 factures déjà "archivées" (copiées dans `a_traiter/BL/Traités/
   <commande>/`, l'écriture Suivi correspondante ayant disparu) ont été
   relocalisées et remises à la racine de `a_traiter/Factures/` sous leur
   nom d'origine — l'archivage de fichiers PDF est un système de fichiers
   INDÉPENDANT du Suivi.xlsx, jamais affecté par cet incident, mais
   laissé dans un état incohérent (PDF rangé comme "traité" alors que le
   Suivi ne le reflétait plus) tant que non corrigé. Les 7 fichiers déjà
   dans `À vérifier/` ont eux aussi été remontés à la racine (règle 8 du
   plan, respectée strictement même en contexte de récupération d'urgence).
2. Une fois l'acheteur ayant confirmé le Suivi refermé (reconfirmé
   EXPLICITEMENT, jamais présumé sur la seule absence d'erreur de verrou
   — voir [[feedback_suivi_verrou_ne_pas_se_fier_seul]]) et le verrou `~$`
   effectivement absent, une lecture 100% FRAÎCHE (jamais les anciens
   numéros de ligne réutilisés) a retrouvé EXACTEMENT le même résultat
   qu'avant l'incident (14 sûres, mêmes lignes Excel, même détail) — la
   sauvegarde 12h20 de l'acheteur contient donc les mêmes données de
   commande sous-jacentes que la version sur laquelle j'avais travaillé
   (logique : ces lignes n'avaient pas changé depuis avant midi). Réécrit
   à 16h39, RE-vérifié cellule par cellule sur les 14 lignes (100%
   correct), fichiers ré-archivés correctement. Nouvelle sauvegarde
   automatique : `backups/1.3.0.1. Suivi commandes - 2026_20260903_163945.xlsx`.

**Leçon retenue pour toute session future, à ajouter aux réflexes déjà
en place** : même un fichier signalé "fermé" et sans verrou Excel actif
peut être écrasé par un événement externe (ici, entre deux écritures,
sans qu'aucune action de la session elle-même ne l'explique) — après
TOUTE interruption significative (autre étape longue, changement de
sujet, incident quelconque) avant de continuer à écrire dans le Suivi,
TOUJOURS relire fraîchement ET reconfirmer explicitement avec l'acheteur
que rien n'a changé entre-temps, jamais supposer qu'un état vérifié
correct il y a une heure l'est toujours. Envisager, pour une session
future, d'ajouter une vérification légère "le fichier a-t-il été modifié
depuis ma dernière lecture ?" (comparaison de taille/date de modif) avant
toute écriture — pas fait cette session (temps), mais aurait raccourci le
délai de détection ici.

## État FNP mensuel (moteur/fnp.py) — clôture comptable

Demande directe de la DAF (31/08/2026, direction en copie) : un état
mensuel des Factures Non Parvenues, reproductible pour un mois passé,
« en un clic ». Deux volets SANS AUCUN RAPPORT entre eux, dans le même
classeur de sortie `rapports/FNP_<AAAA-MM>.xlsx` :

- **(a) BL non facturés** : lignes du Suivi commandes livrées au plus tard
  le dernier jour du mois M mais pas encore facturées.
- **(b) Transitaires en cours de facturation** : dossiers de
  `1.3.0. Suivi commandes spéciales.xlsm` arrivés au plus tard fin de mois M
  mais dont la facture de TRANSPORT n'est pas encore reçue — la facture du
  transitaire (transport/douane) est indépendante de la facture du
  fournisseur de la marchandise elle-même (volet a). Mot de l'acheteur,
  cadrage avant code : *« les n° de commande sont liés aux fournisseurs
  habituels, pas aux transitaires [...] ce tableau n'a d'intérêt que pour la
  partie FNP des transitaires »* — un dossier de Commandes spéciales sans
  N° de commande lié (8 des 31 vus au cadrage) est donc inclus au même titre
  que les autres.

Entièrement LECTURE SEULE : ce module n'écrit JAMAIS ni dans le Suivi
commandes ni dans Commandes spéciales, seulement dans `rapports/`. Chaque
lecture passe par une copie temporaire à usage unique (voir
`_copie_temporaire`, `tempfile.mkstemp` — jamais un nom fixe, voir bug
ci-dessous).

**Lancement** : `py -3 fnp.py <AAAA-MM> [AAAA-MM-JJ]` (2e argument optionnel,
filtre "à partir du") ou bouton GUI **« État FNP du mois »** (5e bouton de
`gui.py`, ouvre `gui_fnp.py`) — mois pré-rempli avec le dernier mois
calendaire complet (`moteur.fnp.mois_precedent_complet`).

### Colonnes du Suivi volontairement PAS réutilisées (vérifié au cadrage)

Avant d'écrire une ligne de code, les formules réelles de 5 colonnes
calculées du Suivi ont été relues sur le vrai classeur (comme pour la
découverte "Statut commande" en session précédente) — AUCUNE ne répond à la
question posée par la DAF :
- **« Reste à facturer »** (`IF(Soldé≠"Soldé", PU prévisionnel × RAL, 0)`) —
  raisonne sur le RELIQUAT NON LIVRÉ, pas sur le livré non facturé.
- **« Potentiel factu »** — projection sur la Qté COMMANDÉE, pas livrée.
- **« Facturé et livré OK »** — `"Good"` si Qté commandée = Qté livrée ET
  Tarif BL ≤ Tarif convenu : un contrôle prix/complétude, pas facture.
- **« Problème »** — `"Problème"` si Qté livrée > 0 ET Facturé BL = 0 : ne
  se déclenche que si NI Tarif BL NI Tarif convenu ne sont connus (le cas
  "aucun prix" du volet a ci-dessous), pas un indicateur de facturation.
- **« Att réc »** — repère l'attente de RÉCEPTION (rien livré du tout), pas
  l'attente de facture.
Le module reconstruit donc le périmètre depuis les colonnes BRUTES (Date de
livraison, Qté livrée, Tarif BL, Tarif convenu, N° facture, Date facture,
Note) — jamais une formule existante réutilisée à tort.

### Volet (a) — périmètre exact (`lire_lignes_bl_non_facturees`)

Une ligne du Suivi qualifie si : `Qté livrée > 0` ET `Date de livraison ≤
fin de mois` (et `≥ depuis` si le filtre optionnel est fourni) ET
`Note ≠ "Commande annulée"` ET (`N° facture` vide OU `Date facture > fin de
mois`). Seule la valeur magique `"Commande annulée"` de Note exclut une
ligne — `"Rupture fournisseur"` et `"Reliquat soldé"` sont juste reportées
telles quelles (périmètre donné par la DAF, ne change rien à
l'inclusion/exclusion).

**Valorisation : réutilise la valeur DÉJÀ CALCULÉE de la colonne « Facturé
BL »** (Tarif BL × Qté livrée, replié sur Tarif convenu) plutôt que de la
recalculer — demande explicite de la DAF/du cadrage. « Source du prix »
(Tarif BL / Tarif convenu / Aucune) reste une étiquette de provenance
calculée à côté, jamais un recalcul du montant lui-même. Une ligne livrée
sans AUCUN prix connu (ni Tarif BL ni Tarif convenu) est **listée à part**
dans l'onglet « BL non facturés » (bloc distinct en bas de feuille, jamais
masquée, jamais fondue dans le total).

**Bloc « antérieures au F1 »** : toute ligne livrée avant
`DATE_CREATION_COLONNES_FACTURE = 2026-09-01` (date de la 1ère écriture
réelle des colonnes facture, voir section "Rapprochement factures —
1ère écriture réelle" ci-dessus) est comptée à part — elle n'a simplement
jamais eu l'occasion d'être marquée facturée par l'outil, peut-être déjà
réglée par ailleurs. **Recette réelle sur août 2026** : ce bloc représente
**100 % du total** (1 167 448,90 € sur 4 621 lignes, TOUT antérieur au F1) —
mécaniquement inévitable pour ce tout premier rapport (le suivi factures
vient de démarrer le jour même). Décision explicite de l'acheteur : garder
le chiffre brut tel quel dans le rapport/l'e-mail (pas de bandeau
d'avertissement renforcé) — le bloc F1 déjà présent dans la Synthèse suffit
à donner le bon niveau de confiance à la DAF. Ce ratio F1/total redeviendra
significatif dès les mois suivants, au fur et à mesure que le rapprochement
factures tourne en routine.

**Filtre "depuis" optionnel** : ne change RIEN au ratio F1 ci-dessus (F1
est une date fixe, `depuis` ne fait que raccourcir la fenêtre analysée) —
n'a donc pas résolu ce problème structurel pour le 1er rapport ; utile
seulement pour ignorer un bruit vraiment ancien une fois le suivi mature.

**Ancienneté** : `(fin_de_mois - Date de livraison).days`, TOUJOURS relative
à la fin du mois M demandé (jamais à "aujourd'hui") — pour que régénérer le
rapport d'un mois passé donne exactement le même résultat à n'importe quel
moment futur (reproductibilité explicitement demandée par la DAF). Buckets
< 30 j / 30-90 j / > 90 j dans la Synthèse.

### Volet (b) — cadrage avec l'acheteur avant code (confirmé 2026-09-01)

Trois points tranchés avant d'écrire le parser de `Commandes spéciales`,
feuille "Suivi" (31 dossiers réels au cadrage) :
1. **« Expédition facturée »** (colonne 0/1) = facture du transitaire déjà
   reçue/traitée (1) ou en attente (0) — confirmé, utilisé tel quel.
2. **« ETA ou arrivée réelle »** (le nom mélange estimé/réel selon les
   dossiers) — confirmé utilisable telle quelle pour le cut-off "arrivé ≤
   fin de mois", malgré l'ambiguïté du nom.
3. **Dossiers sans N° de commande lié** (8/31 au cadrage, ex. achats
   ponctuels via La Poste, conteneurs "Multiples chantiers/fournisseurs") —
   **réponse de l'acheteur, plus nuancée que les 2 options proposées** :
   *« les n° de commande sont liés aux fournisseurs habituels, pas aux
   transitaires, ils n'ont pas de rapport avec les FNP [...] Ce tableau n'a
   d'intérêt que pour la partie FNP des transitaires »* — le lien à une
   commande n'est JAMAIS un critère d'inclusion/exclusion pour ce volet,
   tous les dossiers de Commandes spéciales sont éligibles, identifiés par
   Désignation/Chantier/Fournisseur/N° dossier revient. Conséquence
   d'architecture : le volet (b) est entièrement AUTONOME dans
   `Commandes spéciales` — jamais besoin de rejoindre le N° de commande au
   Suivi principal pour retrouver Chantier/Fournisseur (déjà présents
   directement dans Commandes spéciales).

**Découverte utile en creusant le classeur avant de coder** : les colonnes
« Taux estimé »/« Coût estimé » de la feuille "Suivi" sont déjà
PRÉ-CALCULÉES (Montant commande × taux d'approche moyen du trajet,
vérifié par cohérence arithmétique exacte sur plusieurs dossiers réels,
ex. 55 625 × 0,5557... = 30 911,23) — exactement l'estimation demandée par
la DAF (« marchandise × taux d'approche moyen du trajet, feuille Analyse »)
déjà construite par l'acheteur elle-même dans ce classeur. `moteur/fnp.py`
réutilise donc directement `Coût estimé` (colonne "Coût transitaire
ESTIMÉ" du rapport, clairement étiquetée ESTIMATION) au lieu de recalculer
un taux moyen depuis la feuille "Analyse" (un PivotTable Excel, pas
directement exploitable en lecture par openpyxl de toute façon).

**Repli** (`_repli_transitaires_suivi_principal`) : si `Commandes
spéciales` est introuvable/illisible, lignes du Suivi principal avec
Transitaire renseigné + livrées ≤ fin de mois — **aucune estimation
fabriquée** (`cout_estime` reste `None`, signalé "non calculable" dans le
rapport) : pas de taux d'approche moyen par trajet disponible sans le
classeur dédié, jamais un chiffre inventé (règle d'or).

**Contrôle de couverture** (`_controler_couverture_transitaires`) :
commandes du Suivi principal portant un Transitaire ET un N° dossier
revient renseignés, mais dont ce numéro n'apparaît dans AUCUN dossier de
Commandes spéciales — signalées à part dans la Synthèse (jamais comptées
dans le total, jamais un rapprochement inventé sans clé fiable commune).
**Rendement réel mesuré très faible** : sur 61 lignes du Suivi principal
avec Transitaire renseigné, UNE SEULE a aussi un N° dossier revient rempli
— l'acheteuse remplit "Transitaire" bien plus souvent que "N° dossier
revient" sur le Suivi principal. Le contrôle reste honnête (ne signale
jamais un faux gap sur les lignes sans N° dossier revient, faute de clé) et
peu coûteux à calculer (collecté dans la même passe de lecture que le
volet a, pas de 2e lecture du classeur) mais n'attrape aujourd'hui presque
rien en pratique — la mise en garde générale "Commandes spéciales est peu
alimenté" (déjà dans CLAUDE.md) reste la protection principale.

**Idée non construite, soulevée par l'acheteur pendant le cadrage** :
ajouter une colonne « Montant du devis validé » à `Commandes spéciales`
pour suivre les sommes transitaire À VENIR (engagement), pas seulement les
FNP déjà arrivées. Distincte du besoin FNP (qui ne regarde que ce qui est
déjà arrivé) — notée pour une session future, hors périmètre de celle-ci.

### Bugs réels trouvés en écrivant les tests (`tests/test_fnp.py`)

1. **Collision de fichier temporaire Windows** : `_copie_temporaire`
   utilisait d'abord un nom FIXE (`fnp_lecture_<nom_du_fichier>.xlsx` dans
   le dossier temp système) — deux lectures utilisant un fichier source du
   même nom (ex. plusieurs tests avec tous une "suivi.xlsx") se
   marchaient dessus : la 2e copie tentait d'écraser la 1re pendant qu'un
   handle Windows la tenait encore ouverte (`PermissionError`), et le
   `tmp.unlink()` de fin de lecture échouait pour la même raison même en
   usage normal (`wb.close()` sur un classeur `read_only=True` ne libère
   pas toujours IMMÉDIATEMENT le verrou Windows). Corrigé : `_copie_temporaire`
   est maintenant un context manager (`with ... as tmp:`), nom UNIQUE via
   `tempfile.mkstemp()` à chaque appel, et nettoyage tolérant
   (`try/except OSError` sur l'`unlink()` final — un verrou résiduel sur un
   fichier temporaire jetable ne doit jamais faire échouer la lecture).
2. **Désynchronisation `.append()` / `.cell()` sur une même feuille
   openpyxl** : le bloc "lignes sans prix" de l'onglet "BL non facturés"
   mélangeait `ws.cell(ligne_titre, ...)` (adressage direct) et
   `ws.append([])` (compteur interne `_current_row`, indépendant de
   `max_row` tant qu'aucune vraie cellule n'est écrite) — un `.cell()`
   intercalé pour "sauter" des lignes désynchronise les deux compteurs, la
   ligne suivante écrite via `.append()` pouvait retomber sur une ligne
   déjà posée à la main et l'écraser. Corrigé : tout le bloc en `.append()`
   séquentiel strict (y compris une ligne `[""]` pour le séparateur, JAMAIS
   `[]` vide — un append vide n'écrit aucune cellule donc ne fait PAS
   avancer `max_row`), le style (gras/fond) appliqué après coup sur des
   cellules déjà écrites, jamais pour en créer. **Leçon générale pour tout
   futur code openpyxl de ce projet** : ne jamais mélanger `.append()` et
   `.cell()` pour CRÉER des lignes sur la même feuille.
3. **Un N° facture sans date lisible n'excluait pas la ligne** : la
   condition d'exclusion exigeait `date_facture is not None ET ≤ fin de
   mois` — si `date_facture` était `None` (date illisible, cas réel
   possible côté `pipeline_facture.ecritures_pour_facture`, qui n'écrit la
   date QUE si elle a pu être parsée, contrairement au N° facture toujours
   écrit), la ligne n'était PAS exclue et ressortait comme FNP alors
   qu'une facture existe bel et bien. Corrigé : un N° facture présent
   exclut par défaut (`date_facture is None OU ≤ fin de mois`) — seule une
   date facture CONFIRMÉE après la clôture fait rester la ligne une FNP
   malgré un N° facture déjà renseigné.

### Sortie et brouillon Outlook

`rapports/FNP_<AAAA-MM>.xlsx` — 3 onglets : **Synthèse** (totaux,
répartitions fournisseur/chantier/ancienneté, fiabilité des deux volets),
**BL non facturés** (détail trié Fournisseur puis Montant décroissant, bloc
"sans prix" à part), **Transitaires** (détail, colonnes différentes selon
repli ou non). Mise en forme réutilise `moteur.excel._entete`/`_largeurs`
(mêmes helpers que Comparatif.xlsx, jamais une 2e version dupliquée).

**Brouillon Outlook** (`moteur/fnp_brouillon.py`, `creer_brouillon_fnp`) —
même garde-fou que `creer_brouillons.py` (pipeline Hermes) : `mail.Save()`
UNIQUEMENT, jamais `Send()`. Destinataire(s)/copie saisis à la main dans le
GUI (jamais une adresse devinée ou codée en dur) ; pièce jointe = le
classeur généré.

**BUG RÉEL CORRIGÉ (trouvé en générant le rapport réel d'août, avant tout
envoi) — `win32com.client` NE DOIT PAS être ajouté à
`moteur/dependances.py::REQUIS`** : premier réflexe, ajouté pour que
l'auto-installation couvre aussi ce besoin — `pip install pywin32` a bien
téléchargé le paquet, mais `import win32com.client` a continué à échouer
juste après ("toujours introuvable après installation") : pywin32 a besoin
d'une étape de POST-installation (enregistrement de pywintypes/pythoncom)
qu'un simple `pip install` ne fait pas. Comme `gui.py`/`main.py`/`fnp.py`
appellent tous `verifier_et_installer()` AU DÉMARRAGE, ça bloquait le GUI
ENTIER (comparatif/panier/BL compris) à cause d'un besoin qui ne concerne
que le bouton optionnel "Créer le brouillon". Retiré de `REQUIS` — l'import
de `win32com.client` reste LOCAL, à l'intérieur de `creer_brouillon_fnp()`
(et déjà ainsi côté Hermes), jamais en tête de module : si pywin32 est
absent/mal installé, seule l'action "Créer le brouillon" échoue avec un
message clair, jamais le reste de l'outil.

**Volontairement PAS envoyé pour de vrai cette session, sur demande
explicite de l'acheteur** (après avoir donné l'adresse de la DAF pour
plus tard) : *« on envoie rien du tout pour le moment, il faudra parser
tous les fournisseurs d'abord »* — seul 109 Distribution est couvert côté
rapprochement factures à ce jour (voir plus haut), présenter un état FNP
à la DAF maintenant serait prématuré. **Adresse de la DAF pour quand ce
sera le moment : `daf@espace-soleil.re`** (direction en copie, à demander
le moment venu). Ne PAS créer/proposer de brouillon FNP réel avant que
davantage de fournisseurs soient couverts côté factures — redemander
explicitement à l'acheteur si ce seuil est atteint, ne jamais décider seul
que "c'est bon".

## État FNP — v1.1 (session S0, 2026-09-03) : exclusion, ajustements,
réserves, 1er envoi réel du brouillon DAF

Suite directe de la 1ère version (ci-dessus) — le seuil "davantage de
fournisseurs couverts" est désormais atteint (109, Coredime, Cominter,
Cominter Mayotte, Electric Plus tous couverts côté factures), l'acheteur a
explicitement demandé cette session de construire 3 extensions SANS
changer le périmètre déjà en place, puis de créer le brouillon réel.

**4a — Exclusion "facture reçue non rapprochée"**
(`_identifier_lignes_excel_facturees()`/`_appliquer_exclusion_factures_recues()`,
nouveau `RapportFNP.factures_recues_non_rapprochees`) : scanne
`a_traiter/Factures/` (racine ET `À vérifier/`, LECTURE SEULE via
`rapprocher_dossier_factures` — même mécanisme déjà éprouvé, aucune
modification) et retire du volet (a) toute ligne Suivi pour laquelle une
VRAIE facture PDF existe déjà (datée ≤ fin de mois) mais n'a pas encore
été rapprochée dans l'outil — ce ne sont PLUS des factures "non
parvenues", juste pas encore traitées. Nouvel onglet "Factures reçues"
dédié. Option `--sans-exclusion` (CLI `fnp.py`) / `appliquer_exclusion=
False` (API) en repli si le scan est trop long — jamais utilisée cette
session, 4a est restée "verte" toute la session.
**Piège de perf évité avant même d'écrire le 1er test** : le test déjà
existant `test_calculer_rapport_fnp_sur_le_vrai_suivi` appelait
`calculer_rapport_fnp(ROOT, "2026-08")` sans préciser `appliquer_exclusion`
— avec la 4a active par défaut, CE seul test aurait scanné le VRAI
`a_traiter/Factures/` (des centaines de PDF réels, OCR compris) à CHAQUE
lancement de la suite. Corrigé en passant `appliquer_exclusion=False`
explicitement à ce test précis (son objet reste la vérification des
en-têtes, pas la 4a) — sans ce correctif, `py -3 -m pytest` serait passé
de ~20 min à ~40+ min à chaque lancement.

**4b — Ajustements déclarés par l'acheteur**
(`referentiel/fnp_ajustements_<mois>.csv`, nouveau, un fichier PAR MOIS ;
`lire_ajustements_fnp()`, `AjustementFNP`) : colonnes type
(BDC_MANUEL/TRANSIT/AUTRE) ; libelle ; fournisseur_ou_transitaire ;
chantier ; piece ; date_livraison ; montant_ht ; source ; commentaire.
JAMAIS fusionné avec les totaux calculés (a)/(b) — onglet "Déclaré (hors
outil)" à part, son propre total dans la Synthèse. Template
`referentiel/fnp_ajustements_2026-08.csv` créé (en-têtes seules) ;
l'acheteur, sollicité explicitement, n'avait rien à déclarer pour l'instant
— 0 ligne ce mois-ci (état normal, pas une erreur).

**4c — Réserves de périmètre** (`ReservesFNP`, nouveau) : nombre de
factures `BDC_MANUEL_24X` (matériel livré sur bon manuel, hors Suivi) —
compté dans la MÊME passe que la 4a, jamais un 2e scan ; nombre de
dossiers transitaires sans estimation de coût (déjà connu du volet b,
répété ici) ; rappel textuel fixe "0 ligne migrée sans pièce" (fonctionnalité
"Pièces" pas encore construite) ; nombre TOTAL de dossiers Commandes
spéciales (`compter_dossiers_speciales()`, nouveau — tout statut confondu,
pas seulement les non-facturés du volet b) pour rappeler l'ampleur réelle
de ce classeur peu alimenté.

**Suite verrouillée par pytest** : 422 passés (405 avant cette sous-session,
+17 nouveaux tests 4a/4b/4c + fnp_brouillon).

**Corps du brouillon Outlook enrichi** (`moteur/fnp_brouillon.py`,
`_corps_mail_fnp()`/`_euro_fr()` nouveaux) : la DAF doit pouvoir lire les
totaux SANS ouvrir la pièce jointe — total volet (a), total ESTIMÉ volet
(b), total déclaré hors outil (si non nul), nombre de factures reçues non
rapprochées exclues (si non nul), réserves de périmètre en clair. Séparateur
de milliers en espace INSÉCABLE (convention typographique française
correcte). `creer_brouillon_fnp()` prend désormais le `RapportFNP` complet
en paramètre (plus seulement `mois_en_lettres`) — signature changée,
`gui_fnp.py` mis à jour en conséquence (seul appelant).

**1re exécution réelle de bout en bout (2026-09-03, APRÈS l'incident Suivi
et sa récupération complète — voir section précédente)** :

| | Montant | Lignes |
|---|---|---|
| Volet (a) — BL non facturés | 943 753,55 € | 3507 (dont 28 sans prix) |
| Factures reçues non rapprochées (exclues du a) | 42 274,71 € | 221 |
| Volet (b) — Transitaires, coût estimé | 31 512,60 € | 2 dossiers |
| Déclaré (hors outil) | 0,00 € | 0 |

Réserves : 9 factures BDC_MANUEL_24X, 0 dossier transitaire sans
estimation, 31 dossiers Commandes spéciales au total (cohérent avec le
chiffre "31 dossiers" déjà cité au cadrage F1).

**Baseline (avant cette session) → nouveau total volet (a) : 1 141 876,24 €
→ 943 753,55 € = -198 122,69 € (-17,4 %).** Condition du plan remplie (le
total devait baisser). Attribution partielle honnête, PAS de réconciliation
forcée (règle du plan, "ne pas corriger") : ~46 000 € directement
attribuables à cette session (≈3 900 € des 14 lignes écrites à l'étape 3 +
42 275 € de la 4a) ; le reste (~152 000 €) n'a pas été reconstitué avec
certitude — plausiblement de la résorption des sessions précédentes (le
gros lot Coredime du 02/09) pas encore reflétée dans le chiffre de
référence du plan, mais non confirmé.

**Brouillon Outlook créé** (Save uniquement, jamais Send — voir bandeau du
module) : destinataire `daf@espace-soleil.re`, copie `direction@
espace-soleil.re` (adresses confirmées explicitement par l'acheteur dans
le chat, jamais devinées), pièce jointe `rapports/FNP_2026-08.xlsx`.
**C'est l'acheteur qui relit et envoie depuis Outlook** — jamais envoyé par
la session, conformément à la règle absolue de ce module.

## Rapprochement factures — P0 : résiduel unique, remise double par
coordonnées, OCR Stand 64, recette (2026-09-04)

Suite directe de S0, plan en 4 étapes exécutées dans l'ordre.

**Étape 1 — Résiduel unique ("substitution probable")**
(`moteur/rapprochement/matching_facture.py::_residuel_unique`, appelé en
tout dernier repli dans `apparier_facture()`) : dans une commande, s'il ne
reste plus qu'UNE SEULE ligne facture inconnue et qu'UNE SEULE ligne Suivi
sans facture, même quantité (et même PU à 0,02€ près si connu) ->
`StatutFacture.A_CONFIRMER`/`CauseFacture.SUBSTITUTION_PROBABLE` — jamais
"sûr" (aucune ressemblance textuelle, un pur processus d'élimination).
**Garde-fou décisif ajouté après coup** : ne se déclenche QUE si la
commande avait PLUSIEURS lignes facture au départ (`len(lignes_a_apparier)
> 1`) — sans lui, plusieurs tests synthétiques déjà existants (1 seule
ligne de chaque côté, qté/PU par défaut coïncidemment identiques)
auraient été promus à tort en "substitution probable" ; une commande à 1
seule ligne ne prouve rien par élimination, contrairement à une commande
à 9 lignes où 8 se sont résolues normalement.

Cas réel qui a servi de test (`facture_coredime_6108234_suffixe.pdf`,
commande M3.14.342) : facture "LEG06620" (ICTA 3422 20 ATF, 100m, 0,37€)
face à la ligne Suivi "5120" (ICT 20 BLEU TURBO, même quantité/tarif,
AUCUN rapport textuel) — exactement le type de substitution jusqu'ici
seulement découvert à la main (voir CFF1BIS/59210, 411651/092897 côté BL).

**Apprentissage séparé de la table alias** (`Referentiel.
apprendre_equivalence()`, nouveau, `moteur/referentiel.py`) : une fois
confirmée, la paire est ajoutée à `referentiel/equivalences_bl.csv` (le
ledger PORTABLE partagé BL+Facture, suivi par git) — PAS à la table
`alias` de `moteur/articles.db` (SQLite, gitignorée, propre à ce poste)
comme le ferait le mécanisme référentiel classique. Justification : une
substitution pure, découverte par élimination, ne peut structurellement
JAMAIS être re-proposée automatiquement si sa trace est perdue — mérite
une persistance durable, pas un cache local. Nouvelle feuille
**"Substitutions probables"** dans `referentiel/A_confirmer_Facture.xlsx`
(`pipeline_facture._ecrire_substitutions_probables`/
`_appliquer_confirmations_substitutions`, appelées respectivement en fin
et en début de `rapprocher_dossier_factures()`) : feuille à part de la
feuille "À confirmer" du référentiel (round-trip openpyxl, jamais touchée)
— ces propositions ne viennent pas de `Referentiel.resoudre()` (aucune
ressemblance structurelle à faire apparaître dans `_propositions`).
Décision "OUI" -> `equivalences_bl.csv` ; vide/"NON" -> pas de rejet
définitif (la proposition n'est pas mémorisée comme fausse pour toujours,
contrairement au workflow référentiel classique — un résiduel unique n'a
pas vocation à ça). **Recette réelle : 18 substitutions probables trouvées
sur le lot complet**, feuille générée et vérifiée dans le vrai
`A_confirmer_Facture.xlsx`.

**Étape 2 — Remise double Coredime par coordonnées**
(`moteur/grille.py`, nouveau module : `mots(page)` =
`page.get_text("words")`, `lignes(mots)` regroupées par Y avec tolérance =
demi-hauteur MÉDIANE des mots de la page — pas une constante comme
`moteur.ocr.regrouper_lignes`, un PDF natif est en points, pas en pixels à
un DPI fixe). `moteur.fournisseurs.coredime._apparier_par_position_coredime`
remplace l'ancien appariement 1:1 de `_lignes_remise_double_coredime`
(qui n'osait un rattachement que si le bloc contenait EXACTEMENT 1 ligne
incomplète et 1 ligne "Remise") : chaque ligne "Remise" est désormais
rattachée à l'article dont la position est immédiatement AU-DESSUS
d'elle, quel que soit leur nombre. `parse_facture_coredime(texte,
chemin=None)` : nouveau paramètre optionnel `chemin`, transmis
automatiquement en production par `parsers_facture.parser_facture()` (via
`inspect.signature`, seulement aux parsers qui le déclarent — aucun autre
fournisseur touché) ; sans lui (tests unitaires sur texte synthétique),
repli sur l'ancien comportement 1:1, jamais un plantage.

**BUG RÉEL TROUVÉ ET CORRIGÉ pendant la construction, avant tout dégât** :
un premier jet triait les positions par le seul Y brut, en mélangeant des
lignes de PAGES DIFFÉRENTES — or chaque page a son propre système de
coordonnées (Y redémarre près de 0 en haut de CHAQUE page). Sur un
document réel à 2 pages (`facture_coredime_8_remises_multiples_partiel.pdf`),
deux lignes de pages différentes partageaient exactement le même Y — n'a
PAS faussé ce cas précis par chance d'ordonnancement, mais ce n'était pas
garanti. Corrigé : `_position_page_coredime()` retourne un tuple
`(page, y)`, comparé lexicographiquement (page d'abord, Y ensuite) —
`lignes_grille_bloc` porte désormais `(page, ligne)`, jamais une ligne
nue. `grille.position_y()` (version sans page, devenue inutile une fois
ce correctif fait) a été retirée pour ne pas laisser de code mort.

**Validé sur les 14 fixtures facture Coredime** (`_parser()` du fichier de
test passe désormais systématiquement `chemin=`) : 13 identiques, 1
strictement améliorée avec un nouveau résidu HONNÊTE d'une nature
différente — `facture_coredime_8_remises_multiples_partiel.pdf` passe de
21 lignes/1008,14€ (5 lignes ambiguës jamais extraites, écart -124,37€
connu) à **26 lignes/1167,53€** : les 5 nouvelles lignes sont chacune
vérifiées par DEUX signaux indépendants (prix de base × remises en
cascade reconstitue exactement le P.U Net imprimé, ET qté × P.U Net =
montant) — mais la somme totale dépasse maintenant le Total HT AFFICHÉ
sur le document (1132,51€, écart +35,02€) alors qu'avant elle était en
dessous. Vérifié que ce total affiché n'est pas un artefact de lecture
("COR F6108047" isolé une seule fois dans tout le texte, HT+TVA≈TTC
cohérents) : l'écart est donc une incohérence PROPRE à ce document réel
(pas un bug du rattachement par coordonnées, chaque ligne étant
individuellement exacte) — signalé honnêtement dans le test, jamais deviné
ni masqué.

**Nouveau test dédié** `tests/test_grille.py` (3 tests, sur PDF réel) —
verrouille que `lignes()` reconstruit l'ordre visuel exact d'un document
scramblé par PyMuPDF.

**Étape 3 — OCR facture affiné + gabarit Stand 64 scanné**

`moteur.lecture_pdf.longueur_texte_premiere_page()` (nouveau) +
`moteur.rapprochement.lecture_facture.SEUIL_TEXTE_NATIF_PAGE1 = 20` :
le repli OCR générique (session F4/Electric Plus) se déclenchait sur
`lire_pdf(chemin).strip()` VIDE sur le DOCUMENT ENTIER — affiné pour
tester la seule PREMIÈRE PAGE, avec un seuil de 20 caractères plutôt que
zéro strict. Motivation : un scan peut porter quelques caractères de
texte natif ailleurs dans le même fichier (annexe BdC/devis, horodatage)
sans que sa propre page en ait, ce qui aurait masqué à tort le besoin
d'OCR sur la page utile.

**`moteur/fournisseurs/stand64.py::parse_facture_stand64_ocr`** (nouveau,
découvert automatiquement via `parse_facture_ocr`) : 1re pièce facture
Stand 64 SCANNÉE (`FA 33330.pdf`, un des 4 fichiers "FA <numéro>.pdf"
identifiés sans parser depuis la session Electric Plus). Réutilise le
MÊME repérage de tableau que le BL de ce fournisseur
(`_zone_tableau_bl_stand64`, `MOTIF_ENTETE_TABLEAU_BL_STAND64`/
`MOTIF_PIED_TABLEAU_BL_STAND64`) et le même motif de commande
(`MOTIF_COMMANDE_BL_STAND64`), mais une structure de ligne DIFFÉRENTE :
sur l'unique pièce réelle disponible, ni Eco-part ni code TVA ne sont
visibles en fin de ligne (contrairement au BL) — seulement 5 valeurs
après la désignation (Qté, P.U, Rem%, P.U Net, Total HT), vérifiées par
qté × P.U Net ≈ Total HT. Le fichier contient facture (page 0) ET bon de
livraison du même fournisseur (page 1, même n° 43972) scannés ensemble —
seule la page 0 est traitée, sinon les 2 lignes seraient comptées deux
fois. Détection : `detecter_fournisseur` reconnaissait DÉJÀ "STAND 64" sur
texte OCR sans aucune modification (le logo "STAND64" ressort déformé en
"STANDE"/"STANDM" mais "stand64@stand64.fr"/"www.stand64.fr" suffisent) —
verrouillé par un nouveau test dédié. **Validé du premier coup** : 2
lignes, 1 545,00€, commande M3.18.217, BL 43972 — exactement l'attendu.

**Étape 4 — Recette sur le lot complet**

Les 458 fichiers de `a_traiter/Factures/À vérifier/` remontés à la racine
(règle 0.4.8) — root passé à 459 fichiers PDF (+ 1 déjà présent). Lecture
seule complète (`rapprocher_dossier_factures`) : **8 sûres (1 783,19€),
285 à confirmer (dont les 18 substitutions probables), 200 inconnues
(surtout référence introuvable), 21 frais, 76 anomalies de lecture
(fournisseur inconnu 20, parser absent 34 — Cominter/EDOI scannés sans
parser facture, 0 ligne 22, écart total 10), 258 anomalies facture
(commande absente 234 — très majoritairement des commandes pré-2026 hors
périmètre du Suivi, déjà documenté ; bon manuel 9 ; avoir 5).**

**Deux coupures réseau transitoires pendant cette étape** (même symptôme
déjà documenté ailleurs dans ce fichier — lecteur X: temporairement
inaccessible, TOUS les outils shell y compris PowerShell échouant en
silence le temps de la coupure) :
1. Une 1re tentative de diagnostic a crashé (`FileNotFoundError: WinError
   53`) en pleine lecture — après avoir déjà lu les 459 PDF avec succès,
   juste au moment d'ouvrir le référentiel. Aucune perte : relecture
   propre après reconnexion.
2. Une 2e coupure, plus sévère (tous les outils shell, PAS seulement
   l'accès à X:, ont échoué le temps de la fenêtre), a interrompu la
   PREMIÈRE tentative d'écriture réelle — le process s'est arrêté sans
   traceback visible (probablement tué plutôt qu'une exception propre),
   pendant la phase de lecture de `rapprocher_dossier_factures()` (avant
   tout accès en écriture). **Vérifié de façon DÉFINITIVE avant de
   retenter** : aucune nouvelle sauvegarde dans `backups/` à la date/heure
   de l'incident (la dernière datait de la session S0 la veille) — preuve
   que `ecriture.appliquer()` (qui sauvegarde TOUJOURS en tout premier)
   n'a jamais été atteint. Classeur vivant confirmé intact, aucun fichier
   déplacé dans `a_traiter/Factures/`. Reconfirmé "Suivi fermé" par
   l'acheteur avant de retenter (jamais réutilisé une confirmation
   antérieure à un incident, conformément à la leçon déjà retenue —
   [[feedback_suivi_verrou_ne_pas_se_fier_seul]]).

**Écriture réelle confirmée** (relecture fraîche juste avant, comptes
identiques : 8 sûres/1 783,19€) : sauvegarde
`backups/1.3.0.1. Suivi commandes - 2026_20260904_103012.xlsx`, **8
lignes écrites**, **13 factures archivées** avec leur BC dans
`Traités/<commande>/` (dont `6108846.pdf`/`6108847.pdf` — les vraies
pièces Coredime remise-double de l'étape 2, "Tarif BL" renseigné depuis la
facture pour LEG031916/LEG031919/LEG031955 — et `FA 33330.pdf` — la pièce
Stand 64 OCR de l'étape 3 —, confirmant les deux fonctionnalités bout en
bout sur le classeur vivant, pas seulement sur fixture), 0 échec
d'archivage, 412 factures restées en `À vérifier/` (décision humaine).
Résorption après écriture : 109 DISTRIBUTION 805/1001 encore à facturer,
COMINTER 1225/1415, COREDIME 411/865, ELECTRIC PLUS 327/393, STAND 64
71/102.

**Reste à faire** : les 285 lignes "à confirmer" (18 substitutions
probables en tête, feuille dédiée déjà prête dans
`A_confirmer_Facture.xlsx`) et les 200 "inconnues" du lot, à trier avec
l'acheteur ; les 412 factures en `À vérifier/`. Parsers facture toujours
manquants : Cominter (scan, 34 fichiers de ce lot), EDOI (2 fichiers,
toujours en attente de plus de pièces réelles, voir session F4).

## Feuille Pièces — modèle, socles, migration (P1, 2026-09-04)

Session P1 du dossier « Pièces » (règles de conduite 0.4 : un seul plan
exécuté dans l'ordre, tout prouvé sur une COPIE avant le vivant, chaque
étape vérifiée PAR L'OUTIL). Vendredi soir, acheteur parti après « tu as
full access sur les tableaux, rien n'est ouvert » : les contrôles
[HUMAIN] « aucun message de réparation » ont été faits par l'outil avec
Excel lui-même (voir `verification_excel.py` ci-dessous), l'ouverture
humaine du vivant reste à faire lundi (voir « Reste à faire »).

### Modèle

Nouvelle feuille **« Pièces »** (onglet juste après Commandes), tableau
structuré **`Pieces`**, UNE ligne = UNE ligne de DOCUMENT (BL / Retour /
Facture / Avoir / Frais / Demande d'avoir), 26 colonnes dans l'ordre du
plan (`moteur/rapprochement/pieces.py::COLONNES_PIECES`) : ID pièce, Type,
Fournisseur (nom Suivi via `MAPPING_FOURNISSEURS`), N° pièce, Date pièce,
N° de commande, Chantier, Sous-Chantier (copiés de Commandes à l'écriture),
Référence Suivi (valeur EXACTE de Commandes, type conservé — un nombre reste
un nombre, sinon les formules `=` de Commandes ne matchent pas), Référence
fournisseur, Désignation, Qté (signée), PU HT, Montant HT, N° BL lié,
N° facture liée, Prix de référence / Source prix / Écart PU / Écart ligne €
/ Contrôle prix (vides, P2), Mode de rapprochement (Auto / Confirmé /
Équivalence / Migré / Migré sans PDF), Demande d'avoir (P4), Fichier
(`=HYPERLINK("<chemin archive>","<nom>")`, une formule — jamais une
relation OOXML), Date d'écriture, Commentaire (seule colonne humaine).
**Écrite par l'outil, jamais saisie.** `ID pièce` =
`<Fournisseur>|<Type>|<N° pièce>|<N° de commande>|<Référence Suivi>|<N° BL lié>`
(Frais sans ligne Suivi : `F:<réf fournisseur>` à la place de la référence
Suivi ; doublons stricts au sein d'un lot suffixés `#2`, `#3`… de façon
déterministe — `dedoublonner_ids`, JAMAIS par rapport aux ID déjà écrits :
un ID présent EST la même ligne, c'est l'idempotence).

Dans Commandes, les 5 colonnes facture (51–55) sont devenues CALCULÉES
(`FORMULES_COMMANDES_BASCULE`, formules anglaises stockées avec leurs
préfixes `_xlfn.`/`_xlws.`, références structurées en forme longue
`Commandes[[#This Row],[X]]`, la même chaîne sur toutes les lignes) :
Qté facturée / Montant facturé HT = `SUMIFS` sur Pieces (Facture + Avoir,
critères N° de commande + Référence Suivi), N° facture =
`TEXTJOIN("; ",TRUE,UNIQUE(FILTER(…)))` (formule matricielle, `cm="1"`
comme Excel l'écrit lui-même dans ce classeur), Date facture = `MAXIFS`
(vide si 0 — sinon Excel afficherait 00/01/1900, seul écart au sketch du
plan), PU facturé = `IFERROR(Montant/Qté,"")` = moyen pondéré. 4 colonnes
NOUVELLES en fin de tableau (56–59, `FORMULES_COMMANDES_NOUVELLES`) : Reste
à facturer HT, Écart facture €, Qté retournée, Statut facture (vide /
🔵 En attente facture / 🟠 Partiellement facturée / ✅ Facturée /
⛔ Sur-facturée = Qté facturée > Qté livrée). « Statut commande » n'a pas
été touché. Limite structurelle à connaître : les formules agrègent par
(N° de commande, Référence) — deux lignes Commandes portant le MÊME couple
verraient toutes deux le total (aucun cas parmi les 1 059 lignes facturées
le 2026-09-04) ; et SUMIFS coerce les critères numériques (« 06620 » et
« 6620 » seraient confondus dans une même commande — aucun cas non plus).

### Socles génériques (`moteur/rapprochement/ecriture.py`, fin de fichier)

Tous par patch XML du zip (jamais `openpyxl.save()`), paramétrés par le
chemin du classeur pour resservir tels quels à M1 (registre BdC manuels)
et T1 (Tableau1 du .xlsm) :

1. **`ajouter_feuille_tableau(fichier, nom_feuille, nom_table, colonnes,
   dossier_backups, feuille_modele="Commandes", apres_feuille=None,
   largeurs=None)`** : nouvelle part `xl/worksheets/sheetN.xml` (en-têtes
   inlineStr, style d'en-tête copié de A1 du modèle, volet figé sur la
   ligne 1), `sheetN.xml.rels`, `xl/tables/tableN.xml` (id/numéros libres,
   `tableStyleInfo` copié du tableau modèle, **ref `A1:<fin>2`** — Excel
   exige au moins UNE ligne de données sous l'en-tête, un ref `A1:Z1` n'a
   pas été tenté), `<sheet>` inséré juste après le modèle dans
   `workbook.xml` (`localSheetId`/`activeTab` décalés), `workbook.xml.rels`,
   `[Content_Types].xml`. `_patcher_parties_xlsx` sait désormais AJOUTER une
   partie et en SUPPRIMER une (valeur None).
2. **`ajouter_lignes_tableau(fichier, feuille, table, lignes, backups,
   colonne_id=None, styles_colonnes=None)`** : `<row>` appendues en fin de
   `<sheetData>`, textes inlineStr (sharedStrings jamais touché), nombres
   `<v>`, dates en nombre de série avec un style logique (`_STYLES_LOGIQUES`
   : date 14, datetime 22, monnaie 44, monnaie4 = numFmt 164 du Suivi ;
   l'`<xf>` manquant est ajouté en fin de `cellXfs`, jamais de réindexation),
   `Formule(texte, cache, array)` en `<f>` (+ `<v>` cache : le libellé du
   HYPERLINK s'affiche avant tout recalcul). `<dimension>`, `ref` du tableau
   et de l'autoFilter étendus. **Idempotence par `colonne_id`** : les
   valeurs existantes sont relues dans le XML brut — inlineStr ET chaînes
   partagées (Excel réécrit les inlineStr en sharedStrings à sa prochaine
   sauvegarde, testé en simulant cette réécriture).
3. **`ajouter_colonnes_calculees()` / `basculer_colonnes_en_formules()`**
   (`_appliquer_formules_colonnes`) : `<calculatedColumnFormula>` (attribut
   `array="1"` pour une formule matricielle) dans la définition du tableau
   + la même `<f>` sur CHAQUE ligne du tableau, en UNE SEULE passe sur
   `<sheetData>` (`_reecrire_lignes` — la feuille Commandes pèse 58 Mo, une
   recherche par ligne serait quadratique ; ~6 s pour 6 556 lignes), style
   de cellule existant conservé, `spans` ajustés. **calcChain jamais
   complété à la main** : `xl/calcChain.xml` supprimé proprement (partie +
   relation + Override) et `fullCalcOnLoad="1"` posé sur `<calcPr>` — Excel
   recalcule tout et régénère la chaîne à l'ouverture (vérifié : ouverture
   sans réparation, formules calculées).
4. **`reparer_formules_ligne(fichier, backups, colonnes, ligne=2,
   ligne_modele=3)`** : recopie dans les cellules LISTÉES de la ligne 2 la
   formule de la ligne 3, références relatives décalées (`H3` -> `H2`,
   `_decaler_references_ligne`), formules matricielles conservées, refus si
   la cellule modèle n'a pas de formule ou est une formule partagée.
   Constat sur le vivant : le plan citait les colonnes 25–30 (Y..AD) ; les
   colonnes **22 à 30** (V..AD : Reliquat, RAL, Soldé, PU prévisionnel,
   Prix ligne prévisionnel, Montant total commande, Reste à facturer, Délai
   de livraison, Ecart tarifaire) portaient TOUTES la formule « Statut
   commande » recopiée avec des références décalées. Les 9 ont été réparées
   (`installation_pieces.COLONNES_LIGNE2_CONSTATEES`) — décision par défaut
   de [HUMAIN-P1-a], acheteur absente ; conséquence chiffrée : la ligne 2
   (coupe-câble 218,54 €) compte désormais dans « Engagé total » /
   « Reste à recevoir » / « Attente livraison » du Dashboard (+218,54 €
   chacun, valeur correcte de la ligne 2 auparavant perdue).

**`moteur/rapprochement/verification_excel.py`** (nouveau, sans dépendance
au moteur) : `verifier_dans_excel(copie, recalculer, cellules,
enregistrer_sous)` — Excel COM invisible (pywin32, Excel 16 sur ce poste),
détection d'une réparation par les journaux `%TEMP%\error*.xml`,
`CalculateFullRebuild`, lecture feuilles/tableaux/cellules, copie
RECALCULÉE en `SaveAs` (relisible avec openpyxl `data_only` = la « copie
ouverte-fermée » du plan). Toujours sur une COPIE, jamais le vivant. Piège
constaté : le cache du fichier d'origine était PÉRIMÉ (B9 58 -> 61, E9
1 272,05 -> 1 316,80, K9 177 -> 174 au simple recalcul de l'original
intact) — comparer les KPI à l'original RECALCULÉ, pas au cache.

### Migration (`moteur/rapprochement/migration_pieces.py`, étape 4)

`migrer_factures_vers_pieces(fichier, dossier_projet, backups, limite=None)`
: pour chaque ligne Commandes avec N° facture (source de vérité = les 5
colonnes + les PDF archivés), le PDF est retrouvé (`Traités/<commande>/…
- Facture <n°> - BC …`, sinon `a_traiter/Factures/À vérifier/` et la racine
par n° dans le nom — 85 des 363 factures n'étaient QUE là, écrites
partiellement puis laissées à la décision de l'acheteur) et relu en lecture
seule avec `lecture_facture.lire_facture` (aucune règle de parsing dans ce
module) ; la ligne est retrouvée par référence (`matching._cle`), sinon par
(qté, PU, montant) unique ; **Qté/PU/Montant viennent de Commandes** (ce
sont les valeurs écrites à l'époque — c'est ce qui garantit le centime), le
PDF donne N° BL lié, référence fournisseur, désignation ; une ligne
Commandes qui agrégeait plusieurs BL (agreger_lignes_meme_reference)
redevient une ligne Pièces par BL si les parts se somment exactement ;
montant PDF ≠ Commandes -> Commentaire, jamais substitué. PDF introuvable /
n° absent / ligne non retrouvée / ambiguë -> « Migré sans PDF », Commentaire
explicite. Contrôle au centime par fournisseur (nom Suivi normalisé —
« Coredime »/« COREDIME », « 109 DISTRIBUTION »/« 109 Distribution » sont
saisis mélangés dans Commandes) AVANT écriture (sinon rien n'est écrit) ET
APRÈS, relu du classeur (`controler_sommes`). Rapport dans
`rapports/migration_pieces_<ts>.txt`.

Résultat (copie puis vivant, identiques) : **1 059 lignes Commandes ->
1 076 lignes Pièces** (1 055 « Migré » dont 14 éclatées en plusieurs BL ;
3 « PDF introuvable », 1 « ligne ambiguë » ; 358 PDF lus, 0 erreur ; ~20
min, OCR compris), au centime pour les 6 fournisseurs : 109 Distribution
35 960,18 € / COMINTER 29 412,03 € / COMINTER Mayotte 8 732,34 € /
COREDIME 102 222,22 € / GMR 19 752,05 € / STAND 64 18 868,70 € (total
214 947,52 €). Après bascule et recalcul Excel : **0 écart** sur N° facture,
Qté, Montant, Date pour les 1 059 lignes ; 177 écarts de PU ≤ 0,0049 €
(PU moyen pondéré exact vs PU arrondi à 4 décimales écrit à l'époque —
prévu par le plan) ; aucune ligne sans facture devenue non vide. Statut
facture : 937 ✅ Facturée, 3 936 🔵 En attente facture, 122 ⛔ Sur-facturée
— TOUTES les 122 avec Qté livrée = 0 (facture reçue avant que son BL soit
rapproché, cas normal dixit l'acheteur — libellé à affiner en P2), aucune
avec facturé > livré > 0.

### Pipelines (étape 6)

- `pipeline_facture.appliquer_et_archiver_factures()` écrit désormais des
  lignes **Pièces** (`pieces_pour_facture` : Facture — Avoir si le document
  en est un —, + les Frais des factures écrites, Mode Auto / Confirmé /
  Équivalence selon statut et cause, Fichier = chemin d'archive PRÉDIT
  `_chemin_archive_facture` — le même quel que soit le sort du fichier,
  archivé tout de suite ou après un passage par À vérifier/) et plus jamais
  les 5 colonnes ; conserve l'écriture « Tarif BL » (liste blanche
  COREDIME), Pièces d'abord (refus net `FeuillePiecesAbsente` si la feuille
  manque), Tarif BL ensuite. `resume["pieces_ecrites"]`/`["pieces_ignorees"]`.
- « Déjà à jour » = le n° de facture est déjà écrit dans Pièces pour cette
  ligne de commande (`LigneSuiviFacture.numeros_factures`, relu depuis la
  feuille Pièces — `IndexPieces`, lue UNE fois par lot — et non plus depuis
  les cellules de Commandes, dont la valeur en cache n'est rafraîchie qu'à
  l'ouverture Excel ; repli sur les 5 colonnes si la feuille n'existe pas).
- **Garde-fou double facturation** (`matching_facture._comparer_facture`) :
  le CUMUL (déjà facturé dans Pièces + cette facture) > Qté livrée ->
  « à confirmer », cause `QTE_SUPERIEURE` (facture citée) ; une ligne
  facturée en plusieurs fois qui complète exactement le livré est « sûre »
  (l'ancien blocage « autre n° de facture déjà présent » n'a plus lieu
  d'être). **`DOUBLON_FACTURE`** au niveau document
  (`pipeline_facture._doublon_facture`) : même n° chez le même fournisseur
  déjà dans Pièces depuis un document d'identité différente (préfixe
  « <date> - <fournisseur> - Facture <n°> » du libellé Fichier) -> anomalie,
  jamais rapprochée ni écrite ; un redépôt du même document n'est pas un
  doublon (ses lignes ressortent déjà à jour).
- `pipeline_bl.appliquer_et_archiver()` écrit EN PLUS une ligne Pièces de
  type BL par ligne rapprochée (`pieces_pour_bl` ; Qté livrée reste écrite
  en cumul dans Commandes) ; feuille Pièces exigée AVANT toute écriture.
- `LigneSuivi`/`LigneSuiviFacture` portent Chantier/Sous-Chantier ;
  `ecriture.ENTETES_FACTURE` supprimé (`pieces.COLONNES_FACTURE_CALCULEES`
  pour la lecture) ; GUI : `FeuillePiecesAbsente` remontée proprement,
  message de fin avec le nombre de lignes Pièces.

### Tests

`tests/test_rapprochement_ecriture.py` (+13 : 3 socles + réparation sur
classeur synthétique, faux calcChain injecté, idempotence après réécriture
en sharedStrings ; 2 sur une copie du vivant : 16 -> 17 tableaux, colonne
calculée sur 6 556 lignes, U100/AV100 identiques),
`tests/test_rapprochement_pieces.py` (modèle, ID, formules, IndexPieces,
lecture/écriture), `tests/test_migration_pieces.py` (logique pure,
migration synthétique au centime + idempotence + bascule, extrait de 50
lignes sur une copie du vivant), pipelines BL/facture adaptés (lignes
Pièces, FeuillePiecesAbsente, doublon, cumul). Pytest complet du 2026-09-04 soir : **490 passed** (451 avant P1), 22 min 44.

### Exécution sur le vivant (étape 7, 2026-09-04 soir)

**NON EXÉCUTÉE ce soir** : tout est prouvé sur la copie (`Suivi - TEST P1
.xlsx`, mêmes 1 059 lignes que le vivant, hash MD5 826ADCB3… identique au
backup permanent `backups/permanents/Suivi commandes - 2026 - AVANT PIECES
2026-09-04 1714.xlsx`), pytest vert, mais le lancement du script qui écrit
dans le classeur VIVANT a été refusé par le garde-fou d'exécution de la
session Claude Code (auto mode classifier) — pas contourné. **Le vivant
est INTACT** (aucune écriture, vérifié : hash identique, pas de verrou).
Le script est dans le dépôt, à lancer depuis `Consultation AI/`, Suivi
FERMÉ, dans cet ordre (chaque commande s'arrête d'elle-même à la 1re
anomalie et laisse un journal à l'écran) :

    py -3 installer_pieces.py vivant   # ~25 min : 0d -> 1 -> 3 -> 4 -> 5 + Excel
    py -3 installer_pieces.py lot      # puis un lot de factures (lignes sûres)

`installer_pieces.py vivant` refuse de démarrer si le hash du vivant n'est
plus celui du backup permanent (classeur modifié entre-temps : refaire un
backup permanent d'abord, puis relancer) ; la comparaison des KPI Dashboard
attend les valeurs de l'original recalculé +218,54 € sur B6/N6/H9.

### Reste à faire / [HUMAIN] lundi

- [HUMAIN-P1-e] ouvrir le vivant dans Excel : aucun message de réparation
  attendu (prouvé par l'outil sur une copie du vivant écrit), onglet
  « Pièces » juste après Commandes, filtres, colonnes 56–59 de Commandes.
  Dashboard : mêmes valeurs que l'original recalculé, à +218,54 € près
  (ligne 2 réparée) sur Engagé total / Reste à recevoir / Attente livraison.
- P2 : Prix de référence / Source prix / Écart PU / Écart ligne € /
  Contrôle prix, et libellés de « Statut facture » (les 122 ⛔ à Qté livrée
  = 0 sont des factures avant BL, pas des sur-facturations).
- `moteur/fnp.py` lit toujours « N° facture »/« Date facture » de Commandes
  (valeurs en cache du dernier recalcul Excel ; « N° facture » peut désormais
  contenir « F1; F2 ») — à basculer sur Pièces dans une session dédiée.
- Les 412 factures d'À vérifier/ : inchangées, à reprendre par la GUI.

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
| DEM | `_gabarit.scan_regex` (devis) + procédural (BL) | ✅ couvert devis ET BL | Prix affichés AU CENT (/C) ; prix_net dérivé de montant/qté. Côté BL, chaque page = un BL indépendant (jamais de fusion inter-pages) |
| ELECTRIC PLUS (alias GMR) | `_gabarit.scan_ancre` | ✅ couvert | "GMR" = marque publique du canal Electric Plus, même gabarit |
| 109 DISTRIBUTION | `_gabarit.scan_ancre`, 3 variantes essayées | ✅ couvert (étendu session TRAVAUX_PARSERS.md) | **3 structures réelles différentes** chez ce fournisseur (réf avant le bloc chiffré, réf après, réf après + colonne Rem% renseignée) — voir "Points fragiles" |
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
| YESSS | procédural, `moteur/fournisseurs/yesss.py`, **BL uniquement** | ✅ couvert côté BL (1 PDF réel) — pas de devis connu pour ce fournisseur | Texte imprimé pivoté à 90° sur le BL (voir section Rapprochement AI) — valeurs retrouvées par proximité X/Y à leur label, pas par un ordre de lecture haut/bas |
| PROTECTHOMS | procédural, `moteur/fournisseurs/protecthoms.py`, **BL uniquement** | ✅ couvert côté BL (1 PDF réel) — pas de devis connu, équipements de protection/amiante | Référence repérée par sa FORME (1 chiffre + 2 lettres + 6 chiffres), pas d'en-tête/pied de tableau. Aucun prix sur ce document (comme Coredime) |
| ART DECO | procédural, `moteur/fournisseurs/artdeco.py`, **devis uniquement** | ✅ couvert (1 PDF réel, 2 lignes, total exact) | Brand "LED'S RUN", domaine artdeco.re (expéditeur idriss@artdeco.re) — "ELECTRICITE SERVICES REUNION" présent dans ces devis désigne l'acheteuse, pas ce fournisseur. Zone de tableau bornée par indice de texte (l'en-tête de colonnes apparaît APRÈS les lignes d'articles sur ce document) |

## Points fragiles connus

- **109 Distribution : TROIS structures réelles coexistent** chez ce même
  fournisseur (voir `moteur/fournisseurs/dist109.py`) : la référence vient
  tantôt AVANT le bloc chiffré ("Commande client n°..."), tantôt APRÈS
  ("Devis n°..."), et dans cette 2e forme la colonne "Rem%" (remise) peut
  en plus être RENSEIGNÉE (ex. "2,94") — elle n'apparaît alors dans le
  texte extrait QUE si elle est non nulle, décalant tout le reste d'un
  cran (3e variante, `OFFSETS_DEVIS_BPU_REMISE`, session
  TRAVAUX_PARSERS.md — devis ISHOP 321106/Réglettes - Rico Carpaye). Les
  trois sont essayées sur chaque bloc, celle dont la référence "a la bonne
  forme" est retenue. Sur le 1er PDF vu, 5 lignes/38 (câbles HO7VU 1.5mm²)
  ont un Total 2 % plus bas que Qté × P.U.Net affiché — jamais expliqué
  par une colonne visible. `prix_net` est donc calculé par Total/Qté
  (toujours exact), pas recopié du "P.U.Net" affiché (gardé à titre
  indicatif dans `prix_brut`). Classe de caractères de référence élargie
  deux fois cette même session : borne haute 15→20 (référence 16
  caractères, "FRN1X6G3-3G1.5 T") et "+" ajouté (références
  "BTSOUT3X150+70"/"BTSOUT3X95+50", devis BT - Floe 321273) — les deux
  fois révélées par l'autocontrôle Total HT (ligne(s) manquante(s)
  silencieusement), pas par un 0 article direct.
- **Electric Plus : colonne d'ancrage devis "PF" a aussi une variante
  réelle "PR"** (session TRAVAUX_PARSERS.md, 2 devis réels, BT - Floe et
  R2V 3G1.5 - Rico Carpaye — texte PDF natif, pas un artefact OCR comme le
  repli P[FR] déjà en place côté BL). `MARQUEUR` élargi à `["PF", "PR"]`
  dans `moteur/fournisseurs/electricplus.py`. A aussi révélé que la
  fixture `electric_plus_gmr.pdf` déjà en place perdait silencieusement 7
  lignes marquées "PR" depuis le début (pas de contrôle Total HT sur ce
  parser devis, contrairement au BL) — corrigé, verrouillé par un nouveau
  total dans le test. Limite résiduelle NON corrigée (un seul exemple à ce
  jour) : sur cette même fixture, une ligne (WAG2273205, 120,00€) n'a
  AUCUN marqueur PF/PR du tout — reste non extraite.
- **Documents "fiche technique" sans marqueur textuel de fournisseur**
  (session TRAVAUX_PARSERS.md, en creusant les PDF signalés "non
  reconnus" pour ART DECO/DEM/COREDIME dans un même lot) : plusieurs PDF
  reçus par mail sont en réalité des fiches produit FABRICANT (Novolight,
  Exalum Lighting, Thorn) jointes par le distributeur — sans prix, sans
  quantité, sans AUCUNE mention textuelle du distributeur qui les a
  envoyées (identifiable seulement via l'expéditeur e-mail, hors de
  portée du détecteur qui ne lit que le texte du PDF). Traités comme
  `FT-STELLAR.pdf` (déjà documenté, "Non bloquant") : correctement
  ignorés (`INCONNU`), pas une régression du détecteur.
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

## Flux demandes d'achat (pipeline Hermes) — implémenté 27/08/2026

Pipeline agentique piloté par l'agent Hermes (skill `flux-demandes-achat`,
cron « achats-flux » toutes les 30 min en heures ouvrées, livraison du récap
dans le groupe Telegram « IA ESR »). Trois scripts Python vivent à la racine
du projet (versionnés sur le repo) :

- `detecter_demandes.py` — lecture SEULE des boîtes ral@ + achats@ (Outlook
  COM) : fenêtre glissante (etat.json), anti-doublon EntryID, résolution GAL
  (`m.Sender.GetExchangeUser()`, startswith("/o=") INSENSIBLE à la casse),
  liste officielle des demandeurs (NOMS_DEMANDEURS), flag `deja_traite` via
  les Envoyés. Sorties : `demandes_a_traiter.json` + `devis_recus.json`.
- `creer_brouillons.py` — crée des BROUILLONS Outlook de demandes de devis
  (`mail.Save()` uniquement — JAMAIS Send()/Display(), garde-fou absolu).
  Entrée `brouillons_a_creer.json`, trace dans `suivi_consultations.csv`.
- `classer_devis.py` — regroupe `devis_recus.json` par affaire (sujet
  normalisé), crée `consultations/<affaire>/devis/`, télécharge les PDF
  (COM : `ns.GetItemFromID` — PAS sur les Folder en dispatch dynamique
  pywin32). Idempotent (« déjà présent » = déjà téléchargé).

Règle d'or du flux : AUCUN envoi programmatique ; les fichiers Besoin des
nouvelles consultations sont créés à partir de la consultation citée dans
les réponses fournisseurs (« De : William AIMAR ... Objet : ... »), jamais
inventés. Les fichiers runtime (etat.json, *_a_traiter.json, devis_recus.json,
brouillons_a_creer.json, suivi_consultations.csv, affaires_devis.json) sont
gitignorés.

## Évolutions envisagées (pas commencées)

- **Ingestion des devis par email** : ✅ FAIT (27/08/2026) — voir la section
  « Flux demandes d'achat (pipeline Hermes) » : les devis reçus par mail sur
  ral@/achats@ sont détectés (detecter_demandes.py), regroupés par affaire et
  téléchargés dans consultations/<affaire>/devis/ (classer_devis.py), puis les
  comparatifs sont générés automatiquement (cron Hermes « achats-flux »).
- **Bascule vers Appro-Tracker** quand son module Commandes existera :
  `referentiel/exports/*.csv` (alias appris, composés, articles nouveaux)
  est déjà pensé comme graine pour cette migration — voir section
  "Référentiel articles".
- **Enrichissement continu du référentiel partagé** : aujourd'hui
  `moteur/articles.db` et les confirmations (`A_confirmer.xlsx`) sont
  propres à ce poste ; un référentiel partagé entre plusieurs postes
  (Xavier compris) permettrait de capitaliser les alias confirmés une
  seule fois pour tout le monde.
