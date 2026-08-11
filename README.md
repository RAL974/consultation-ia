# Consultation AI

Outil qui automatise la consultation fournisseurs : tu déposes un besoin
chantier et les devis PDF reçus, il lit tout, rapproche les offres entre
fournisseurs, et te sort un tableau comparatif pour choisir le mieux-disant
en quelques clics — puis un panier prêt à coller dans le Suivi commandes.

Ce document explique comment l'installer et l'utiliser sans rien connaître
à la programmation.

## 1. Installer (une seule fois)

1. Installe Python depuis [python.org](https://www.python.org/downloads/)
   (bouton "Download Python"). **Pendant l'installation, coche bien la case
   "Add python.exe to PATH"** avant de cliquer sur "Install Now" — sans ça,
   l'outil ne trouvera pas Python.
2. Double-clique sur `lancer_gui.bat`. La première fois, il installe seul
   les quelques bibliothèques dont il a besoin (une fenêtre de console
   s'ouvre le temps de l'installation, c'est normal) puis affiche la
   fenêtre de l'outil.

Rien d'autre à installer. Les fois suivantes, `lancer_gui.bat` s'ouvre
directement.

## 2. Utiliser

### Créer une consultation

Chaque affaire (chantier, appel d'offres...) vit dans **son propre
dossier**, sous `consultations/` :

```
consultations/
  2026-08_Doujani/
    Besoin Doujani.txt        <- le besoin (voir modèle Besoin_exemple.txt)
    devis/                    <- dépose ici les PDF de devis reçus
      DEVIS - RAVATE.pdf
      DEVIS - GMR.pdf
```

- Crée un dossier sous `consultations/` (le nom est libre ; préfixer par la
  date, ex. `2026-08_Doujani`, aide juste à s'y retrouver plus tard — ce
  n'est pas obligatoire).
- Dépose le fichier besoin directement dans ce dossier (voir
  `Besoin_exemple.txt` à la racine du projet pour le format : une ligne par
  article, `Demande d'origine ; Référence ; Quantité`). Si tu n'as pas de
  besoin précis, tu peux aussi ne rien mettre : l'outil comparera alors
  toutes les lignes des devis, sans rapprochement à une demande.
- Crée un sous-dossier `devis/` dans ce même dossier et dépose-y tous les
  PDF de devis reçus pour cette affaire.

### Générer le comparatif

1. Double-clique sur `lancer_gui.bat`.
2. "Parcourir…" → choisis le dossier de la consultation
   (`consultations/2026-08_Doujani`, par ex.).
3. Clique sur **"Générer le comparatif"**. Le journal en bas de fenêtre
   affiche ce qui se passe (PDF lus, fournisseur détecté, anomalies...).
4. Un fichier `Comparatif ....xlsx` est créé dans le sous-dossier
   `resultats/` de la consultation. Clique sur "Ouvrir le comparatif" pour
   le voir directement.

### Décider et commander

1. Dans le Comparatif ouvert dans Excel, la colonne **"Fournisseur
   retenu"** est déjà pré-remplie au moins cher (mieux-disant) sur chaque
   ligne, mais reste modifiable : clique sur une cellule pour choisir un
   autre fournisseur dans la liste déroulante si tu préfères.
2. **Enregistre** le fichier Excel (Ctrl+S).
3. Reviens dans l'outil, clique sur **"Générer le panier"** (le champ
   "Comparatif" est déjà rempli automatiquement).
4. Un fichier `Panier ....xlsx` est créé dans `resultats/`. Vérifie
   l'onglet **"Non commandées"** (rien n'y est perdu en silence : lignes
   sans offre, ou sans décision prise), puis colle le reste dans la feuille
   "Commandes" du Suivi commandes.

Tu peux fermer l'outil et y revenir plus tard : en re-choisissant le même
dossier de consultation, le Comparatif déjà généré est retrouvé tout seul.

## 3. Lire le Comparatif — que veulent dire les couleurs ?

| Couleur | Sur quelle colonne | Ce que ça veut dire |
|---|---|---|
| 🟩 Vert | Prix d'un fournisseur | C'est le prix retenu comme meilleur-disant sur cette ligne. |
| 🟥 Rouge | "Meilleur prix" | Aucune offre reçue pour cette ligne du besoin — à traiter à part (devis manquant, référence introuvable...). |
| 🟥 Rouge | Prix d'un fournisseur | Ce prix s'écarte fortement (×6 ou plus) des autres offres — souvent une erreur d'unité (ex. prix au mètre confondu avec prix à la barre). **À vérifier sur le PDF avant de retenir cette offre.** |
| 🟧 Orange | "Base" (unité de comparaison) | La conversion vers l'unité commune (€/m, €/unité...) est incertaine, faute de connaître le conditionnement exact. Le prix est affiché tel quel, à vérifier. |
| 🟨 Jaune | "Qté" de l'offre retenue | La quantité de l'offre retenue ne correspond pas exactement à la quantité demandée dans le besoin — à vérifier avant de valider. |

Aucune ligne n'est jamais supprimée ou masquée automatiquement : une ligne
signalée reste visible, c'est à toi de trancher.

## 4. En cas de problème

Chaque génération écrit un **journal** dans le sous-dossier `resultats/`
de la consultation traitée : `journal_comparatif_....log` ou
`journal_panier_....log`. Il contient tout le détail de ce qui a été lu, y
compris les PDF qui ont posé problème et pourquoi.

Si la fenêtre elle-même ne s'ouvre pas au double-clic sur
`lancer_gui.bat`, un fichier `gui_erreur.log` est écrit à la racine du
projet.

**En cas de souci, c'est un de ces deux fichiers qu'il faut envoyer** pour
obtenir de l'aide.

## 5. Autres outils

- `py -3 audit.py` (ligne de commande) génère `resultats/Audit_BDD.xlsx` :
  un contrôle qualité de la base d'équivalences articles (base/BDD_articles.csv).
- `referentiel/A_confirmer.xlsx`, quand il existe après une génération,
  liste des rapprochements de références proposés entre fournisseurs, à
  valider (colonne "Décision" : `OUI` / `NON` / ou coller la bonne clé).
  Les décisions prises sont appliquées à la génération suivante.

Ces deux points sont détaillés dans `CLAUDE.md`, à l'usage d'une prochaine
session de développement sur ce projet.
