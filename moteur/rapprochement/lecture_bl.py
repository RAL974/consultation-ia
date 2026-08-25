"""
Lecture des BL déposés dans a_traiter/BL/ : OCR (moteur/ocr.py) + détection
fournisseur (moteur/detecteur.py, réutilisé tel quel) + parser BL dédié
(moteur/rapprochement/parsers_bl.py) — étape 1 de Rapprochement AI (session
R2, voir CLAUDE.md).

Même tolérance aux pannes que moteur/lecture_pdf.py pour les devis : un BL
illisible, d'un fournisseur non reconnu, ou d'un fournisseur reconnu mais
sans parser BL, ne bloque JAMAIS le traitement des autres fichiers du lot.

Détection PAR PAGE (pas seulement par fichier entier) : un même scan groupé
peut mélanger PLUSIEURS FOURNISSEURS ET plusieurs commandes (cas réel
rencontré 2 fois en session, traité à la main jusqu'ici — voir CLAUDE.md,
"détection de fournisseur par page"). `lire_bl()` détecte le fournisseur de
CHAQUE page individuellement ; si toutes les pages détectées s'accordent (ou
qu'une seule est reconnue), le comportement est STRICTEMENT celui d'avant
(détection sur le texte entier du document, un seul appel au parser — le cas
de loin le plus courant). Seulement si plusieurs fournisseurs DIFFÉRENTS
ressortent, le document est découpé par groupes de pages et chaque groupe
est parsé indépendamment (voir _parser_groupe_fournisseur)."""

from pathlib import Path

from moteur.detecteur import detecter_fournisseur
from moteur.ocr import lignes_ocr, mots_document
from moteur.rapprochement.parsers_bl import parser_bl

EXTENSIONS_SUPPORTEES = (".pdf", ".jpg", ".jpeg", ".png")


def _texte_page(mots_page: list[dict]) -> str:
    return "\n".join(lignes_ocr(mots_page))


def _parser_groupe_fournisseur(fournisseur: str, indices: list[int], mots_par_page: list[list[dict]]):
    """Parse le sous-ensemble de pages `indices` (0-based, DANS L'ORDRE
    d'apparition dans le fichier, pas forcément contigu) pour UN seul
    fournisseur. Retourne (bons, anomalies) — chaque bon a TOUJOURS
    `bl.pages` renseigné, remappé vers les indices ORIGINAUX du fichier
    (jamais None ici, contrairement au chemin mono-fournisseur historique :
    laisser None ferait archiver TOUT le fichier, y compris les pages des
    AUTRES fournisseurs — voir moteur.rapprochement.pipeline_bl).

    Un sous-ensemble à PLUSIEURS pages est d'abord tenté en UN SEUL appel —
    nécessaire pour les fournisseurs qui répartissent eux-mêmes un même BL
    sur 2 pages via moteur.ocr.pages_par_identifiant (109 Distribution,
    Cominter Ouest, Electric Plus : ils renseignent bl.pages eux-mêmes,
    même pour un fichier à un seul BL). Si le résultat ne renseigne PAS
    bl.pages (fournisseur qui traite toute son entrée comme UN SEUL
    document, ex. Coredime/Ravate/Stand 64/DEM), l'appel est refait PAGE
    PAR PAGE : regrouper plusieurs pages non contiguës du même fournisseur
    sous un seul BL risquerait de fusionner à tort deux commandes
    RÉELLEMENT différentes (cas réel constaté en session — voir CLAUDE.md,
    "2e occurrence... doc07205620260824145119.pdf") — mieux vaut
    sous-découper (au pire une info incomplète, réexaminée à la main) que
    mélanger deux BL sous une seule commande."""

    sous_ensemble = [mots_par_page[i] for i in indices]
    resultat = parser_bl(fournisseur, sous_ensemble)

    if resultat is None:
        pages_lisibles = ", ".join(str(i + 1) for i in indices)
        return [], [f"Page(s) {pages_lisibles} : fournisseur {fournisseur} reconnu mais pas encore de parser BL"]

    bons = resultat if isinstance(resultat, list) else [resultat]

    if len(indices) > 1 and len(bons) == 1 and bons[0].pages is None:
        bons_tous, anomalies_tous = [], []
        for i in indices:
            b, a = _parser_groupe_fournisseur(fournisseur, [i], mots_par_page)
            bons_tous.extend(b)
            anomalies_tous.extend(a)
        return bons_tous, anomalies_tous

    for bl in bons:
        bl.pages = list(indices) if bl.pages is None else [indices[j] for j in bl.pages]

    return bons, []


def lire_bl(chemin):
    """OCR + parse un seul fichier. Retourne (liste_de_BonLivraison,
    liste_de_raisons_en_clair) — la liste de raisons est vide en cas de
    succès complet ; jamais d'exception laissée remonter pour une raison
    "attendue" (fournisseur inconnu, pas de parser).

    Une LISTE de BonLivraison, pas un seul : certains fournisseurs
    (Cominter Ouest, 109 Distribution, Electric Plus) scannent parfois
    PLUSIEURS BL à la suite dans un même fichier — leur parse_bl() retourne
    alors une liste ; les autres retournent un seul BonLivraison, normalisé
    ici en liste à un élément. Un même fichier peut aussi mélanger
    PLUSIEURS FOURNISSEURS (voir docstring du module) : dans ce cas, `bons`
    peut contenir des BonLivraison de fournisseurs différents, ET
    `raisons` peut être non vide MÊME si `bons` ne l'est pas (une partie du
    fichier résolue, une autre page en anomalie — ex. fournisseur sans
    parser BL)."""

    chemin = Path(chemin)

    mots_par_page = mots_document(chemin)

    fournisseurs_page = [detecter_fournisseur(_texte_page(mots)) for mots in mots_par_page]
    distincts = {f for f in fournisseurs_page if f != "INCONNU"}

    if len(distincts) <= 1:
        # Chemin historique, INCHANGÉ : un seul fournisseur ressort de
        # l'ensemble du document (ou aucun) -> détection sur le texte
        # ENTIER (pas seulement page par page, pour ne rien perdre d'un
        # fournisseur qui ne se révèle que sur UNE page parmi plusieurs,
        # ex. une page de garde différente des suivantes) et un seul appel
        # au parser sur toutes les pages.
        texte = "\n".join(_texte_page(mots) for mots in mots_par_page)

        fournisseur = detecter_fournisseur(texte)

        if fournisseur == "INCONNU":
            return [], ["Fournisseur non reconnu (OCR)"]

        resultat = parser_bl(fournisseur, mots_par_page)

        if resultat is None:
            return [], [f"Fournisseur {fournisseur} reconnu mais pas encore de parser BL"]

        bons = resultat if isinstance(resultat, list) else [resultat]

        for bl in bons:
            bl.fichier = chemin.name

        return bons, []

    # Plusieurs fournisseurs détectés sur des pages différentes du MÊME
    # fichier — scan groupé par l'acheteur (cas réel rencontré 2 fois en
    # session, voir docstring du module). Regroupe les pages par
    # fournisseur (ordre d'apparition dans le fichier conservé pour
    # chaque groupe, PAS forcément contigu — un fournisseur peut
    # réapparaître plus loin dans le fichier), chaque groupe traité
    # indépendamment : un échec sur un groupe (fournisseur sans parser,
    # page illisible) n'empêche jamais les autres groupes d'être résolus.
    groupes_pages: dict[str, list[int]] = {}
    for i, f in enumerate(fournisseurs_page):
        groupes_pages.setdefault(f, []).append(i)

    bons_tous = []
    anomalies = []

    for fournisseur, indices in groupes_pages.items():

        if fournisseur == "INCONNU":
            pages_lisibles = ", ".join(str(i + 1) for i in indices)
            anomalies.append(f"Page(s) {pages_lisibles} : fournisseur non reconnu (OCR)")
            continue

        try:
            bons_groupe, anomalies_groupe = _parser_groupe_fournisseur(fournisseur, indices, mots_par_page)
        except Exception as e:
            pages_lisibles = ", ".join(str(i + 1) for i in indices)
            anomalies.append(f"Page(s) {pages_lisibles} ({fournisseur}) : erreur de lecture ({e})")
            continue

        for bl in bons_groupe:
            bl.fichier = chemin.name

        bons_tous.extend(bons_groupe)
        anomalies.extend(anomalies_groupe)

    return bons_tous, anomalies


def analyser_dossier(dossier):
    """Lit tous les BL de `dossier`. Retourne (bons_de_livraison, anomalies)
    — anomalies : liste de (nom_fichier, raison_en_clair)."""

    dossier = Path(dossier)

    if not dossier.is_dir():
        print(f"\n{dossier} n'existe pas encore (aucun BL déposé) — rien à lire.\n")
        return [], []

    fichiers = sorted(
        f for f in dossier.iterdir()
        if f.is_file() and f.suffix.lower() in EXTENSIONS_SUPPORTEES
    )

    print(f"\n{len(fichiers)} BL trouvé(s) dans {dossier}\n")

    bons_de_livraison = []
    anomalies = []

    for fichier in fichiers:

        print("-" * 60)
        print(f"Lecture : {fichier.name}")

        try:
            bons, raisons = lire_bl(fichier)
        except Exception as e:
            print(f"!! Erreur de lecture, ignoré : {e}")
            anomalies.append((fichier.name, f"Erreur de lecture ({e})"))
            continue

        for raison in raisons:
            print(f"!! {raison}")
            anomalies.append((fichier.name, raison))

        if len(bons) > 1:
            print(f"{len(bons)} BL détectés dans ce fichier (scan groupé).")

        for bl in bons:
            print(
                f"Fournisseur : {bl.fournisseur} — {len(bl.lignes)} ligne(s), "
                f"commande {bl.numero_commande or '?'}, BL {bl.numero_bl or '?'}"
            )

            if not bl.lignes:
                anomalies.append((fichier.name, f"Aucune ligne extraite (BL {bl.numero_bl or '?'})"))

            bons_de_livraison.append(bl)

    print("=" * 60)
    print("Résumé de la lecture des BL")
    print("=" * 60)
    print(f"{len(bons_de_livraison)} BL lu(s) sur {len(fichiers)}.")
    if anomalies:
        print(f"{len(anomalies)} anomalie(s) :")
        for nom, raison in anomalies:
            print(f"  - {nom} : {raison}")
    else:
        print("Aucune anomalie.")
    print("=" * 60)

    return bons_de_livraison, anomalies
