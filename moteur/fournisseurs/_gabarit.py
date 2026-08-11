"""
Moteur partagé pour les gabarits fournisseurs.

Objectif : séparer, dans chaque parser, le MOTEUR (boucle, ancrage, calcul
d'offsets) du GABARIT (motifs regex, marqueurs, ordre des champs — propres à
un fournisseur). Le gabarit d'un fournisseur reste dans son fichier
(moteur/fournisseurs/<nom>.py), sous un bandeau "# --- GABARIT ---" : seules
les constantes changent d'un fournisseur à l'autre, la boucle vient d'ici.

Deux primitives couvrent les structures réellement observées dans les devis
réels (voir tests/fixtures/) :

- scan_ancre()  : un marqueur répété dans le texte (ex. "Réf. FNR :" chez
                  Ravate, "PF" chez Electric Plus), et des champs à des
                  positions FIXES relatives au marqueur (positives ou
                  négatives). Un bloc dont un champ tombe hors texte est
                  écarté (comme le try/except IndexError d'origine).
- scan_regex()  : un regex plein-texte par ligne (Coredime, DEM), groupes
                  nommés laissés au gabarit à interpréter.

Les fournisseurs à structure vraiment procédurale (Cominter : deux formats
de PDF à essayer ; 109 Distribution : remontée arrière à longueur variable
selon la présence d'un éco-part) NE PASSENT PAS par ce moteur — les forcer
dans ces primitives risquerait de changer un comportement non couvert par
les PDF réels actuellement disponibles. Leurs constantes (regex, marqueurs)
sont simplement alignées sur le même bandeau "# --- GABARIT ---" pour la
cohérence de lecture, la boucle reste explicite dans leur fichier.

Règle d'or (voir CLAUDE.md) : aucune primitive ici n'invente une règle de
parsing — chacune reproduit exactement un mécanisme déjà observé dans un
vrai PDF, verrouillé par tests/test_parsers.py.
"""

import re


# ----------------------------------------------------------------------
# Table partagée — mots-clés de disponibilité
# ----------------------------------------------------------------------
# Vocabulaire commun observé chez plusieurs fournisseurs. Le MÉCANISME de
# détection (recherche libre dans une zone de texte chez Ravate, préfixe de
# désignation chez Coredime...) reste propre à chaque gabarit : seul le
# vocabulaire est mutualisé ici, pour n'avoir qu'un seul endroit où
# l'étendre.
DISPONIBILITES = ("AEC", "SUR COMMANDE", "STOCK", "DISPO")


# ----------------------------------------------------------------------
# Primitive 1 — ancrage + offsets fixes (Ravate, Electric Plus)
# ----------------------------------------------------------------------
def scan_ancre(lignes: list[str], marqueurs, offsets: dict) -> list[dict]:
    """
    Pour chaque ligne de `lignes` égale (insensible à la casse) à l'un des
    `marqueurs`, construit un dict {nom_champ: ligne_brute_ou_None} en
    lisant, pour chaque (nom_champ, offset) de `offsets`, la ligne à
    l'indice (indice_marqueur + offset). offset peut être négatif (champ
    avant le marqueur) ou positif (après).

    Un champ dont l'offset tombe hors du texte vaut None : au gabarit
    appelant de décider s'il écarte le bloc (comme le ferait un
    try/except IndexError).

    Retourne la liste des blocs, dans l'ordre des marqueurs rencontrés,
    ainsi que l'indice du marqueur (clé "_i") pour un éventuel scan
    complémentaire (ex. disponibilité, voir `disponibilite_apres`).
    """
    if isinstance(marqueurs, str):
        marqueurs = [marqueurs]
    marqueurs_maj = {m.upper() for m in marqueurs}

    n = len(lignes)
    blocs = []

    for i, ligne in enumerate(lignes):

        if ligne.strip().upper() not in marqueurs_maj:
            continue

        bloc = {"_i": i}

        for nom_champ, offset in offsets.items():
            idx = i + offset
            bloc[nom_champ] = lignes[idx] if 0 <= idx < n else None

        blocs.append(bloc)

    return blocs


def disponibilite_apres(lignes: list[str], depart: int, marqueurs,
                         mots_cles=DISPONIBILITES) -> str:
    """
    Scanne `lignes` à partir de l'indice `depart`, jusqu'au prochain
    marqueur (ou la fin du texte), et retourne la PREMIÈRE ligne brute
    contenant l'un de `mots_cles` (recherche libre, style Ravate : c'est la
    ligne entière qui est retournée, pas seulement le mot-clé — le devis
    peut y ajouter du contexte). "" si rien trouvé.
    """
    if isinstance(marqueurs, str):
        marqueurs = [marqueurs]
    marqueurs_maj = {m.upper() for m in marqueurs}

    disponibilite = ""
    j = depart
    n = len(lignes)

    while j < n and lignes[j].strip().upper() not in marqueurs_maj:
        lg = lignes[j].upper()
        if any(mot in lg for mot in mots_cles):
            disponibilite = lignes[j]
        j += 1

    return disponibilite


def diviser_qte_unite(brut: str, unite_defaut: str = "UN"):
    """'15 UN' -> (15.0, 'UN') ; '15' -> (15.0, unite_defaut)."""
    from moteur.outils import to_float

    morceaux = (brut or "").split()
    quantite = to_float(morceaux[0]) if morceaux else 0.0
    unite = morceaux[1] if len(morceaux) > 1 else unite_defaut

    return quantite, unite


# ----------------------------------------------------------------------
# Primitive 2 — un regex par ligne (Coredime, DEM)
# ----------------------------------------------------------------------
def scan_regex(lignes: list[str], motif: "re.Pattern") -> list[tuple[int, "re.Match"]]:
    """
    Applique `motif` à chaque ligne, retourne (indice, correspondance) pour
    chaque ligne qui matche — l'indice permet au gabarit appelant d'aller
    lire une ligne voisine si un champ (ex. désignation) se trouve sur la
    ligne suivante plutôt que dans le regex lui-même (cas DEM).
    """
    resultat = []
    for i, ligne in enumerate(lignes):
        m = motif.match(ligne)
        if m:
            resultat.append((i, m))
    return resultat
