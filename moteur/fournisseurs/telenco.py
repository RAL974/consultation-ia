"""
Parser TELENCO (Telenco Réunion — 44 rue Mahatma Gandhi, La Possession).

Écrit sur 3 vrais PDF (tests/fixtures/telenco_*.pdf).
"""

import re

from moteur.modele import Article
from moteur.outils import to_float, lignes_propres

# --- GABARIT (Telenco) ------------------------------------------------------
# Bloc :
#     Référence     (entier SANS virgule, ex. "16475", "0403" — se distingue
#                     des champs numériques qui ont toujours 2 décimales)
#     Désignation   (1 à 2 lignes)
#     Qté           (ex. "5,00")
#     Prix unitaire (ex. "7,66")
#     Montant       (ex. "38,30")
MOTIF_REF = re.compile(r"^\d+$")
MOTIF_NUM = re.compile(r"^\d+,\d{2}$")
MOTIF_DEVIS = r"N°\s*(SQFR\S+)"
# --- fin GABARIT -------------------------------------------------------------


def parse_telenco(texte: str) -> list[Article]:

    articles = []

    m = re.search(MOTIF_DEVIS, texte)
    devis = m.group(1) if m else ""

    lignes = lignes_propres(texte)
    n = len(lignes)

    i = 0
    while i < n:

        if not MOTIF_REF.match(lignes[i]):
            i += 1
            continue

        ref = lignes[i]

        # Désignation : lignes non numériques jusqu'aux 3 valeurs chiffrées
        j = i + 1
        desig = []
        while j < n and len(desig) < 3 and not MOTIF_NUM.match(lignes[j]):
            desig.append(lignes[j])
            j += 1

        if j + 2 >= n or not (
            MOTIF_NUM.match(lignes[j])
            and MOTIF_NUM.match(lignes[j + 1])
            and MOTIF_NUM.match(lignes[j + 2])
        ):
            i += 1
            continue

        quantite = to_float(lignes[j])
        prix_unitaire = to_float(lignes[j + 1])
        montant = to_float(lignes[j + 2])

        articles.append(
            Article(
                fournisseur="TELENCO",
                devis=devis,
                reference_fournisseur=ref,
                reference_distributeur="",
                designation=" ".join(desig).strip(),
                quantite=quantite,
                unite="UN",
                prix_brut=prix_unitaire,
                prix_net=prix_unitaire,
                montant=montant,
            )
        )

        i = j + 3

    return articles


# Déclaration pour le chargement automatique
FOURNISSEURS = ['TELENCO']
parse = parse_telenco
