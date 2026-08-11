"""
Parser Legrand (export de prix PROJET, "Legrand Internal#").

Ce n'est pas un devis distributeur mais l'export tarif projet Legrand
(souvent transmis via un distributeur). Bloc article :
    Code Produit (réf)
    Description
    Quantité Dema (entier)
    Quantité Arrondie (entier)
    Prix Unitaire (x €)
    Montant Total (y €)
La fin du document porte le total puis un compteur : on s'arrête dès
qu'une ligne attendue comme référence est en fait un montant.
"""

import re

from moteur.modele import Article
from moteur.outils import to_float, lignes_propres

_EURO = re.compile(r"^[\d\s.,]+€$")


def parse_legrand(texte: str) -> list[Article]:

    articles = []

    m = re.search(r"REU\s*\d+-([^\n]+)", texte)
    devis = m.group(1).strip()[:40] if m else "LEGRAND PROJET"

    lignes = lignes_propres(texte)
    n = len(lignes)

    # Début après l'en-tête
    debut = 0
    for k, l in enumerate(lignes):
        if l.startswith("Montant Total"):
            debut = k + 1
            break

    i = debut
    while i + 5 < n:
        ref = lignes[i]

        # Fin : un montant à la place d'une référence = total du document
        if _EURO.match(ref):
            break

        desc = lignes[i + 1]
        qte_dema = lignes[i + 2]
        qte_arr = lignes[i + 3]
        prix = lignes[i + 4]
        montant = lignes[i + 5]

        if not (_EURO.match(prix) and _EURO.match(montant)):
            i += 1
            continue

        quantite = to_float(qte_arr)
        prix_net = to_float(prix.replace("€", ""))
        montant_v = to_float(montant.replace("€", ""))

        articles.append(
            Article(
                fournisseur="LEGRAND",
                devis=devis,
                reference_fournisseur=ref,
                reference_distributeur="",
                designation=desc,
                quantite=quantite,
                unite="UN",
                prix_brut=prix_net,
                prix_net=prix_net,
                montant=montant_v,
            )
        )
        i += 6

    return articles


FOURNISSEURS = ["LEGRAND"]
parse = parse_legrand
