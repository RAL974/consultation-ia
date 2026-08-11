"""
Parser Importer (Importer, St-Paul).

Colonnes en ordre INVERSÉ dans le texte, un article = 8 lignes :
    Code (réf, ex. 421-D18C220)
    TVA (8,50)
    Montant HT
    P.U.Net
    % Rem
    P.U. HT
    Qté
    Description
Sections "EN STOCK :" / "SUR COMMANDE :" à ignorer.
"""

import re

from moteur.modele import Article
from moteur.outils import to_float, lignes_propres

_REF = re.compile(r"^\d{3}-\S+$")


def parse_importer(texte: str) -> list[Article]:

    articles = []

    m = re.search(r"\n(DE\d{6,})\n", texte)
    devis = m.group(1) if m else ""

    lignes = lignes_propres(texte)
    n = len(lignes)

    for i, l in enumerate(lignes):

        if not _REF.match(l):
            continue

        # 7 lignes suivantes attendues
        if i + 7 >= n:
            continue

        try:
            tva = lignes[i + 1]
            montant = to_float(lignes[i + 2])
            prix_net = to_float(lignes[i + 3])
            # i+4 = % Rem, i+5 = P.U. HT
            prix_brut = to_float(lignes[i + 5])
            quantite = to_float(lignes[i + 6])
            designation = lignes[i + 7]
        except Exception:
            continue

        # Garde-fou : montant ~ qté x prix net
        if quantite and prix_net and abs(montant - quantite * prix_net) > max(0.05 * montant, 1):
            continue

        articles.append(
            Article(
                fournisseur="IMPORTER",
                devis=devis,
                reference_fournisseur=l,
                reference_distributeur="",
                designation=designation,
                quantite=quantite,
                unite="UN",
                prix_brut=prix_brut,
                prix_net=prix_net,
                montant=montant,
            )
        )

    return articles


FOURNISSEURS = ["IMPORTER"]
parse = parse_importer
