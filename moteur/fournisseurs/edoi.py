"""
Parser EDOI (Sonepar Mayotte — B.P 244 ZI Kaweni, contact.edoi@sonepar.fr).

Écrit sur 3 vrais PDF (tests/fixtures/edoi_*.pdf).
"""

import re

from moteur.modele import Article
from moteur.outils import to_float, lignes_propres
from moteur.fournisseurs._gabarit import scan_regex

# --- GABARIT (EDOI) ---------------------------------------------------------
# Un article = une ligne complète, colonnes à largeur fixe :
#     LEG401802      12sem XL3 160 COMPLET ISOLANT 2R         9   UN            659,0971    5931,87 6
#     LEG404926      DISPO PEIGNE POUR 13 APPAREILS 1P+N     13   UN              5,1979      67,57 6
# Référence, [disponibilité : "DISPO" ou délai "Nsem"], Désignation, Qté,
# Unité (toujours "UN" observé), Prix net HT (4 décimales), Montant, TVA.
MOTIF_LIGNE = re.compile(
    r"^([A-Z][A-Z0-9]+)\s+"       # référence
    r"(DISPO|\d+sem)?\s*"          # disponibilité (facultative)
    r"(.+?)\s+"                    # désignation
    r"(\d+)\s+"                    # quantité
    r"UN\s+"                       # unité
    r"([\d,]+)\s+"                 # prix net HT
    r"([\d,]+)\s+"                 # montant
    r"\d\s*$"                      # code TVA
)
MOTIF_DEVIS = r"EDO\s*(B\d+)"
# --- fin GABARIT -------------------------------------------------------------


def parse_edoi(texte: str) -> list[Article]:

    articles = []

    m = re.search(MOTIF_DEVIS, texte)
    devis = m.group(1) if m else ""

    for _i, m in scan_regex(lignes_propres(texte), MOTIF_LIGNE):

        ref = m.group(1)
        disponibilite = m.group(2) or ""
        designation = m.group(3).strip()
        quantite = to_float(m.group(4))
        prix_net = to_float(m.group(5))
        montant = to_float(m.group(6))

        articles.append(
            Article(
                fournisseur="EDOI",
                devis=devis,
                reference_fournisseur=ref,
                reference_distributeur="",
                designation=designation,
                quantite=quantite,
                unite="UN",
                prix_brut=prix_net,
                prix_net=prix_net,
                montant=montant,
                disponibilite=disponibilite,
            )
        )

    return articles


# Déclaration pour le chargement automatique
FOURNISSEURS = ['EDOI']
parse = parse_edoi
