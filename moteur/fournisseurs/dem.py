import re

from moteur.modele import Article
from moteur.fournisseurs._gabarit import scan_regex

# --- GABARIT (DEM) ---------------------------------------------------------
# Structure du texte extrait :
#     LEG411651        15.00 U    8058.85/C    1208.83 1
#     DX3-ID 2P 63A A 30MA TGA          <- désignation sur la ligne SUIVANTE
#
# ATTENTION : les prix DEM sont exprimés AU CENT (/C).
# Le prix unitaire réel est donc montant / quantité.
# Les nombres utilisent le point décimal (128.52).
MOTIF_DEVIS = r"N°\s*(\d+)\s+du"
MOTIF_LIGNE = re.compile(
    r"^\s*([A-Z][A-Z0-9]+)\s+"      # référence
    r"([\d.]+)\s+"                   # quantité
    r"([A-Z]{1,3})\s+"               # unité (U...)
    r"([\d.]+)/C\s+"                 # prix au cent
    r"([\d.]+)\s+"                   # total HT
    r"(\d)\s*$"                      # code TVA
)
# --- fin GABARIT -------------------------------------------------------------


def parse_dem(texte: str) -> list[Article]:

    articles = []

    devis = ""
    m = re.search(MOTIF_DEVIS, texte)
    if m:
        devis = m.group(1)

    lignes = texte.splitlines()

    for i, m in scan_regex(lignes, MOTIF_LIGNE):

        ref = m.group(1)
        quantite = float(m.group(2))
        unite = m.group(3)
        prix_cent = float(m.group(4))
        montant = float(m.group(5))

        # Prix unitaire réel : montant / quantité (les prix sont au cent)
        prix_net = round(montant / quantite, 4) if quantite else 0.0
        prix_brut = round(prix_cent / 100, 4)

        # Désignation sur la ligne suivante
        designation = lignes[i + 1].strip() if i + 1 < len(lignes) else ""

        articles.append(
            Article(
                fournisseur="DEM",
                devis=devis,
                reference_fournisseur=ref,
                reference_distributeur="",
                designation=designation,
                quantite=quantite,
                unite=unite,
                prix_brut=prix_brut,
                prix_net=prix_net,
                montant=montant,
            )
        )

    return articles


# Déclaration pour le chargement automatique
FOURNISSEURS = ['DEM']
parse = parse_dem
