"""
Parser COMPTOIR DU CABLING (SARL Comptoir du Cabling, Sainte-Marie —
comptoir@comptoirducabling.com).

Repris sur 4 vrais PDF (tests/fixtures/comptoir_cabling_*.pdf) après un
premier essai abandonné (structure mal comprise avec un seul exemple, voir
CLAUDE.md) : le 2e PDF, plus simple (2 lignes, sans numéro de lot), a permis
de confirmer la structure EXACTE ci-dessous.
"""

import re

from moteur.modele import Article
from moteur.outils import to_float, lignes_propres

# --- GABARIT (Comptoir du Cabling) ------------------------------------------
# Bloc, parfois précédé d'un numéro de lot ignorable ("8.1", "9.12"...) :
#     Référence            ex. "42U600X600", "U7LITE"
#     TVA %                toujours "0,00" observé
#     Montant HT
#     P.U. HT
#     Qté
#     Désignation          1 à 2 lignes, jusqu'à la référence suivante, un
#                          numéro de lot, ou un mot-clé de pied de page
MOTIF_REF = re.compile(r"^(?=.*[A-Z])[A-Z0-9]{3,}$")  # ex. "42U600X600" (chiffres devant)
MOTIF_NUM = re.compile(r"^\d+(?:[.,]\d{1,2})?$")
MOTIF_LOT = re.compile(r"^\d+\.\d+$")
MOTS_PIED_DE_PAGE = ("Code", "Taux", "Base", "TVA", "Total", "Coordonnées",
                     "Document", "Devis", "Clause", "Pour le client")
MOTIF_DEVIS = r"\bDE\d{6,}\b"
MOTIF_TOTAL = r"Total HT remisé\s*\n\s*([\d\s]+,\d{2})"
# --- fin GABARIT -------------------------------------------------------------


def _autocontrole_total(texte: str, articles: list[Article]) -> None:
    m = re.search(MOTIF_TOTAL, texte)
    if not m:
        return
    total_pdf = to_float(m.group(1))
    total_extrait = round(sum(a.montant for a in articles), 2)
    if abs(total_pdf - total_extrait) > 0.02:
        print(
            f"!! COMPTOIR DU CABLING : Total HT remisé du PDF ({total_pdf:.2f}€) "
            f"!= somme des lignes extraites ({total_extrait:.2f}€) — à vérifier."
        )


def parse_comptoir_cabling(texte: str) -> list[Article]:

    articles = []

    m = re.search(MOTIF_DEVIS, texte)
    devis = m.group(0) if m else ""

    lignes = lignes_propres(texte)
    n = len(lignes)

    i = 0
    while i < n:

        if not MOTIF_REF.match(lignes[i]) or lignes[i].startswith(MOTS_PIED_DE_PAGE):
            i += 1
            continue

        # Les 4 valeurs chiffrées doivent suivre IMMÉDIATEMENT la référence
        if i + 4 >= n or not all(MOTIF_NUM.match(lignes[i + k]) for k in range(1, 5)):
            i += 1
            continue

        ref = lignes[i]
        tva = to_float(lignes[i + 1])
        montant = to_float(lignes[i + 2])
        prix_unitaire = to_float(lignes[i + 3])
        quantite = to_float(lignes[i + 4])

        desig = []
        j = i + 5
        while (
            j < n and len(desig) < 2
            and not MOTIF_REF.match(lignes[j])
            and not MOTIF_LOT.match(lignes[j])
            and not lignes[j].startswith(MOTS_PIED_DE_PAGE)
            and not MOTIF_NUM.match(lignes[j])
        ):
            desig.append(lignes[j])
            j += 1

        designation = " ".join(desig).strip()

        if designation:
            articles.append(
                Article(
                    fournisseur="COMPTOIR DU CABLING",
                    devis=devis,
                    reference_fournisseur=ref,
                    reference_distributeur="",
                    designation=designation,
                    quantite=quantite,
                    unite="UN",
                    prix_brut=prix_unitaire,
                    prix_net=prix_unitaire,
                    montant=montant,
                )
            )

        i = j

    _autocontrole_total(texte, articles)

    return articles


# Déclaration pour le chargement automatique
FOURNISSEURS = ['COMPTOIR DU CABLING']
parse = parse_comptoir_cabling
