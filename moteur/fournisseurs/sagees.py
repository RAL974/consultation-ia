"""
Parser SAGEES (SAGEES Réunion, Saint-Paul — sagees.reunion@sagees.com).

Fournisseur découvert cette session (absent de la liste de priorité
initiale), ajouté à la demande de l'acheteur. Écrit sur plusieurs vrais PDF
(tests/fixtures/sagees_*.pdf).

PAS de colonne "Référence" dans ce devis : les articles ne sont identifiés
que par leur désignation (rapprochement automatique donc plus faible pour
ce fournisseur — attendu, pas un défaut du parser).
"""

import re

from moteur.modele import Article
from moteur.outils import to_float, lignes_propres

# --- GABARIT (SAGEES) --------------------------------------------------------
# Bloc ancré sur le marqueur "V0" (seul sur sa ligne, une fois par article) :
#     Désignation         (avant "V0", 1 à 2 lignes — la 2e ligne, quand
#                          elle existe, est une note de disponibilité/stock)
#     Montant             (juste avant "V0")
#     V0                  (marqueur)
#     Qté
#     Prix unitaire
MARQUEUR = "V0"
MOTIF_MONTANT = re.compile(r"^[\d\s]+,\d{2}$")
MOTS_PIED_DE_PAGE = ("Taux", "Taxe", "Base", "Total", "Acompte", "Escompte",
                     "Port", "NET A PAYER", "EUR", "SARL", "N°", "-")
MOTIF_DEVIS = r"\b(DR\d{6})\b"
MOTIF_TOTAL = r"Total HT\s*\n?.*?\n?\s*([\d\s]+,\d{2})"
# --- fin GABARIT -------------------------------------------------------------


def _autocontrole_total(texte: str, articles: list[Article]) -> None:
    m = re.search(MOTIF_TOTAL, texte)
    if not m:
        return
    total_pdf = to_float(m.group(1))
    total_extrait = round(sum(a.montant for a in articles), 2)
    if abs(total_pdf - total_extrait) > 0.02:
        print(
            f"!! SAGEES : Total HT du PDF ({total_pdf:.2f}€) != somme des "
            f"lignes extraites ({total_extrait:.2f}€) — à vérifier."
        )


def parse_sagees(texte: str) -> list[Article]:

    articles = []

    m = re.search(MOTIF_DEVIS, texte)
    devis = m.group(1) if m else ""

    lignes = lignes_propres(texte)
    n = len(lignes)

    for i, ligne in enumerate(lignes):

        if ligne != MARQUEUR:
            continue

        if i == 0 or not MOTIF_MONTANT.match(lignes[i - 1]):
            continue

        montant = to_float(lignes[i - 1])

        if i + 2 >= n:
            continue

        qte_brute = lignes[i + 1]
        prix_brute = lignes[i + 2]

        if not (MOTIF_MONTANT.match(qte_brute) and MOTIF_MONTANT.match(prix_brute)):
            continue

        quantite = to_float(qte_brute)
        prix_unitaire = to_float(prix_brute)

        # Remonte jusqu'à 2 lignes de désignation (la note de dispo, si
        # présente, est la ligne juste avant le montant)
        j = i - 2
        desig = []
        while j >= 0 and len(desig) < 2 and not lignes[j].startswith(MOTS_PIED_DE_PAGE) \
                and lignes[j] != MARQUEUR and not MOTIF_MONTANT.match(lignes[j]):
            desig.insert(0, lignes[j])
            j -= 1

        designation = " ".join(desig).strip()

        if not designation:
            continue

        articles.append(
            Article(
                fournisseur="SAGEES",
                devis=devis,
                reference_fournisseur="",
                reference_distributeur="",
                designation=designation,
                quantite=quantite,
                unite="UN",
                prix_brut=prix_unitaire,
                prix_net=prix_unitaire,
                montant=montant,
            )
        )

    _autocontrole_total(texte, articles)

    return articles


# Déclaration pour le chargement automatique
FOURNISSEURS = ['SAGEES']
parse = parse_sagees
