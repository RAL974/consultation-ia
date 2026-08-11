import re

from moteur.modele import Article
from moteur.fournisseurs._gabarit import scan_ancre

# --- GABARIT (Electric Plus Réunion — marque publique du canal GMR) --------
# Structure du texte extrait, un champ par ligne, ancrée sur "PF" :
#     Référence          (i-5)
#     Désignation        (i-4)
#     Quantité           (i-3)
#     Prix unit. HT      (i-2)
#     P.U. net HT        (i-1)
#     PF                 (i)
#     Montant HT         (i+1)
#     Code TVA           (i+2)
MARQUEUR = "PF"
OFFSETS = {
    "reference_fournisseur": -5,
    "designation": -4,
    "quantite": -3,
    "prix_brut": -2,
    "prix_net": -1,
    "montant": 1,
}
MOTIF_DEVIS = r"^\s*(\d{7})\s*$"
MOTIF_REF = re.compile(r"^[A-Z0-9]{4,}$")
# --- fin GABARIT -------------------------------------------------------------


def _f(v: str) -> float:
    return float(v.replace(" ", "").replace(" ", "").replace(",", "."))


def parse_electricplus(texte: str) -> list[Article]:

    articles = []

    devis = ""
    m = re.search(MOTIF_DEVIS, texte, re.MULTILINE)
    if m:
        devis = m.group(1)

    lignes = [l.rstrip() for l in texte.splitlines()]

    for bloc in scan_ancre(lignes, MARQUEUR, OFFSETS):

        try:
            ref = bloc["reference_fournisseur"].strip()
            designation = bloc["designation"].strip()

            quantite = _f(bloc["quantite"])
            prix_brut = _f(bloc["prix_brut"])
            prix_net = _f(bloc["prix_net"])
            montant = _f(bloc["montant"])

        except (ValueError, IndexError, AttributeError):
            continue

        if not MOTIF_REF.match(ref):
            continue

        articles.append(
            Article(
                fournisseur="ELECTRIC PLUS",
                devis=devis,
                reference_fournisseur=ref,
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


# Déclaration pour le chargement automatique
FOURNISSEURS = ['ELECTRIC PLUS', 'GMR']
parse = parse_electricplus
