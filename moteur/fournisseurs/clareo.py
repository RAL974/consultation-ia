"""
Parser Clareo (Clareo Lighting, Paris).

Bloc article :
    Référence (ex. TUB.101700)
    Nom du produit (1-2 lignes, jusqu'à "Couleur :")
    Couleur : ... / Flux ... / Impact ... / Voir la fiche produit   (specs, ignorées)
    Qté (entier)
    PU public HT (x €)
    Remise métier (53% ou -)
    Remise except. (5% ou -)
    PU net HT (y €)
    TVA 20%
    Total HT (z €)
On ancre sur "TVA 20%".
"""

import re

from moteur.modele import Article
from moteur.outils import to_float, lignes_propres

_REF = re.compile(r"^[A-Z]{2,4}\.[A-Z0-9]+$")


def parse_clareo(texte: str) -> list[Article]:

    articles = []

    m = re.search(r"bon de commande\s*:\s*(\S+)", texte)
    devis = m.group(1) if m else ""

    lignes = lignes_propres(texte)
    n = len(lignes)

    # Repérer les positions des références et des ancres "TVA 20%"
    refs = [(k, l) for k, l in enumerate(lignes) if _REF.match(l)]

    for k, l in enumerate(lignes):

        if not l.startswith("TVA "):
            continue

        try:
            prix_net = to_float(lignes[k - 1].replace("€", ""))
            # remontée : k-1 punet, k-2 rem_exc, k-3 rem_met, k-4 pu_public, k-5 qté
            quantite = to_float(lignes[k - 5])
            total = to_float(lignes[k + 1].replace("€", ""))
        except Exception:
            continue

        if not quantite:
            continue

        # Référence = dernière réf rencontrée avant cette ancre
        ref = ""
        nom = []
        for rk, rl in refs:
            if rk < k:
                ref = rl
                ref_pos = rk
            else:
                break
        # Nom = lignes entre la réf et la première spec "Couleur :"
        if ref:
            p = ref_pos + 1
            while p < k and not lignes[p].startswith(("Couleur", "Flux", "Impact", "Voir")):
                nom.append(lignes[p])
                p += 1

        articles.append(
            Article(
                fournisseur="CLAREO",
                devis=devis,
                reference_fournisseur=ref,
                reference_distributeur="",
                designation=" ".join(nom),
                quantite=quantite,
                unite="UN",
                prix_brut=prix_net,
                prix_net=prix_net,
                montant=total,
            )
        )

    return articles


FOURNISSEURS = ["CLAREO"]
parse = parse_clareo
