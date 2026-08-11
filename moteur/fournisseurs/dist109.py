"""
Parser 109 DISTRIBUTION.

RÉÉCRIT sur un vrai PDF (`tests/fixtures/109_distribution.pdf`) : l'ancien
gabarit supposait un ordre de champs jamais observé en réalité (0 article
extrait — voir CLAUDE.md, "points fragiles" de la session précédente).

DEUX VARIANTES réelles coexistent chez ce fournisseur (même marque, mais
109 Est/Sud/Ouest/Nord partagent apparemment plusieurs gabarits de devis) :
- "commande" (ex. tests/fixtures/109_distribution.pdf) : la référence vient
  AVANT le bloc chiffré, motif "Commande client n° ... du ...".
- "devis_bpu" (ex. devis "e.s BPU Région...") : la référence vient APRÈS,
  motif "Devis n° ... du ...". Découverte via l'autocontrôle Total HT
  (voir plus bas) : la variante "commande" y donnait 0 article, l'écart
  entre Total HT du PDF et somme extraite l'a immédiatement signalé.
Les deux sont essayées sur chaque bloc ancré sur le code TVA ; le premier
dont la référence "a la forme attendue" est retenu.
"""

import re

from moteur.modele import Article
from moteur.outils import to_float, lignes_propres
from moteur.fournisseurs._gabarit import scan_ancre

# --- GABARIT (109 Distribution) --------------------------------------------
# Bloc ancré sur le code TVA (C0 à C9, un seul mot sur sa ligne).
MARQUEURS_TVA = tuple(f"C{d}" for d in range(10))

# Variante "commande" : Référence, Désignation, Total, Qté, TVA, P.U.Net, [code]
OFFSETS_COMMANDE = {
    "reference_fournisseur": -4,
    "designation": -3,
    "total": -2,
    "qte": -1,
    "pu_net_affiche": 1,
}
MOTIF_REF_COMMANDE = re.compile(r"^\d{6,}-")  # ex. "45001001-PRYSM-H07VU1.5B"

# Variante "devis_bpu" : Désignation, Qté, TVA, Total, P.U.Net, Référence
OFFSETS_DEVIS_BPU = {
    "designation": -2,
    "qte": -1,
    "total": 1,
    "pu_net_affiche": 2,
    "reference_fournisseur": 3,
}
MOTIF_REF_DEVIS_BPU = re.compile(r"^[A-Z0-9./]{5,15}$")  # ex. "LK4288", "CR1/3G2.5"

MOTIF_DEVIS = r"(?:Commande client|Devis) n[°o]\s*([\d\s]+?)\s+du"
MOTIF_TOTAL_HT = r"Total HT\s*\n\s*([\d\s]+,\d{2})\s*(?:EUR|€)"
# --- fin GABARIT -------------------------------------------------------------


def _autocontrole_total_ht(texte: str, articles: list[Article]) -> None:
    """Contrôle additif propre à 109 Distribution : le PDF affiche un
    Total HT global -> le comparer à la somme des lignes extraites est le
    meilleur signal possible qu'aucune ligne n'a été oubliée ou mal lue,
    ou que la variante de gabarit essayée ne correspond pas à ce PDF."""

    m = re.search(MOTIF_TOTAL_HT, texte)
    if not m:
        return

    total_pdf = to_float(m.group(1))
    total_extrait = round(sum(a.montant for a in articles), 2)

    if abs(total_pdf - total_extrait) > 0.02:
        print(
            f"!! 109 DISTRIBUTION : Total HT du PDF ({total_pdf:.2f}€) "
            f"!= somme des lignes extraites ({total_extrait:.2f}€) "
            f"— une ligne a peut-être été oubliée ou mal lue, ou ce PDF "
            f"correspond à une 3e variante de gabarit jamais vue."
        )


def _bloc_vers_article(bloc, devis, ref) -> Article | None:

    if bloc["qte"] is None or bloc["total"] is None or bloc["pu_net_affiche"] is None:
        return None

    quantite = to_float(bloc["qte"])
    total = to_float(bloc["total"])
    pu_net_affiche = to_float(bloc["pu_net_affiche"])

    if not quantite:
        return None

    return Article(
        fournisseur="109 DISTRIBUTION",
        devis=devis,
        reference_fournisseur=ref,
        reference_distributeur="",
        designation=(bloc["designation"] or "").strip(),
        quantite=quantite,
        unite="UN",
        prix_brut=pu_net_affiche,
        prix_net=round(total / quantite, 4),
        montant=total,
    )


def parse_109(texte: str) -> list[Article]:

    articles = []

    m = re.search(MOTIF_DEVIS, texte)
    devis = m.group(1).replace(" ", "") if m else ""

    lignes = lignes_propres(texte)

    blocs_commande = scan_ancre(lignes, MARQUEURS_TVA, OFFSETS_COMMANDE)
    blocs_devis_bpu = scan_ancre(lignes, MARQUEURS_TVA, OFFSETS_DEVIS_BPU)

    for bloc_c, bloc_d in zip(blocs_commande, blocs_devis_bpu):

        ref_c = (bloc_c["reference_fournisseur"] or "").strip()
        ref_d = (bloc_d["reference_fournisseur"] or "").strip()

        if MOTIF_REF_COMMANDE.match(ref_c):
            article = _bloc_vers_article(bloc_c, devis, ref_c)
        elif MOTIF_REF_DEVIS_BPU.match(ref_d):
            article = _bloc_vers_article(bloc_d, devis, ref_d)
        else:
            article = None

        if article:
            articles.append(article)

    _autocontrole_total_ht(texte, articles)

    return articles


# Déclaration pour le chargement automatique
FOURNISSEURS = ['109 DISTRIBUTION']
parse = parse_109
