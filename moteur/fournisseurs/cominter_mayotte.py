"""
Parser COMINTER MAYOTTE (Cominter, Mamoudzou — contact@cominter.yt).

Entité distincte de Cominter Réunion (SIRET, adresse et e-mail différents) :
la question posée en début de session ("le format est-il identique ?") a une
réponse claire — NON. La structure de devis est différente de celle de
`moteur/fournisseurs/cominter.py` (v1/v2), d'où ce module séparé.

Écrit sur 1 vrai PDF (tests/fixtures/cominter_mayotte.pdf) : à confirmer sur
un 2e devis Mayotte si l'acheteur en dépose un (remise différente,
plusieurs pages...).

LIMITE CONNUE (voir CLAUDE.md) : sur ce PDF, le DERNIER article de la page
("L76565") a ses valeurs numériques (qté, prix, remise, montant) extraites
AVANT sa référence/désignation dans le flux de texte, au lieu d'après comme
partout ailleurs — un bloc orphelin apparaît en tête de document
("38,00"/"16,78"/"30%"/"446,35€") qui lui appartient réellement (38 x
16,78 x 70 % = 446,35€, exact). Cause probable : ordre d'extraction PyMuPDF
différent pour le dernier bloc d'une page. Un seul exemple observé -> pas
de règle générale à en tirer (règle d'or) ; l'article correspondant est
signalé "bloc incomplet" plutôt que rattaché à tort. Total du devis
(9 630,36€) donc supérieur de 446,35€ à la somme des lignes extraites —
visible en comparant au PDF, pas de contrôle automatique fiable possible
sur ce seul exemple.
"""

import re

from moteur.modele import Article
from moteur.outils import to_float, lignes_propres

# --- GABARIT (Cominter Mayotte) ---------------------------------------------
# Bloc ancré sur le marqueur "Unité" (toujours seul sur sa ligne, une fois
# par article) :
#     Référence           (avant "Unité", motif ^L\d+$ ou "ZZ" — code
#                           générique pour un article non catalogué, la
#                           vraie référence se retrouve alors en tête de
#                           la désignation)
#     Désignation          (1 à 3 lignes, jusqu'à "Unité")
#     Unité                (marqueur)
#     Qté
#     Px unitaire
#     [Remise %]            (facultative : absente sur certaines lignes,
#                            alors Montant = Qté x Px unitaire directement)
#     Montant (€)
MARQUEUR = "Unité"
MOTIF_REF = re.compile(r"^L\d+$|^ZZ$")
MOTIF_MONEY = re.compile(r"^[\d\s]+,\d{2}\s*€$")
MOTIF_DEVIS = r"Devis\s*:\s*([A-Z0-9]+)"
# --- fin GABARIT -------------------------------------------------------------


def parse_cominter_mayotte(texte: str) -> list[Article]:

    articles = []

    m = re.search(MOTIF_DEVIS, texte)
    devis = m.group(1) if m else ""

    lignes = lignes_propres(texte)
    n = len(lignes)

    for i, ligne in enumerate(lignes):

        if ligne != MARQUEUR:
            continue

        # Remonte jusqu'à la référence précédente
        j = i - 1
        while j >= 0 and not MOTIF_REF.match(lignes[j]):
            j -= 1

        if j < 0:
            continue

        ref = lignes[j]
        designation = " ".join(lignes[j + 1:i]).strip()

        # Avance jusqu'au montant (€) : qté, puis 1 (px unitaire seul) ou
        # 2 (px unitaire + remise %) valeurs intermédiaires.
        k = i + 1
        if k >= n:
            continue
        qte_brute = lignes[k]

        valeurs = []
        k += 1
        while k < n and not MOTIF_MONEY.match(lignes[k]) and len(valeurs) < 2:
            valeurs.append(lignes[k])
            k += 1

        if k >= n or not MOTIF_MONEY.match(lignes[k]):
            print(f"Erreur lecture article COMINTER MAYOTTE ({ref}) : bloc incomplet")
            continue

        montant = to_float(lignes[k])
        quantite = to_float(qte_brute)
        prix_unitaire = to_float(valeurs[0]) if valeurs else 0.0

        articles.append(
            Article(
                fournisseur="COMINTER MAYOTTE",
                devis=devis,
                reference_fournisseur=ref,
                reference_distributeur="",
                designation=designation,
                quantite=quantite,
                unite="UN",
                prix_brut=prix_unitaire,
                prix_net=round(montant / quantite, 4) if quantite else 0.0,
                montant=montant,
            )
        )

    return articles


# Déclaration pour le chargement automatique
FOURNISSEURS = ['COMINTER MAYOTTE']
parse = parse_cominter_mayotte
