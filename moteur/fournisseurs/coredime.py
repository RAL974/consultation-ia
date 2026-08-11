import re

from moteur.modele import Article
from moteur.outils import to_float as _f, chercher_devis
from moteur.fournisseurs._gabarit import scan_regex

# --- GABARIT (Coredime) ---------------------------------------------------
# Ligne type (colonnes à espaces variables, espace possible en tête) :
#     LEG411651   DISPO DX3-ID 2P 63A A 30MA TGA   15  U*            42,0000  630,00 6
#     SCHDZ5CA162 DISPO EMBOUT REP C.16MM2        100  40%  0,7600    0,4560   45,60 6
#
# - DISPO / AEC en tête de désignation = disponibilité
# - Remise éventuelle (40%) suivie du prix de base, avant le prix net
# - Unités : U* M* B* UN (quirk d'origine : seuls U*/M* sont normalisés en
#   UN/MT ; B* reste "B" tel quel — comportement préservé, pas "corrigé")
# - Les lignes ECOTAXE (ECO-...) sont ignorées
# - Garde-fou existant : la ligne est écartée (silencieusement) si
#   qté×prix net s'écarte de plus de 5 % (ou 1€) du montant lu — redondant
#   depuis l'ajout de l'autocontrôle global (moteur/autocontrole.py), mais
#   laissé tel quel : le retirer changerait quelles lignes sont gardées.
MOTIF_DEVIS = r"COR\s+B\d+"
MOTIF_LIGNE = re.compile(
    r"^\s*([A-Z][A-Z0-9]+)\s+"                  # référence
    r"(.+?)\s+"                                  # désignation (avec DISPO/AEC)
    r"(?:\*{3}\s+)?"                              # *** éventuel (ecotaxe)
    r"([\d]+)\s+"                                 # quantité
    r"(?:([UMB]\*|UN|BTE)\s+)?"                   # unité (parfois absente)
    r"(?:(\d+)%\s+([\d,]+)\s+)?"                  # remise % + prix base (facultatif)
    r"([\d,]+)\s+"                                # prix net
    r"([\d\s,]+?)\s+"                             # montant
    r"(\d)\s*$"                                   # code TVA
)
# --- fin GABARIT -----------------------------------------------------------


def parse_coredime(texte: str):

    articles = []

    devis = chercher_devis(texte, MOTIF_DEVIS)

    for _i, m in scan_regex(texte.splitlines(), MOTIF_LIGNE):

        ref = m.group(1)

        if ref.startswith("ECO"):
            continue

        designation = m.group(2).strip()
        quantite = float(m.group(3))
        unite_brute = m.group(4) or "UN"
        unite = unite_brute[0] if unite_brute.endswith("*") else unite_brute
        prix_base = _f(m.group(6)) if m.group(6) else None
        prix_net = _f(m.group(7))
        montant = _f(m.group(8))

        dispo = ""

        if designation.startswith("DISPO "):
            dispo = "DISPO"
            designation = designation[6:]

        elif designation.startswith("AEC "):
            dispo = "AEC"
            designation = designation[4:]

        # Garde-fou : le montant doit correspondre à qté x prix net (à 5 % près)
        if quantite and prix_net and abs(montant - quantite * prix_net) > max(
            0.05 * montant, 1.0
        ):
            continue

        articles.append(
            Article(
                fournisseur="COREDIME",
                devis=devis,
                reference_fournisseur=ref,
                reference_distributeur="",
                designation=designation,
                quantite=quantite,
                unite="UN" if unite == "U" else ("MT" if unite == "M" else unite),
                prix_brut=prix_base if prix_base else prix_net,
                prix_net=prix_net,
                montant=montant,
                disponibilite=dispo,
            )
        )

    return articles


# Déclaration pour le chargement automatique
FOURNISSEURS = ['COREDIME']
parse = parse_coredime
