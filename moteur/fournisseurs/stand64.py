"""
Parser Stand 64 (STAND 64, luminaires - prescription).

Format difficile : la référence est souvent un placeholder (ZARTICLENP =
article non prédéfini par le commercial, absent du catalogue fournisseur ;
ZARTICLETVA = éco-contribution). Le produit réel n'est décrit que dans le
texte. Colonnes en ordre inversé, nombre de colonnes numériques variable
(éco-part présente ou non).

Ancre fiable : après la description viennent, dans l'ordre,
    Qté, code TVA (C0/C1/C4/C7), Total HT, P.U Net, [autres nombres], Référence

**Placeholder ZARTICLENP/"Alternative:ZARTICLENP" : la vraie référence
fabricant, quand elle existe, est alors donnée SEULEMENT entre parenthèses
en fin de désignation** (ex. "...SD-WOOD RING SUSPENSION ... IP20 BOIS
(KUBIA-ART00031180)"), pas dans le champ référence — cas réel, chantier
Cosinus (`tests/fixtures/stand64_cosinus.pdf`). Sans extraire ce code, TOUTES
les lignes personnalisées d'un même devis partagent la même "référence" et
s'écrasent entre elles au comparateur (34 lignes -> 5 conservées, constaté
sur ce PDF) : `_reference_reelle()` la récupère quand le champ référence
imprimé est un tel placeholder ; sinon la référence imprimée fait foi, comme
avant.

PARSER MEILLEUR EFFORT : à contrôler visuellement. Les lignes
"ECO CONTRIBUTION" (ZARTICLETVA) sont ignorées. Un garde-fou
qté x P.U Net ≈ Total HT écarte les lectures douteuses.
"""

import re

from moteur.modele import Article
from moteur.outils import to_float, lignes_propres

_NUM = re.compile(r"^\s*\d[\d\s]*,\d{2}\s*$")
_TVA = re.compile(r"^C\d$")
_MARQUEUR = re.compile(r"^(REPERE|REPRE|VARIANTE|DETECTEUR|PRIX NETS)", re.I)
_CODE_INTEGRE = re.compile(r"\(([A-Z0-9][A-Z0-9\-\.]{2,})\)\s*$")


def _reference_reelle(ref: str, designation: str) -> str:
    """Récupère le code fabricant intégré en fin de désignation quand la
    référence imprimée est un placeholder générique ("ZARTICLENP",
    "Alternative:ZARTICLENP"...) — sinon la référence imprimée telle quelle."""

    if "ZARTICLE" not in ref.upper():
        return ref

    m = _CODE_INTEGRE.search(designation)
    return m.group(1) if m else ref


def parse_stand64(texte: str) -> list[Article]:

    articles = []

    m = re.search(r"Devis n°\s*([\d\s]+?)\s+du", texte)
    devis = m.group(1).replace(" ", "") if m else ""

    lignes = lignes_propres(texte)
    n = len(lignes)

    # Début après l'en-tête "Description"
    debut = 0
    for k, l in enumerate(lignes):
        if l == "Description":
            debut = k + 1
            break

    i = debut
    desc_buffer = []

    while i < n:
        l = lignes[i]

        # Pied de page légal répété à chaque page : sauter jusqu'à la
        # reprise du tableau ("Description") sur la page suivante, ou finir.
        if l.startswith(("Le transfert", "Total Eco-part", "Total HT",
                         "Mode de paiement", "Signature")):
            desc_buffer = []
            p = i + 1
            while p < n and lignes[p] != "Description":
                p += 1
            if p >= n:
                break
            i = p + 1
            continue

        # Marqueurs de repère / variante : réinitialisent la description
        if _MARQUEUR.match(l):
            desc_buffer = []
            i += 1
            continue

        # Début d'un bloc chiffré : Qté (nombre) suivi d'un code TVA
        if _NUM.match(l) and i + 1 < n and _TVA.match(lignes[i + 1].strip()):
            try:
                quantite = to_float(l)
                total = to_float(lignes[i + 2])
                prix_net = to_float(lignes[i + 3])

                # Avancer jusqu'à la référence (1re ligne non numérique)
                j = i + 4
                while j < n and _NUM.match(lignes[j]):
                    j += 1
                ref = lignes[j] if j < n else ""
            except Exception:
                desc_buffer = []
                i += 1
                continue

            designation = " ".join(desc_buffer).strip()

            # Ignorer l'éco-contribution
            est_eco = ref.upper() == "ZARTICLETVA" or "ECO CONTRIBUTION" in designation.upper()

            # Garde-fou cohérence
            coherent = quantite and prix_net and abs(
                total - quantite * prix_net
            ) <= max(0.02 * total, 1)

            if not est_eco and designation and coherent:
                articles.append(
                    Article(
                        fournisseur="STAND 64",
                        devis=devis,
                        reference_fournisseur=_reference_reelle(ref, designation),
                        reference_distributeur="",
                        designation=designation,
                        quantite=quantite,
                        unite="UN",
                        prix_brut=prix_net,
                        prix_net=prix_net,
                        montant=total,
                    )
                )

            desc_buffer = []
            i = j + 1
            continue

        # Sinon : ligne de description
        desc_buffer.append(l)
        i += 1

    return articles


FOURNISSEURS = ["STAND 64"]
parse = parse_stand64
