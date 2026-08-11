"""
Parser COMINTER (et COMINTER MAYOTTE).

Le texte extrait du PDF est éclaté : chaque champ sur sa propre ligne,
dans l'ordre :
    Référence
    Désignation (1 à 2 lignes)
    Montant HT €
    Quantité
    Prix Unit € [remise %]
    Code TVA (chiffre)
    Prix Net €
    Conditionnement (Unité / Boîte de 1 / Sachet de / Barre)
"""

import re

from moteur.modele import Article

# --- GABARIT (Cominter / Cominter Mayotte) ---------------------------------
# Deux formats de PDF coexistent chez Cominter (v1, v2 ci-dessous) : une
# logique procédurale réelle de "quel format essayer", pas une donnée — elle
# reste en Python explicite (voir CLAUDE.md). Constantes alignées sur le
# même bandeau que les autres fournisseurs, par cohérence de lecture.
_REF = re.compile(r"^[A-Z][A-Z0-9/\-]{3,}$")
_MONEY = re.compile(r"^\d[\d\s]*,\d{2}\s*€$")
_MONEY_REM = re.compile(r"^(\d[\d\s]*,\d{2})\s*€(?:\s+(\d+)%)?$")
_QTE = re.compile(r"^\d[\d\s]*,\d{2}$")

_UNITES = {
    "UNITÉ": "UN",
    "UNITE": "UN",
    "BOÎTE": "BTE",
    "BOITE": "BTE",
    "SACHET": "SCH",
    "BARRE": "BARRE",
}
# --- fin GABARIT (suite : _CDT et les deux stratégies plus bas) ------------


def _f(v: str) -> float:
    return float(v.replace(" ", "").replace("\u202f", "").replace("€", "").replace(",", "."))


def _unite(ligne: str) -> str:

    mot = ligne.strip().upper().split()

    if not mot:
        return ""

    return _UNITES.get(mot[0], mot[0])


def _parse_cominter_v1(texte: str) -> list[Article]:

    articles = []

    devis = ""

    m = re.search(r"Devis\s*:\s*([A-Z]{3}\d+)", texte)

    if m:
        devis = m.group(1)

    lignes = [l.rstrip() for l in texte.splitlines()]

    i = 0
    n = len(lignes)

    while i < n:

        ligne = lignes[i].strip()

        # Une référence candidate, hors éco-taxes
        if not _REF.match(ligne) or ligne.startswith("ECO"):
            i += 1
            continue

        ref = ligne

        # Désignation : lignes jusqu'au montant (max 4 lignes)
        j = i + 1
        designation = []

        while j < n and j <= i + 4 and not _MONEY.match(lignes[j].strip()):
            designation.append(lignes[j].strip())
            j += 1

        if j >= n or not _MONEY.match(lignes[j].strip()):
            i += 1
            continue

        try:

            montant = _f(lignes[j].strip())

            quantite = _f(lignes[j + 1].strip())

            m_prix = _MONEY_REM.match(lignes[j + 2].strip())
            prix_brut = _f(m_prix.group(1)) if m_prix else 0.0

            # lignes[j+3] = code TVA (1 chiffre)
            prix_net = _f(lignes[j + 4].strip())

            unite = _unite(lignes[j + 5]) if j + 5 < n else ""

        except (ValueError, AttributeError, IndexError):
            i += 1
            continue

        articles.append(
            Article(
                fournisseur="COMINTER",
                devis=devis,
                reference_fournisseur=ref,
                reference_distributeur="",
                designation=" ".join(designation),
                quantite=quantite,
                unite=unite,
                prix_brut=prix_brut,
                prix_net=prix_net,
                montant=montant,
            )
        )

        i = j + 5

    return articles



_CDT = ("Unité", "Barre", "Boîte de 1", "Sachet de", "Rouleau", "Mètre")


def _parse_cominter_v2(texte: str) -> list[Article]:
    """
    Nouveau format Cominter : par article, ordre
        Qté, Px unitaire, Montant HT€, [code TVA], Désignation(1-2 l),
        Cdt (Unité/Barre...), Référence
    On s'ancre sur la ligne Cdt : la référence suit, la désignation et le
    bloc prix précèdent.
    """
    articles = []

    m = re.search(r"Devis\s*:\s*([A-Z]{3}\d+)", texte)
    devis = m.group(1) if m else ""

    lignes = [l.rstrip() for l in texte.splitlines()]
    n = len(lignes)

    for i, l in enumerate(lignes):

        if l.strip() not in _CDT and not l.strip().startswith(("Boîte", "Sachet", "Barre", "Unité", "Rouleau")):
            continue

        if i + 1 >= n:
            continue

        ref = lignes[i + 1].strip()
        if not _REF.match(ref) or ref.startswith("ECO"):
            continue

        # Remonter : désignation jusqu'au montant €
        j = i - 1
        design = []
        montant = None
        while j >= 0 and j > i - 6:
            lg = lignes[j].strip()
            if _MONEY.match(lg):
                montant = _f(lg)
                break
            design.append(lg)
            j -= 1

        if montant is None or j < 2:
            continue

        # Une éventuelle ligne "code TVA" (1 chiffre) juste après le montant
        if design and re.fullmatch(r"\d", design[-1]):
            design.pop()

        try:
            prix = _f(lignes[j - 1].strip())
            quantite = _f(lignes[j - 2].strip())
        except (ValueError, IndexError):
            continue

        articles.append(
            Article(
                fournisseur="COMINTER",
                devis=devis,
                reference_fournisseur=ref,
                reference_distributeur="",
                designation=" ".join(reversed(design)),
                quantite=quantite,
                unite=_unite(lignes[i]),
                prix_brut=prix,
                prix_net=prix,
                montant=montant,
            )
        )

    return articles


def parse_cominter(texte: str) -> list[Article]:
    """Essaie les deux formats et retourne le plus complet."""
    v1 = _parse_cominter_v1(texte)
    v2 = _parse_cominter_v2(texte)
    return v1 if len(v1) >= len(v2) else v2



# "COMINTER MAYOTTE" est géré par moteur/fournisseurs/cominter_mayotte.py :
# entité distincte (SIRET, adresse, structure de devis différents), voir
# le bandeau GABARIT de ce module.
FOURNISSEURS = ['COMINTER']
parse = parse_cominter
