"""
Fonctions utilitaires communes aux parsers fournisseurs.

Mutualise ce qui était dupliqué dans chaque parser : conversion des
nombres au format français, nettoyage des lignes, recherche du numéro
de devis.
"""

import re


def to_float(valeur: str) -> float:
    """'1 234,56' -> 1234.56 ; '' -> 0.0"""
    if valeur is None:
        return 0.0
    v = (
        str(valeur)
        .replace("\u202f", "")
        .replace("\xa0", "")
        .replace(" ", "")
        .replace("€", "")
        .replace(",", ".")
        .strip()
    )
    try:
        return float(v)
    except ValueError:
        return 0.0


def lignes_propres(texte: str) -> list[str]:
    """Découpe le texte en lignes non vides, espaces de bord retirés."""
    return [l.strip() for l in texte.splitlines() if l.strip()]


def chercher_devis(texte: str, motif: str) -> str:
    """Retourne le premier numéro de devis correspondant au motif, ou ''."""
    m = re.search(motif, texte)
    return m.group(0) if m else ""
