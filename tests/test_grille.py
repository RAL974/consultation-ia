"""
moteur/grille.py (étape 2, voir CLAUDE.md) — sur un vrai PDF (règle d'or du
projet), pas du texte inventé. Valide que mots()/lignes() reconstruisent
l'ordre VISUEL réel d'un document dont le flux PyMuPDF natif est scramblé
(moteur.fournisseurs.coredime, GABARIT FACTURE : une ligne "Remise" imprimée
à un endroit du texte totalement déconnecté de sa ligne d'article, mais
toujours juste EN DESSOUS d'elle visuellement).
"""

from pathlib import Path

import fitz

from moteur import grille

FIXTURES = Path(__file__).parent / "fixtures"


def test_lignes_reconstruit_lordre_visuel_reel():
    """facture_coredime_6108846_remise_double_x3.pdf : 3 articles + leur
    confirmation de quantité + leur ligne "Remise" forment, dans le texte
    PyMuPDF natif, un flux scramblé (le bloc [référence+désignation+qté]
    ne précède pas directement son "Remise" dans le texte linéaire) — mais
    visuellement, la ligne "Remise" de chaque article est TOUJOURS
    immédiatement en dessous de sa confirmation de quantité, elle-même
    juste en dessous de la ligne article. Vérifié ici en reconstruisant les
    6 lignes utiles et leur ORDRE Y exact (verrouille aussi le calcul de
    tolérance par demi-hauteur médiane, qui doit regrouper "LEG031916 COLSON
    NOIR 9X262 400 U* 0,3400" en une seule ligne sans la scinder ni la
    fusionner avec sa voisine)."""

    doc = fitz.open(FIXTURES / "facture_coredime_6108846_remise_double_x3.pdf")
    page = doc[0]

    lignes = grille.lignes(grille.mots(page))
    lignes_utiles = [
        " ".join(m["texte"] for m in ligne)
        for ligne in lignes
        if any(ref in " ".join(m["texte"] for m in ligne) for ref in ("LEG0319", "Remise"))
    ]

    assert lignes_utiles == [
        "LEG031916 COLSON NOIR 9X262 400 U* 0,3400",
        "Remise 35,00+31,00% 0,1525 61,00 1",
        "LEG031919 COLSON NOIR 9X357 400 U* 0,4600",
        "Remise 35,00+27,00% 0,2183 87,32 1",
        "LEG031955 EMBASE NOIRE A CHEVILLE 500 U* 0,2200",
        "Remise 35,00+32,00% 0,0972 48,60 6",
    ]

    doc.close()


def test_lignes_vide_sur_page_sans_mots():
    assert grille.lignes([]) == []


def test_mots_extrait_le_texte_et_les_coordonnees():
    doc = fitz.open(FIXTURES / "facture_coredime_6108846_remise_double_x3.pdf")
    page = doc[0]

    mots = grille.mots(page)

    assert len(mots) > 0
    assert all({"texte", "x0", "y0", "x1", "y1"} <= set(m) for m in mots)
    assert any(m["texte"] == "LEG031916" for m in mots)

    doc.close()
