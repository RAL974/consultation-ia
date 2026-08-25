"""
Tests unitaires de moteur/ocr.py sur des mots construits à la main (pas
d'OCR réel ici — voir tests/test_parsers_bl_*.py pour les vrais PDF).
"""

import re

from moteur.ocr import grouper_pages_par_identifiant, pages_par_identifiant


def _mot(texte):
    return {"texte": texte, "score": 0.9, "x0": 0, "y0": 0, "x1": 10, "y1": 10}


def test_grouper_pages_par_identifiant_pages_consecutives_meme_id():
    """Une page sans identifiant (pied de page qui déborde) rejoint le
    groupe précédent plutôt que d'en ouvrir un nouveau — cas réel Cominter
    Ouest, un même BL sur 2 pages."""

    pages = [
        [_mot("Bon de livraison : OBL108056")],
        [_mot("Total HT 123,45")],  # pas d'identifiant sur cette page
    ]

    groupes = grouper_pages_par_identifiant(pages, re.compile(r"OBL(\d{5,7})"))

    assert len(groupes) == 1
    assert len(groupes[0]) == 2


def test_grouper_pages_par_identifiant_pages_avec_id_different():

    pages = [
        [_mot("Bon de livraison : OBL108135")],
        [_mot("Bon de livraison : OBL107196")],
    ]

    groupes = grouper_pages_par_identifiant(pages, re.compile(r"OBL(\d{5,7})"))

    assert len(groupes) == 2
    assert len(groupes[0]) == 1
    assert len(groupes[1]) == 1


def test_grouper_pages_par_identifiant_compare_le_groupe_captant_pas_le_texte_entier():
    """Cas réel : l'OCR lit "OBL108056" sur une page et "0BL108056" (O
    confondu avec 0) sur une autre page du MÊME document — la comparaison
    doit se faire sur les chiffres captés (groupe 1), pas le texte brut du
    match, sinon un même BL est coupé en deux à tort."""

    pages = [
        [_mot("Bon de livraison : OBL108056")],
        [_mot("Bon de livraison : 0BL108056")],
    ]

    groupes = grouper_pages_par_identifiant(pages, re.compile(r"[0O]BL(\d{5,7})", re.IGNORECASE))

    assert len(groupes) == 1
    assert len(groupes[0]) == 2


def test_pages_par_identifiant_memes_groupes_que_grouper_pages_mais_en_indices():
    """pages_par_identifiant() doit produire EXACTEMENT le même
    regroupement que grouper_pages_par_identifiant(), mais en indices de
    page (0-based) plutôt qu'en mots — utilisé pour découper le PDF
    source par BL individuel (voir moteur.rapprochement.pipeline_bl)."""

    pages = [
        [_mot("Bon de livraison : OBL108135")],
        [_mot("Bon de livraison : OBL107196")],
        [_mot("Total HT 123,45")],  # pas d'identifiant -> rejoint le groupe précédent
    ]
    motif = re.compile(r"OBL(\d{5,7})")

    assert pages_par_identifiant(pages, motif) == [[0], [1, 2]]
