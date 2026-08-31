"""
Tests de non-régression du parsing BL YESSS ÉLECTRIQUE (Rapprochement AI,
nouveau fournisseur — un seul vrai BL vu à ce jour,
tests/fixtures/bl_yesss_1.pdf, copie de a_traiter/BL/). Même esprit que
tests/test_parsers_bl_coredime.py.
"""

import functools

from moteur.ocr import mots_document, texte_ocr
from moteur.detecteur import detecter_fournisseur
from moteur.fournisseurs.yesss import parse_bl_yesss
from moteur.rapprochement.modele_bl import LigneBL

from conftest import FIXTURES


@functools.lru_cache(maxsize=None)
def _mots(nom):
    return mots_document(FIXTURES / nom)


def test_detection_fournisseur_bl_yesss():
    assert detecter_fournisseur(texte_ocr(FIXTURES / "bl_yesss_1.pdf")) == "YESSS"


def test_parse_bl_yesss_1():
    """Structure inhabituelle : texte pivoté à 90°, chaque valeur retrouvée
    par proximité à son label plutôt que par un ordre de lecture haut/bas
    classique (voir bandeau de moteur/fournisseurs/yesss.py). Quantité
    déduite de Montant/Prix net (78.13/39.07), pas lue directement (la
    cellule Qté est ambiguë sur ce document, un emplacement de 2e ligne
    vide "0" juste à côté de la vraie valeur "2")."""

    bl = parse_bl_yesss(_mots("bl_yesss_1.pdf"))

    assert bl.fournisseur == "YESSS"
    assert bl.numero_commande == "M4.273"
    assert bl.numero_bl == "CAM/040759"
    assert bl.date_bl == "24/08/2026"

    assert bl.lignes == [
        LigneBL(
            reference_fournisseur="411651",
            designation="DX3-ID 2P 63AA 30MA TGA",
            quantite_livree=2.0, prix_net=39.07, montant=78.13,
        ),
    ]


def test_parse_bl_yesss_2_commande_bc_no_et_date_groupee():
    """2e vrai BL vu (`BL M4.276.pdf`) : DEUX formats différents du 1er
    document sur la MÊME structure de gabarit.

    Commande imprimée "BC N°M4.276" (un seul mot OCR, sans espace entre
    "N°" et la valeur) — différent du 1er document ("N° commande M4.273",
    "commande" en toutes lettres). Le label "N° commande" existe bien
    aussi ici, mais comme label de colonne isolé, sans valeur adjacente.

    Date imprimée "25 aout 2026" en un seul mot bien formé (jour, mois et
    année ensemble) — sur le 1er document, le jour ressortait comme un mot
    séparé du mois+année. MOTIF_MOIS_ANNEE_BL_YESSS capture désormais un
    jour optionnel directement dans le même mot, avant de retomber sur la
    recherche par proximité si absent (comportement du 1er document
    inchangé, voir test_parse_bl_yesss_1)."""

    bl = parse_bl_yesss(_mots("bl_yesss_2_commande_bc_no_et_date_groupee.pdf"))

    assert bl.fournisseur == "YESSS"
    assert bl.numero_commande == "M4.276"
    assert bl.numero_bl == "CAM/040794"
    assert bl.date_bl == "25/08/2026"

    assert bl.lignes == [
        LigneBL(
            reference_fournisseur="E2BPP",
            designation="Emetteur 2 Canaux Radio Radio",
            quantite_livree=1.0, prix_net=44.78, montant=44.78,
        ),
    ]
