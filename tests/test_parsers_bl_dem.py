"""
Tests de non-régression du parsing BL DEM (Rapprochement AI — fournisseur
déjà couvert côté devis, nouveau côté BL), à partir de VRAIS BL scannés
(tests/fixtures/bl_dem_*.pdf, copies de a_traiter/BL/). Même esprit que
tests/test_parsers_bl_coredime.py.
"""

import functools

from moteur.ocr import mots_document, texte_ocr
from moteur.detecteur import detecter_fournisseur
from moteur.fournisseurs.dem import parse_bl_dem
from moteur.rapprochement.modele_bl import LigneBL

from conftest import FIXTURES


@functools.lru_cache(maxsize=None)
def _mots(nom):
    return mots_document(FIXTURES / nom)


def test_detection_fournisseur_bl_dem():
    assert detecter_fournisseur(texte_ocr(FIXTURES / "bl_dem_2_six_lignes.pdf")) == "DEM"


def test_parse_bl_dem_1_deux_pages_reste_a_livrer():
    """Cas réel (commande M3.14.363) : le fichier contient DEUX pages, CHACUNE
    un bon de livraison DEM DIFFÉRENT (n° de BL et date différents sur
    chaque page — pas un même BL étalé sur 2 pages) : parse_bl_dem()
    retourne donc une LISTE de 2 BonLivraison.

    La page 2 (BL 706990, du 20/08) contient en plus une ligne "Reste à
    livrer" : la référence FILSYT15P0.9T y est listée SANS prix ni montant
    (250.00 U, désignation collée sur la même ligne visuelle faute de
    place prise par les colonnes de prix vides) — exclue. Cette même
    référence/quantité réapparaît PRICÉE sur la page 1 (BL 706992, du
    24/08, la livraison qui la solde) : preuve concrète que l'exclusion
    est correcte."""

    bl_page1, bl_page2 = parse_bl_dem(_mots("bl_dem_1_deux_pages_reste_a_livrer.pdf"))

    assert bl_page1.numero_commande == "M3.14.363"
    assert bl_page1.numero_bl == "706992"
    assert bl_page1.date_bl == "24/08/2026"
    assert bl_page1.total_ht_affiche == 237.5
    assert bl_page1.lignes == [
        LigneBL(reference_fournisseur='FILSYT15P0.9T', designation='SYT1 5P0.9 T500M', quantite_livree=250.0, prix_net=0.95, montant=237.5),
    ]

    assert bl_page2.numero_commande == "M3.14.363"
    assert bl_page2.numero_bl == "706990"
    assert bl_page2.date_bl == "20/08/2026"
    assert bl_page2.total_ht_affiche == 87.5
    assert bl_page2.lignes == [
        LigneBL(reference_fournisseur='FILSYT110P0.9T5', designation='SYT1 1OPAIRES 0,9 AWG20', quantite_livree=50.0, prix_net=1.75, montant=87.5),
    ]


def test_parse_bl_dem_2_six_lignes():

    [bl] = parse_bl_dem(_mots("bl_dem_2_six_lignes.pdf"))

    assert bl.numero_commande == "M3.23.046"
    assert bl.numero_bl == "706994"
    assert bl.date_bl == "24/08/2026"
    assert bl.total_ht_affiche == 702.66

    assert bl.lignes == [
        LigneBL(reference_fournisseur='FILVK16BT', designation='VK16BLEUT500', quantite_livree=20.0, prix_net=2.4, montant=48.0),
        LigneBL(reference_fournisseur='FILVK16RT', designation='VK 16 ROUGE T500', quantite_livree=20.0, prix_net=2.4, montant=48.0),
        LigneBL(reference_fournisseur='LEG600049', designation='GRIFFEPROF40MM', quantite_livree=30.0, prix_net=0.31, montant=9.3),
        LigneBL(reference_fournisseur='LEG031490', designation='SORT.DE CABLES 20/32A95X95MM', quantite_livree=10.0, prix_net=5.41, montant=54.1),
        LigneBL(reference_fournisseur='EUR53037', designation='KITPTCTREE27D67+DOUILLE', quantite_livree=20.0, prix_net=5.911, montant=118.22),
        LigneBL(reference_fournisseur='EUR53098', designation='BTECTRE DCL SAILLIE + PITON +DOUILLEE27', quantite_livree=80.0, prix_net=5.313, montant=425.04),
    ]

    total_extrait = round(sum(l.montant for l in bl.lignes), 2)
    assert abs(bl.total_ht_affiche - total_extrait) <= 0.02
