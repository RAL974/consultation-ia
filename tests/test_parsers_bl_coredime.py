"""
Tests de non-régression du parsing BL Coredime (Rapprochement AI, session
R2 suite — fournisseur n°2), à partir de VRAIS BL scannés
(tests/fixtures/bl_coredime_*.pdf, copies de a_traiter/BL/). Même esprit
que tests/test_parsers_bl_dist109.py.

Particularité Coredime : ces BL réels n'affichent AUCUN prix (réglé à la
facture) — pas d'autocontrôle Total HT possible ici, contrairement à 109
Distribution.
"""

import functools

from moteur.ocr import mots_document, texte_ocr
from moteur.detecteur import detecter_fournisseur
from moteur.fournisseurs.coredime import parse_bl_coredime
from moteur.rapprochement.modele_bl import LigneBL

from conftest import FIXTURES


@functools.lru_cache(maxsize=None)
def _mots(nom):
    return mots_document(FIXTURES / nom)


def test_detection_fournisseur_bl_coredime():
    assert detecter_fournisseur(texte_ocr(FIXTURES / "bl_coredime_1.pdf")) == "COREDIME"


def test_parse_bl_coredime_1_separateur_ocr_perdu():
    """Cas réel : l'OCR a lu "BC123097" (séparateur perdu) au lieu de
    "BC 123.097" — repli sur MOTIF_COMMANDE_COREDIME_SANS_SEPARATEUR
    (découpage 3+3)."""

    bl = parse_bl_coredime(_mots("bl_coredime_1.pdf"))

    assert bl.fournisseur == "COREDIME"
    assert bl.numero_commande == "123.097"
    assert bl.numero_bl == "CORB032399"
    assert bl.date_bl == "06/08/2026"

    assert bl.lignes == [
        LigneBL(reference_fournisseur='LEG031919', designation='COLSON NOIR 9X357', quantite_livree=100.0, prix_net=None, montant=None),
        LigneBL(reference_fournisseur='LEG031955', designation='EMBASENOIREACHEVILLE', quantite_livree=100.0, prix_net=None, montant=None),
    ]


def test_parse_bl_coredime_2_quantite_corrompue_par_la_coche():
    """Cas réel important : la coche "livré" imprimée dans la cellule
    Quantité est lue par l'OCR comme un chiffre collé au vrai nombre
    ("300" -> "3007"). La quantité correcte (300) est retrouvée via la
    mention de confirmation "<qté> X 1 unite" imprimée plus loin sur la
    même ligne ("39% 300x1unite"), pas via la cellule Quantité elle-même."""

    bl = parse_bl_coredime(_mots("bl_coredime_2.pdf"))

    assert bl.numero_commande == "129.049"
    assert bl.numero_bl == "CORB032820"
    assert bl.date_bl == "11/08/2026"

    assert bl.lignes == [
        LigneBL(reference_fournisseur='227059360', designation='BOITEDE100DRIVATP12', quantite_livree=300.0, prix_net=None, montant=None),
    ]


def test_parse_bl_coredime_3():
    bl = parse_bl_coredime(_mots("bl_coredime_3.pdf"))

    assert bl.numero_commande == "131.153"
    assert bl.numero_bl == "CORB032442"
    assert bl.date_bl == "06/08/2026"

    assert bl.lignes == [
        LigneBL(reference_fournisseur='LEG031916', designation='COLSONNOIR9X262', quantite_livree=400.0, prix_net=None, montant=None),
        LigneBL(reference_fournisseur='LEG031919', designation='COLSONNOIR9X357', quantite_livree=400.0, prix_net=None, montant=None),
        LigneBL(reference_fournisseur='LEG031955', designation='EMBASENOIREACHEVILLE', quantite_livree=500.0, prix_net=None, montant=None),
    ]


def test_parse_bl_coredime_4_numero_commande_code_chantier_ref_a_livrer_directement():
    """N° de commande au format "M2.16.011" (code chantier, comme le cas
    109 Distribution trouvé en recette) + mention "Ref à livrer
    directement" : l'article EST bien livré (seul le prix est différé à la
    facture), ne doit PAS être exclu."""

    bl = parse_bl_coredime(_mots("bl_coredime_4.pdf"))

    assert bl.numero_commande == "M2.16.011"
    assert bl.numero_bl == "CORB032040.2"
    assert bl.date_bl == "05/08/2026"

    assert bl.lignes == [
        LigneBL(reference_fournisseur='705001039', designation='TPC50ROUGEC25M', quantite_livree=50.0, prix_net=None, montant=None),
    ]


def test_parse_bl_coredime_5_ref_alphanumerique_et_ligne_ecotaxe_ignoree():
    """Référence "LBCLASTD02" (8 lettres + 2 chiffres, plus permissive que
    les codes vus jusque-là) + une ligne ECO-23 (écotaxe, pas un article
    commandable) qui ne doit pas ressortir comme une ligne livrée."""

    bl = parse_bl_coredime(_mots("bl_coredime_5.pdf"))

    assert bl.numero_commande == "M3.23.031"
    assert bl.numero_bl == "CORB032112"
    assert bl.date_bl == "05/08/2026"

    assert bl.lignes == [
        LigneBL(reference_fournisseur='LEG411651', designation='DX3-ID2P63AA30MATGA', quantite_livree=15.0, prix_net=None, montant=None),
        LigneBL(reference_fournisseur='LBCLASTD02', designation='E278.5W840LMX5P19590005', quantite_livree=12.0, prix_net=None, montant=None),
    ]


def test_parse_bl_coredime_6_designation_sur_deux_lignes():
    """Cas réel : la désignation déborde sur une 2e ligne ("Autoris.
    OMA2500081", numéro d'autorisation) APRÈS la ligne qté/unité — cette
    2e ligne ne doit pas être prise pour un article à part (aucune
    référence valide en 1ère cellule), et la quantité "1" mal lue "i x 1
    unite" par l'OCR doit quand même être retrouvée (0 ligne extraite
    avant correction)."""

    bl = parse_bl_coredime(_mots("bl_coredime_6.pdf"))

    assert bl.numero_commande == "139.107"
    assert bl.numero_bl == "CORB032551"
    assert bl.date_bl == "06/08/2026"

    assert bl.lignes == [
        LigneBL(reference_fournisseur='LEG040596', designation='DLMOSAIC2CDFLASHROUGE', quantite_livree=1.0, prix_net=None, montant=None),
    ]


def test_parse_bl_coredime_7_reste_a_livrer_multi_page():
    """1er cas réel de BL sur PLUSIEURS PAGES avec "Reste à livrer" —
    BUG RÉEL CORRIGÉ : `reste_a_livrer` était un drapeau global à tout le
    document ; une fois activé en page 1, il restait actif pour la page 2
    aussi, excluant à tort des lignes pourtant livrées (page 2 réimprime
    les mêmes articles que "reste à livrer" en page 1, mais CETTE FOIS
    avec leur propre confirmation "<qté> x 1 unite" — donc bien livrés).
    Réinitialisé désormais à chaque page.

    2e bug réel corrigé sur ce même document : "10" (quantité de
    LEG031490) lu "lo" par l'OCR (le "1" de tête confondu avec un "l"
    collé au "0" suivant, pas un "l" isolé comme le cas déjà connu) —
    élargit le remplacement à un "l" en tout début de mot suivi d'un
    chiffre ou d'un "o"/"O"."""

    bl = parse_bl_coredime(_mots("bl_coredime_7_reste_a_livrer_multi_page.pdf"))

    assert bl.numero_commande == "M3.23.043"

    assert bl.lignes == [
        LigneBL(reference_fournisseur='LEG411651', designation='DX3-ID2P63AA30MATGA', quantite_livree=11.0, prix_net=None, montant=None),
        LigneBL(reference_fournisseur='LEG405209', designation='KIT 1OBORNES SORTIEHAUTE', quantite_livree=2.0, prix_net=None, montant=None),
        LigneBL(reference_fournisseur='LEG031490', designation='SORT.DE CABLES 20/32A', quantite_livree=10.0, prix_net=None, montant=None),
        LigneBL(reference_fournisseur='LBCLASTD02', designation='E278.5W840LMX5P19590005', quantite_livree=12.0, prix_net=None, montant=None),
    ]


def test_parse_bl_coredime_8_entete_code_aricle():
    """Cas réel (commande 142.041) : l'en-tête "Code article" est lu
    "Code aricle" par l'OCR (le "T" disparu) — l'ancre de début de tableau
    ne matchait alors plus DU TOUT, faisant disparaître TOUTE la ligne (0
    ligne extraite pour un article pourtant simple, confirmation "10 X 1
    unite" bien lisible par ailleurs). "T" rendu optionnel dans l'ancre."""

    bl = parse_bl_coredime(_mots("bl_coredime_8_entete_code_aricle.pdf"))

    assert bl.numero_commande == "142.041"
    assert bl.numero_bl == "CORB034121"

    assert bl.lignes == [
        LigneBL(reference_fournisseur='LBCLASTD02', designation='E278.5W840LMX5P19590005', quantite_livree=10.0, prix_net=None, montant=None),
    ]
