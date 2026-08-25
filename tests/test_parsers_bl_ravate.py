"""
Tests de non-régression du rapprochement Ravate Elec (Rapprochement AI —
nouveau fournisseur, à partir de VRAIES pièces scannées, tests/fixtures/
bl_ravate_*.jpg, copies de a_traiter/BL/).

Les 2 seules pièces réelles disponibles à ce jour sont TOUTES DEUX du
"ravatelec" (branding visible sur le document) — Ravate Pro reste NON
couvert côté BL faute de pièce réelle (voir moteur/fournisseurs/ravate.py).
"""

import functools

from moteur.ocr import mots_document, texte_ocr
from moteur.detecteur import detecter_fournisseur
from moteur.fournisseurs.ravate import parse_bl_ravate
from moteur.rapprochement.modele_bl import LigneBL

from conftest import FIXTURES


@functools.lru_cache(maxsize=None)
def _mots(nom):
    return mots_document(FIXTURES / nom)


def test_detection_fournisseur_bl_ravate():
    assert detecter_fournisseur(texte_ocr(FIXTURES / "bl_ravate_1.jpg")) == "RAVATE"


def test_parse_bl_ravate_1_cable_au_metre():
    """Câble facturé au mètre : la quantité (100) n'apparaît PAS dans la
    ligne chiffrée elle-même mais sur la ligne suivante ("R2V3G2.5T1 |
    :MT:100.00") — déduite ici de Montant / Px net (115,00 / 1,15 = 100),
    jamais lue directement. La vraie Référence fournisseur ("R2V3G2.5T1")
    est elle aussi sur cette ligne SUIVANTE (pas précédente comme dans les
    autres cas, voir bandeau GABARIT BL) — repli look-ahead, confirmé en
    recette réelle (commande 123.095 : Suivi attend bien "R2V3G2.5T1", pas
    "44200019" le Code Art)."""

    bl = parse_bl_ravate(_mots("bl_ravate_1.jpg"))

    assert bl.fournisseur == "RAVATE"
    assert bl.numero_commande == "123.095"
    assert bl.date_bl == "06/08/2026"
    assert bl.total_ht_affiche == 115.0

    assert bl.lignes == [
        LigneBL(reference_fournisseur='R2V3G2.5T1', designation='/CABLE U1O00 RO2VCU 3G2.ST1O00', quantite_livree=100.0, prix_net=1.15, montant=115.0),
    ]


def test_parse_bl_ravate_2_quantite_inline_avec_unite():
    """Article à l'unité : la quantité (3) est CETTE FOIS directement dans
    la ligne chiffrée (cellule ":UN: 3,00:"), contrairement au cas
    précédent — vérifie que la déduction Montant/Px net (57,57 / 19,19 = 3)
    donne le même résultat sans dépendre de cette cellule. La Référence
    fournisseur ("069864") est elle aussi sur la ligne SUIVANTE, comme pour
    bl_ravate_1 (repli look-ahead)."""

    bl = parse_bl_ravate(_mots("bl_ravate_2.jpg"))

    assert bl.numero_commande == "M3.23.036"
    assert bl.date_bl == "10/08/2026"
    assert bl.total_ht_affiche == 57.57

    # "1.683 Kg" (poids du colis) est imprimé sur la même ligne visuelle que
    # la désignation, donc capturé avec elle — sans conséquence pour le
    # rapprochement (seuls référence/qté/prix sont utilisés).
    assert bl.lignes == [
        LigneBL(reference_fournisseur='069864', designation='POUSSOIR INV LUM P-E BLANC EN 1.683 Kg', quantite_livree=3.0, prix_net=19.19, montant=57.57),
    ]


def test_parse_bl_ravate_3_plusieurs_lignes_code_art_et_ref_fournisseur_separes():
    """6 articles réels : sur ce document, le "Code Art" (interne Ravate,
    ex. "100017709") est sur la ligne désignation, la vraie "Référence
    fournisseur" (ex. "VK16BT") en tête de la ligne chiffrée — la
    désignation contient donc le Code Art en préfixe (sans conséquence,
    non utilisé pour le rapprochement)."""

    bl = parse_bl_ravate(_mots("bl_ravate_3_multi_lignes.pdf"))

    assert bl.numero_commande == "M3.23.033"
    assert bl.date_bl == "05/08/2026"
    assert bl.total_ht_affiche == 314.45

    refs_et_montants = [(l.reference_fournisseur, l.quantite_livree, l.montant) for l in bl.lignes]
    assert refs_et_montants == [
        ('VK16BT', 20.0, 45.0),
        ('VK16RT', 20.0, 45.0),
        ('404926', 30.0, 64.5),
        ('600002', 10.0, 44.5),
        ('VU1.5VJ', 1.0, 19.5),
        ('069864', 5.0, 95.95),
    ]


def test_parse_bl_ravate_4_ref_fournisseur_en_tete_de_ligne_chiffree():
    """Même article que bl_ravate_2 (POUSSOIR, 069864), mais sur ce scan
    c'est directement la Référence fournisseur ("069864L") qui est en tête
    de la ligne chiffrée — pas le Code Art comme sur bl_ravate_2 — la
    logique doit utiliser cette référence directement, sans chercher de
    ligne isolée."""

    bl = parse_bl_ravate(_mots("bl_ravate_4_ref_fournisseur_avant_ligne_chiffree.pdf"))

    assert bl.numero_commande == "M3.23.036"
    assert bl.date_bl == "10/08/2026"

    assert bl.lignes == [
        LigneBL(reference_fournisseur='069864L', designation='100151156 /POUSSOIRINVLUM P-EBLANCEN 1.683 Kg', quantite_livree=3.0, prix_net=19.19, montant=57.57),
    ]


def test_parse_bl_ravate_5_virgule_lue_asterisque():
    """Cas réel : la virgule du Px Brut est lue "*" par l'OCR ("99*91" au
    lieu de "99,91") — sans tolérance, la ligne chiffrée n'était pas
    reconnue (ses 4 dernières cellules ne parsaient pas toutes comme un
    montant) et la ligne entière disparaissait silencieusement."""

    bl = parse_bl_ravate(_mots("bl_ravate_5_virgule_lue_asterisque.pdf"))

    assert bl.numero_commande == "M3.10.182"
    assert bl.total_ht_affiche == 877.5

    assert bl.lignes == [
        LigneBL(reference_fournisseur='R2V5G10', designation='44200039 CABLEU1QOORO2V CU5G10T500M', quantite_livree=130.0, prix_net=6.75, montant=877.5),
    ]


def test_parse_bl_ravate_6_huit_lignes():
    """8 articles réels sur un même BL, dont une ligne dont la désignation
    déborde sur la ligne d'anomalie "DISPOS A CONFIRMER DEM NORD" (repli du
    fournisseur pour un article en rupture, imprimé juste avant la ligne
    désignation suivante) — n'empêche pas la bonne extraction de la ligne."""

    bl = parse_bl_ravate(_mots("bl_ravate_6_huit_lignes.pdf"))

    assert bl.numero_commande == "M2.16.010"
    assert bl.total_ht_affiche == 1139.75

    refs_et_montants = [(l.reference_fournisseur, l.montant) for l in bl.lignes]
    assert refs_et_montants == [
        ('XVR1IISTI', 580.0),
        ('NFT820', 63.49),
        ('BDF925F', 190.11),
        ('GD213A', 26.5),
        ('GP213P', 17.9),
        ('SBN440', 88.75),
        ('KJ125B', 35.0),
        ('R2V5G6', 138.0),
    ]


def test_parse_bl_ravate_7_tiret_dans_la_commande():
    """Cas réel (commande 131.162) : la commande collée directement après
    l'année sans espace ("2026131-162") contient un TIRET — sans "-" dans
    la classe de caractères du motif de capture, celui-ci s'arrêtait juste
    avant et le \\s*$ final ne matchait plus DU TOUT, faisant échouer toute
    la ligne (n° de commande ET exclusion de cette ligne hors désignation,
    qui réutilise le même motif — la ligne entière se retrouvait accumulée
    en tête de la désignation du 1er article)."""

    bl = parse_bl_ravate(_mots("bl_ravate_7_tiret_dans_commande.jpg"))

    assert bl.numero_commande == "131.162"
    assert bl.numero_bl == "00312608LV0124"
    assert bl.date_bl == "12/08/2026"

    assert bl.lignes == [
        LigneBL(reference_fournisseur='A270285', designation='100015885 CHEV TAP-VIS 5/0/25 SEAU 2500', quantite_livree=1.0, prix_net=70.0, montant=70.0),
    ]


def test_parse_bl_ravate_8_confusion_ocr_8_b_et_qte_unite_collees():
    """Cas réel (commande M3.10.182, même famille que l'exemple ci-dessus) :
    "8" lu "B" par l'OCR, y compris DANS le n° de commande lui-même
    ("M3.10.1B2" pour "M3.10.182" — confirmé par le nom du fichier BL).
    Cette pièce a aussi la quantité+unité collées dans UNE SEULE cellule
    juste avant Px Brut (":MT:130,00:"), décalant la fenêtre des 4
    dernières cellules — repli sur les 3 dernières (Px Brut, Px Net,
    Montant) — et Px Net/Remises collés SANS espace dans une même cellule
    ("6,75:1288,30") — repli combo dédié. Qté déduite de Montant/Px Net =
    877,50/6,75 = 130 pile, cohérent avec le "130,00" imprimé dans la
    cellule qté+unité (jamais lue directement, juste une confirmation)."""

    bl = parse_bl_ravate(_mots("bl_ravate_8_confusion_ocr_8_b_dans_commande.jpg"))

    assert bl.numero_commande == "M3.10.182"
    assert bl.numero_bl == "00312608LV0187"
    assert bl.date_bl == "18/08/2026"
    assert bl.total_ht_affiche == 877.5

    assert bl.lignes == [
        LigneBL(reference_fournisseur='R2V5G10', designation='44200039 CABLE U1QOO RO2V CU SG1O T50OM', quantite_livree=130.0, prix_net=6.75, montant=877.5),
    ]


def test_autocontrole_total_ht_coherent_sur_les_bl_ravate():

    noms = [
        "bl_ravate_1.jpg", "bl_ravate_2.jpg", "bl_ravate_3_multi_lignes.pdf",
        "bl_ravate_5_virgule_lue_asterisque.pdf", "bl_ravate_6_huit_lignes.pdf",
        "bl_ravate_8_confusion_ocr_8_b_dans_commande.jpg",
    ]
    for nom in noms:
        bl = parse_bl_ravate(_mots(nom))
        total_extrait = round(sum(l.montant for l in bl.lignes), 2)
        assert abs(bl.total_ht_affiche - total_extrait) <= 0.02, nom
