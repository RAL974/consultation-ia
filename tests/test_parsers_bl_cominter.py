"""
Tests de non-régression du parsing BL Cominter Ouest (Rapprochement AI,
session R2 suite — fournisseur n°4), à partir de VRAIS BL scannés
(tests/fixtures/bl_cominter_*.pdf, copies de a_traiter/BL/).

Particularité signalée par l'acheteur AVANT tout code, décisive : un même
fichier peut contenir PLUSIEURS BL scannés à la suite (jusqu'à 8 vus en
session), mais un même BL peut aussi déborder sur 2 pages — parse_bl_cominter
retourne donc une LISTE de BonLivraison, pas un seul (voir
moteur.ocr.grouper_pages_par_identifiant, testé séparément dans
tests/test_ocr.py).
"""

import functools

from moteur.ocr import mots_document, texte_ocr
from moteur.detecteur import detecter_fournisseur
from moteur.fournisseurs.cominter import parse_bl_cominter
from moteur.rapprochement.modele_bl import LigneBL

from conftest import FIXTURES


@functools.lru_cache(maxsize=None)
def _mots(nom):
    return mots_document(FIXTURES / nom)


def test_detection_fournisseur_bl_cominter():
    assert detecter_fournisseur(texte_ocr(FIXTURES / "bl_cominter_1.pdf")) == "COMINTER"


def test_parse_bl_cominter_1():
    [bl] = parse_bl_cominter(_mots("bl_cominter_1.pdf"))

    assert bl.fournisseur == "COMINTER"
    assert bl.numero_commande == "M3.23.030"
    assert bl.numero_bl == "OBL108106"
    assert bl.date_bl == "06/08/2026"

    assert bl.lignes == [
        LigneBL(reference_fournisseur='IMP221-E26', designation='Emboutssimplespourlanguetteporte-reperes,', quantite_livree=1.0, prix_net=9.27, montant=9.27),
        LigneBL(reference_fournisseur='L6008.', designation='DooxiePlaque1posteBlanc', quantite_livree=35.0, prix_net=0.58, montant=20.3),
        LigneBL(reference_fournisseur='BLM68--7', designation='Boiteencastrement67MultimateriauxNoAirPt', quantite_livree=22.0, prix_net=3.45, montant=75.9),
        LigneBL(reference_fournisseur='L600049', designation='DooxiegriffeRapidoprofondeur4Omm', quantite_livree=30.0, prix_net=0.29, montant=8.7),
        LigneBL(reference_fournisseur='ECLHUB047', designation='Hublot0265mmE27220VACIP65BlancRond', quantite_livree=10.0, prix_net=13.7, montant=137.0),
        LigneBL(reference_fournisseur='L69731L', designation='Prise2P+T saillie Plexo gris', quantite_livree=10.0, prix_net=5.83, montant=58.3),
        LigneBL(reference_fournisseur='L69831L', designation='Prise2P+TencastrePlexogris', quantite_livree=10.0, prix_net=7.49, montant=74.9),
        LigneBL(reference_fournisseur='E55317', designation='AppliqueSdbTubeLED+PC445mm7W400Lm', quantite_livree=10.0, prix_net=32.13, montant=321.3),
        LigneBL(reference_fournisseur='L405209', designation='BornedeconnexionDx3-ID63A', quantite_livree=10.0, prix_net=4.21, montant=42.1),
    ]


def test_parse_bl_cominter_2_reliquat_meme_commande_bl_different():
    """Cas réel : même commande M3.23.030 que bl_cominter_1, mais un BL
    (OBL) différent — reliquat/2e livraison de la même commande."""

    [bl] = parse_bl_cominter(_mots("bl_cominter_2.pdf"))

    assert bl.numero_commande == "M3.23.030"
    assert bl.numero_bl == "OBL108251"
    assert bl.date_bl == "11/08/2026"

    assert bl.lignes == [
        LigneBL(reference_fournisseur='BLM680527', designation='Boiteencastrement 67MultimateriauxNoAirPt Ctre', quantite_livree=6.0, prix_net=3.45, montant=20.7),
        LigneBL(reference_fournisseur='L4052', designation='BornedeconnexionDx3-ID63A', quantite_livree=5.0, prix_net=4.21, montant=21.05),
    ]


def test_parse_bl_cominter_3_code_tva_colle_au_montant():
    """Cas réel : le code TVA final est parfois collé à la cellule Montant
    ("41,10E 2" au lieu de deux cellules séparées) selon la ligne."""

    [bl] = parse_bl_cominter(_mots("bl_cominter_3.pdf"))

    assert bl.numero_commande == "M3.23.034"
    assert bl.numero_bl == "OBL108187"
    assert bl.date_bl == "10/08/2026"

    assert bl.lignes == [
        LigneBL(reference_fournisseur='L411651', designation='Interdifferentiel2P63AtypeAarriveehautaviset', quantite_livree=1.0, prix_net=42.15, montant=42.15),
        LigneBL(reference_fournisseur='L405205', designation='BornedeconnexionDx3-ID63A', quantite_livree=3.0, prix_net=4.21, montant=12.63),
        LigneBL(reference_fournisseur='L404926', designation='PeignedalimentationP+N13modules', quantite_livree=10.0, prix_net=2.23, montant=22.3),
        LigneBL(reference_fournisseur='L600801', designation='DooxiePlaque1posteBlanc', quantite_livree=3.0, prix_net=0.58, montant=1.74),
        LigneBL(reference_fournisseur='E55317', designation='AppliqueSdbTubeLED+PC445mm7W400Lm', quantite_livree=3.0, prix_net=32.13, montant=96.39),
        LigneBL(reference_fournisseur='ECLHUB047', designation='Hublot0265mmE27220VACIP65BlancRond', quantite_livree=3.0, prix_net=13.7, montant=41.1),
        LigneBL(reference_fournisseur='SY0029651', designation='LampeLEDTOLEDOGLSA608W806LM4000K', quantite_livree=1.0, prix_net=11.59, montant=11.59),
        LigneBL(reference_fournisseur='L69831L', designation='Prise2P+TencastrePlexogris', quantite_livree=2.0, prix_net=7.49, montant=14.98),
        LigneBL(reference_fournisseur='L69731L', designation='Prise2P+Tsaillie Plexogris', quantite_livree=2.0, prix_net=5.83, montant=11.66),
    ]


def test_parse_bl_cominter_4_remise_et_px_net_colles_dans_la_meme_cellule():
    """Cas réel : la remise % et le Px net sont collés dans la MÊME
    cellule OCR ("30% 110,67" au lieu de deux cellules séparées) — le Px
    net (110,67) doit être extrait de cette cellule combinée, pas confondu
    avec le Px unitaire (158,10) d'avant."""

    [bl] = parse_bl_cominter(_mots("bl_cominter_4.pdf"))

    assert bl.numero_commande == "M3.10.171"
    assert bl.numero_bl == "OBL108154"
    assert bl.date_bl == "07/08/2026"

    assert bl.lignes == [
        LigneBL(reference_fournisseur='AEA9P22610', designation='IDT401P+N10AA9P22610', quantite_livree=1.0, prix_net=13.97, montant=13.97),
        LigneBL(reference_fournisseur='AEA9Y13625', designation='VIGIITG401P+N25A300MAA9Y13625', quantite_livree=1.0, prix_net=110.67, montant=110.67),
        LigneBL(reference_fournisseur='ECLLED778820', designation='PlafonnierLEDHublotrond20W2000lm@300mm', quantite_livree=1.0, prix_net=25.59, montant=25.59),
    ]


def test_parse_bl_cominter_5_meme_bl_sur_2_pages_un_seul_document():
    """Cas réel : le tableau d'articles tient sur la page 1, le pied de
    page des totaux déborde sur la page 2 (même n° OBL sur les deux) —
    doit ressortir comme UN SEUL BonLivraison, pas deux. Verrouille aussi
    la quantité de la ligne L86001L à 15.0 pile (pas 15,01 : cette valeur
    doit venir de la cellule Qté imprimée "15,00", jamais d'une division
    Montant/Px net qui introduit du bruit d'arrondi sur les lignes
    remisées — voir bandeau GABARIT BL de moteur/fournisseurs/cominter.py)."""

    bls = parse_bl_cominter(_mots("bl_cominter_5_meme_bl_2pages.pdf"))

    assert len(bls) == 1
    assert bls[0].numero_bl == "OBL108056"
    assert bls[0].date_bl == "05/08/2026"
    assert len(bls[0].lignes) == 16

    quantites = {l.reference_fournisseur: l.quantite_livree for l in bls[0].lignes}
    assert quantites["L86001L"] == 15.0


def test_parse_bl_cominter_6_plusieurs_bl_dans_un_seul_fichier():
    """Cas réel signalé par l'acheteur AVANT tout code : 8 BL différents
    scannés à la suite dans un seul fichier PDF — doit ressortir comme 8
    BonLivraison distincts, chacun avec son propre n° de commande/OBL."""

    bls = parse_bl_cominter(_mots("bl_cominter_6_multi_bl_8pages.pdf"))

    assert len(bls) == 8

    numeros_bl = [bl.numero_bl for bl in bls]
    assert numeros_bl == [
        "OBL108135", "OBL107196", "OBL107273", "OBL108041",
        "OBL108042", "OBL107471", "OBL107938", "OBL108293",
    ]

    dates = [bl.date_bl for bl in bls]
    assert dates == [
        "06/08/2026", "", "07/07/2026", "04/08/2026",
        "04/08/2026", "15/07/2026", "31/07/2026", "12/08/2026",
    ]

    # "131.155" pas "E131.155" : l'OCR colle le mot de chantier au n° de
    # commande sans espace ("LAGOURGUE131.155") — bug réel corrigé qui
    # faisait disparaître tout ce BL du rapprochement (commande introuvable
    # dans le Suivi).
    commandes = [bl.numero_commande for bl in bls]
    assert commandes == [
        "130.036", "", "13.015.394", "M4.260",
        "M4.258", "108.271", "123.083", "131.155",
    ]

    nb_lignes = [len(bl.lignes) for bl in bls]
    assert nb_lignes == [3, 4, 1, 3, 1, 1, 1, 8]


def test_parse_bl_cominter_6_commande_m4_260_quantites_entieres_et_ligne_recuperee():
    """Cas réel signalé par l'acheteur (recette) : ce BL avait produit des
    quantités livrées non entières pour des articles à l'unité (29,96
    interrupteur au lieu de 30) et perdu silencieusement sa 3e ligne
    (PLW11643, référence renvoyée par l'OCR sur sa PROPRE ligne, après la
    ligne désignation+prix). Les deux sont corrigés : quantités entières
    (cellule Qté imprimée, pas Montant/Px net) et raccord référence/ligne
    suivante."""

    bls = parse_bl_cominter(_mots("bl_cominter_6_multi_bl_8pages.pdf"))
    [bl_m4_260] = [bl for bl in bls if bl.numero_commande == "M4.260"]

    assert bl_m4_260.lignes == [
        LigneBL(reference_fournisseur='L86101L', designation='ASLMECANISMEInterrupteurva-et-vientblanc', quantite_livree=30.0, prix_net=4.17, montant=124.95),
        LigneBL(reference_fournisseur='L86120L', designation='ASLMECANISMEDoubleva-et-vientblanc', quantite_livree=6.0, prix_net=10.7, montant=64.18),
        LigneBL(reference_fournisseur='PLW11643', designation='PointdecentreDCL-avecconnect.etcrochetfixat.-Pour', quantite_livree=30.0, prix_net=15.37, montant=460.95),
    ]

    [bl_m4_258] = [bl for bl in bls if bl.numero_commande == "M4.258"]
    assert bl_m4_258.lignes == [
        LigneBL(reference_fournisseur='PLW11821', designation='Cadre simpleAppareillage SaillieLeg-Pourmoulure', quantite_livree=150.0, prix_net=2.62, montant=392.7),
    ]


def test_parse_bl_cominter_7_remise_et_px_net_colles_sans_espace():
    """BUG RÉEL CORRIGÉ (signalé par l'acheteur — "nous avons le prix
    unitaire, puis la remise, puis le prix net, aucun écart") : contrairement
    à test_parse_bl_cominter_4 ("30% 110,67", AVEC un espace), ce document
    colle la remise et le Px net SANS AUCUN espace ("30%435,94",
    "30%106,50") — le motif combo (qui exigeait \\s+ entre le % et le prix)
    ne matchait alors jamais, et la boucle retombait à tort sur le Px
    UNITAIRE de la cellule précédente (622,77€ au lieu du vrai Px net
    435,94€ pour CAELK2766 ; 152,14€ au lieu de 106,50€ pour CAEACB4V) —
    silencieusement, sans lever d'anomalie (montant et quantité restaient
    cohérents avec le mauvais prix pris isolément)."""

    [bl] = parse_bl_cominter(_mots("bl_cominter_7_remise_sans_espace.pdf"))

    # Le n° de commande n'est pas lisible directement sur ce document (il
    # est déduit du contenu par le rapprochement, voir matching.
    # deduire_commande_par_contenu — hors périmètre du parser lui-même).
    assert bl.numero_commande == ""
    assert bl.numero_bl == "OBL108110"
    assert bl.date_bl == "06/08/2026"

    lignes_par_ref = {l.reference_fournisseur: l for l in bl.lignes}

    assert lignes_par_ref["CAELK2766"] == LigneBL(
        reference_fournisseur='CAELK2766', designation='BAIELINK+600X60027U',
        quantite_livree=1.0, prix_net=435.94, montant=435.94,
    )
    assert lignes_par_ref["CAEACB4V"] == LigneBL(
        reference_fournisseur='CAEACB4V',
        designation='Bloc de 4 ventilateurs pour baie de brassage 19"',
        quantite_livree=1.0, prix_net=106.5, montant=106.5,
    )
