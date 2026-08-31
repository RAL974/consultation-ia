"""
Tests de non-régression du rapprochement Electric Plus / GMR (Rapprochement
AI, session R2 suite — fournisseur n°3), à partir de VRAIES FACTURES
scannées (tests/fixtures/bl_electricplus_*.pdf, copies de a_traiter/BL/).

Particularité : GMR n'envoie jamais de bon de livraison séparé (confirmé
par l'acheteur) — le rapprochement se fait directement à partir de ses
factures. Comme 109 Distribution, ces factures affichent un "Total HT" :
autocontrôle possible (contrairement à Coredime).
"""

import functools

from moteur.ocr import mots_document, texte_ocr
from moteur.detecteur import detecter_fournisseur
from moteur.fournisseurs.electricplus import parse_bl_electricplus
from moteur.rapprochement.modele_bl import LigneBL

from conftest import FIXTURES


@functools.lru_cache(maxsize=None)
def _mots(nom):
    return mots_document(FIXTURES / nom)


def test_detection_fournisseur_bl_electricplus():
    assert detecter_fournisseur(texte_ocr(FIXTURES / "bl_electricplus_1.pdf")) == "ELECTRIC PLUS"


def test_parse_bl_electricplus_1():
    [bl] = parse_bl_electricplus(_mots("bl_electricplus_1.pdf"))

    assert bl.fournisseur == "ELECTRIC PLUS"
    assert bl.numero_commande == "M3.10.172"
    assert bl.numero_bl == "1207031"
    assert bl.date_bl == "07/08/2026"
    assert bl.total_ht_affiche == 279.22

    assert bl.lignes == [
        LigneBL(reference_fournisseur='LEG062525', designation='BAES EVACIP43 SATI CONNECTE', quantite_livree=6.0, prix_net=40.43, montant=242.58),
        LigneBL(reference_fournisseur='LEG411617', designation='DISJ DX3-ID2P 40A30MATG', quantite_livree=1.0, prix_net=36.64, montant=36.64),
    ]


def test_parse_bl_electricplus_2():
    [bl] = parse_bl_electricplus(_mots("bl_electricplus_2.pdf"))

    assert bl.numero_commande == "M3.23.035"
    assert bl.numero_bl == "1207039"
    assert bl.date_bl == "10/08/2026"
    assert bl.total_ht_affiche == 62.16

    assert bl.lignes == [
        LigneBL(reference_fournisseur='LEG004107', designation='SONNERIE 230-240V4VA 3,.00', quantite_livree=3.0, prix_net=11.27, montant=33.81),
        LigneBL(reference_fournisseur='LEG069601L', designation='PLEXOBOITIER1P2XEMBOUT1E', quantite_livree=3.0, prix_net=1.82, montant=5.46),
        LigneBL(reference_fournisseur='LEG600018', designation='DOOXIE POUSSOIRLUM SYM SONN', quantite_livree=3.0, prix_net=7.63, montant=22.89),
    ]


def test_parse_bl_electricplus_3_reference_coupee_en_deux_cellules():
    """Cas réel : la référence "LEG004107" est lue coupée en deux cellules
    OCR ("LE" + "4107") sur ce scan précis — reconstituée en "LE4107". Le
    cœur numérique ("4107") reste identique à la vraie référence
    "LEG004107" (les zéros de tête "004" sont perdus mais le cœur ne
    change pas), donc le rapprochement contre le Suivi (moteur.base.
    coeur_numerique) fonctionne quand même malgré la référence imparfaite —
    voir tests/test_rapprochement_matching.py."""

    [bl] = parse_bl_electricplus(_mots("bl_electricplus_3.pdf"))

    assert bl.numero_commande == "M3.23.032"
    assert bl.numero_bl == "1206995"
    assert bl.date_bl == "05/08/2026"
    assert bl.total_ht_affiche == 103.6

    assert bl.lignes == [
        LigneBL(reference_fournisseur='LE4107', designation='SONNERIE230-240V4VA', quantite_livree=5.0, prix_net=11.27, montant=56.35),
        LigneBL(reference_fournisseur='LEG069601L', designation='PLEXOBOITIER1P2XEMBOUT1E', quantite_livree=5.0, prix_net=1.82, montant=9.1),
        LigneBL(reference_fournisseur='LEG600018', designation='DOOXIE POUSSOIR LUM SYM SONN', quantite_livree=5.0, prix_net=7.63, montant=38.15),
    ]


def test_parse_bl_electricplus_4():
    [bl] = parse_bl_electricplus(_mots("bl_electricplus_4.pdf"))

    assert bl.numero_commande == "M3.14.354"
    assert bl.numero_bl == "1207011"
    assert bl.date_bl == "06/08/2026"
    assert bl.total_ht_affiche == 11.86

    assert bl.lignes == [
        LigneBL(reference_fournisseur='PLA11584', designation='ANGLE PLAT MOULUREKEVA', quantite_livree=2.0, prix_net=3.22, montant=6.44),
        LigneBL(reference_fournisseur='PLA11587', designation='TE DERIV-POUR MOULURE KEVA', quantite_livree=1.0, prix_net=2.76, montant=2.76),
        LigneBL(reference_fournisseur='PLA11582', designation='ANGLE INTVARIASOUPLE +OU-7D', quantite_livree=1.0, prix_net=2.66, montant=2.66),
    ]


def test_parse_bl_electricplus_6_montant_derive_quand_qte_absente():
    """1 seule ligne, quantité "67.00" avec un POINT décimal (pas la
    virgule habituelle) — la quantité livrée est de toute façon déduite de
    Montant / P.U. net (comme 109 Distribution), donc insensible à ce
    genre de variante OCR sur la cellule Qté elle-même."""

    [bl] = parse_bl_electricplus(_mots("bl_electricplus_6.pdf"))

    assert bl.numero_commande == "M3.15.398"
    assert bl.numero_bl == "1206959"
    assert bl.total_ht_affiche == 771.17

    assert bl.lignes == [
        LigneBL(reference_fournisseur='SCHS521089', designation='ODACE COUR2P+T1 DOUB 1', quantite_livree=67.0, prix_net=11.51, montant=771.17),
    ]


def test_parse_bl_electricplus_7_prix_o_confondu_avec_zero():
    """Cas réel : le prix "0,53 PF" est lu "O,53 PF" (lettre O au lieu du
    chiffre 0) sur la 2e ligne — sans le repli O/0, cette ligne entière
    disparaissait silencieusement (ancre "PF" jamais trouvée)."""

    [bl] = parse_bl_electricplus(_mots("bl_electricplus_7.pdf"))

    assert bl.numero_commande == "135.042"
    assert bl.total_ht_affiche == 42.33

    assert bl.lignes == [
        LigneBL(reference_fournisseur='SCHIMT50620', designation='TUBE TULIPE GRIS D20MM PAR 3M', quantite_livree=51.0, prix_net=0.3, montant=15.3),
        LigneBL(reference_fournisseur='SCHIMT50625', designation='TUBE TULIPE GRIS D25MMPAR3M', quantite_livree=51.0, prix_net=0.53, montant=27.03),
    ]


def test_parse_bl_electricplus_8_prix_r_confondu_avec_f():
    """Cas réel : le prix "1,93 PF" est lu "1,93 PR" (lettre F confondue
    avec R, même famille que la confusion O/0 déjà corrigée sur cette
    ancre) — sans ce repli, l'ancre "PF" n'était jamais trouvée et TOUTE la
    facture ressortait à 0 ligne malgré un Total HT affiché de 482,50€."""

    [bl] = parse_bl_electricplus(_mots("bl_electricplus_8.pdf"))

    assert bl.numero_commande == "162.002"
    assert bl.total_ht_affiche == 482.5

    assert bl.lignes == [
        LigneBL(reference_fournisseur='CAB015010T500', designation='RO2V-CU 5G2,5 T500', quantite_livree=250.0, prix_net=1.93, montant=482.5),
    ]


def test_autocontrole_total_ht_coherent_sur_toutes_les_factures():

    noms = [f"bl_electricplus_{i}.pdf" for i in range(1, 9)]
    for nom in noms:
        [bl] = parse_bl_electricplus(_mots(nom))
        total_extrait = round(sum(l.montant for l in bl.lignes), 2)
        assert abs(bl.total_ht_affiche - total_extrait) <= 0.02, nom


def test_parse_bl_electricplus_9_deux_factures_dans_un_seul_fichier():
    """BUG RÉEL CORRIGÉ (recette réelle, signalé par l'acheteur) : ce
    fichier de 2 pages contient en réalité DEUX FACTURES Electric Plus
    totalement distinctes (commande/n° facture/date différents par page) —
    avant le découpage par page, tout était traité comme un seul document :
    le n° de commande et le total venaient de la page 0, mais les LIGNES
    d'articles étaient accumulées sur les deux pages sans distinction (la
    ligne PLA11527 de la page 0 disparaissait même silencieusement, sa
    cellule prix n'ayant pas le suffixe "PF" habituel — voir cette ligne
    dans les assertions ci-dessous)."""

    bls = parse_bl_electricplus(_mots("bl_electricplus_9_multi_facture_2pages.pdf"))
    assert len(bls) == 2

    facture_m4_263, facture_142_036 = bls

    assert facture_m4_263.numero_commande == "M4.263"
    assert facture_m4_263.numero_bl == "1207019"
    assert facture_m4_263.date_bl == "06/08/2026"
    assert facture_m4_263.total_ht_affiche == 160.8
    assert facture_m4_263.pages == [0]
    assert facture_m4_263.lignes == [
        LigneBL(reference_fournisseur='PLA11527', designation='TEDERIV-POURMOULUREKEVA', quantite_livree=60.0, prix_net=2.68, montant=160.8),
    ]

    assert facture_142_036.numero_commande == "142.036"
    assert facture_142_036.numero_bl == "1207127"
    assert facture_142_036.date_bl == "14/08/2026"
    assert facture_142_036.total_ht_affiche == 180.64
    assert facture_142_036.pages == [1]
    assert facture_142_036.lignes == [
        LigneBL(reference_fournisseur='LEG030654', designation='POINT CENTRE DCL', quantite_livree=20.0, prix_net=7.15, montant=143.0),
        LigneBL(reference_fournisseur='LEG401332', designation='PORTEDRIVIAOPAQUECOFFRET', quantite_livree=3.0, prix_net=12.09, montant=36.27),
        LigneBL(reference_fournisseur='LEG600323', designation='DOOXIE SORTIECABLE IP21 BLANC', quantite_livree=1.0, prix_net=1.37, montant=1.37),
    ]

    # Chaque facture retombe exactement sur son propre Total HT — preuve
    # que le mélange page0/page1 est bien résolu (avant le correctif :
    # 160,80€ affiché contre 180,64€ extrait, un mélange des deux
    # factures).
    for facture in bls:
        total_extrait = round(sum(l.montant for l in facture.lignes), 2)
        assert abs(facture.total_ht_affiche - total_extrait) <= 0.02


def test_parse_bl_electricplus_10_entete_fusionne_avec_1er_article():
    """Cas réel (fichier multi-fournisseurs, commande M4.269) : l'en-tête
    de colonnes ("DESIGNATION QTE PRIX UNIT.HT...") s'est retrouvé groupé
    par l'OCR sur la MÊME ligne visuelle que la référence+désignation du
    1er article ("PLA11525 EMBTMOULUREKEVA32MMX12MM DESIGNATION QTE...").
    Sans correctif, cette ligne échouait entièrement (cellules[-2]/[-1] =
    "MONTANT HT"/"TVA", pas des nombres) et la référence du 1er article se
    retrouvait perdue — la ligne suivante (ses propres nombres) récupérait
    à tort "PVC B ARTIC" (un bout de désignation) comme référence.
    _zone_tableau_electricplus() détache désormais les cellules qui
    précèdent le mot d'en-tête et les reporte sur la ligne suivante.

    Une 3e ligne chiffrée existe sur ce document (montant 56,70€, Total HT
    global 86,40€) mais SANS aucune référence/désignation adjacente
    (regroupement Y de l'OCR défavorable, ou véritablement absente sur ce
    document) — cellules[0] vaut alors la quantité elle-même ("70,00MTR"),
    jamais une vraie référence : _ligne_vers_article_electricplus() refuse
    désormais de produire une ligne dans ce cas plutôt que d'écrire une
    fausse référence — l'écart de Total HT (86,40€ affiché vs 29,70€
    extrait) signale honnêtement qu'une ligne manque."""

    [bl] = parse_bl_electricplus(_mots("bl_electricplus_10_entete_fusionne_1er_article.pdf"))

    assert bl.numero_commande == "M4.269"
    assert bl.total_ht_affiche == 86.4

    assert bl.lignes == [
        LigneBL(reference_fournisseur='PLA11525', designation='EMBTMOULUREKEVA32MMX12MM PVC B ARTIC', quantite_livree=10.0, prix_net=0.55, montant=5.5),
        LigneBL(reference_fournisseur='CAB012000T', designation='RO2V-CU 2X1,5', quantite_livree=10.0, prix_net=2.42, montant=24.2),
    ]

    total_extrait = round(sum(l.montant for l in bl.lignes), 2)
    assert abs(bl.total_ht_affiche - total_extrait) > 0.02
