"""
Tests de non-régression du parsing BL 109 Distribution (Rapprochement AI,
session R2), à partir de VRAIS BL scannés (tests/fixtures/bl_dist109_*.pdf,
copies de a_traiter/BL/ — voir CLAUDE.md, tableau de flux R1). Comme pour
les devis (tests/test_parsers.py) : sortie ACTUELLE figée champ par champ,
plus lente que les tests devis (OCR réel, ~10-30s/fixture).

parse_bl_109() retourne une LISTE (pas un seul BonLivraison) : un même
fichier peut contenir PLUSIEURS BL scannés à la suite ("scans en masse",
voir test_parse_bl_dist109_8_multi_bl_8pages) — même principe que Cominter
Ouest (moteur.ocr.pages_par_identifiant).
"""

import functools

from moteur.ocr import mots_document, texte_ocr
from moteur.detecteur import detecter_fournisseur
from moteur.fournisseurs.dist109 import parse_bl_109
from moteur.rapprochement.modele_bl import LigneBL

from conftest import FIXTURES


@functools.lru_cache(maxsize=None)
def _mots(nom):
    return mots_document(FIXTURES / nom)


def test_detection_fournisseur_bl_dist109():
    assert detecter_fournisseur(texte_ocr(FIXTURES / "bl_dist109_1.pdf")) == "109 DISTRIBUTION"


def test_parse_bl_dist109_1():
    [bl] = parse_bl_109(_mots("bl_dist109_1.pdf"))

    assert bl.fournisseur == "109 DISTRIBUTION"
    assert bl.numero_commande == "123.096"
    assert bl.numero_bl == "735136"
    assert bl.date_bl == "06/08/2026"
    assert bl.total_ht_affiche == 64.0

    assert bl.lignes == [
        LigneBL(reference_fournisseur='81000298', designation='KitBoulonPVCMBMetatraySACHETDE100', quantite_livree=2.0, prix_net=32.0, montant=64.0),
    ]


def test_parse_bl_dist109_2():
    [bl] = parse_bl_109(_mots("bl_dist109_2.pdf"))

    assert bl.numero_commande == "131.156"
    assert bl.numero_bl == "736366"
    assert bl.date_bl == "10/08/2026"
    assert bl.total_ht_affiche == 108.4

    assert bl.lignes == [
        LigneBL(reference_fournisseur='52302', designation='COUVERCLEFINITIONAVECVISDIAM80', quantite_livree=32.0, prix_net=0.95, montant=30.4),
        LigneBL(reference_fournisseur='53041', designation='CouvercledappliqueDCL85mmpourboite simple', quantite_livree=32.0, prix_net=1.5, montant=48.0),
        LigneBL(reference_fournisseur='70002', designation='BARETTEDOMINO6MM2', quantite_livree=20.0, prix_net=1.5, montant=30.0),
    ]

    # Cas réel important : la coche "livré" imprimée dans la cellule Qté est
    # lue par l'OCR comme un chiffre collé au vrai nombre ("32" -> "327"),
    # ce qui rendrait une lecture directe de la cellule fausse. La quantité
    # est donc déduite de Total / P.U.Net (30,40 / 0,95 = 32), pas de la
    # cellule Qté elle-même — la 2e ligne (327 -> 32) est le cas qui l'a
    # révélé.
    assert bl.lignes[1].quantite_livree == 32.0


def test_parse_bl_dist109_3_quantite_manquante_a_ocr():
    """Sur ce BL, l'OCR ne détecte AUCUNE cellule Qté du tout (case à
    cocher probablement fusionnée avec une autre cellule) : la ligne reste
    exploitable uniquement parce que la quantité est déduite de Total /
    P.U.Net plutôt que lue directement — sinon la ligne serait perdue."""

    [bl] = parse_bl_109(_mots("bl_dist109_3.pdf"))

    assert bl.numero_commande == "132.008"
    assert bl.numero_bl == "736611"
    assert bl.total_ht_affiche == 70.0

    assert bl.lignes == [
        LigneBL(reference_fournisseur='CAP856574', designation='MANCHONPREDALLEHT65ENTR74(PAR50)', quantite_livree=50.0, prix_net=1.4, montant=70.0),
    ]


def test_parse_bl_dist109_4():
    [bl] = parse_bl_109(_mots("bl_dist109_4.pdf"))

    assert bl.numero_commande == "142.031"
    assert bl.numero_bl == "735956"
    assert bl.date_bl == "10/08/2026"
    assert bl.total_ht_affiche == 170.0

    assert bl.lignes == [
        LigneBL(reference_fournisseur='86101L', designation='INTERRUPTEURVAETVIENT-POUSSOIROTEO', quantite_livree=20.0, prix_net=4.6, montant=92.0),
        LigneBL(reference_fournisseur='086127L', designation='Prise de courant avec terre-composable - saillie', quantite_livree=20.0, prix_net=3.9, montant=78.0),
    ]


def test_parse_bl_dist109_5_numero_commande_code_chantier():
    """N°Réf.Client au format "M3.14.353" (préfixe lettre + 3 groupes) au
    lieu du format habituel "123.096" (2 groupes) — cas réel trouvé lors de
    la recette avec l'acheteur (session R2 suite), qui a fait échouer
    l'extraction jusqu'à l'élargissement de MOTIF_COMMANDE_BL."""

    [bl] = parse_bl_109(_mots("bl_dist109_5.pdf"))

    assert bl.numero_commande == "M3.14.353"
    assert bl.numero_bl == "734931"
    assert bl.date_bl == "05/08/2026"
    assert bl.total_ht_affiche == 25.0

    assert bl.lignes == [
        LigneBL(reference_fournisseur='R2V5G1.5TECC', designation='CoupeCABLEU1000R2V5G1.5mm2-T500', quantite_livree=20.0, prix_net=1.25, montant=25.0),
    ]


def test_parse_bl_dist109_6_entete_avec_accent():
    """BUG RÉEL CORRIGÉ (recette réelle) : l'OCR a lu l'en-tête du tableau
    "Reférencearticle" AVEC un accent (RÉFÉRENCE) au lieu de la forme sans
    accent habituelle — l'ancre MOTIF_ENTETE_TABLEAU_BL ne matchait jamais,
    faisant disparaître TOUT le tableau (0 ligne extraite alors qu'un Total
    HT de 1750€ était bien affiché). Confirme aussi que la quantité réelle
    est 10 (pas 1) : le total 1750€ ne s'explique que par 175€ x 10."""

    [bl] = parse_bl_109(_mots("bl_dist109_6.pdf"))

    assert bl.numero_commande == "M3.14.360"
    assert bl.numero_bl == "737426"
    assert bl.date_bl == "13/08/2026"
    assert bl.total_ht_affiche == 1750.0

    assert bl.lignes == [
        LigneBL(reference_fournisseur='401086', designation='PRO4OW-PROJECTEURLEDSOLAIREEXTERIEUR4OW-MURALOUS', quantite_livree=10.0, prix_net=175.0, montant=1750.0),
    ]


def test_parse_bl_dist109_7_eco_part_intercalee_et_lignes_sans_code_tva():
    """BUG RÉEL CORRIGÉ (recette réelle) : deux défauts distincts sur ce
    document.
    1. Sur la ligne 302304, une cellule "Eco-part" (0.63€, 2 décimales)
       s'intercale entre le vrai P.U.Net (22€, 5 décimales) et le code TVA
       -> l'ancien code prenait l'éco-part pour le P.U.Net, donnant une
       quantité totalement fausse (104,76 au lieu de 3).
    2. 3 lignes sur 7 (302831, 302632, 50033) n'ont AUCUNE cellule de code
       TVA lue par l'OCR -> silencieusement ignorées avant ce correctif.
    Le Total HT affiché (136,96€) correspond exactement à la somme des 7
    lignes une fois toutes correctement extraites."""

    [bl] = parse_bl_109(_mots("bl_dist109_7.pdf"))

    assert bl.numero_commande == "M2.17.005"
    assert bl.numero_bl == "736690"
    assert bl.date_bl == "11/08/2026"
    assert bl.total_ht_affiche == 136.96

    assert bl.lignes == [
        LigneBL(reference_fournisseur='302304', designation='SHARKCCTPROJECTEURNOIR30WIP66-IK043000lm', quantite_livree=3.0, prix_net=22.0, montant=66.0),
        LigneBL(reference_fournisseur='302831', designation='HOLLYWOODIl1200mm32W(28-24-20W)4480lm-IP66-IK10-3CCTMulti-p', quantite_livree=1.0, prix_net=27.5, montant=27.5),
        LigneBL(reference_fournisseur='302632', designation='PRORADMANCHONCONNECTEURNOIR3PIP68ALEVIER2X3POLES', quantite_livree=1.0, prix_net=3.9, montant=3.9),
        LigneBL(reference_fournisseur='50033', designation='BOITEDEDERIVATIONETANCHEIP55B0X80X40', quantite_livree=4.0, prix_net=1.2, montant=4.8),
        LigneBL(reference_fournisseur='16041508', designation='TUBEIRODIAM20-3M', quantite_livree=24.0, prix_net=0.45, montant=10.8),
        LigneBL(reference_fournisseur='R2V3G1.5TECC', designation='CoupeCABLEU1000R2V3G1,5mm2-T500', quantite_livree=26.0, prix_net=0.76, montant=19.76),
        LigneBL(reference_fournisseur='R9PXH213', designation='RESI9-PEIGNEMONOBLOC-1P+N-63A-13MODULES-CACHEDENTS', quantite_livree=1.0, prix_net=4.2, montant=4.2),
    ]


def test_parse_bl_dist109_8_multi_bl_8pages():
    """Cas réel signalé par l'acheteur (nouveau lot, "scans en masse") : 8
    BL différents scannés à la suite dans un seul fichier — doit ressortir
    comme 8 BonLivraison distincts, chacun avec son propre n° de
    commande/BL et ses propres indices de page (bl.pages). Verrouille
    aussi le 2e bug trouvé sur ce fichier : le repère de fin de tableau
    "Total Eco-part HT" est parfois lu de travers par l'OCR ("Tatal
    Eco-part HT", voire pire) sur 2 des 8 pages — sans repli sur "Total
    HT" (repère resté fiable partout), ces 2 BL ressortaient à 0 ligne
    malgré un Total HT bien affiché."""

    bls = parse_bl_109(_mots("bl_dist109_8_multi_bl_8pages.pdf"))

    assert len(bls) == 8
    assert [bl.pages for bl in bls] == [[0], [1], [2], [3], [4], [5], [6], [7]]

    commandes = [bl.numero_commande for bl in bls]
    assert commandes == [
        "142.033", "M3.10.175", "M3.10.175", "M3.10.175",
        "132.008", "142.031", "143.187", "123.098",
    ]

    # BUG RÉEL CORRIGÉ (recette réelle, voir CLAUDE.md "bons de retour") :
    # la page 1 est en réalité un RETOUR (n°25894), pas une 2e page du BL
    # 737760 — avant le correctif, MOTIF_BL_NUMERO_DATE matchait à tort la
    # référence "Bon de livraison n°737760" présente DANS le corps du
    # retour, fusionnant les deux à tort (risque de compter deux fois la
    # même quantité si le retour avait été traité comme une livraison).
    numeros_bl = [bl.numero_bl for bl in bls]
    assert numeros_bl == ["737748", "25894", "737851", "737760", "36611", "735956", "734672", "735348"]

    types = [bl.type_document for bl in bls]
    assert types == ["BL", "RETOUR", "BL", "BL", "BL", "BL", "BL", "BL"]

    assert bls[1].numero_bl_origine == "737760"
    assert bls[1].lignes == [
        LigneBL(reference_fournisseur='R9PRA263', designation='RESI9-INTERDIFFERENTIEL-2P-63A-30MA-TYPEA-PEIGNABLE-AL', quantite_livree=1.0, prix_net=58.5, montant=58.5),
    ]

    # Pages 6 et 7 (734672, 735348) : repère "Total Eco-part HT" mal lu par
    # l'OCR, repli sur "Total HT" — sans ce correctif, 0 ligne extraite.
    assert bls[6].lignes == [
        LigneBL(reference_fournisseur='J08849', designation='CHACHEBORNE-CS-ZLBMIZHBM123-L177', quantite_livree=4.0, prix_net=18.0, montant=72.0),
        LigneBL(reference_fournisseur='1SEP620013R3000', designation='ZLBM3-3P-M12', quantite_livree=4.0, prix_net=485.0, montant=1940.0),
    ]
    assert bls[7].lignes == [
        LigneBL(reference_fournisseur='9894', designation='SCIECLOCHEBETOND67', quantite_livree=1.0, prix_net=125.0, montant=125.0),
    ]


def test_parse_bl_dist109_9_retour_seul():
    """Le même document que la page 1 de bl_dist109_8, isolé dans son
    propre fichier — confirme la détection "Retour n° X du date" et
    l'extraction du numéro du BL d'origine qu'il annule."""

    [bl] = parse_bl_109(_mots("bl_dist109_9_retour.pdf"))

    assert bl.type_document == "RETOUR"
    assert bl.numero_commande == "M3.10.175"
    assert bl.numero_bl == "25894"
    assert bl.date_bl == "14/08/2026"
    assert bl.numero_bl_origine == "737760"
    assert bl.total_ht_affiche == 58.5

    assert bl.lignes == [
        LigneBL(reference_fournisseur='R9PRA263', designation='RESI9-INTERDIFFERENTIEL-2P-63A-30MA-TYPEA-PEIGNABLE-AL', quantite_livree=1.0, prix_net=58.5, montant=58.5),
    ]


def test_parse_bl_dist109_10_bl_avec_retour_associe():
    """Le BL 737760 lui-même (que le retour n°25894 annule en partie sur
    l'article R9PRA263) : reste un type "BL" normal au niveau du parsing
    — c'est au niveau du rapprochement (pipeline_bl.rapprocher_dossier)
    que la ligne R9PRA263 est exclue de l'écriture, pas ici."""

    [bl] = parse_bl_109(_mots("bl_dist109_10_bl_avec_retour_associe.pdf"))

    assert bl.type_document == "BL"
    assert bl.numero_commande == "M3.10.175"
    assert bl.numero_bl == "737760"
    assert bl.numero_bl_origine == ""
    assert bl.total_ht_affiche == 360.6

    assert bl.lignes == [
        LigneBL(reference_fournisseur='R9PRC263', designation='RESI9-INTER DIFFERENTIEL-2P-63A-3OMA-TYPEAC-PEIGNABLE-AI', quantite_livree=2.0, prix_net=48.0, montant=96.0),
        LigneBL(reference_fournisseur='R9PRA263', designation='RESI9-INTERDIFFERENTIEL-2P-63A-30MA-TYPEA-PEIGNABLE-ALI', quantite_livree=1.0, prix_net=58.5, montant=58.5),
        LigneBL(reference_fournisseur='R9PFC620', designation='RESI9-DISJONCTEURMODULAIRE-1P+N-20A-COURBEC-PEIGNABL', quantite_livree=1.0, prix_net=6.5, montant=6.5),
        LigneBL(reference_fournisseur='R9PFC616', designation='RESI9-DISJONCTEURMODULAIRE-1P+N-16A-COURBEC-PEIGNABL', quantite_livree=7.0, prix_net=6.8, montant=47.6),
        LigneBL(reference_fournisseur='7171621', designation='EMBOUTDECABLAGE16mm?BEIGE(sachetde100)', quantite_livree=1.0, prix_net=11.0, montant=11.0),
        LigneBL(reference_fournisseur='61401', designation='PRISE2P+T45X45OPTIMA', quantite_livree=12.0, prix_net=2.95, montant=35.4),
        LigneBL(reference_fournisseur='BC6AFSTL8', designation='CONNECTEURRJ45CAT6ABLINDE360°sanSOutil', quantite_livree=24.0, prix_net=3.95, montant=94.8),
        LigneBL(reference_fournisseur='BC451C8', designation='PLASTRON45X45(Pour1connecteurRJ45)', quantite_livree=12.0, prix_net=0.9, montant=10.8),
    ]


def test_parse_bl_dist109_11_livraison_conforme_apres_retour():
    """BL 737851 : la vraie livraison conforme de R9PRA263, le lendemain
    du retour."""

    [bl] = parse_bl_109(_mots("bl_dist109_11.pdf"))

    assert bl.type_document == "BL"
    assert bl.numero_commande == "M3.10.175"
    assert bl.numero_bl == "737851"
    assert bl.total_ht_affiche == 58.5

    assert bl.lignes == [
        LigneBL(reference_fournisseur='R9PRA263', designation='RESI9-INTERDIFFERENTIEL-2P-63A-3OMA-TYPEA-PEIGNABLE-ALI', quantite_livree=1.0, prix_net=58.5, montant=58.5),
    ]


def test_autocontrole_total_ht_coherent_sur_tous_les_bl():
    """Le total HT affiché sur chaque BL correspond exactement à la somme
    des lignes extraites — aucune anomalie ne doit être signalée (voir
    _autocontrole dans parse_bl_109, imprime un avertissement sinon)."""

    noms = (
        "bl_dist109_1.pdf", "bl_dist109_2.pdf", "bl_dist109_3.pdf", "bl_dist109_4.pdf",
        "bl_dist109_5.pdf", "bl_dist109_6.pdf", "bl_dist109_7.pdf", "bl_dist109_8_multi_bl_8pages.pdf",
        "bl_dist109_9_retour.pdf", "bl_dist109_10_bl_avec_retour_associe.pdf", "bl_dist109_11.pdf",
    )
    for nom in noms:
        for bl in parse_bl_109(_mots(nom)):
            total_extrait = round(sum(l.montant for l in bl.lignes), 2)
            assert abs(bl.total_ht_affiche - total_extrait) <= 0.02, f"{nom} (BL {bl.numero_bl})"
