"""
Comparateur multi-fournisseurs — migré depuis test_pipeline.py (racine).

Couvre : le pipeline parse -> export ne plante pas (fumée), et la
régression V1.9 (consolidation par cœur numérique : une offre isolée par
un code interne déduit d'un seul fournisseur, ou par une réf absente de la
BDD, doit quand même se rattacher à la ligne du besoin qui partage son cœur
numérique — cas Vauban : L410707 Cominter isolé alors que 410707 / LEG410707
étaient rapprochés).
"""

import tempfile

from moteur.fournisseurs.ravate import parse_ravate
from moteur.fournisseurs.coredime import parse_coredime
from moteur.besoin import lire_besoin
from moteur.excel import exporter_comparatif
from moteur.modele import Article, LigneBesoin
from moteur.comparateur import comparer, _codes_compatibles
from moteur.normalisation import code_interne_auto
from moteur.base import BaseArticles
from moteur.referentiel import Referentiel

from conftest import ROOT

TEXTE_RAVATE = """RAVATE ELEC
Devis 00312607DV0262
Réf. FNR :
411651
41001713
DX3 ID 2P63A A 30MA
15 UN
78,14
40,72
610,80
EXO
AEC
Réf. FNR :
406774
41002001
DNX3 4500 1P+N C16
40 UN
12,50
6,10
244,00
EXO
STOCK
Réf. FNR :
980100
41009999
COFFRET DRIVIA 13M 2R
5 UN
60,00
35,00
175,00
EXO
SUR COMMANDE
"""

TEXTE_COREDIME = """COREDIME
COR B123456
LEG411651 DISPO DX3-ID 2P 63A A 30MA TGA 15 U* 39,9000 598,50 6
LEG406774 AEC DISJ DNX3 1P+N C16 40 U* 6,4500 258,00 6
LEG552088 DISPO GOULOTTE 40X25 2 M* 4,1000 8,20 6
"""


def _art(fourn, ref, des, prix):
    return Article(fournisseur=fourn, devis="TEST", reference_fournisseur=ref,
                   reference_distributeur="", designation=des, quantite=1,
                   unite="UN", prix_brut=prix, prix_net=prix, montant=prix)


def test_pipeline_fumee(tmp_path):
    """Parse (texte synthétique) -> comparateur -> export .xlsx, sans planter."""
    articles = parse_ravate(TEXTE_RAVATE) + parse_coredime(TEXTE_COREDIME)
    assert len(articles) == 6

    besoin = lire_besoin(ROOT / "besoins")
    exporter_comparatif(articles, tmp_path, besoin)


def test_code_interne_type_diff_apres_sensibilite():
    # Le type différentiel (AC/A/B/F) est lu même placé APRÈS la sensibilité
    assert code_interne_auto("Disjoncteur diff DX3 1P+N 25A 4.5/6K 30MA AC") \
        == "INTERDIFF_1P+N_25A_AC_30MA"


def test_consolidation_par_coeur_v19():
    articles = [
        _art("RAVATE", "410707", "DISJ DX3 1P+NG C25 4,5/6K AC", 70.0),
        _art("ELECTRIC PLUS", "LEG410707", "Disj diff DX3 1P+N 25A", 124.88),
        _art("COMINTER", "L410707", "Disjoncteur diff DX3 1P+N 25A 4.5/6K 30MA AC", 224.46),
    ]
    besoin = [LigneBesoin(
        designation="Disj diff DX³ 4500/6kA 1P+N C 25A 30mA Type AC",
        reference="410707", quantite=1)]

    # Base vide (réf absente de la BDD = le cas réel) : Cominter part en
    # CI:INTERDIFF..., Ravate/Electric Plus en cœur 410707 -> la
    # consolidation doit réunir les trois sous la même ligne de besoin.
    base = BaseArticles(tempfile.mkdtemp())
    resultat = comparer(articles, besoin, base)

    assert sorted(resultat["lignes"][0]["offres"]) == ["COMINTER", "ELECTRIC PLUS", "RAVATE"]
    assert resultat["hors_besoin"] == []


def test_consolidation_par_coeur_casse_code_interne_bdd(tmp_path):
    """
    Régression réelle (chantier Doujani, cf. test_codes_compatibles_
    insensible_a_la_casse) rejouée via comparer() de bout en bout, avec le
    même partage des rôles que le pipeline réel :
    - RAVATE (référence nue "406776") est rapproché par le RÉFÉRENTIEL
      (alias exact importé de la BDD) -> "REF:406776", sans jamais passer
      par base.groupe() ;
    - Electric Plus (référence préfixée "LEG406776", pas un alias exact
      connu) retombe sur base.groupe() -> "CI:DISJ_1P+N_C25A_6KA" (le
      "Code interne" de la BDD, MAJUSCULÉ à l'import par moteur/base.py).
    Deux clés différentes, dont une "kA"/"KA" ne diffèrent QUE par la
    casse : sans le fix, l'offre Electric Plus finissait en "Hors besoin"
    au lieu de rejoindre la ligne "406776" du besoin.
    """
    csv_bdd = tmp_path / "BDD_articles.csv"
    csv_bdd.write_text(
        "Référence;Désignation;Fournisseur;Fabricant;Catégorie;"
        "Tarif approximatif;CONCAT;Clé_Réf;Code interne\n"
        "406776;Disj DNX3 4500/6kA 1P+N C 25A;COMINTER;Legrand;"
        "Disjoncteurs;9,68;406776;406776;DISJ_1P+N_C25A_6kA\n",
        encoding="utf-8-sig",
    )
    base = BaseArticles(tmp_path / "base")
    base.importer_bdd(csv_bdd)

    referentiel = Referentiel(tmp_path / "moteur")
    referentiel.importer_bdd(csv_bdd)

    articles = [
        _art("RAVATE", "406776", "DNX3 1P+NG C25 4500A/6KA", 6.30),
        _art("ELECTRIC PLUS", "LEG406776", "DNX3 1P+NG C25 4500A/6KA", 6.30),
    ]
    besoin = [LigneBesoin(
        designation="Disj DNX³ 4500/6kA 1P+N C 25A", reference="406776", quantite=8)]

    resultat = comparer(articles, besoin, base, referentiel)
    referentiel.fermer()

    assert sorted(resultat["lignes"][0]["offres"]) == ["ELECTRIC PLUS", "RAVATE"]
    assert resultat["hors_besoin"] == []


def test_codes_compatibles_garde_fous():
    # A (type) jamais confondu avec AC
    assert not _codes_compatibles("INTERDIFF_2P_63A_AC_30MA", "INTERDIFF_2P_63A_A_30MA")
    # "?" sert de joker (type inconnu d'un côté)
    assert _codes_compatibles("INTERDIFF_2P_63A_AC_30MA", "INTERDIFF_2P_63A_?_30MA")
    # commodité (AMP_LED, EMBOUT_) : jamais de faux ami
    assert not _codes_compatibles("", "AMP_LED_E27_4000K")
    # familles différentes : jamais compatibles
    assert not _codes_compatibles("DISJ_1P+N_C16A", "EMBOUT_16")


def test_codes_compatibles_insensible_a_la_casse():
    # Régression réelle (chantier Doujani) : code_interne_auto() écrit
    # toujours le pouvoir de coupure en "kA" minuscule, mais moteur/base.py
    # MAJUSCULE le "Code interne" saisi dans base/BDD_articles.csv à
    # l'import -> "...6KA" (BDD) contre "...6kA" (auto) pour le MÊME
    # article. Sans insensibilité à la casse, les offres 406773/406774/
    # 406776 de COMINTER ("L406773"...) et Electric Plus ("LEG406776")
    # étaient rejetées de la ligne de besoin et finissaient en "Hors besoin"
    # alors que RAVATE (référence nue, rapprochée via le référentiel) y
    # apparaissait normalement.
    assert _codes_compatibles("DISJ_1P+N_C25A_6kA", "DISJ_1P+N_C25A_6KA")
    assert _codes_compatibles("DISJ_1P+N_C25A_6KA", "DISJ_1P+N_C25A_6kA")
