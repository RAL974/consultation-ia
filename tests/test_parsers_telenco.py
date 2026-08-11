"""TELENCO -- 3 vrais PDF (tests/fixtures/telenco_*.pdf)."""

from moteur.lecture_pdf import lire_pdf
from moteur.modele import Article
from moteur.detecteur import detecter_fournisseur
from moteur.fournisseurs.telenco import parse_telenco

from conftest import FIXTURES


def _texte(nom):
    return lire_pdf(FIXTURES / nom)


def test_detection_telenco():
    for nom in ("telenco_1.pdf", "telenco_2.pdf", "telenco_3.pdf"):
        assert detecter_fournisseur(_texte(nom)) == "TELENCO"


def test_parse_telenco_1():
    articles = parse_telenco(_texte("telenco_1.pdf"))
    assert articles == [
        Article(fournisseur='TELENCO', devis='SQFR01-00057092-1', reference_fournisseur='1008041', reference_distributeur='', designation='Tripack Premium MILWAUKEE perfo perceuse meuleuse 18V 2x5Ah + sac - V2', quantite=1.0, unite='UN', prix_brut=667.47, prix_net=667.47, montant=667.47, disponibilite=''),
    ]
    assert round(sum(a.montant for a in articles), 2) == 667.47


def test_parse_telenco_2():
    articles = parse_telenco(_texte("telenco_2.pdf"))
    assert articles == [
        Article(fournisseur='TELENCO', devis='SQFR01-00058010-1', reference_fournisseur='16475', reference_distributeur='', designation='Protection épissure thermo-rétractable Ø2,4mm Lg 45mm /Qté 100', quantite=5.0, unite='UN', prix_brut=7.66, prix_net=7.66, montant=38.3, disponibilite=''),
        Article(fournisseur='TELENCO', devis='SQFR01-00058010-1', reference_fournisseur='10841', reference_distributeur='', designation='Valise à outils TED Optima', quantite=1.0, unite='UN', prix_brut=70.7, prix_net=70.7, montant=70.7, disponibilite=''),
        Article(fournisseur='TELENCO', devis='SQFR01-00058010-1', reference_fournisseur='16476', reference_distributeur='', designation='Protection épissure thermo-rétractable Ø2,4mm Lg 61mm /Qté 100', quantite=5.0, unite='UN', prix_brut=7.66, prix_net=7.66, montant=38.3, disponibilite=''),
        Article(fournisseur='TELENCO', devis='SQFR01-00058010-1', reference_fournisseur='8247', reference_distributeur='', designation='Alcool isopropylique TED _Bouteille 1Lt', quantite=2.0, unite='UN', prix_brut=6.18, prix_net=6.18, montant=12.36, disponibilite=''),
        Article(fournisseur='TELENCO', devis='SQFR01-00058010-1', reference_fournisseur='4803', reference_distributeur='', designation='Lingette sèche pour FO Kimwipes EXL /Qté 280', quantite=3.0, unite='UN', prix_brut=9.79, prix_net=9.79, montant=29.37, disponibilite=''),
        Article(fournisseur='TELENCO', devis='SQFR01-00058010-1', reference_fournisseur='34004', reference_distributeur='', designation='Stylo nettoyage pour férules TED équipement 2,5mm', quantite=1.0, unite='UN', prix_brut=55.55, prix_net=55.55, montant=55.55, disponibilite=''),
        Article(fournisseur='TELENCO', devis='SQFR01-00058010-1', reference_fournisseur='15289', reference_distributeur='', designation='Localisateur visuel défauts optique TED 10mW', quantite=1.0, unite='UN', prix_brut=70.38, prix_net=70.38, montant=70.38, disponibilite=''),
    ]
    assert round(sum(a.montant for a in articles), 2) == 314.96


def test_parse_telenco_3():
    articles = parse_telenco(_texte("telenco_3.pdf"))
    assert articles == [
        Article(fournisseur='TELENCO', devis='SQFR01-00058233-1', reference_fournisseur='10154', reference_distributeur='', designation='Dégraissant FO_Bidon 3,8L', quantite=1.0, unite='UN', prix_brut=61.81, prix_net=61.81, montant=61.81, disponibilite=''),
        Article(fournisseur='TELENCO', devis='SQFR01-00058233-1', reference_fournisseur='4906', reference_distributeur='', designation='Pince à dénuder 3 positions (Ø3mm - 900µm - 250µm)', quantite=2.0, unite='UN', prix_brut=22.56, prix_net=22.56, montant=45.12, disponibilite=''),
        Article(fournisseur='TELENCO', devis='SQFR01-00058233-1', reference_fournisseur='4907', reference_distributeur='', designation='Ciseaux pour kevlar', quantite=2.0, unite='UN', prix_brut=20.6, prix_net=20.6, montant=41.2, disponibilite=''),
        Article(fournisseur='TELENCO', devis='SQFR01-00058233-1', reference_fournisseur='0403', reference_distributeur='', designation='Collier noir PA6.6 TED - 150x3,6mm /Qté 100', quantite=1.0, unite='UN', prix_brut=2.02, prix_net=2.02, montant=2.02, disponibilite=''),
    ]
    assert round(sum(a.montant for a in articles), 2) == 150.15

