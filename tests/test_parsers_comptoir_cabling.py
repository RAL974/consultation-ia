"""
COMPTOIR DU CABLING -- 4 vrais PDF (tests/fixtures/comptoir_cabling_*.pdf).

Un premier essai (1 seul PDF) avait ete abandonne (structure mal comprise :
30 lignes pour 335 euros au lieu de 2090 reels). Avec 4 PDF reels, la
structure fixe (Ref, TVA%, Montant, P.U., Qte, Designation) a pu etre
confirmee -- voir le bandeau GABARIT de moteur/fournisseurs/comptoir_cabling.py.

comptoir_cabling_1.pdf reste incomplet de 55,51 euros / 2089,76 (quelques
references longues coupees sur 2 lignes par l'extraction PDF, non
recollees -- pas assez d'exemples pour une regle fiable, regle d'or) :
assertion sur la somme ACTUELLE, pas sur le total du PDF, pour ne pas
perdre cet ecart de vue.
"""

from moteur.lecture_pdf import lire_pdf
from moteur.modele import Article
from moteur.detecteur import detecter_fournisseur
from moteur.fournisseurs.comptoir_cabling import parse_comptoir_cabling

from conftest import FIXTURES


def _texte(nom):
    return lire_pdf(FIXTURES / nom)


def test_detection_comptoir_cabling():
    for nom in (
        "comptoir_cabling_1.pdf", "comptoir_cabling_2.pdf",
        "comptoir_cabling_3.pdf", "comptoir_cabling_4.pdf",
    ):
        assert detecter_fournisseur(_texte(nom)) == "COMPTOIR DU CABLING"


def test_parse_comptoir_cabling_1():
    articles = parse_comptoir_cabling(_texte("comptoir_cabling_1.pdf"))
    assert articles == [
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='42U600X600', reference_distributeur='', designation='Baie\xa042U\xa0600x600 2\xa0paires\xa0de\xa0montants\xa0+\xa02\xa0ventilo\xa0+\xa04 roulettes', quantite=1.0, unite='UN', prix_brut=475.0, prix_net=475.0, montant=475.0, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='PLATEAU500', reference_distributeur='', designation="Plateau\xa0rackable\xa01U\xa0500m\xa04\xa0points\xa0de\xa0fixations\xa0\xa0pattes réglables\xa0jusqu'à\xa0800mm", quantite=1.0, unite='UN', prix_brut=39.0, prix_net=39.0, montant=39.0, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='42U800X800', reference_distributeur='', designation='Baie\xa042U\xa0800x800 2\xa0paires\xa0de\xa0montants\xa0+\xa01\xa0cabling\xa0+\xa04 ventilo\xa0+\xa04\xa0roulettes', quantite=1.0, unite='UN', prix_brut=655.0, prix_net=655.0, montant=655.0, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='PLATEAU500', reference_distributeur='', designation="Plateau\xa0rackable\xa01U\xa0500m\xa04\xa0points\xa0de\xa0fixations\xa0\xa0pattes réglables\xa0jusqu'à\xa0800mm", quantite=1.0, unite='UN', prix_brut=39.0, prix_net=39.0, montant=39.0, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='12U600X600', reference_distributeur='', designation='Coffret kit\xa0mural 12U\xa0600x600\xa0\xa02\xa0paires\xa0de\xa0montants\xa0+\xa0portes latérales\xa0+\xa0plaque\xa0de\xa0fond\xa0(640\xa0de\xa0hauteur)', quantite=1.0, unite='UN', prix_brut=145.0, prix_net=145.0, montant=145.0, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='18U600X600', reference_distributeur='', designation='Coffret kit\xa0mural 18U\xa0600x600\xa0+\xa02\xa0paires\xa0de\xa0montants\xa0+\xa0portes latérales\xa0+\xa0plaque\xa0de\xa0fond\xa0(905\xa0hauteur)', quantite=1.0, unite='UN', prix_brut=169.0, prix_net=169.0, montant=169.0, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='15U600PIVO', reference_distributeur='', designation='Coffret\xa015U\xa0600x600\xa0double\xa0partie\xa0pivotant', quantite=1.0, unite='UN', prix_brut=178.0, prix_net=178.0, montant=178.0, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='PLATEAU300', reference_distributeur='', designation='Plateau\xa0rackable\xa01U\xa0300m', quantite=1.0, unite='UN', prix_brut=22.9, prix_net=22.9, montant=22.9, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='BA6PC', reference_distributeur='', designation="Bandeau\xa0d'alimentation\xa06PC +\xa0inter", quantite=1.0, unite='UN', prix_brut=39.0, prix_net=39.0, montant=39.0, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='OBTU1U', reference_distributeur='', designation='Obturateur\xa01U', quantite=1.0, unite='UN', prix_brut=5.58, prix_net=5.58, montant=5.58, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='OBTU2U', reference_distributeur='', designation='Obturateur\xa02U', quantite=1.0, unite='UN', prix_brut=7.47, prix_net=7.47, montant=7.47, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='OBTU3U', reference_distributeur='', designation='Obturateur\xa03U', quantite=1.0, unite='UN', prix_brut=8.0, prix_net=8.0, montant=8.0, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='SCD', reference_distributeur='', designation='Tiroir\xa0optique\xa01U\xa024SC\xa0duplex +\xa048 smooves +\xa02\xa0cassettes', quantite=1.0, unite='UN', prix_brut=46.99, prix_net=46.99, montant=46.99, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='CDC2070336', reference_distributeur='', designation='Traversée\xa0fibre\xa0optique\xa0SC\xa0Multimode\xa0duplex', quantite=1.0, unite='UN', prix_brut=1.3, prix_net=1.3, montant=1.3, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='CDCOM4SCS', reference_distributeur='', designation='Traversée\xa0fibre\xa0optique\xa0SC\xa0\xa0OM4\xa0\xa0Multimode\xa0duplex', quantite=1.0, unite='UN', prix_brut=1.3, prix_net=1.3, montant=1.3, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='CDC0000008', reference_distributeur='', designation='Pigtail\xa050/ 125\xa0OM3\xa0SC\xa0Multimode\xa02m', quantite=1.0, unite='UN', prix_brut=3.49, prix_net=3.49, montant=3.49, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='PIGTOM4SC', reference_distributeur='', designation='Pigtail\xa050/ 125\xa0OM4 SC\xa0Multimode\xa02m', quantite=1.0, unite='UN', prix_brut=3.9, prix_net=3.9, montant=3.9, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='CDC0000003', reference_distributeur='', designation='Jarretière\xa0optique\xa0OM3\xa0SC/ SC\xa0duplex\xa0multimode\xa050/ 125\xa0- 2m', quantite=1.0, unite='UN', prix_brut=8.0, prix_net=8.0, montant=8.0, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='CDCPSCHN', reference_distributeur='', designation='Panneau\xa0de\xa0brassage\xa024\xa0ports\xa0RJ45\xa0Keystone\xa0/ \xa0Schneider', quantite=1.0, unite='UN', prix_brut=24.5, prix_net=24.5, montant=24.5, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='CDC45X45', reference_distributeur='', designation='Face\xa0avant\xa045x45\xa0pour\xa01\xa0connecteur keystone\xa0 / \xa0CDC', quantite=1.0, unite='UN', prix_brut=1.65, prix_net=1.65, montant=1.65, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='S10WE', reference_distributeur='', designation='Cordons\xa0de\xa0brassage\xa0RJ45\xa0cat6a\xa010G\xa0S/ FTP\xa01\xa0metre\xa0\xa0\xa0\xa0\xa0/ Schneider', quantite=1.0, unite='UN', prix_brut=5.95, prix_net=5.95, montant=5.95, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='S20WE', reference_distributeur='', designation='Cordons\xa0de\xa0brassage\xa0RJ45\xa0cat6a\xa010G\xa0S/ FTP\xa02\xa0metres\xa0\xa0\xa0\xa0\xa0/ Schneider', quantite=1.0, unite='UN', prix_brut=8.03, prix_net=8.03, montant=8.03, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='S30WE', reference_distributeur='', designation='Cordons\xa0de\xa0brassage\xa0RJ45\xa0cat6a\xa010G\xa0S/ FTP\xa03 metres\xa0\xa0\xa0\xa0\xa0/ Schneider', quantite=1.0, unite='UN', prix_brut=8.8, prix_net=8.8, montant=8.8, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='OM3X6', reference_distributeur='', designation='Fibre\xa0optique\xa06\xa0FO\xa0structure\xa0libre\xa0renforcé\xa050/ 125\xa0OM3\xa0LSZH Int/ Ext\xa0au\xa0mètre', quantite=1.0, unite='UN', prix_brut=1.65, prix_net=1.65, montant=1.65, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='CDC2070336', reference_distributeur='', designation='Traversée\xa0fibre\xa0optique\xa0SC\xa0Multimode\xa0duplex', quantite=1.0, unite='UN', prix_brut=1.3, prix_net=1.3, montant=1.3, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='CDC0000008', reference_distributeur='', designation='Pigtail\xa050/ 125\xa0OM3\xa0SC\xa0Multimode\xa02m', quantite=1.0, unite='UN', prix_brut=3.49, prix_net=3.49, montant=3.49, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='CDC0000001', reference_distributeur='', designation='Jarretière\xa0optique\xa0OM3\xa0SC/ LC\xa0duplex\xa0multimode\xa050/ 125\xa0- 2m', quantite=1.0, unite='UN', prix_brut=8.0, prix_net=8.0, montant=8.0, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='SCD', reference_distributeur='', designation='Tiroir\xa0optique\xa01U\xa024SC\xa0duplex +\xa048\xa0smooves\xa0+\xa02\xa0cassettes', quantite=1.0, unite='UN', prix_brut=46.99, prix_net=46.99, montant=46.99, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='CDCOM4SCS', reference_distributeur='', designation='Traversée\xa0fibre\xa0optique\xa0SC\xa0\xa0OM4\xa0\xa0Multimode\xa0duplex', quantite=1.0, unite='UN', prix_brut=1.3, prix_net=1.3, montant=1.3, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='PIGTOM4SC', reference_distributeur='', designation='Pigtail\xa050/ 125\xa0OM4 SC\xa0Multimode\xa02m', quantite=1.0, unite='UN', prix_brut=3.9, prix_net=3.9, montant=3.9, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='BA8PC', reference_distributeur='', designation="Bandeau\xa0d'alimentation\xa08PC\xa0+\xa0inter", quantite=1.0, unite='UN', prix_brut=44.0, prix_net=44.0, montant=44.0, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='CDCPSCHN', reference_distributeur='', designation='Panneau\xa0de\xa0brassage\xa024\xa0ports\xa0RJ45\xa0Keystone\xa0/ \xa0Schneider', quantite=1.0, unite='UN', prix_brut=24.5, prix_net=24.5, montant=24.5, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='CTE3P9', reference_distributeur='', designation='Câble\xa0SYT+1\xa0AWG\xa020\xa03P\xa09/ 10\xa0100%\xa0cuivre\xa0gris\xa0au\xa0mètre', quantite=1.0, unite='UN', prix_brut=0.72, prix_net=0.72, montant=0.72, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00029123', reference_fournisseur='CR1PYRO2M', reference_distributeur='', designation='Câble\xa0CR1-C1\xa0\xa02X2.5mm\xa0Pyro\xa0au\xa0mètre', quantite=1.0, unite='UN', prix_brut=1.54, prix_net=1.54, montant=1.54, disponibilite=''),
    ]
    assert round(sum(a.montant for a in articles), 2) == 2034.25


def test_parse_comptoir_cabling_2():
    articles = parse_comptoir_cabling(_texte("comptoir_cabling_2.pdf"))
    assert articles == [
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='NU400218', reference_distributeur='', designation='Plaque\xa0Unica\xa0pour\xa0boîtier\xa0saillie\xa0simple\xa0/ \xa0Schneider', quantite=1.0, unite='UN', prix_brut=1.86, prix_net=1.86, montant=1.86, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='NU7002PC', reference_distributeur='', designation='Support\xa0Unica\xa0pour\xa0boîtier\xa0saillie\xa0simple\xa0/ \xa0Schneider', quantite=1.0, unite='UN', prix_brut=1.39, prix_net=1.39, montant=1.39, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='NU840218', reference_distributeur='', designation='Boite\xa0Unica\xa0saillie\xa045x45\xa0simple\xa0/ \xa0Schneider', quantite=1.0, unite='UN', prix_brut=3.6, prix_net=3.6, montant=3.6, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='NU411418', reference_distributeur='', designation='Plaque\xa0Unica\xa0pour\xa0boîtier\xa0saillie\xa0double\xa0/ \xa0Schneider', quantite=1.0, unite='UN', prix_brut=4.61, prix_net=4.61, montant=4.61, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='NU7004PC', reference_distributeur='', designation='Support\xa0Unica\xa0pour\xa0boîtier\xa0saillie\xa0double\xa0/ \xa0Schneider', quantite=1.0, unite='UN', prix_brut=3.39, prix_net=3.39, montant=3.39, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='NU840418', reference_distributeur='', designation='Boite\xa0Unica\xa0saillie\xa0double\xa0\xa0/ \xa0Schneider', quantite=1.0, unite='UN', prix_brut=6.82, prix_net=6.82, montant=6.82, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='NU411618', reference_distributeur='', designation='Plaque\xa0Unica\xa0pour\xa0boîtier\xa0saillie\xa0triple\xa0/ \xa0Schneider', quantite=1.0, unite='UN', prix_brut=6.22, prix_net=6.22, montant=6.22, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='NU7006C', reference_distributeur='', designation='Support\xa0Unica\xa0zamak\xa0pour\xa0boîtier\xa0saillie\xa0triple\xa0/ \xa0Schneider', quantite=1.0, unite='UN', prix_brut=5.79, prix_net=5.79, montant=5.79, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='NU840618', reference_distributeur='', designation='Boite\xa0Unica\xa0saillie\xa0triple\xa0/ \xa0Schneider', quantite=1.0, unite='UN', prix_brut=11.64, prix_net=11.64, montant=11.64, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='CDC45X45', reference_distributeur='', designation='Face\xa0avant\xa045x45\xa0pour\xa01\xa0connecteur keystone\xa0 / \xa0CDC', quantite=1.0, unite='UN', prix_brut=1.65, prix_net=1.65, montant=1.65, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='42U800X1000', reference_distributeur='', designation='Baie\xa042U\xa0800x1000\xa0\xa02\xa0paires\xa0de\xa0montants\xa0+\xa01\xa0cabling\xa0+\xa04 ventilo\xa0+\xa04\xa0roulettes', quantite=1.0, unite='UN', prix_brut=720.0, prix_net=720.0, montant=720.0, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='CDCPSCHN', reference_distributeur='', designation='Panneau\xa0de\xa0brassage\xa024\xa0ports\xa0RJ45\xa0Keystone\xa0/ \xa0Schneider', quantite=3.0, unite='UN', prix_brut=24.5, prix_net=24.5, montant=73.5, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='BA8PC', reference_distributeur='', designation="Bandeau\xa0d'alimentation\xa08PC\xa0+\xa0inter", quantite=1.0, unite='UN', prix_brut=44.0, prix_net=44.0, montant=44.0, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='PLATEAU300', reference_distributeur='', designation='Plateau\xa0rackable\xa01U\xa0300m', quantite=1.0, unite='UN', prix_brut=22.9, prix_net=22.9, montant=22.9, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='PLATEAU500', reference_distributeur='', designation="Plateau\xa0rackable\xa01U\xa0500m\xa04\xa0points\xa0de\xa0fixations\xa0\xa0pattes réglables\xa0jusqu'à\xa0800mm", quantite=1.0, unite='UN', prix_brut=39.0, prix_net=39.0, montant=39.0, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='CDCPSCHN', reference_distributeur='', designation='Panneau\xa0de\xa0brassage\xa024\xa0ports\xa0RJ45\xa0Keystone\xa0/ \xa0Schneider', quantite=6.0, unite='UN', prix_brut=24.5, prix_net=24.5, montant=147.0, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='BA8PC', reference_distributeur='', designation="Bandeau\xa0d'alimentation\xa08PC\xa0+\xa0inter", quantite=6.0, unite='UN', prix_brut=44.0, prix_net=44.0, montant=264.0, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='PLATEAU300', reference_distributeur='', designation='Plateau\xa0rackable\xa01U\xa0300m', quantite=6.0, unite='UN', prix_brut=22.9, prix_net=22.9, montant=137.4, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='PLATEAU500', reference_distributeur='', designation="Plateau\xa0rackable\xa01U\xa0500m\xa04\xa0points\xa0de\xa0fixations\xa0\xa0pattes réglables\xa0jusqu'à\xa0800mm", quantite=6.0, unite='UN', prix_brut=39.0, prix_net=39.0, montant=234.0, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='SCD', reference_distributeur='', designation='Tiroir\xa0optique\xa01U\xa024SC\xa0duplex +\xa048\xa0smooves\xa0+\xa02\xa0cassettes', quantite=6.0, unite='UN', prix_brut=46.99, prix_net=46.99, montant=281.94, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='CDC2070336', reference_distributeur='', designation='Traversée\xa0fibre\xa0optique\xa0SC\xa0Multimode\xa0duplex', quantite=18.0, unite='UN', prix_brut=1.3, prix_net=1.3, montant=23.4, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='CDC0000008', reference_distributeur='', designation='Pigtail\xa050/ 125\xa0OM3\xa0SC\xa0Multimode\xa02m', quantite=36.0, unite='UN', prix_brut=3.49, prix_net=3.49, montant=125.64, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='CDC0000003', reference_distributeur='', designation='Jarretière\xa0optique\xa0OM3\xa0SC/ SC\xa0duplex\xa0multimode\xa050/ 125\xa0- 2m', quantite=18.0, unite='UN', prix_brut=8.0, prix_net=8.0, montant=144.0, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='LUECA', reference_distributeur='', designation='Fibre\xa0optique\xa0 6\xa0FO\xa0unitube\xa0renforcé\xa050/ 125\xa0OM3\xa0LSZH Int/ Ext\xa0au\xa0mètre\xa0/ \xa0LEVITON', quantite=400.0, unite='UN', prix_brut=1.65, prix_net=1.65, montant=660.0, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='CDCOM4SCS', reference_distributeur='', designation='Traversée\xa0fibre\xa0optique\xa0SC\xa0\xa0OM4\xa0\xa0Multimode\xa0duplex', quantite=18.0, unite='UN', prix_brut=1.3, prix_net=1.3, montant=23.4, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028376', reference_fournisseur='PIGTOM4SC', reference_distributeur='', designation='Pigtail\xa050/ 125\xa0OM4 SC\xa0Multimode\xa02m', quantite=36.0, unite='UN', prix_brut=3.9, prix_net=3.9, montant=140.4, disponibilite=''),
    ]
    assert round(sum(a.montant for a in articles), 2) == 3127.55


def test_parse_comptoir_cabling_3():
    articles = parse_comptoir_cabling(_texte("comptoir_cabling_3.pdf"))
    assert articles == [
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028528', reference_fournisseur='U7LITE', reference_distributeur='', designation="Point d'accès Wifi 7 Lite U7-Lite / UBIQUITI", quantite=5.0, unite='UN', prix_brut=156.0, prix_net=156.0, montant=780.0, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028528', reference_fournisseur='SW0908POE', reference_distributeur='', designation='Switch 8 ports full POE+ 250 m 10/ 100/ 1000 +2 port gigabit 120w', quantite=1.0, unite='UN', prix_brut=129.0, prix_net=129.0, montant=129.0, disponibilite=''),
    ]
    assert round(sum(a.montant for a in articles), 2) == 909.0


def test_parse_comptoir_cabling_4():
    articles = parse_comptoir_cabling(_texte("comptoir_cabling_4.pdf"))
    assert articles == [
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028723', reference_fournisseur='DIVERS', reference_distributeur='', designation='LIEN\xa0ABONNE\xa0AVEC\xa0DTIO\xa01FO\xa0G657\xa030\xa0m', quantite=4.0, unite='UN', prix_brut=30.0, prix_net=30.0, montant=120.0, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028723', reference_fournisseur='DIVERS', reference_distributeur='', designation='LIEN\xa0ABONNE\xa0AVEC\xa0DTIO\xa01FO\xa0G657\xa040\xa0m', quantite=8.0, unite='UN', prix_brut=35.0, prix_net=35.0, montant=280.0, disponibilite=''),
        Article(fournisseur='COMPTOIR DU CABLING', devis='DE00028723', reference_fournisseur='DIVERS', reference_distributeur='', designation='LIEN\xa0ABONNE\xa0AVEC\xa0DTIO\xa01FO\xa0G657\xa090\xa0m', quantite=5.0, unite='UN', prix_brut=55.0, prix_net=55.0, montant=275.0, disponibilite=''),
    ]
    assert round(sum(a.montant for a in articles), 2) == 675.0

