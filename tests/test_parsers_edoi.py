"""EDOI — 3 vrais PDF (tests/fixtures/edoi_*.pdf), verrouille la sortie."""

from moteur.lecture_pdf import lire_pdf
from moteur.modele import Article
from moteur.detecteur import detecter_fournisseur
from moteur.fournisseurs.edoi import parse_edoi

from conftest import FIXTURES


def _texte(nom):
    return lire_pdf(FIXTURES / nom)


def test_detection_edoi():
    for nom in ("edoi_1.pdf", "edoi_2.pdf", "edoi_3.pdf"):
        assert detecter_fournisseur(_texte(nom)) == "EDOI"


def test_parse_edoi_1():
    articles = parse_edoi(_texte("edoi_1.pdf"))
    assert articles == [
        Article(fournisseur='EDOI', devis='B228071', reference_fournisseur='LEG401802', reference_distributeur='', designation='XL3 160 COMPLET ISOLANT 2R', quantite=9.0, unite='UN', prix_brut=659.0971, prix_net=659.0971, montant=5931.87, disponibilite='12sem'),
        Article(fournisseur='EDOI', devis='B228071', reference_fournisseur='LEG404905', reference_distributeur='', designation='BORNE ARRIVEE PEIGNEP+N', quantite=6.0, unite='UN', prix_brut=10.8345, prix_net=10.8345, montant=65.01, disponibilite='12sem'),
        Article(fournisseur='EDOI', devis='B228071', reference_fournisseur='LEG404926', reference_distributeur='', designation='PEIGNE POUR 13 APPAREILS 1P+N', quantite=13.0, unite='UN', prix_brut=5.1979, prix_net=5.1979, montant=67.57, disponibilite='DISPO'),
        Article(fournisseur='EDOI', devis='B228071', reference_fournisseur='LEG406434', reference_distributeur='', designation='DX3 IS 2P 32A', quantite=8.0, unite='UN', prix_brut=37.1227, prix_net=37.1227, montant=296.98, disponibilite='12sem'),
        Article(fournisseur='EDOI', devis='B228071', reference_fournisseur='LEG406490', reference_distributeur='', designation='DX3 IS 4P 125A', quantite=1.0, unite='UN', prix_brut=236.3591, prix_net=236.3591, montant=236.36, disponibilite='12sem'),
        Article(fournisseur='EDOI', devis='B228071', reference_fournisseur='LEG406773', reference_distributeur='', designation='DNX3 1P+NG C10 4500A/6KA', quantite=33.0, unite='UN', prix_brut=11.7988, prix_net=11.7988, montant=389.36, disponibilite='DISPO'),
        Article(fournisseur='EDOI', devis='B228071', reference_fournisseur='LEG411524', reference_distributeur='', designation='DX3-ID 2P 25A AC 300MA', quantite=1.0, unite='UN', prix_brut=165.5157, prix_net=165.5157, montant=165.52, disponibilite='12sem'),
        Article(fournisseur='EDOI', devis='B228071', reference_fournisseur='LEG406774', reference_distributeur='', designation='DNX3 1P+NG C16 4500A/6KA', quantite=21.0, unite='UN', prix_brut=11.8918, prix_net=11.8918, montant=249.73, disponibilite='DISPO'),
        Article(fournisseur='EDOI', devis='B228071', reference_fournisseur='LEG411525', reference_distributeur='', designation='DX3-ID 2P 40A AC 300MA', quantite=1.0, unite='UN', prix_brut=162.9295, prix_net=162.9295, montant=162.93, disponibilite='12sem'),
        Article(fournisseur='EDOI', devis='B228071', reference_fournisseur='LEG411613', reference_distributeur='', designation='DX3-ID 2P 25A AC 300MA TG', quantite=1.0, unite='UN', prix_brut=163.9329, prix_net=163.9329, montant=163.93, disponibilite='12sem'),
        Article(fournisseur='EDOI', devis='B228071', reference_fournisseur='LEG003901', reference_distributeur='', designation='COFFRET TELECOMMANDE 300 BLOCS', quantite=17.0, unite='UN', prix_brut=180.2054, prix_net=180.2054, montant=3063.49, disponibilite='12sem'),
        Article(fournisseur='EDOI', devis='B228071', reference_fournisseur='LEG020051', reference_distributeur='', designation='OBTURATEUR DECOUPABLE 24 MOD', quantite=12.0, unite='UN', prix_brut=17.3023, prix_net=17.3023, montant=207.63, disponibilite='12sem'),
        Article(fournisseur='EDOI', devis='B228071', reference_fournisseur='LEG020252', reference_distributeur='', designation='PORTE GALBEE METALLIQUE H450', quantite=9.0, unite='UN', prix_brut=409.2003, prix_net=409.2003, montant=3682.8, disponibilite='12sem'),
        Article(fournisseur='EDOI', devis='B228071', reference_fournisseur='LEG406776', reference_distributeur='', designation='DNX3 1P+NG C25 4500A/6KA', quantite=8.0, unite='UN', prix_brut=18.7625, prix_net=18.7625, montant=150.1, disponibilite='DISPO'),
    ]
    assert round(sum(a.montant for a in articles), 2) == 14833.28


def test_parse_edoi_2():
    articles = parse_edoi(_texte("edoi_2.pdf"))
    assert articles == [
        Article(fournisseur='EDOI', devis='B228073', reference_fournisseur='LEG076576', reference_distributeur='', designation='RJ45 C6A STP MOSAIC 2 M', quantite=38.0, unite='UN', prix_brut=17.43, prix_net=17.43, montant=662.34, disponibilite='DISPO'),
        Article(fournisseur='EDOI', devis='B228073', reference_fournisseur='LEG078778', reference_distributeur='', designation='HDMI PRE-CONNECT.1M BL MOSAIC', quantite=22.0, unite='UN', prix_brut=88.9314, prix_net=88.9314, montant=1956.49, disponibilite='12sem'),
    ]
    assert round(sum(a.montant for a in articles), 2) == 2618.83


def test_parse_edoi_3():
    articles = parse_edoi(_texte("edoi_3.pdf"))
    assert articles == [
        Article(fournisseur='EDOI', devis='B228077', reference_fournisseur='GFO031818', reference_distributeur='', designation='R2V 5G16 T500', quantite=100.0, unite='UN', prix_brut=16.0694, prix_net=16.0694, montant=1606.94, disponibilite='12sem'),
        Article(fournisseur='EDOI', devis='B228077', reference_fournisseur='GFO026518', reference_distributeur='', designation='R2V 3G1 5 T500', quantite=1050.0, unite='UN', prix_brut=1.183, prix_net=1.183, montant=1242.15, disponibilite='DISPO'),
        Article(fournisseur='EDOI', devis='B228077', reference_fournisseur='GFO031005', reference_distributeur='', designation='R2V 5G1 5 C100', quantite=200.0, unite='UN', prix_brut=1.9069, prix_net=1.9069, montant=381.38, disponibilite='DISPO'),
        Article(fournisseur='EDOI', devis='B228077', reference_fournisseur='GFO026718', reference_distributeur='', designation='R2V 3G2 5 T500', quantite=350.0, unite='UN', prix_brut=1.8118, prix_net=1.8118, montant=634.13, disponibilite='DISPO'),
    ]
    assert round(sum(a.montant for a in articles), 2) == 3864.60
