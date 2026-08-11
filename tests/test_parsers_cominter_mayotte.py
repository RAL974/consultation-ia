"""
COMINTER MAYOTTE — 1 vrai PDF (tests/fixtures/cominter_mayotte.pdf).

18/19 articles réels sont extraits : le 19e ("L76565", dernier de la page)
a ses valeurs numériques extraites AVANT sa référence dans le flux PyMuPDF
(cas jamais vu ailleurs, voir le bandeau GABARIT de cominter_mayotte.py) —
signalé "bloc incomplet" plutôt que rattaché au hasard. Total réel du devis
9 630,36€, somme des 18 lignes extraites 9 184,01€ (écart 446,35€ = les
valeurs orphelines de L76565, vérifié à la main : 38 x 16,78 x 70 % =
446,35€ exact).
"""

from moteur.lecture_pdf import lire_pdf
from moteur.modele import Article
from moteur.detecteur import detecter_fournisseur
from moteur.fournisseurs.cominter_mayotte import parse_cominter_mayotte

from conftest import FIXTURES


def _texte(nom):
    return lire_pdf(FIXTURES / nom)


def test_detection_cominter_mayotte():
    assert detecter_fournisseur(_texte("cominter_mayotte.pdf")) == "COMINTER MAYOTTE"


def test_parse_cominter_mayotte():
    articles = parse_cominter_mayotte(_texte("cominter_mayotte.pdf"))
    assert articles == [
        Article(fournisseur='COMINTER MAYOTTE', devis='MDEC22462', reference_fournisseur='L76576', reference_distributeur='', designation='Mosaic Prise RJ45 Catégorie 6A STP à blindage métal 2 modules blanc', quantite=9.0, unite='UN', prix_brut=454.62, prix_net=318.2344, montant=2864.11, disponibilite=''),
        Article(fournisseur='COMINTER MAYOTTE', devis='MDEC22462', reference_fournisseur='L401802', reference_distributeur='', designation='XL3 160 complet métal 2 Rangé', quantite=20.0, unite='UN', prix_brut=3.81, prix_net=2.667, montant=53.34, disponibilite=''),
        Article(fournisseur='COMINTER MAYOTTE', devis='MDEC22462', reference_fournisseur='L404905', reference_distributeur='', designation='Borne d arrivee pour cable 4mm² à 25mm²  Neutre', quantite=13.0, unite='UN', prix_brut=6.15, prix_net=4.3054, montant=55.97, disponibilite=''),
        Article(fournisseur='COMINTER MAYOTTE', devis='MDEC22462', reference_fournisseur='L404926', reference_distributeur='', designation='Peigne d alimentation P+N 13 modules DISPO', quantite=10.0, unite='UN', prix_brut=14.3, prix_net=14.3, montant=143.0, disponibilite=''),
        Article(fournisseur='COMINTER MAYOTTE', devis='MDEC22462', reference_fournisseur='L406434', reference_distributeur='', designation='Inter-sectionneur de tête DX³-IS - vis/vis - 2P  32 A', quantite=1.0, unite='UN', prix_brut=97.27, prix_net=97.27, montant=97.27, disponibilite=''),
        Article(fournisseur='COMINTER MAYOTTE', devis='MDEC22462', reference_fournisseur='L406490', reference_distributeur='', designation='DX3 IS 4P 125A', quantite=33.0, unite='UN', prix_brut=15.31, prix_net=10.717, montant=353.66, disponibilite=''),
        Article(fournisseur='COMINTER MAYOTTE', devis='MDEC22462', reference_fournisseur='L406773', reference_distributeur='', designation='Disjoncteur DNX3 1P+NG 10A   4.5 KA DISPO', quantite=21.0, unite='UN', prix_brut=15.31, prix_net=10.7171, montant=225.06, disponibilite=''),
        Article(fournisseur='COMINTER MAYOTTE', devis='MDEC22462', reference_fournisseur='L406774', reference_distributeur='', designation='Disjoncteur DNX3 1P+NG 16A   4.5 KA DISPO', quantite=8.0, unite='UN', prix_brut=13.55, prix_net=13.55, montant=108.4, disponibilite=''),
        Article(fournisseur='COMINTER MAYOTTE', devis='MDEC22462', reference_fournisseur='L406776', reference_distributeur='', designation='DNX3 1P+NG C25 4500A 1M p97 406776', quantite=1.0, unite='UN', prix_brut=146.8, prix_net=102.76, montant=102.76, disponibilite=''),
        Article(fournisseur='COMINTER MAYOTTE', devis='MDEC22462', reference_fournisseur='L410704', reference_distributeur='', designation='Disjoncteur diff DX3 1P+N 10A 4.5/6K 30MA AC', quantite=4.0, unite='UN', prix_brut=141.42, prix_net=98.995, montant=395.98, disponibilite=''),
        Article(fournisseur='COMINTER MAYOTTE', devis='MDEC22462', reference_fournisseur='L410705', reference_distributeur='', designation='Disjoncteur diff DX3 1P+N 16A 4.5/6K 30MA AC', quantite=1.0, unite='UN', prix_brut=152.92, prix_net=107.04, montant=107.04, disponibilite=''),
        Article(fournisseur='COMINTER MAYOTTE', devis='MDEC22462', reference_fournisseur='L410707', reference_distributeur='', designation='Disjoncteur diff DX3 1P+N 25A 4.5/6K 30MA AC', quantite=1.0, unite='UN', prix_brut=72.31, prix_net=50.62, montant=50.62, disponibilite=''),
        Article(fournisseur='COMINTER MAYOTTE', devis='MDEC22462', reference_fournisseur='L411524', reference_distributeur='', designation='Interrupteur diff 2P 25A 300MA AC', quantite=1.0, unite='UN', prix_brut=74.57, prix_net=52.2, montant=52.2, disponibilite=''),
        Article(fournisseur='COMINTER MAYOTTE', devis='MDEC22462', reference_fournisseur='L411525', reference_distributeur='', designation='Interrupteur diff 2P 40A 300MA AC', quantite=11.0, unite='UN', prix_brut=79.63, prix_net=55.7409, montant=613.15, disponibilite=''),
        Article(fournisseur='COMINTER MAYOTTE', devis='MDEC22462', reference_fournisseur='ZZ', reference_distributeur='', designation='411613 DX3-ID 2P 25A AC 300MA TG', quantite=12.0, unite='UN', prix_brut=217.7, prix_net=152.39, montant=1828.68, disponibilite=''),
        Article(fournisseur='COMINTER MAYOTTE', devis='MDEC22462', reference_fournisseur='L03901', reference_distributeur='', designation='Télécom. modulaire SATI connecté 300 Blocs DISPO', quantite=12.0, unite='UN', prix_brut=11.26, prix_net=7.8817, montant=94.58, disponibilite=''),
        Article(fournisseur='COMINTER MAYOTTE', devis='MDEC22462', reference_fournisseur='L20051', reference_distributeur='', designation='Obturateur XL 160/400', quantite=9.0, unite='UN', prix_brut=195.25, prix_net=195.25, montant=1757.25, disponibilite=''),
        Article(fournisseur='COMINTER MAYOTTE', devis='MDEC22462', reference_fournisseur='L20252', reference_distributeur='', designation='Porte Galbée tole 2 Rgées  XL160', quantite=22.0, unite='UN', prix_brut=12.77, prix_net=12.77, montant=280.94, disponibilite=''),
    ]
    # Écart connu et documenté (voir docstring du module) : le 19e article
    # (L76565) est absent, ses valeurs étant orphelines dans le flux texte.
    assert round(sum(a.montant for a in articles), 2) == 9184.01
