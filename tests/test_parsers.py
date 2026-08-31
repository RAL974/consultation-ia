"""
Tests de non-régression des parsers fournisseurs, à partir de VRAIS PDF
(tests/fixtures/). Chaque test verrouille la sortie ACTUELLE du parser,
champ par champ : toute différence après une refonte doit être justifiée
explicitement, jamais silencieuse.
"""

from moteur.lecture_pdf import lire_pdf
from moteur.modele import Article
from moteur.detecteur import detecter_fournisseur
from moteur.fournisseurs.ravate import parse_ravate
from moteur.fournisseurs.coredime import parse_coredime
from moteur.fournisseurs.dem import parse_dem
from moteur.fournisseurs.dist109 import parse_109
from moteur.fournisseurs.cominter import parse_cominter
from moteur.fournisseurs.electricplus import parse_electricplus
from moteur.fournisseurs.artdeco import parse_artdeco

from conftest import FIXTURES


def _texte(nom):
    return lire_pdf(FIXTURES / nom)


def test_detection_fournisseurs():
    assert detecter_fournisseur(_texte("ravate.pdf")) == "RAVATE"
    assert detecter_fournisseur(_texte("coredime.pdf")) == "COREDIME"
    assert detecter_fournisseur(_texte("dem.pdf")) == "DEM"
    assert detecter_fournisseur(_texte("109_distribution.pdf")) == "109 DISTRIBUTION"
    assert detecter_fournisseur(_texte("cominter.pdf")) == "COMINTER"
    assert detecter_fournisseur(_texte("electric_plus_gmr.pdf")) == "ELECTRIC PLUS"
    assert detecter_fournisseur(_texte("artdeco_1.pdf")) == "ART DECO"


def test_parse_ravate():
    articles = parse_ravate(_texte("ravate.pdf"))
    assert articles == [
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='VU1.5B', reference_distributeur='100025371', designation='CABLE HO7 VU 1X1.5 BLEU C100', quantite=9.0, unite='COU', prix_brut=44.0, prix_net=18.25, montant=164.25, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='VU1.5R', reference_distributeur='100025375', designation='CABLE HO7 VU 1X1.5 ROUGE C100', quantite=8.0, unite='COU', prix_brut=44.0, prix_net=18.25, montant=146.0, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='VU1.5O', reference_distributeur='100025374', designation='CABLE HO7VU 1X1.5 ORANGE C100M', quantite=5.0, unite='COU', prix_brut=44.0, prix_net=18.25, montant=91.25, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='VU1.5N', reference_distributeur='100025373', designation='CABLE HO7 VU 1X1.5 NOIR C100', quantite=2.0, unite='COU', prix_brut=44.0, prix_net=18.25, montant=36.5, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='VU1.5V', reference_distributeur='100025376', designation='CABLE HO7 VU 1X1.5 VIOLET C100', quantite=1.0, unite='COU', prix_brut=44.0, prix_net=18.25, montant=18.25, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='VU2.5VJ', reference_distributeur='100025390', designation='CABLE HO7 VU 1X2.5 V-J C100', quantite=1.0, unite='COU', prix_brut=65.0, prix_net=29.25, montant=29.25, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='VU2.5B', reference_distributeur='100025387', designation='CABLE HO7VU 1X2.5 BLEU C100M', quantite=11.0, unite='COU', prix_brut=64.0, prix_net=29.25, montant=321.75, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='VU2.5R', reference_distributeur='100025389', designation='CABLE HO7 VU 1X2.5 ROUGE C100', quantite=14.0, unite='COU', prix_brut=65.0, prix_net=29.25, montant=409.5, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='05120', reference_distributeur='100014224', designation='ICT 20 BLEU TURBO G-ROUL 100M', quantite=5.0, unite='ROU', prix_brut=240.0, prix_net=69.36, montant=346.8, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='05125', reference_distributeur='46000317', designation='ICT 25 BLEU TURBO G-ROUL 100M', quantite=2.0, unite='ROU', prix_brut=226.52, prix_net=69.32, montant=138.64, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='CAP856574', reference_distributeur='100060841', designation='MANCHON PREDALLE HT65 AXE74', quantite=100.0, unite='UN', prix_brut=3.23, prix_net=1.31, montant=131.0, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='P02000', reference_distributeur='100042318', designation='POT SIBTOP 2000', quantite=200.0, unite='UN', prix_brut=1.24, prix_net=0.5, montant=100.0, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='P01760', reference_distributeur='100084416', designation='POT DE DESCENTE RH176+COUVERC', quantite=300.0, unite='UN', prix_brut=2.98, prix_net=1.21, montant=363.0, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='P0104300', reference_distributeur='100084408', designation='DISTANCIER HORIZONTAL ENT 100', quantite=900.0, unite='UN', prix_brut=0.35, prix_net=0.14, montant=126.0, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='20PTT298V', reference_distributeur='100169717', designation='ICT 20 PREF PTT 298 4P5/10 E', quantite=5.0, unite='ROU', prix_brut=185.0, prix_net=50.0, montant=250.0, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='TUBFFPTTGRD3', reference_distributeur='100151314', designation='ICT 20 PREF PTT GRADE 3 S+ VER', quantite=6.0, unite='ROU', prix_brut=260.0, prix_net=136.34, montant=818.04, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='09131', reference_distributeur='100050091', designation='PREF D25 VERT GRADE 2 TV 4P', quantite=2.0, unite='ROU', prix_brut=505.0, prix_net=152.81, montant=305.62, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='TUBFF1X16I20', reference_distributeur='100278514', designation='ICT 20 PREF 1X16 VJ', quantite=2.0, unite='ROU', prix_brut=1060.0, prix_net=291.92, montant=583.84, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='09004', reference_distributeur='100004713', designation='ICT 16 PREF BLEU 3G1.5 B.R V/J', quantite=2.0, unite='ROU', prix_brut=320.0, prix_net=86.0, montant=172.0, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='TUBFF3G6', reference_distributeur='100151309', designation='ICT 25 PREF 3X6.0  R B VJ 50M', quantite=2.0, unite='ROU', prix_brut=545.0, prix_net=164.92, montant=329.84, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='20515-132', reference_distributeur='100151310', designation='ICT 20 PREF 5X1.5  R B VJ M N', quantite=6.0, unite='ROU', prix_brut=475.0, prix_net=142.12, montant=852.72, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='20515-134', reference_distributeur='100292993', designation='ICT 20 PREF 5X1.5  R B VJ M O', quantite=3.0, unite='ROU', prix_brut=465.0, prix_net=139.13, montant=417.39, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='PREG 20', reference_distributeur='100292979', designation='ICT 20 + TIRE AIGUILLE GRIS', quantite=1.0, unite='ROU', prix_brut=97.0, prix_net=27.0, montant=27.0, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='05125', reference_distributeur='46000317', designation='ICT 25 BLEU TURBO G-ROUL 100M', quantite=7.0, unite='ROU', prix_brut=226.52, prix_net=67.77, montant=474.39, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='06632', reference_distributeur='46000360', designation='ICT 32 ROULEAU DE 50M', quantite=1.0, unite='ROU', prix_brut=140.0, prix_net=41.89, montant=41.89, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='PREG 40C', reference_distributeur='100292986', designation='ICTA 40+TIRE AIGUILLE GRIS 25M', quantite=6.0, unite='ROU', prix_brut=97.5, prix_net=29.17, montant=175.02, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='06650', reference_distributeur='100009099', designation='ICT 50 ROULEAU DE 50M', quantite=2.0, unite='ROU', prix_brut=297.0, prix_net=135.0, montant=270.0, disponibilite=''),
        Article(fournisseur='RAVATE', devis='00312607DV0407', reference_fournisseur='60399-00001-01', reference_distributeur='79701191', designation='ADHESIF PVC ORANG 50X33ML TESA', quantite=5.0, unite='PCE', prix_brut=6.5, prix_net=4.0, montant=20.0, disponibilite=''),
    ]


def test_parse_coredime():
    articles = parse_coredime(_texte("coredime.pdf"))
    assert len(articles) == 31
    assert articles[0] == Article(fournisseur='COREDIME', devis='COR B424261', reference_fournisseur='GFO000305', reference_distributeur='', designation='H07VU 1 5 BLEU 5012 C100', quantite=900.0, unite='MT', prix_brut=0.2, prix_net=0.2, montant=180.0, disponibilite='DISPO')
    assert articles[-1] == Article(fournisseur='COREDIME', devis='COR B424261', reference_fournisseur='MMM85298', reference_distributeur='', designation='EASY TAPE ORANGE  30M X 50MM', quantite=5.0, unite='UN', prix_brut=7.6, prix_net=7.6, montant=38.0, disponibilite='DISPO')
    # Cas notable : une ligne sans mot-clé DISPO/AEC en tête -> disponibilité vide
    sans_dispo = [a for a in articles if a.reference_fournisseur == 'COU220215400062DISPO']
    assert sans_dispo == [Article(fournisseur='COREDIME', devis='COR B424261', reference_fournisseur='COU220215400062DISPO', reference_distributeur='', designation='PREFILE NOIR 20/100 1G16 V/J H', quantite=200.0, unite='MT', prix_brut=3.1, prix_net=3.1, montant=620.0, disponibilite='')]


def test_parse_dem():
    # Les 8 premières lignes (références à POINT : FILVU1.5B...) étaient
    # SILENCIEUSEMENT PERDUES avant le correctif de MOTIF_LIGNE (la classe
    # de la référence n'admettait pas le point) : ce fixture ne rendait que
    # 9 lignes sur 18, sans qu'aucune anomalie ne soit levée. La somme
    # extraite vaut désormais exactement le Total H.T. imprimé (2822,00 €).
    articles = parse_dem(_texte("dem.pdf"))
    assert round(sum(a.montant for a in articles), 2) == 2822.00
    assert articles == [
        Article(fournisseur='DEM', devis='821035', reference_fournisseur='FILVU1.5B', reference_distributeur='', designation='VU 1.5 BLEU C100', quantite=900.0, unite='U', prix_brut=0.195, prix_net=0.195, montant=175.5, disponibilite=''),
        Article(fournisseur='DEM', devis='821035', reference_fournisseur='FILVU1.5R', reference_distributeur='', designation='VU 1.5 ROUGE C100', quantite=800.0, unite='U', prix_brut=0.195, prix_net=0.195, montant=156.0, disponibilite=''),
        Article(fournisseur='DEM', devis='821035', reference_fournisseur='FILVU1.5O', reference_distributeur='', designation='VU 1.5 ORANGE C100', quantite=500.0, unite='U', prix_brut=0.195, prix_net=0.195, montant=97.5, disponibilite=''),
        Article(fournisseur='DEM', devis='821035', reference_fournisseur='FILVU1.5N', reference_distributeur='', designation='VU 1.5 NOIR C100', quantite=200.0, unite='U', prix_brut=0.195, prix_net=0.195, montant=39.0, disponibilite=''),
        Article(fournisseur='DEM', devis='821035', reference_fournisseur='FILVU1.5V', reference_distributeur='', designation='VU 1.5 VIOLET C100', quantite=100.0, unite='U', prix_brut=0.2, prix_net=0.2, montant=20.0, disponibilite=''),
        Article(fournisseur='DEM', devis='821035', reference_fournisseur='FILVU2.5VJ', reference_distributeur='', designation='VU 2.5 VERT/JAUNE C100', quantite=100.0, unite='U', prix_brut=0.3, prix_net=0.3, montant=30.0, disponibilite=''),
        Article(fournisseur='DEM', devis='821035', reference_fournisseur='FILVU2.5B', reference_distributeur='', designation='VU 2.5 BLEU C100', quantite=1100.0, unite='U', prix_brut=0.3, prix_net=0.3, montant=330.0, disponibilite=''),
        Article(fournisseur='DEM', devis='821035', reference_fournisseur='FILVU2.5R', reference_distributeur='', designation='VU 2.5 ROUGE C100', quantite=1400.0, unite='U', prix_brut=0.3, prix_net=0.3, montant=420.0, disponibilite=''),
        Article(fournisseur='DEM', devis='821035', reference_fournisseur='CAPCAP309284', reference_distributeur='', designation='100 BORNES DE CONNEXION STAND  2 ENTREES POUR FILS RIGIDES 0,', quantite=800.0, unite='U', prix_brut=0.07, prix_net=0.07, montant=56.0, disponibilite=''),
        Article(fournisseur='DEM', devis='821035', reference_fournisseur='CAPCAP309285', reference_distributeur='', designation='100 BORNES DE CONNEXION STAND  3 ENTREES POUR FILS RIGIDES 0,', quantite=500.0, unite='U', prix_brut=0.085, prix_net=0.085, montant=42.5, disponibilite=''),
        Article(fournisseur='DEM', devis='821035', reference_fournisseur='CAPCAP309286', reference_distributeur='', designation='100 BORNES DE CONNEXION STAND  4 ENTREES POUR FILS RIGIDES 0,', quantite=300.0, unite='U', prix_brut=0.095, prix_net=0.095, montant=28.5, disponibilite=''),
        Article(fournisseur='DEM', devis='821035', reference_fournisseur='CAPCAP309287', reference_distributeur='', designation='50 BORNES DE CONNEXION STAND 5  ENTREES POUR FILS RIGIDES 0,5', quantite=100.0, unite='U', prix_brut=0.105, prix_net=0.105, montant=10.5, disponibilite=''),
        Article(fournisseur='DEM', devis='821035', reference_fournisseur='CAPCAP856574', reference_distributeur='', designation='MANCHON POUR PRE-DALLE HAUTEUR  65 ENTRAXE 74', quantite=100.0, unite='U', prix_brut=1.41, prix_net=1.41, montant=141.0, disponibilite=''),
        Article(fournisseur='DEM', devis='821035', reference_fournisseur='EUR52097', reference_distributeur='', designation="POT SIMPLE AIR'METIC            67mm PROF 50mm", quantite=100.0, unite='U', prix_brut=0.48, prix_net=0.48, montant=48.0, disponibilite=''),
        Article(fournisseur='DEM', devis='821035', reference_fournisseur='TUBFFPTT', reference_distributeur='', designation='FOURREAU FILE PTT ICT20', quantite=500.0, unite='U', prix_brut=0.43, prix_net=0.43, montant=215.0, disponibilite=''),
        Article(fournisseur='DEM', devis='821035', reference_fournisseur='TUBFFTV', reference_distributeur='', designation='FOURREAU FILE TV17VATC', quantite=200.0, unite='U', prix_brut=0.55, prix_net=0.55, montant=110.0, disponibilite=''),
        Article(fournisseur='DEM', devis='821035', reference_fournisseur='TUBFF3G1.5ICT16', reference_distributeur='', designation='FOURREAU FILE 3G1.5 ICT16', quantite=200.0, unite='U', prix_brut=0.8, prix_net=0.8, montant=160.0, disponibilite=''),
        Article(fournisseur='DEM', devis='821035', reference_fournisseur='TUBFF3G6', reference_distributeur='', designation='FF 3G6 RGE B V/J C50', quantite=200.0, unite='U', prix_brut=3.7125, prix_net=3.7125, montant=742.5, disponibilite=''),
    ]


def test_parse_109_distribution():
    articles = parse_109(_texte("109_distribution.pdf"))
    assert articles == [
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='45001001-PRYSM-H07VU1.5B', reference_distributeur='', designation='CABLE HO7VU 1.5mm2 BLEU C100M', quantite=900.0, unite='UN', prix_brut=0.25, prix_net=0.245, montant=220.5, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='45001001-PRYSM-H07VU1.5R', reference_distributeur='', designation='CABLE HO7VU 1.5mm2 ROUGE C100M', quantite=800.0, unite='UN', prix_brut=0.25, prix_net=0.245, montant=196.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='45001001-PRYSM-H07VU1.5O', reference_distributeur='', designation='CABLE HO7VU 1.5mm2 ORANGE C100M', quantite=500.0, unite='UN', prix_brut=0.25, prix_net=0.245, montant=122.5, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='45001001-PRYSM-H07VU1.5N', reference_distributeur='', designation='CABLE HO7VU 1.5mm2 NOIR C100M', quantite=200.0, unite='UN', prix_brut=0.25, prix_net=0.245, montant=49.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='45001001-PRYSM-H07VU1.5V', reference_distributeur='', designation='CABLE HO7VU 1.5mm2 VIOLET', quantite=100.0, unite='UN', prix_brut=0.25, prix_net=0.245, montant=24.5, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='45001001-PRYSM-H07VU2.5B', reference_distributeur='', designation='CABLE HO7VU2.5mm2 BLEU', quantite=1100.0, unite='UN', prix_brut=0.41, prix_net=0.41, montant=451.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='45001001-PRYSM-H07VU2.5VJ', reference_distributeur='', designation='CABLE HO7VU2.5mm2 VERT/JAUNE C100M', quantite=100.0, unite='UN', prix_brut=0.41, prix_net=0.41, montant=41.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='45001001-PRYSM-H07VU2.5R', reference_distributeur='', designation='CABLE HO7VU2.5mm2 ROUGE C100M', quantite=1400.0, unite='UN', prix_brut=0.41, prix_net=0.41, montant=574.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='50001001-COURA-10041540', reference_distributeur='', designation='GAINE ICTA PRELUB DIAM 20 C100M', quantite=500.0, unite='UN', prix_brut=0.25, prix_net=0.25, montant=125.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='50001001-COURA-10041940', reference_distributeur='', designation='GAINE ICTA PRELUB DIAM 25 C100M', quantite=200.0, unite='UN', prix_brut=0.38, prix_net=0.38, montant=76.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='10004002-S.I.B-P07132', reference_distributeur='', designation='CONNEXION RAPIDE 2 TROUS MINI BOITE DE 150PCS (WAGO', quantite=8.0, unite='UN', prix_brut=7.5, prix_net=7.5, montant=60.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='10004002-S.I.B-P07133', reference_distributeur='', designation='CONNEXION RAPIDE 3 TROUS MINI BOITE DE 120PCS (WAGO', quantite=5.0, unite='UN', prix_brut=7.0, prix_net=7.0, montant=35.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='10004002-SIB-P07135', reference_distributeur='', designation='CONNEXION RAPIDE 5 TROUS MINI BOITE DE 80PCS (WAGO)', quantite=6.0, unite='UN', prix_brut=8.0, prix_net=8.0, montant=48.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='10004002-SIB-P07138', reference_distributeur='', designation='CONNEXION RAPIDE 8 TROUS MINI BOITE DE 60PCS (WAGO)', quantite=2.0, unite='UN', prix_brut=11.0, prix_net=11.0, montant=22.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='10005006-NOL-CAP856574', reference_distributeur='', designation='MANCHON PREDALLE HT 65 ENTR 74 (PAR 50)', quantite=100.0, unite='UN', prix_brut=1.35, prix_net=1.35, montant=135.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='10003005-S.I.B-P01811', reference_distributeur='', designation='FLEX.180 3 ICT20+ 1 ICT25 LIDIC : 032796', quantite=200.0, unite='UN', prix_brut=0.65, prix_net=0.65, montant=130.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='10003005-S.I.B-P02000', reference_distributeur='', designation='BOITIER SIP TOP 2000', quantite=200.0, unite='UN', prix_brut=0.45, prix_net=0.45, montant=90.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='10003007-S.I.B-P03200', reference_distributeur='', designation='COUVERCLE DE POSE ( DALLE PLEINE )', quantite=200.0, unite='UN', prix_brut=0.26, prix_net=0.26, montant=52.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='10003005-S.I.B-P0176000', reference_distributeur='', designation='RH176 JAUNE+ COUVERCLE', quantite=300.0, unite='UN', prix_brut=1.0, prix_net=1.0, montant=300.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='10003007-S.I.B-P01052', reference_distributeur='', designation='ANNEAU A VIS UNIVERSEL (BAGUE)', quantite=400.0, unite='UN', prix_brut=0.3, prix_net=0.3, montant=120.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='10003007-S.I.B-P0104300', reference_distributeur='', designation='DISTANCIER HORIZONTALE 100MM LIDIC', quantite=900.0, unite='UN', prix_brut=0.18, prix_net=0.18, montant=162.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='10003007-S.I.B-P0102001', reference_distributeur='', designation='COUVERCLE DE POSE', quantite=200.0, unite='UN', prix_brut=0.26, prix_net=0.26, montant=52.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='10003001-SIB-P16840', reference_distributeur='', designation='POT SIMPLE PLACO  D68-PROF.40 LIDIC', quantite=100.0, unite='UN', prix_brut=0.24, prix_net=0.24, montant=24.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='50001005-COURA-20080035', reference_distributeur='', designation='GAINE PREFILE CABLE TELEPHONE 1X4P SERIE 298', quantite=500.0, unite='UN', prix_brut=0.52, prix_net=0.52, montant=260.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='50001005-COURA-20080043', reference_distributeur='', designation='GAINE PREFILEE CABLE INFORMATIQUE 1X4 CAT 6', quantite=600.0, unite='UN', prix_brut=0.93, prix_net=0.93, montant=558.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='50001005-COURA-20080033', reference_distributeur='', designation='GAINE PREFILE CABLE TELEVISION 17 VATC', quantite=200.0, unite='UN', prix_brut=0.41, prix_net=0.41, montant=82.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='50001005-PM-F2R160T', reference_distributeur='', designation='GAINE PREFILE 16mm² V/J  / DIAM 20', quantite=200.0, unite='UN', prix_brut=3.05, prix_net=3.05, montant=610.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='50001005-PM-F2U15BRT', reference_distributeur='', designation='GAINE PREFILE 3G1.5 mm² / DIAM 20 B-R-VJ', quantite=200.0, unite='UN', prix_brut=0.89, prix_net=0.89, montant=178.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='50001005-PM-F2U25BRT', reference_distributeur='', designation='GAINE PREFILE 3G2.5 mm² / DIAM 20 B-R-VJ', quantite=200.0, unite='UN', prix_brut=1.3, prix_net=1.3, montant=260.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='50001005-PM-F3R60BRT50', reference_distributeur='', designation='GAINE PREFILE 3G6 mm² / DIAM 25 R-B-VJ', quantite=200.0, unite='UN', prix_brut=3.05, prix_net=3.05, montant=610.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='50001005-PM-F2U15BMNRT', reference_distributeur='', designation='GAINE PREFILEE 5G1.5 DIAM 20 B-R-VJ-M-N', quantite=600.0, unite='UN', prix_brut=1.35, prix_net=1.35, montant=810.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='50001005-PM-F2U15RVVOO', reference_distributeur='', designation='GAINE PREFILEE 5X1.5 DIAM 20 1RO-2VIO-2 OR', quantite=300.0, unite='UN', prix_brut=1.35, prix_net=1.35, montant=405.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='50001001-COURA-10041540', reference_distributeur='', designation='GAINE ICTA PRELUB DIAM 20 C100M', quantite=100.0, unite='UN', prix_brut=0.25, prix_net=0.25, montant=25.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='50001001-COURA-10041940', reference_distributeur='', designation='GAINE ICTA PRELUB DIAM 25 C100M', quantite=700.0, unite='UN', prix_brut=0.38, prix_net=0.38, montant=266.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='50001001-COURA-10042724', reference_distributeur='', designation='GAINE ICTA PRELUB DIAM 32 C50M', quantite=100.0, unite='UN', prix_brut=0.75, prix_net=0.75, montant=75.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='50001001-COURA-10043324', reference_distributeur='', designation='GAINE ICTA PRELUB DIAM 40 C50M', quantite=150.0, unite='UN', prix_brut=1.3, prix_net=1.3, montant=195.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='50001001-COURA-10043924', reference_distributeur='', designation='GAINE ICTA PRELUB DIAM 50 C50M', quantite=150.0, unite='UN', prix_brut=2.5, prix_net=2.5, montant=375.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='433307', reference_fournisseur='55003004-FIXAT-PVCORANGE', reference_distributeur='', designation='PVC ORANGE 2800 50MM X 33M', quantite=5.0, unite='UN', prix_brut=4.2, prix_net=4.2, montant=21.0, disponibilite=''),
    ]
    # Recoupement : la somme des lignes retombe exactement sur le "Total HT"
    # affiché par le PDF (7 839,50 EUR) -> aucune ligne oubliée ou dupliquée.
    assert round(sum(a.montant for a in articles), 2) == 7839.50


def test_parse_109_distribution_variante_devis_bpu():
    # 2e variante réelle du même fournisseur : la référence vient APRÈS le
    # bloc chiffré (pas avant) — voir le bandeau GABARIT de dist109.py.
    articles = parse_109(_texte("109_distribution_devis_bpu.pdf"))
    assert articles == [
        Article(fournisseur='109 DISTRIBUTION', devis='317718', reference_fournisseur='LK4288', reference_distributeur='', designation='BAIE LINK+  800 X 800 - 42U', quantite=1.0, unite='UN', prix_brut=820.0, prix_net=820.0, montant=820.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='317718', reference_fournisseur='LK4268', reference_distributeur='', designation='BAIE LINK+ 600X800 42U', quantite=1.0, unite='UN', prix_brut=695.0, prix_net=695.0, montant=695.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='317718', reference_fournisseur='LKOFF12UP600', reference_distributeur='', designation='COFFRET BAIE MURAL 19P 12U PROF 600', quantite=1.0, unite='UN', prix_brut=190.14, prix_net=190.14, montant=190.14, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='317718', reference_fournisseur='49BM6DDP', reference_distributeur='', designation='BLOC ALIM ALU 6 PRISES DD PROTEGER EQUIP DISJONCTEUR DIFF P', quantite=1.0, unite='UN', prix_brut=115.0, prix_net=115.0, montant=115.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='317718', reference_fournisseur='49BM8IPPM', reference_distributeur='', designation='BLOC ALIM ALU 8 PRISES + INTER EQUIP INTERRUPTEUR PROTEGE', quantite=1.0, unite='UN', prix_brut=29.0, prix_net=29.0, montant=29.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='317718', reference_fournisseur='49BM8PM', reference_distributeur='', designation='BLOC ALIM 19" ALU 8 PRISES', quantite=1.0, unite='UN', prix_brut=24.5, prix_net=24.5, montant=24.5, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='317718', reference_fournisseur='F5554SH5', reference_distributeur='', designation='CABLE CAT6A F/FTP 1x4P ZH 555MHZ T500M', quantite=1.0, unite='UN', prix_brut=0.7, prix_net=0.7, montant=0.7, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='317718', reference_fournisseur='F5554SHC5', reference_distributeur='', designation='CABLE CAT6A F/FTP 1x4P LSZH Cca- 555MHz T500', quantite=1.0, unite='UN', prix_brut=0.95, prix_net=0.95, montant=0.95, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='317718', reference_fournisseur='BC6AFSTL8', reference_distributeur='', designation='CONNECTEUR RJ45 CAT6A BLINDE 360° sans outil', quantite=1.0, unite='UN', prix_brut=3.95, prix_net=3.95, montant=3.95, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='317718', reference_fournisseur='ECORD6ASF01MSH', reference_distributeur='', designation='CORDON CAT6A S/FTP 1M LSZH', quantite=1.0, unite='UN', prix_brut=2.5, prix_net=2.5, montant=2.5, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='317718', reference_fournisseur='ECORD6ASF02MSH', reference_distributeur='', designation='CORDON CAT6A S/FTP 2M LSZH', quantite=1.0, unite='UN', prix_brut=4.3, prix_net=4.3, montant=4.3, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='317718', reference_fournisseur='ECORD6ASF03MSH', reference_distributeur='', designation='CORDON CAT6A S/FTP 3M LSZH', quantite=1.0, unite='UN', prix_brut=5.1, prix_net=5.1, montant=5.1, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='317718', reference_fournisseur='SYT120G5', reference_distributeur='', designation='CABLE TELEPHONIQUE NUMERIQUE-SYT1 PLUS AWG 20 - 1P 9/10 - T5', quantite=1.0, unite='UN', prix_brut=0.25, prix_net=0.25, montant=0.25, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='317718', reference_fournisseur='SYT220G5', reference_distributeur='', designation='CABLE TELEPHONIQUE NUMERIQUE-SYT1 PLUS AWG 20 - 2P 9/10 - T5', quantite=1.0, unite='UN', prix_brut=0.52, prix_net=0.52, montant=0.52, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='317718', reference_fournisseur='SYT520G5', reference_distributeur='', designation='CABLE TELEPHONIQUE NUMERIQUE-SYT1 PLUS AWG 20 - 5P 9/10 - T5', quantite=1.0, unite='UN', prix_brut=1.05, prix_net=1.05, montant=1.05, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='317718', reference_fournisseur='VATC17T', reference_distributeur='', designation='CABLE 17 VATC COAXIAL CLASS B - T500', quantite=1.0, unite='UN', prix_brut=0.23, prix_net=0.23, montant=0.23, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='317718', reference_fournisseur='CR1/2X2.5', reference_distributeur='', designation='CABLE CR1-C1 NA 2x2.5(PYRO)', quantite=1.0, unite='UN', prix_brut=1.55, prix_net=1.55, montant=1.55, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='317718', reference_fournisseur='CR1/2X1.5', reference_distributeur='', designation='CABLE CR1-C1 NA 2X1.5 mm2 ( PYRO )', quantite=1.0, unite='UN', prix_brut=0.9, prix_net=0.9, montant=0.9, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='317718', reference_fournisseur='CRCT209', reference_distributeur='', designation='CABLE CR1-C1 NA 2 PAIRES 0.9 ( PYRO ) T500M', quantite=1.0, unite='UN', prix_brut=1.55, prix_net=1.55, montant=1.55, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='317718', reference_fournisseur='CR1/5G1.5', reference_distributeur='', designation='CABLE CR1-C1 NA 5G1.5 (PYRO)', quantite=1.0, unite='UN', prix_brut=2.35, prix_net=2.35, montant=2.35, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='317718', reference_fournisseur='CR1/3G1.5', reference_distributeur='', designation='CABLE CR1-C1 NA 3G1,5 (PYRO) T500M', quantite=1.0, unite='UN', prix_brut=1.55, prix_net=1.55, montant=1.55, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='317718', reference_fournisseur='CR1/3G2.5', reference_distributeur='', designation='CABLE CR1-C1 NA 3G2.5 (PYRO)', quantite=1.0, unite='UN', prix_brut=2.3, prix_net=2.3, montant=2.3, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='317718', reference_fournisseur='P04389', reference_distributeur='', designation='BARETTE COUP.PLAST. LIDIC : 034076', quantite=1.0, unite='UN', prix_brut=6.0, prix_net=6.0, montant=6.0, disponibilite=''),
    ]
    assert round(sum(a.montant for a in articles), 2) == 1909.39


def test_parse_109_distribution_variante_devis_bpu_reference_longue():
    # Même variante "devis_bpu" que ci-dessus, PAS une 3e variante malgré
    # le message d'anomalie initial (devis ISHOP Saint-Denis n° 321051) :
    # la référence "FRN1X6G3-3G1.5 T" fait 16 caractères, au-delà de la
    # borne à 15 de MOTIF_REF_DEVIS_BPU — élargie à 20.
    articles = parse_109(_texte("109_distribution_2_ref_longue_devis_bpu.pdf"))
    assert articles == [
        Article(fournisseur='109 DISTRIBUTION', devis='321051', reference_fournisseur='FRN1X6G3-3G1.5 T', reference_distributeur='', designation='CABLE FR-N1X6G3-U 3G1,5 T500', quantite=100.0, unite='UN', prix_brut=1.2, prix_net=1.2, montant=120.0, disponibilite=''),
    ]
    assert round(sum(a.montant for a in articles), 2) == 120.00


def test_parse_109_distribution_variante_devis_bpu_remise():
    # Variante "devis_bpu" avec colonne Rem% renseignée (devis ISHOP
    # 321106/Réglettes - Rico Carpaye) : "Rem%" (ex. "2,94") s'intercale
    # entre Total et P.U.Net quand elle est non nulle, décalant tout le
    # reste d'un cran — voir OFFSETS_DEVIS_BPU_REMISE dans dist109.py.
    articles = parse_109(_texte("109_distribution_3_devis_bpu_remise.pdf"))
    assert articles == [
        Article(fournisseur='109 DISTRIBUTION', devis='321106', reference_fournisseur='302758', reference_distributeur='', designation='REGLETTE VEGAS III 32W 4000K-IP66-IK08-CLII 4760lm DET HF ON-OFF', quantite=14.0, unite='UN', prix_brut=45.0, prix_net=45.0, montant=630.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='321106', reference_fournisseur='302759', reference_distributeur='', designation='REGLETTE VEGAS III 44W 4000K-IP66-IK08-CLII 6440lm DET HF ON-OFF', quantite=14.0, unite='UN', prix_brut=52.0, prix_net=52.0, montant=728.0, disponibilite=''),
    ]
    assert round(sum(a.montant for a in articles), 2) == 1358.00


def test_parse_109_distribution_reference_avec_plus():
    # Trouvé en vérifiant BT - Floe (devis 321273) après le correctif
    # ci-dessus : "BTSOUT3X150+70"/"BTSOUT3X95+50" contiennent un "+",
    # absent de la classe de caractères de MOTIF_REF_DEVIS_BPU — 2 lignes
    # sur 5 manquaient silencieusement (1 539,00€ sur 1 626,40€), révélé
    # par l'autocontrôle Total HT.
    articles = parse_109(_texte("109_distribution_4_ref_avec_plus.pdf"))
    assert articles == [
        Article(fournisseur='109 DISTRIBUTION', devis='321273', reference_fournisseur='BTSOUT3X150+70', reference_distributeur='', designation='CABLE BT SOUTERRAIN 3x150 + 70 - NFC33210 T250M', quantite=60.0, unite='UN', prix_brut=18.4, prix_net=18.4, montant=1104.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='321273', reference_fournisseur='BTSOUT3X95+50', reference_distributeur='', designation='CABLE BT SOUTERRAIN 3x95 + 50 - NFC33210 T 250M', quantite=30.0, unite='UN', prix_brut=14.5, prix_net=14.5, montant=435.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='321273', reference_fournisseur='C-9-92KL', reference_distributeur='', designation='COLLIER COLSON 9 X 355', quantite=2.0, unite='UN', prix_brut=14.8, prix_net=14.8, montant=29.6, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='321273', reference_fournisseur='C-9-62KL', reference_distributeur='', designation='COLLIER COLSON 9 X 265', quantite=2.0, unite='UN', prix_brut=11.0, prix_net=11.0, montant=22.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='321273', reference_fournisseur='6880544', reference_distributeur='', designation='GAINE REMONTEE POTEAU PVC -  GPC 90 GRIS 2.75M', quantite=2.0, unite='UN', prix_brut=17.9, prix_net=17.9, montant=35.8, disponibilite=''),
    ]
    assert round(sum(a.montant for a in articles), 2) == 1626.40


def test_parse_electricplus():
    # Les 7 premières lignes (CAB0001*0C100, câbles HO7V-U) étaient
    # SILENCIEUSEMENT PERDUES avant l'élargissement de MARQUEUR à ["PF",
    # "PR"] (voir bandeau GABARIT electricplus.py) : leur colonne
    # d'ancrage affiche "PR" au lieu de "PF" sur ce même document réel,
    # jamais remarqué avant que 2 devis entiers (BT-Floe, R2V 3G1.5)
    # ressortent à 0 article pour la même raison. Total encore incomplet
    # de 120,00 € (WAG2273205, ligne SANS AUCUN marqueur PF/PR imprimé —
    # un seul exemple à ce jour, non corrigé, voir "Points fragiles").
    articles = parse_electricplus(_texte("electric_plus_gmr.pdf"))
    assert round(sum(a.montant for a in articles), 2) == 3894.15
    assert articles == [
        Article(fournisseur='ELECTRIC PLUS', devis='1108800', reference_fournisseur='CAB000100C100', reference_distributeur='', designation='HO7V-U 1,5 BLEU CIEL      C100', quantite=900.0, unite='UN', prix_brut=0.22, prix_net=0.22, montant=198.0, disponibilite=''),
        Article(fournisseur='ELECTRIC PLUS', devis='1108800', reference_fournisseur='CAB000110C100', reference_distributeur='', designation='HO7V-U 1,5 ROUGE          C100', quantite=800.0, unite='UN', prix_brut=0.22, prix_net=0.22, montant=176.0, disponibilite=''),
        Article(fournisseur='ELECTRIC PLUS', devis='1108800', reference_fournisseur='CAB000130C100', reference_distributeur='', designation='HO7V-U 1,5 NOIR           C100', quantite=200.0, unite='UN', prix_brut=0.27, prix_net=0.27, montant=54.0, disponibilite=''),
        Article(fournisseur='ELECTRIC PLUS', devis='1108800', reference_fournisseur='CAB000170C100', reference_distributeur='', designation='HO7V-U 1,5 VIOLET         C100', quantite=100.0, unite='UN', prix_brut=0.22, prix_net=0.22, montant=22.0, disponibilite=''),
        Article(fournisseur='ELECTRIC PLUS', devis='1108800', reference_fournisseur='CAB000220C100', reference_distributeur='', designation='HO7V-U 2,5 V/J            C100', quantite=100.0, unite='UN', prix_brut=0.45, prix_net=0.45, montant=45.0, disponibilite=''),
        Article(fournisseur='ELECTRIC PLUS', devis='1108800', reference_fournisseur='CAB000200C100', reference_distributeur='', designation='HO7V-U 2,5 BLEU CIEL      C100', quantite=1100.0, unite='UN', prix_brut=0.36, prix_net=0.36, montant=396.0, disponibilite=''),
        Article(fournisseur='ELECTRIC PLUS', devis='1108800', reference_fournisseur='CAB000210C100', reference_distributeur='', designation='HO7V-U 2,5 ROUGE          C100', quantite=1400.0, unite='UN', prix_brut=0.36, prix_net=0.36, montant=504.0, disponibilite=''),
        Article(fournisseur='ELECTRIC PLUS', devis='1108800', reference_fournisseur='COU13021540', reference_distributeur='', designation='FLEXPRO+ NOIR TAGP 20/100', quantite=500.0, unite='UN', prix_brut=0.25, prix_net=0.25, montant=125.0, disponibilite=''),
        Article(fournisseur='ELECTRIC PLUS', devis='1108800', reference_fournisseur='COU13021940', reference_distributeur='', designation='FLEXPRO+ NOIR TAGP 25/100', quantite=200.0, unite='UN', prix_brut=0.4, prix_net=0.4, montant=80.0, disponibilite=''),
        Article(fournisseur='ELECTRIC PLUS', devis='1108800', reference_fournisseur='WAG2273202', reference_distributeur='', designation='BORNE WAGO 2273 - 2 X 0,5 A 2,', quantite=800.0, unite='UN', prix_brut=0.12, prix_net=0.12, montant=96.0, disponibilite=''),
        Article(fournisseur='ELECTRIC PLUS', devis='1108800', reference_fournisseur='WAG2273203', reference_distributeur='', designation='BORNE WAGO 2273 - 3 X 0,5 A 2,', quantite=500.0, unite='UN', prix_brut=0.14, prix_net=0.14, montant=70.0, disponibilite=''),
        Article(fournisseur='ELECTRIC PLUS', devis='1108800', reference_fournisseur='WAG2273208', reference_distributeur='', designation='BORNE WAGO 2273 - 8 X 0,5 A 2,', quantite=100.0, unite='UN', prix_brut=0.29, prix_net=0.29, montant=29.0, disponibilite=''),
        Article(fournisseur='ELECTRIC PLUS', devis='1108800', reference_fournisseur='COU20080035', reference_distributeur='', designation='PREFILCO V 20/100 ADSL GRADE 1', quantite=500.0, unite='UN', prix_brut=0.47, prix_net=0.47, montant=235.0, disponibilite=''),
        Article(fournisseur='ELECTRIC PLUS', devis='1108800', reference_fournisseur='COU20080033', reference_distributeur='', designation='PREFILCO V 20/100 17VATCA', quantite=200.0, unite='UN', prix_brut=0.35, prix_net=0.35, montant=70.0, disponibilite=''),
        Article(fournisseur='ELECTRIC PLUS', devis='1108800', reference_fournisseur='COU20020003', reference_distributeur='', designation='PREFILCO N 20/100 3G1.5 BRV/J', quantite=200.0, unite='UN', prix_brut=0.77, prix_net=0.77, montant=154.0, disponibilite=''),
        Article(fournisseur='ELECTRIC PLUS', devis='1108800', reference_fournisseur='COU20020020', reference_distributeur='', designation='PREFILCO N 20/100 3G2.5 BRV/J', quantite=200.0, unite='UN', prix_brut=1.09, prix_net=1.09, montant=218.0, disponibilite=''),
        Article(fournisseur='ELECTRIC PLUS', devis='1108800', reference_fournisseur='COU220215400096', reference_distributeur='', designation='PREFILE N 20/100 5X1.5 MNOOR', quantite=300.0, unite='UN', prix_brut=2.64, prix_net=2.64, montant=792.0, disponibilite=''),
        Article(fournisseur='ELECTRIC PLUS', devis='1108800', reference_fournisseur='COU13021540', reference_distributeur='', designation='FLEXPRO+ NOIR TAGP 20/100', quantite=100.0, unite='UN', prix_brut=0.25, prix_net=0.25, montant=25.0, disponibilite=''),
        Article(fournisseur='ELECTRIC PLUS', devis='1108800', reference_fournisseur='COU13021940', reference_distributeur='', designation='FLEXPRO+ NOIR TAGP 25/100', quantite=700.0, unite='UN', prix_brut=0.4, prix_net=0.4, montant=280.0, disponibilite=''),
        Article(fournisseur='ELECTRIC PLUS', devis='1108800', reference_fournisseur='COU12022724', reference_distributeur='', designation='ICTA SP NOIR TAP 32/50', quantite=100.0, unite='UN', prix_brut=0.82, prix_net=0.82, montant=82.0, disponibilite=''),
        Article(fournisseur='ELECTRIC PLUS', devis='1108800', reference_fournisseur='COU10043324', reference_distributeur='', designation='ICTA SPOT G TAG 40/50 - ICTA', quantite=150.0, unite='UN', prix_brut=1.47, prix_net=1.47, montant=220.5, disponibilite=''),
        Article(fournisseur='ELECTRIC PLUS', devis='1108800', reference_fournisseur='AGI393500', reference_distributeur='', designation='RUBAN ADH ETANCHE PARE-VAPEUR', quantite=5.0, unite='UN', prix_brut=4.53, prix_net=4.53, montant=22.65, disponibilite=''),
    ]


def test_parse_electricplus_marqueur_pr():
    # Devis réels (BT - Floe, D1109436.pdf) où la colonne ancrée sur "PF"
    # sur les autres devis de ce fournisseur affiche "PR" à la place —
    # même position, mêmes calculs (qté x prix_net = montant vérifié
    # exact) : ressortait à 0 article avant l'élargissement de MARQUEUR.
    articles = parse_electricplus(_texte("electric_plus_gmr_2_marqueur_pr.pdf"))
    assert articles == [
        Article(fournisseur='ELECTRIC PLUS', devis='1109436', reference_fournisseur='LEG031919', reference_distributeur='', designation='COLLIER COLSON NOIR 9X357', quantite=200.0, unite='UN', prix_brut=0.16, prix_net=0.16, montant=32.0, disponibilite=''),
        Article(fournisseur='ELECTRIC PLUS', devis='1109436', reference_fournisseur='LEG031916', reference_distributeur='', designation='COLLIER COLSON NOIR 9X262', quantite=200.0, unite='UN', prix_brut=0.12, prix_net=0.12, montant=24.0, disponibilite=''),
    ]


def test_parse_electricplus_marqueur_pr_2():
    # 2e devis réel (R2V 3G1.5 - Rico Carpaye, D1109369.pdf), même symptôme.
    articles = parse_electricplus(_texte("electric_plus_gmr_3_marqueur_pr.pdf"))
    assert articles == [
        Article(fournisseur='ELECTRIC PLUS', devis='1109369', reference_fournisseur='CAB013000T500', reference_distributeur='', designation='RO2V-CU 3G1,5             T500', quantite=500.0, unite='UN', prix_brut=0.87, prix_net=0.87, montant=435.0, disponibilite=''),
    ]


def test_parse_cominter():
    articles = parse_cominter(_texte("cominter.pdf"))
    assert len(articles) == 17
    assert articles[0] == Article(fournisseur='COMINTER', devis='ODE270211', reference_fournisseur='MELVS08103', reference_distributeur='', designation='Coffret - L600 - IP30 - 9M - RAL9003', quantite=7.0, unite='UN', prix_brut=275.3, prix_net=192.71, montant=1348.98, disponibilite='')
    assert articles[-1] == Article(fournisseur='COMINTER', devis='ODE270211', reference_fournisseur='L03901', reference_distributeur='', designation='Télécom. modulaire SATI connecté 300 Blocs', quantite=1.0, unite='UN', prix_brut=165.92, prix_net=116.14, montant=116.14, disponibilite='')
    refs = [a.reference_fournisseur for a in articles]
    assert refs == [
        'MELVS08103', 'MELVS08123', 'MELVS03001', 'MELVS03203', 'MELVS03220',
        'L404926', 'L406434', 'L406490', 'L406773', 'L406774', 'L406776',
        'L410704', 'L410705', 'L410707', 'L411524', 'L411525', 'L03901',
    ]


def test_parse_artdeco():
    # Nouveau fournisseur (session TRAVAUX_PARSERS.md, devis 617004507,
    # Réglettes - Rico Carpaye) — voir bandeau GABARIT de artdeco.py :
    # "ELECTRICITE SERVICES REUNION" dans le nom du fichier désigne
    # l'acheteuse, pas ce fournisseur ; le vrai vendeur est "LED'S RUN"
    # (domaine artdeco.re). Désignation reconstruite depuis 2 lignes (le
    # texte déborde sur une ligne de continuation après tout le bloc
    # chiffré, décalage propre à ce gabarit).
    articles = parse_artdeco(_texte("artdeco_1.pdf"))
    assert articles == [
        Article(fournisseur='ART DECO', devis='617004507', reference_fournisseur='5470', reference_distributeur='', designation='Rampe led Novolight Porto 48w + DET (4000k) 1200mm IP65 (ecotaxe 0,20€)', quantite=14.0, unite='UN', prix_brut=110.0, prix_net=60.0, montant=840.0, disponibilite=''),
        Article(fournisseur='ART DECO', devis='617004507', reference_fournisseur='5471', reference_distributeur='', designation='Télécommande RF Novolight pour Porto détecteur (ec otaxe 0,08€)', quantite=1.0, unite='UN', prix_brut=9.22, prix_net=6.45, montant=6.45, disponibilite=''),
    ]
    # Recoupement : la somme des lignes retombe exactement sur le "Sous
    # Total :" HT affiché par le PDF (846,45 EUR).
    assert round(sum(a.montant for a in articles), 2) == 846.45
