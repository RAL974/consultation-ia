"""
SAGEES -- 3 vrais PDF confirmes (tests/fixtures/sagees_1/2/3.pdf), format
"V0" (voir bandeau GABARIT de moteur/fournisseurs/sagees.py).

2 AUTRES formats reels existent chez ce meme fournisseur, NON couverts
(voir CLAUDE.md) : un format avec colonne Reference + Conditionnement
('ESPACE SOLEIL CLINIQUE MAYOTTE.pdf') et un format 'bordereau de prix'
en lettre de couverture avec prix en euros suffixes ('ML20260615 ES.pdf').
"""

from moteur.lecture_pdf import lire_pdf
from moteur.modele import Article
from moteur.detecteur import detecter_fournisseur
from moteur.fournisseurs.sagees import parse_sagees

from conftest import FIXTURES


def _texte(nom):
    return lire_pdf(FIXTURES / nom)


def test_detection_sagees():
    for nom in ("sagees_1.pdf", "sagees_2.pdf", "sagees_3.pdf"):
        assert detecter_fournisseur(_texte(nom)) == "SAGEES"


def test_parse_sagees_1():
    articles = parse_sagees(_texte("sagees_1.pdf"))
    assert articles == [
        Article(fournisseur='SAGEES', devis='DR011352', reference_fournisseur='', reference_distributeur='', designation='PINCE COUPE CABLE CUIVRE & ALU 5 UNITES DISPONIBLE', quantite=1.0, unite='UN', prix_brut=34.21, prix_net=34.21, montant=34.21, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR011352', reference_fournisseur='', reference_distributeur='', designation='COUPE CABLE A CREMAILLERE 34MM 5 UNITES DISPONIBLE', quantite=1.0, unite='UN', prix_brut=307.48, prix_net=307.48, montant=307.48, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR011352', reference_fournisseur='', reference_distributeur='', designation='COUPE CABLE A CREMAILLERRE 52MM 2 UNITES DISPONIBLE', quantite=1.0, unite='UN', prix_brut=335.42, prix_net=335.42, montant=335.42, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR011352', reference_fournisseur='', reference_distributeur='', designation='PINCE A DÉGAINER COMPOSITE 20 VARIABLE MODÈLE B 7 UNITES DISPONIBLE', quantite=1.0, unite='UN', prix_brut=338.19, prix_net=338.19, montant=338.19, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR011352', reference_fournisseur='', reference_distributeur='', designation='APP A CERCLER FEUILLARD A LEVIER 5 UNITES DISPONIBLE', quantite=1.0, unite='UN', prix_brut=144.36, prix_net=144.36, montant=144.36, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR011352', reference_fournisseur='', reference_distributeur='', designation='CLE MUTIPLE PLIABLE 3 UNITES DISPONIBLE', quantite=1.0, unite='UN', prix_brut=56.18, prix_net=56.18, montant=56.18, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR011352', reference_fournisseur='', reference_distributeur='', designation='COUTEAU C1 INJECTÉ - LAME ISOLÉE 12 UNITES DISPONIBLE', quantite=1.0, unite='UN', prix_brut=19.75, prix_net=19.75, montant=19.75, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR011352', reference_fournisseur='', reference_distributeur='', designation='TAPIS ISOLANT POUR POSTE 1 PAR 1 M EP 3 MM 10 UNITES DISPONIBLE', quantite=1.0, unite='UN', prix_brut=72.05, prix_net=72.05, montant=72.05, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR011352', reference_fournisseur='', reference_distributeur='', designation='CLE TRIANGLE 11 9 UNITES DISPONIBLE', quantite=1.0, unite='UN', prix_brut=10.49, prix_net=10.49, montant=10.49, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR011352', reference_fournisseur='', reference_distributeur='', designation='DÉTECTEUR TENSION BT 8 SEUILS 9 UNITES DISPONIBLE', quantite=1.0, unite='UN', prix_brut=121.07, prix_net=121.07, montant=121.07, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR011352', reference_fournisseur='', reference_distributeur='', designation='TOURNEVIS ISOLE EMBOUT PLAT TYPE B 6.5X150 2 UNITES DISPONIBLE', quantite=1.0, unite='UN', prix_brut=9.98, prix_net=9.98, montant=9.98, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR011352', reference_fournisseur='', reference_distributeur='', designation='TOURNEVIS ISOLE EMBOUT PLAT TYPE B 8X150 2 UNITES DISPONIBLE', quantite=1.0, unite='UN', prix_brut=13.65, prix_net=13.65, montant=13.65, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR011352', reference_fournisseur='', reference_distributeur='', designation='TOURNEVIS ISOLE EMBOUT POZIDRIVE 4X75 2 UNITES DISPONIBLE', quantite=1.0, unite='UN', prix_brut=7.32, prix_net=7.32, montant=7.32, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR011352', reference_fournisseur='', reference_distributeur='', designation='TOURNEVIS ISOLE EMBOUT POZIDRIVE 6X125', quantite=1.0, unite='UN', prix_brut=10.86, prix_net=10.86, montant=10.86, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR011352', reference_fournisseur='', reference_distributeur='', designation='TOURNEVIS ISOLE EMBOUT POZIDRIVE 8X150 2 UNITES DISPONIBLE', quantite=1.0, unite='UN', prix_brut=11.52, prix_net=11.52, montant=11.52, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR011352', reference_fournisseur='', reference_distributeur='', designation='TOURNEVIS ISOLE EMBOUT POZIDRIVE 10X200 2 UNITES DISPONIBLE', quantite=1.0, unite='UN', prix_brut=16.15, prix_net=16.15, montant=16.15, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR011352', reference_fournisseur='', reference_distributeur='', designation='DISQUE APPAREIL CONDAMNE PVC 8 UNITES DISPONIBLE', quantite=1.0, unite='UN', prix_brut=3.1, prix_net=3.1, montant=3.1, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR011352', reference_fournisseur='', reference_distributeur='', designation='COFFRET A CLIQUET 3/8 MODELE P8A 2 UNITES DISPONIBLE', quantite=1.0, unite='UN', prix_brut=577.02, prix_net=577.02, montant=577.02, disponibilite=''),
    ]
    assert round(sum(a.montant for a in articles), 2) == 2088.8


def test_parse_sagees_2():
    articles = parse_sagees(_texte("sagees_2.pdf"))
    assert articles == [
        Article(fournisseur='SAGEES', devis='DR011343', reference_fournisseur='', reference_distributeur='', designation='REMBT -  RAC 240² RESTE 2 UNITES EN STOCK - COMMANDE EN COURS - PAS DE DATE DE RECEPTION', quantite=2.0, unite='UN', prix_brut=54.9, prix_net=54.9, montant=109.8, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR011343', reference_fournisseur='', reference_distributeur='', designation='CONNECTEUR MALT KZ 2-95', quantite=2.0, unite='UN', prix_brut=10.83, prix_net=10.83, montant=21.66, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR011343', reference_fournisseur='', reference_distributeur='', designation="ETIQUETTES - D'IDENTIFICATION DES CÂBLES   (LOT DE 50)", quantite=1.0, unite='UN', prix_brut=29.45, prix_net=29.45, montant=29.45, disponibilite=''),
    ]
    assert round(sum(a.montant for a in articles), 2) == 160.91


def test_parse_sagees_3():
    articles = parse_sagees(_texte("sagees_3.pdf"))
    assert articles == [
        Article(fournisseur='SAGEES', devis='DR010805', reference_fournisseur='', reference_distributeur='', designation='Détecteur de Défaut Directionnel et Ampérométrique Alimentation secourue C13-100 - 5A / 14Ah', quantite=1.0, unite='UN', prix_brut=63090.3, prix_net=63090.3, montant=63090.3, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR010805', reference_fournisseur='', reference_distributeur='', designation='LIAISON HT CSD/CSE 50²ALU 3x9M', quantite=2.0, unite='UN', prix_brut=1156.37, prix_net=1156.37, montant=2312.74, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR010805', reference_fournisseur='', reference_distributeur='', designation='Transformateur sec enrobé ECO2021 Enveloppe IP315 + Emballage SEI4c', quantite=2.0, unite='UN', prix_brut=35890.0, prix_net=35890.0, montant=71780.0, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR010805', reference_fournisseur='', reference_distributeur='', designation='Ns1600N debro 4P 2.0 + epanouisseurs 33622 + Mx 230V + OF/SD', quantite=2.0, unite='UN', prix_brut=6168.0, prix_net=6168.0, montant=12336.0, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR010805', reference_fournisseur='', reference_distributeur='', designation='FUSIBLES   63 A 24 KV AVEC PERCUTEUR LOT 3', quantite=4.0, unite='UN', prix_brut=432.0, prix_net=432.0, montant=1728.0, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR010805', reference_fournisseur='', reference_distributeur='', designation='4 coupe-circuit bipolaire + fusibles - 2A / 10A + VIGIREX', quantite=2.0, unite='UN', prix_brut=1280.5, prix_net=1280.5, montant=2561.0, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR010805', reference_fournisseur='', reference_distributeur='', designation='TORE FERMÉ POUR PROTECTION DIFF DIAMÈTRE 80 MM', quantite=1.0, unite='UN', prix_brut=104.28, prix_net=104.28, montant=104.28, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR010805', reference_fournisseur='', reference_distributeur='', designation='KIT/ACC/POSTE', quantite=1.0, unite='UN', prix_brut=1638.73, prix_net=1638.73, montant=1638.73, disponibilite=''),
        Article(fournisseur='SAGEES', devis='DR010805', reference_fournisseur='', reference_distributeur='', designation='KIT/AFF/POSTE', quantite=1.0, unite='UN', prix_brut=82.84, prix_net=82.84, montant=82.84, disponibilite=''),
    ]
    assert round(sum(a.montant for a in articles), 2) == 155633.89

