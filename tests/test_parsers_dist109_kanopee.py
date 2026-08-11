"""
109 DISTRIBUTION — 2e PDF réel (chantier Kanopée CDC), ancien gabarit
("109 Est/Sud/Nord/Ouest", légalement "109 Holding") qui ne contient
JAMAIS le mot "DISTRIBUTION" dans le corps du devis, seulement dans le
pied de page légal — la détection (moteur/detecteur.py) le manquait
entièrement ("Fournisseur non reconnu, PDF ignoré").

Structure "devis_bpu" déjà connue, mais 3 lignes/18 ont une référence
contenant un tiret ("M-160101") ou un espace ("GOUJON 8X70"), hors du
format alphanumérique strict jusqu'ici supposé — elles étaient
silencieusement perdues (le contrôle Total HT du PDF l'a révélé : 0,02€
d'écart accepté, ici 1 593€ constatés). Total HT retombe exactement une
fois les 18 lignes extraites."""

from moteur.detecteur import detecter_fournisseur
from moteur.fournisseurs.dist109 import parse_109
from moteur.modele import Article

from conftest import FIXTURES


def _texte(nom):
    import fitz
    doc = fitz.open(FIXTURES / nom)
    return "\n".join(p.get_text() for p in doc)


def test_detecte_109_distribution_sans_le_mot_distribution():
    assert detecter_fournisseur(_texte("dist109_kanopee.pdf")) == "109 DISTRIBUTION"


def test_parse_dist109_kanopee_references_avec_tiret_et_espace():
    articles = parse_109(_texte("dist109_kanopee.pdf"))

    assert articles == [
        Article(fournisseur='109 DISTRIBUTION', devis='169692', reference_fournisseur='2/10070', reference_distributeur='', designation='DALLE PERFOREE BPE 60X100 PVCM1UV 7035', quantite=99.0, unite='UN', prix_brut=8.9, prix_net=8.9, montant=881.1, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='169692', reference_fournisseur='2/10072', reference_distributeur='', designation='DALLE PERFOREE BPE 60X200 PVCM1UV 7035', quantite=150.0, unite='UN', prix_brut=15.0, prix_net=15.0, montant=2250.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='169692', reference_fournisseur='2/10073', reference_distributeur='', designation='DALLE PERFOREE BPE 60X300 PVCM1UV 7035', quantite=75.0, unite='UN', prix_brut=16.5, prix_net=16.5, montant=1237.5, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='169692', reference_fournisseur='2/10084', reference_distributeur='', designation='ECLISSE JUBPE 60 PVCUV 7035', quantite=220.0, unite='UN', prix_brut=1.2, prix_net=1.2, montant=264.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='169692', reference_fournisseur='2/10017', reference_distributeur='', designation='BOULON 8X 20 PVC 7035', quantite=4.0, unite='UN', prix_brut=29.9, prix_net=29.9, montant=119.6, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='169692', reference_fournisseur='M-160101', reference_distributeur='', designation='CDC TOL H48 OBR2-99-MAGNELIS PRE-ECLISSE', quantite=126.0, unite='UN', prix_brut=8.0, prix_net=8.0, montant=1008.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='169692', reference_fournisseur='M-160151', reference_distributeur='', designation='CDC TOL H48 OBR2-147-MAGNELIS PRE-ECLISSE', quantite=30.0, unite='UN', prix_brut=9.5, prix_net=9.5, montant=285.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='169692', reference_fournisseur='915210', reference_distributeur='', designation='CDC FIL H100 OBCLICK-L100-ZnALI', quantite=90.0, unite='UN', prix_brut=5.9, prix_net=5.9, montant=531.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='169692', reference_fournisseur='915220', reference_distributeur='', designation='CDC FIL H200 OBCLICK-L200-ZnALI', quantite=102.0, unite='UN', prix_brut=7.5, prix_net=7.5, montant=765.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='169692', reference_fournisseur='915230', reference_distributeur='', designation='CDC FIL H300 OBCLICK-L300-ZnALI', quantite=102.0, unite='UN', prix_brut=11.0, prix_net=11.0, montant=1122.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='169692', reference_fournisseur='252715', reference_distributeur='', designation='CCA 100 PENDARD EN C L150-GAC', quantite=313.0, unite='UN', prix_brut=6.5, prix_net=6.5, montant=2034.5, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='169692', reference_fournisseur='252720', reference_distributeur='', designation='CCA 150 PENDARD EN C L200-GAC', quantite=30.0, unite='UN', prix_brut=7.0, prix_net=7.0, montant=210.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='169692', reference_fournisseur='252725', reference_distributeur='', designation='CCA 200 PENDARD EN C L250-GAC', quantite=250.0, unite='UN', prix_brut=7.5, prix_net=7.5, montant=1875.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='169692', reference_fournisseur='252735', reference_distributeur='', designation='CCA 300 PENDARD EN C L350-GAC', quantite=173.0, unite='UN', prix_brut=9.0, prix_net=9.0, montant=1557.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='169692', reference_fournisseur='250002', reference_distributeur='', designation='CLAME OB20 -DC', quantite=200.0, unite='UN', prix_brut=0.4, prix_net=0.4, montant=80.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='169692', reference_fournisseur='250003', reference_distributeur='', designation='CLAME OB30 - DC', quantite=200.0, unite='UN', prix_brut=0.45, prix_net=0.45, montant=90.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='169692', reference_fournisseur='260630', reference_distributeur='', designation='BOULON TRCC 6 X 30 - GEOMET BOITE DE 100', quantite=5.0, unite='UN', prix_brut=21.0, prix_net=21.0, montant=105.0, disponibilite=''),
        Article(fournisseur='109 DISTRIBUTION', devis='169692', reference_fournisseur='GOUJON 8X70', reference_distributeur='', designation="GOUJON D'ANCRAGE AC ZG M8X70", quantite=600.0, unite='UN', prix_brut=0.5, prix_net=0.5, montant=300.0, disponibilite=''),
    ]

    # Total HT du PDF (14 714,70€) retombe exactement : aucune ligne perdue.
    assert round(sum(a.montant for a in articles), 2) == 14714.70
