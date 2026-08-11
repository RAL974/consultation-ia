"""
IMPORELEC -- vrai PDF (195 lignes, tests/fixtures/imporelec.pdf). Fixture
trop volumineuse pour une liste exhaustive comme les autres tests de ce
dossier (voir CLAUDE.md) : verrouillée par nombre de lignes + somme des
montants (recoupée avec le "Sous Total" du PDF avant remise globale) +
les 3 premieres et 3 dernieres lignes + les 6 lignes touchees par le bug
connu (sous-total intermediaire du PDF pris a tort pour une reference,
voir le bandeau de moteur/fournisseurs/imporelec.py).
"""

from moteur.lecture_pdf import lire_pdf
from moteur.modele import Article
from moteur.detecteur import detecter_fournisseur
from moteur.fournisseurs.imporelec import parse_imporelec

from conftest import FIXTURES


def _texte(nom):
    return lire_pdf(FIXTURES / nom)


def test_detection_imporelec():
    assert detecter_fournisseur(_texte("imporelec.pdf")) == "IMPORELEC"


def test_parse_imporelec_volume_et_total():
    articles = parse_imporelec(_texte("imporelec.pdf"))
    assert len(articles) == 195
    # Sous Total du PDF avant remise globale (2.10%) : 44 675,70€ -- la
    # somme des lignes retombe dessus a 0,01€ pres (arrondi sur 195 lignes).
    assert abs(round(sum(a.montant for a in articles), 2) - 44675.71) < 0.02


def test_parse_imporelec_premieres_lignes():
    articles = parse_imporelec(_texte("imporelec.pdf"))
    assert articles[:3] == [
        Article(fournisseur='IMPORELEC', devis='807625', reference_fournisseur='A9P71602', reference_distributeur='A9P71602', designation='Acti9 iDT40K disjoncteur mod 1P+N 2A courbe C 4500A/4,5kA', quantite=1.0, unite='UN', prix_brut=14.94, prix_net=14.94, montant=14.94, disponibilite=''),
        Article(fournisseur='IMPORELEC', devis='807625', reference_fournisseur='A9P71606', reference_distributeur='A9P71606', designation='Acti9 iDT40K disjoncteur mod 1P+N 6A courbe C 4500A/4,5kA', quantite=1.0, unite='UN', prix_brut=14.54, prix_net=14.54, montant=14.54, disponibilite=''),
        Article(fournisseur='IMPORELEC', devis='807625', reference_fournisseur='A9P71610', reference_distributeur='A9P71610', designation='Acti9 iDT40K disjoncteur mod 1P+N 10A courbe C 4500A/4,5kA', quantite=1.0, unite='UN', prix_brut=11.32, prix_net=11.32, montant=11.32, disponibilite=''),
    ]


def test_parse_imporelec_dernieres_lignes():
    articles = parse_imporelec(_texte("imporelec.pdf"))
    assert articles[-3:] == [
        Article(fournisseur='IMPORELEC', devis='807625', reference_fournisseur='LE1D18P7', reference_distributeur='LE1D18P7', designation='TeSys LE1D - démarreur en coffret - 18A - bobine 230Vca', quantite=1.0, unite='UN', prix_brut=69.89, prix_net=69.89, montant=69.89, disponibilite=''),
        Article(fournisseur='IMPORELEC', devis='807625', reference_fournisseur='LRD14', reference_distributeur='LRD14', designation='TeSys LRD - relais de protection thermique - 7..10A - classe 10A', quantite=1.0, unite='UN', prix_brut=40.2, prix_net=40.2, montant=40.2, disponibilite=''),
        Article(fournisseur='IMPORELEC', devis='807625', reference_fournisseur='LRD21', reference_distributeur='LRD21', designation='TeSys LRD - relais de protection thermique - 12..18A - classe 10A', quantite=1.0, unite='UN', prix_brut=44.16, prix_net=44.16, montant=44.16, disponibilite=''),
    ]


def test_parse_imporelec_bug_sous_total_connu():
    """Verrouille le bug connu (voir moteur/fournisseurs/imporelec.py) :
    ne pas le "corriger" silencieusement sans s'en rendre compte.    """
    articles = parse_imporelec(_texte("imporelec.pdf"))
    corrompues = [a for a in articles if "€" in a.reference_fournisseur]
    assert corrompues == [
        Article(fournisseur='IMPORELEC', devis='807625', reference_fournisseur='406.94 €', reference_distributeur='', designation='22 A9F77216 A9F74216 Acti9, iC60N disjoncteur 2P 16A courbe C', quantite=2.0, unite='UN', prix_brut=21.65, prix_net=21.65, montant=43.3, disponibilite=''),
        Article(fournisseur='IMPORELEC', devis='807625', reference_fournisseur='4 425.35 €', reference_distributeur='', designation='59 A9P24606 A9P24606 Acti9 iDT40N disjoncteur mod 1P+N C 6A 6000A/10kA', quantite=1.0, unite='UN', prix_brut=30.66, prix_net=30.66, montant=30.66, disponibilite=''),
        Article(fournisseur='IMPORELEC', devis='807625', reference_fournisseur='22 431.55 €', reference_distributeur='', designation='89 C16F4TM160 C16F4TM160 ComPacT NSX160F - disjoncteur - TM-D 160A - 4P4D - 36kA - montage fi', quantite=1.0, unite='UN', prix_brut=503.04, prix_net=503.04, montant=503.04, disponibilite=''),
        Article(fournisseur='IMPORELEC', devis='807625', reference_fournisseur='36 641.63 €', reference_distributeur='119', designation='A9F95232 A9F95232 Acti9, iC60L disjoncteur 2P 32A courbe K', quantite=2.0, unite='UN', prix_brut=74.01, prix_net=74.01, montant=148.02, disponibilite=''),
        Article(fournisseur='IMPORELEC', devis='807625', reference_fournisseur='41 711.52 €', reference_distributeur='157', designation='A9S61220 A9S61220 ISW INTER A VOYANT 2P 20A', quantite=1.0, unite='UN', prix_brut=15.73, prix_net=15.73, montant=15.73, disponibilite=''),
        Article(fournisseur='IMPORELEC', devis='807625', reference_fournisseur='43 963.97 €', reference_distributeur='183', designation='A9C22722 A9C22722 Acti9, iCT contacteur 20A 2NO  230...240VCA 50Hz', quantite=1.0, unite='UN', prix_brut=29.96, prix_net=29.96, montant=29.96, disponibilite=''),
    ]
