"""
COMINTER — 3e format réel (chantier Kanopée CDC, devis DV121328/DV124395),
absent des 2 formats déjà connus (v1/v2 dans moteur/fournisseurs/cominter.py) :
la référence vient AVANT la désignation (v2 la met après le conditionnement).
Les 2 PDF retombaient sur 0 article extrait avant l'ajout de v3.

Les deux totaux extraits retombent exactement sur le Total HT affiché par
chaque PDF (6 010,49€ et 9 515,39€) — aucune ligne oubliée."""

from moteur.fournisseurs.cominter import parse_cominter
from moteur.modele import Article

from conftest import FIXTURES


def _texte(nom):
    import fitz
    doc = fitz.open(FIXTURES / nom)
    return "\n".join(p.get_text() for p in doc)


def test_parse_cominter_kanopee_v3():
    articles = parse_cominter(_texte("cominter_kanopee.pdf"))

    assert articles == [
        Article(fournisseur='COMINTER', devis='DV121328', reference_fournisseur='VECC60100/1', reference_distributeur='', designation='Chemin de cables ZMKBSCL 60 larg 100 Barre 3 MT', quantite=42.0, unite='BARRE', prix_brut=25.41, prix_net=25.41, montant=1067.22, disponibilite=''),
        Article(fournisseur='COMINTER', devis='DV121328', reference_fournisseur='VECC60150/1', reference_distributeur='', designation='Chemin de cables ZMKBSCL 60 larg 150 Barre 3 MT', quantite=10.0, unite='BARRE', prix_brut=30.41, prix_net=30.41, montant=304.10, disponibilite=''),
        Article(fournisseur='COMINTER', devis='DV121328', reference_fournisseur='VECF60100/1', reference_distributeur='', designation='Chemin de câbles en fil ZAVFUL 60 larg 100  Barre 3 MT', quantite=30.0, unite='BARRE', prix_brut=15.12, prix_net=15.12, montant=453.60, disponibilite=''),
        Article(fournisseur='COMINTER', devis='DV121328', reference_fournisseur='VECF60200/1', reference_distributeur='', designation='Chemin de câbles en fil ZAVFUL 60 larg 200  Barre 3 MT', quantite=34.0, unite='BARRE', prix_brut=17.18, prix_net=17.18, montant=584.12, disponibilite=''),
        Article(fournisseur='COMINTER', devis='DV121328', reference_fournisseur='VECF60300/1', reference_distributeur='', designation='Chemin de câbles en fil ZAVFUL 60 larg 300  Barre 3 MT', quantite=34.0, unite='BARRE', prix_brut=21.03, prix_net=21.03, montant=715.02, disponibilite=''),
        Article(fournisseur='COMINTER', devis='DV121328', reference_fournisseur='VECHD170150', reference_distributeur='', designation='Console Comega en C 170X150 GAG universelle', quantite=30.0, unite='UN', prix_brut=5.93, prix_net=5.93, montant=177.90, disponibilite=''),
        Article(fournisseur='COMINTER', devis='DV121328', reference_fournisseur='VECHD170200', reference_distributeur='', designation='Console Comega en C 170X200 GAG universelle', quantite=250.0, unite='UN', prix_brut=6.26, prix_net=6.26, montant=1565.00, disponibilite=''),
        Article(fournisseur='COMINTER', devis='DV121328', reference_fournisseur='VECHD170300', reference_distributeur='', designation='Console Comega en C 170X300 GAG universelle', quantite=173.0, unite='UN', prix_brut=6.61, prix_net=6.61, montant=1143.53, disponibilite=''),
    ]

    assert round(sum(a.montant for a in articles), 2) == 6010.49


def test_parse_cominter_kanopee_fca_v3():
    articles = parse_cominter(_texte("cominter_kanopee_fca.pdf"))

    assert len(articles) == 12
    assert round(sum(a.montant for a in articles), 2) == 9515.39
    # Deux lignes VECC60100/1 et deux lignes VECHD170150 à des quantités et
    # prix DIFFÉRENTS (commandes distinctes dans le même devis) : ne
    # doivent pas s'écraser entre elles.
    vecc60100 = [a for a in articles if a.reference_fournisseur == "VECC60100/1"]
    assert len(vecc60100) == 2
    assert {a.quantite for a in vecc60100} == {33.0, 42.0}
