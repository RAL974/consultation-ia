"""
Tests de non-régression du parsing BL PROTECTHOMS (Rapprochement AI,
nouveau fournisseur — un seul vrai BL vu à ce jour,
tests/fixtures/bl_protecthoms_1.pdf, copie de a_traiter/BL/). Même esprit
que tests/test_parsers_bl_yesss.py.
"""

import functools

from moteur.ocr import mots_document, texte_ocr
from moteur.detecteur import detecter_fournisseur
from moteur.fournisseurs.protecthoms import parse_bl_protecthoms
from moteur.rapprochement.modele_bl import LigneBL

from conftest import FIXTURES


@functools.lru_cache(maxsize=None)
def _mots(nom):
    return mots_document(FIXTURES / nom)


def test_detection_fournisseur_bl_protecthoms():
    assert detecter_fournisseur(texte_ocr(FIXTURES / "bl_protecthoms_1.pdf")) == "PROTECTHOMS"


def test_parse_bl_protecthoms_1():
    """Équipements de protection individuelle/amiante, pas du matériel
    électrique — fournisseur bien présent dans la liste Fournisseurs du
    Suivi (confirmé), donc en périmètre malgré la nature différente des
    articles. Tableau simple "Reference produit | Designation | Quantites
    | Reste à livrer" : chaque ligne visuelle est déjà un article complet,
    repéré par la FORME de sa référence (1 chiffre + 2 lettres + 6
    chiffres) plutôt que par un en-tête/pied de tableau. "Quantites" est
    la quantité livrée sur CE bon ; "Reste à livrer" (present uniquement
    sur la 2e ligne ici, 20) est purement informatif, jamais soustrait."""

    bl = parse_bl_protecthoms(_mots("bl_protecthoms_1.pdf"))

    assert bl.fournisseur == "PROTECTHOMS"
    assert bl.numero_commande == "M3.15.399"
    assert bl.numero_bl == "BL097191"
    assert bl.date_bl == "22/08/2026"

    assert bl.lignes == [
        LigneBL(reference_fournisseur='2VU043003', designation='COMBIWEECOVERMAX1AMIANT.L', quantite_livree=68.0, prix_net=None, montant=None),
        LigneBL(reference_fournisseur='2VU043004', designation='COMBIWEECOVERMAX1AMIANT.XL', quantite_livree=30.0, prix_net=None, montant=None),
        LigneBL(reference_fournisseur='6SP040000', designation="SACPEIMPRIMEAMIANTEblancI'unite", quantite_livree=250.0, prix_net=None, montant=None),
        LigneBL(reference_fournisseur='6SP041100', designation="BIGBAGONU4POINTS'AMIANTE'", quantite_livree=30.0, prix_net=None, montant=None),
        LigneBL(reference_fournisseur='3GE030010', designation='GANTS87-190T.9.5/10LOTDE12PAIRES', quantite_livree=13.0, prix_net=None, montant=None),
        LigneBL(reference_fournisseur='2VU061300', designation='COUVREBOTTESTYVECK(lapaire)', quantite_livree=150.0, prix_net=None, montant=None),
        LigneBL(reference_fournisseur='5CO010100', designation="SCELLED'IDENTIFICATIONPROTECTHOMS", quantite_livree=30.0, prix_net=None, montant=None),
        LigneBL(reference_fournisseur='1MA010701', designation='PAIREDEFILTRESP3POUR7000ET9000-9030-01', quantite_livree=300.0, prix_net=None, montant=None),
    ]
