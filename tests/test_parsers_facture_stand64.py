"""
Parser facture Stand 64 (moteur.fournisseurs.stand64.parse_facture_stand64)
— session F4 suite (2026-09-02+), voir CLAUDE.md. Sur de VRAIS PDF
(tests/fixtures/, jamais de texte inventé — règle d'or du projet), déposés
directement par l'acheteur (~25 pièces, texte natif — jamais de scan,
contrairement au BL du même fournisseur).
"""

from pathlib import Path

from moteur.lecture_pdf import lire_pdf
from moteur.fournisseurs.stand64 import parse_facture_stand64

FIXTURES = Path(__file__).parent / "fixtures"


def _parser(nom_fixture):
    texte = lire_pdf(FIXTURES / nom_fixture)
    return parse_facture_stand64(texte)


def test_parse_facture_stand64_1_simple():
    """Cas de base : 3 lignes, aucune remise, aucune éco-part."""

    f = _parser("facture_stand64_1_simple.pdf")

    assert f.fournisseur == "STAND 64"
    assert f.numero_facture == "34556"
    assert f.date_facture == "31/08/2026"
    assert f.date_echeance == "15/10/2026"
    assert f.numeros_commande == ["M4.270"]
    assert f.numeros_bl == ["45618"]
    assert f.total_ht_affiche == 133.0

    assert len(f.lignes) == 3
    l0 = f.lignes[0]
    assert l0.reference_fournisseur == "WESTI-COMET-KITLUM-N"
    assert l0.designation == "COMET KIT LUMIERE 2X40W NOIR"
    assert l0.quantite_facturee == 1.0
    assert l0.prix_unitaire_ht == 7.0
    assert l0.montant_ht == 7.0

    assert round(sum(l.montant_ht for l in f.lignes), 2) == 133.0


def test_parse_facture_stand64_2_remises_variables_commande_point_decale():
    """Bug réel corrigé : un mot de désignation tout en MAJUSCULES sans
    chiffre ni tiret ("CHAINETTE", fin d'une désignation qui déborde sur 2
    lignes) était pris à tort pour une référence — toutes les vraies
    références contiennent au moins un tiret. Le n° de commande de
    l'en-tête ("BC N°M2.220.78") a un point décalé par rapport à notre
    propre BON DE COMMANDE en pied de facture ("M2.22.078", mêmes chiffres
    "22078") : celui-ci est préféré."""

    f = _parser("facture_stand64_2_remises_variables_commande_point_decale.pdf")

    assert f.numero_facture == "33707"
    assert f.numeros_commande == ["M2.22.078"]
    assert f.numeros_bl == ["44401"]
    assert f.total_ht_affiche == 3615.0

    assert len(f.lignes) == 6
    refs = {l.reference_fournisseur: l for l in f.lignes}

    # Remise 5,56% : Total HT = Qté x P.U Net (91,00 = 94,00 x (1-0,0556)).
    assert refs["WESTI-73044"].quantite_facturee == 20.0
    assert refs["WESTI-73044"].prix_unitaire_ht == 91.0
    assert refs["WESTI-73044"].montant_ht == 1820.0
    assert refs["WESTI-78017"].designation == "COMET 5 PALES Ø132 2X40W E14 BLANC + ERABLE/BLANC + CHAINETTE"
    assert refs["WESTI-78017"].montant_ht == 980.0
    # Ligne gratuite (Total HT nul, jamais rejetée pour autant).
    assert refs["WESTI-99993"].quantite_facturee == 50.0
    assert refs["WESTI-99993"].montant_ht == 0.0

    assert round(sum(l.montant_ht for l in f.lignes), 2) == 3615.0


def test_parse_facture_stand64_3_ecopart_variable_par_ligne():
    """LIMITE LA PLUS IMPORTANTE trouvée : l'Eco-part est présente ou non
    LIGNE PAR LIGNE au sein d'un MÊME document (2 lignes sur 3 avec
    Eco-part renseignée, 1 sans) — un compte de valeurs numériques fixe
    après le code TVA aurait décalé une ligne sur deux. L'ancrage sur la
    référence (fiable dans les deux sens : Qté/P.U Net à position fixe
    juste après le code TVA, Total HT/désignation retrouvés en remontant
    depuis la référence) résout ça sans avoir besoin de connaître ce
    compte."""

    f = _parser("facture_stand64_3_ecopart_variable_par_ligne.pdf")

    assert f.numero_facture == "34184"
    assert f.numeros_commande == ["M3.14.338"]
    assert f.numeros_bl == ["44745"]
    assert f.total_ht_affiche == 1312.20

    assert len(f.lignes) == 3
    l0, l1, l2 = f.lignes
    assert l0.reference_fournisseur == "UPSHI-FOCUS58A-3P-35W-30-36-BL"
    assert l0.quantite_facturee == 14.0
    assert l0.montant_ht == 952.0
    assert l1.reference_fournisseur == "ACB-T37640N"
    assert l1.designation == "ZOOM PROJECTEUR SUR RAIL TRI 1XGU10 NOIR"
    assert l1.quantite_facturee == 8.0
    assert l1.montant_ht == 344.0
    assert l2.reference_fournisseur == "LEDVA-818392"
    assert l2.quantite_facturee == 3.0
    assert l2.montant_ht == 16.2

    assert round(sum(l.montant_ht for l in f.lignes), 2) == 1312.20


def test_parse_facture_stand64_4_reference_avec_plus():
    """Référence contenant un "+" ("ELIOT-ES52-2678-BLC+BLC", déjà connue
    côté BL du même fournisseur, commande M2.5.126) — bug réel corrigé,
    le motif référence n'acceptait pas ce caractère."""

    f = _parser("facture_stand64_4_reference_avec_plus.pdf")

    assert f.numero_facture == "34558"
    assert f.numeros_commande == ["M2.5.126"]

    assert len(f.lignes) == 1
    l0 = f.lignes[0]
    assert l0.reference_fournisseur == "ELIOT-ES52-2678-BLC+BLC"
    assert l0.quantite_facturee == 18.0
    assert l0.montant_ht == 1710.0

    assert round(sum(l.montant_ht for l in f.lignes), 2) == 1710.0
