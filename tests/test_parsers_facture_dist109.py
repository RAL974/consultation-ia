"""
Parser facture 109 Distribution (moteur.fournisseurs.dist109.parse_facture_109)
— voir CLAUDE.md, session F2, "Volet 2"/"Rapprochement factures". Sur de
VRAIS PDF (tests/fixtures/, jamais de texte inventé — règle d'or du projet),
extraits du lot de cadrage de 79 vraies factures (79/79 exactes sur le
Total HT affiché).
"""

from pathlib import Path

from moteur.lecture_pdf import lire_pdf
from moteur.fournisseurs.dist109 import parse_facture_109
from moteur.rapprochement.matching_facture import est_bdc_manuel_24x

FIXTURES = Path(__file__).parent / "fixtures"


def _parser(nom_fixture):
    texte = lire_pdf(FIXTURES / nom_fixture)
    return parse_facture_109(texte)


def test_parse_facture_dist109_1_simple_bc_direct():
    """N°Réf.Client au format direct du Suivi ("123.075"), 1 seul BL."""

    f = _parser("facture_dist109_1_simple.pdf")

    assert f.fournisseur == "109 DISTRIBUTION"
    assert f.numero_facture == "360311"
    assert f.date_facture == "15/07/2026"
    assert f.date_echeance == "29/08/2026"
    assert f.numeros_commande == ["123.075"]
    assert f.numeros_bl == ["724331"]
    assert f.type_document == "FACTURE"
    assert f.total_ht_affiche == 534.0

    assert len(f.lignes) == 2

    l0 = f.lignes[0]
    assert l0.reference_fournisseur == "P03101"
    assert l0.designation == "POT POINT DE CENTRE ( DALLE PLEINE )"
    assert l0.quantite_facturee == 200.0
    assert l0.prix_unitaire_ht == 0.6
    assert l0.montant_ht == 120.0
    assert l0.numero_bl == "724331"
    assert l0.numero_commande == ""  # pas encore résolu au parsing, voir matching_facture

    l1 = f.lignes[1]
    assert l1.reference_fournisseur == "F2U15RVVOO"
    assert l1.quantite_facturee == 300.0
    assert l1.prix_unitaire_ht == 1.38
    assert l1.montant_ht == 414.0

    assert round(sum(l.montant_ht for l in f.lignes), 2) == f.total_ht_affiche


def test_parse_facture_dist109_2_bc_interne_non_exploitable():
    """N°Réf.Client au format INTERNE 109 ("BC 241659") — ne correspond pas
    au format du Suivi, numeros_commande doit rester VIDE (résolution
    laissée à matching_facture/pipeline_facture, jamais devinée ici)."""

    f = _parser("facture_dist109_2_bc_interne.pdf")

    assert f.numero_facture == "360310"
    assert f.numeros_commande == []
    assert f.numeros_commande_bruts == ["BC 241659"]  # voir CauseFacture.BDC_MANUEL_24X
    assert est_bdc_manuel_24x(f.numeros_commande_bruts)
    assert f.numeros_bl == ["723657"]
    assert f.total_ht_affiche == 87.5

    assert len(f.lignes) == 2
    assert [l.reference_fournisseur for l in f.lignes] == ["P07312", "515509"]
    assert round(sum(l.montant_ht for l in f.lignes), 2) == f.total_ht_affiche


def test_parse_facture_dist109_5_bon_manuel_bc_24x():
    """2e confirmation réelle du format "bon manuel" (voir test 2 et
    CauseFacture.BDC_MANUEL_24X, moteur.rapprochement.matching_facture) :
    N°Réf.Client = "BC 241766" (préfixe BC + 24 + 4 chiffres) — commande
    passée sur un carnet papier, structurellement non rattachable à une
    ligne du Suivi, jamais devinée."""

    f = _parser("facture_109_362777_bdc_manuel.pdf")

    assert f.numero_facture == "362777"
    assert f.numeros_commande == []
    assert f.numeros_commande_bruts == ["BC 241766"]
    assert est_bdc_manuel_24x(f.numeros_commande_bruts)
    assert f.numeros_bl == ["732796"]
    assert f.total_ht_affiche == 62.0

    assert len(f.lignes) == 1
    assert f.lignes[0].reference_fournisseur == "CDS764F"
    assert round(sum(l.montant_ht for l in f.lignes), 2) == f.total_ht_affiche


def test_parse_facture_dist109_6_multi_bl_meme_reference_bases_de_lagregation():
    """Fixture support de la correction 1a (agrégation multi-BL, voir
    moteur.rapprochement.matching_facture.agreger_lignes_meme_reference et
    tests/test_rapprochement_matching_facture.py) — verrouille ICI
    seulement le PARSING (11 lignes brutes réparties sur 3 BL, P03200 et
    F2U15RVVOO chacun répartis sur 2 BL), l'agrégation elle-même est
    testée côté matching."""

    f = _parser("facture_109_362840_multi_bl_meme_ref.pdf")

    assert f.numero_facture == "362840"
    assert f.numeros_commande == ["123.089"]
    assert f.numeros_bl == ["731835", "731846", "731934"]
    assert f.total_ht_affiche == 1972.0

    assert len(f.lignes) == 11
    refs = [l.reference_fournisseur for l in f.lignes]
    assert refs.count("P03200") == 2
    assert refs.count("F2U15RVVOO") == 2
    assert round(sum(l.montant_ht for l in f.lignes), 2) == f.total_ht_affiche


def test_parse_facture_dist109_7_partielle_facturation_inferieure_a_la_livraison():
    """Fixture "partielle" (voir CLAUDE.md session S0) : une facturation
    PARTIELLE d'un BL réel, sans bug de parsing — le rapprochement doit la
    classer QTE_PARTIELLE (voir test_rapprochement_matching_facture.py),
    pas un correctif de ce parser."""

    f = _parser("facture_109_362763_partielle.pdf")

    assert f.numero_facture == "362763"
    assert f.numeros_commande == ["M3.23.024"]
    assert f.numeros_bl == ["729174"]
    assert f.total_ht_affiche == 135.2

    assert len(f.lignes) == 3
    assert round(sum(l.montant_ht for l in f.lignes), 2) == f.total_ht_affiche


def test_parse_facture_dist109_3_multi_bl_meme_commande():
    """3 BL cités, tous sous la MÊME commande (cas le plus courant des
    factures multi-BL réelles, 8/9 sur le lot de cadrage) — chaque ligne
    porte le n° du bloc de BL auquel elle appartient."""

    f = _parser("facture_dist109_3_multi_bl_meme_commande.pdf")

    assert f.numero_facture == "360366"
    assert f.numeros_commande == ["129.034"]
    assert f.numeros_bl == ["726634", "726648", "726699"]
    assert f.total_ht_affiche == 427.0

    assert len(f.lignes) == 5
    assert [l.numero_bl for l in f.lignes] == [
        "726634", "726634", "726634", "726648", "726699",
    ]
    assert [l.reference_fournisseur for l in f.lignes] == [
        "625160", "SILICONE", "50033", "CABA1.5C100M", "50033",
    ]
    assert round(sum(l.montant_ht for l in f.lignes), 2) == f.total_ht_affiche


def test_parse_facture_dist109_4_multi_bl_multi_commande_ecopart():
    """Cas réel le plus complexe du lot de cadrage (Facture_365533.pdf) :
    N°Réf.Client illisible tel quel ("ORANE; 132.008; 132.00", tronqué),
    3 BL couvrant en réalité 2 commandes DIFFÉRENTES (voir CLAUDE.md),
    ET la seule facture du lot à afficher un Eco-part non nul sur une ligne
    (4,41€, décale Total/Description/TVA d'un cran — voir _bloc_ligne_facture).

    Verrouille aussi l'absence de la ligne FANTÔME que produisait une 1ère
    version du parser (scan non borné par bloc de BL) : le pied de page
    "Code/Taux/Base HT/Montant TVA" contient lui aussi des codes TVA
    isolés (C1, C4) qui, non bornés, formaient un faux article
    "référence 8,50, qté 718,41" — exactement 3 lignes attendues ici,
    pas 4."""

    f = _parser("facture_dist109_4_multi_bl_multi_commande_ecopart.pdf")

    assert f.numero_facture == "365533"
    assert f.numeros_commande == []  # N°Réf.Client non exploitable tel quel
    assert f.numeros_bl == ["706544", "736589", "736611"]
    assert f.total_ht_affiche == 854.0

    assert len(f.lignes) == 3

    l0 = f.lignes[0]  # ligne avec Eco-part non nul
    assert l0.reference_fournisseur == "302889"
    assert l0.quantite_facturee == 21.0
    assert l0.prix_unitaire_ht == 34.0
    assert l0.montant_ht == 714.0
    assert l0.numero_bl == "706544"

    assert f.lignes[1].numero_bl == "736589"
    assert f.lignes[2].numero_bl == "736611"
    assert f.lignes[1].reference_fournisseur == f.lignes[2].reference_fournisseur == "CAP856574"

    assert round(sum(l.montant_ht for l in f.lignes), 2) == f.total_ht_affiche
