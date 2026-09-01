"""
Parser facture Coredime (moteur.fournisseurs.coredime.parse_facture_coredime)
— voir CLAUDE.md, session F4. Sur de VRAIS PDF (tests/fixtures/, jamais de
texte inventé — règle d'or du projet), extraits des 70 vraies factures
transmises par Prisca LEBLÉ (comptable) pour Coredime, juillet 2026.
"""

from pathlib import Path

from moteur.lecture_pdf import lire_pdf
from moteur.fournisseurs.coredime import parse_facture_coredime

FIXTURES = Path(__file__).parent / "fixtures"


def _parser(nom_fixture):
    texte = lire_pdf(FIXTURES / nom_fixture)
    return parse_facture_coredime(texte)


def test_parse_facture_coredime_1_simple():
    """Cas de base : 1 seul bloc "BON D'EXPEDITION", 2 lignes propres,
    Total HT exact."""

    f = _parser("facture_coredime_1_simple.pdf")

    assert f.fournisseur == "COREDIME"
    assert f.type_document == "FACTURE"
    assert f.numero_facture == "6107293"
    assert f.date_facture == "03/07/2026"
    assert f.numeros_commande == ["123.077"]
    assert f.numeros_bl == ["B028249"]
    assert f.total_ht_affiche == 287.0

    assert len(f.lignes) == 2
    l0, l1 = f.lignes
    assert l0.reference_fournisseur == "LEG06620"
    assert l0.designation == "ICTA 3422 20 ATF STANDARD 100M"
    assert l0.quantite_facturee == 500.0
    assert l0.prix_unitaire_ht == 0.35
    assert l0.montant_ht == 175.0
    assert l0.numero_bl == "B028249"
    assert l1.reference_fournisseur == "LEG06625"
    assert l1.montant_ht == 112.0

    assert round(sum(l.montant_ht for l in f.lignes), 2) == f.total_ht_affiche


def test_parse_facture_coredime_2_multi_bl_meme_commande():
    """2 blocs "BON D'EXPEDITION" (B028558.1 / .2), même commande
    M3.23.020 — numeros_commande ne garde que la valeur DISTINCTE (1 seule
    ici), chaque ligne garde son propre numero_bl (voir bandeau du
    module : la résolution de commande par bloc est déjà générique côté
    pipeline_facture, aucune modif nécessaire)."""

    f = _parser("facture_coredime_2_multi_bl.pdf")

    assert f.numeros_commande == ["M3.23.020"]
    assert f.numeros_bl == ["B028558.1", "B028558.2"]
    assert f.total_ht_affiche == 241.25
    assert len(f.lignes) == 3

    assert [l.numero_bl for l in f.lignes] == ["B028558.1", "B028558.1", "B028558.2"]
    assert f.lignes[2].reference_fournisseur == "LEG086147"
    assert f.lignes[2].montant_ht == 76.35

    assert round(sum(l.montant_ht for l in f.lignes), 2) == f.total_ht_affiche


def test_parse_facture_coredime_3_avoir_jamais_extrait_en_detail():
    """AVOIR (détecté via la ligne de métadonnées "...;Avoir;...", 1er
    exemple réel de ce projet) — type_document="AVOIR", header renseigné,
    mais AUCUNE ligne extraite (format numérique différent, sans intérêt :
    un AVOIR n'est de toute façon jamais rapproché automatiquement, voir
    moteur.rapprochement.pipeline_facture)."""

    f = _parser("facture_coredime_3_avoir.pdf")

    assert f.type_document == "AVOIR"
    assert f.numero_facture == "6108972"
    assert f.date_facture == "17/08/2026"
    assert f.lignes == []
    assert f.total_ht_affiche is None


def test_parse_facture_coredime_4_reference_avec_tiret():
    """Référence contenant un tiret ("WAG221-425") — bug réel corrigé
    (voir bandeau MOTIF_LIGNE_FACTURE_COREDIME) : exclue par la 1ère
    version de la regex, faisant perdre TOUTE la ligne."""

    f = _parser("facture_coredime_4_reference_tiret.pdf")

    assert len(f.lignes) == 1
    l0 = f.lignes[0]
    assert l0.reference_fournisseur == "WAG221-425"
    assert l0.designation == "BORNE WAGO 221 GREEN 5X4MM"
    assert l0.quantite_facturee == 100.0
    assert l0.montant_ht == 90.86
    assert f.total_ht_affiche == 90.86
    assert round(sum(l.montant_ht for l in f.lignes), 2) == f.total_ht_affiche


def test_parse_facture_coredime_5_reference_numerique_et_ligne_apres_important():
    """2 bugs réels corrigés sur ce même document : (1) référence purement
    numérique ("227060133", une vraie référence article, pas un code de
    fret) — la regex exigeait une lettre en tête ; (2) la 2e ligne
    ("SIBP16840") est imprimée APRÈS le repère de contenu "----- IMPORTANT
    -----" dans le flux PyMuPDF scramblé — bornage corrigé pour se caler
    sur la vraie limite de page (métadonnées), jamais un repère de
    contenu (voir bandeau du module)."""

    f = _parser("facture_coredime_5_reference_numerique.pdf")

    assert len(f.lignes) == 2
    assert f.lignes[0].reference_fournisseur == "227060133"
    assert f.lignes[0].montant_ht == 395.0
    assert f.lignes[1].reference_fournisseur == "SIBP16840"
    assert f.lignes[1].montant_ht == 315.0
    assert f.total_ht_affiche == 710.0
    assert round(sum(l.montant_ht for l in f.lignes), 2) == f.total_ht_affiche


def test_parse_facture_coredime_6_remise_double_appariee_sans_ambiguite():
    """"Remise 35,00+26,00%" (double remise) imprimée sur une ligne
    totalement DISJOINTE de sa référence/désignation/quantité dans le flux
    PyMuPDF (positions différentes, pas seulement une ligne d'écart) — 1
    SEULE ligne incomplète + 1 SEULE ligne "Remise" dans ce bloc : aucune
    ambiguïté, appariées automatiquement (voir
    _lignes_remise_double_coredime). Vérifie aussi que le garde-fou
    qté×prix_net fonctionne sur une ligne ainsi reconstruite."""

    f = _parser("facture_coredime_6_remise_double_appariee.pdf")

    assert len(f.lignes) == 2
    refs = {l.reference_fournisseur: l for l in f.lignes}
    assert refs["LEG030804"].quantite_facturee == 14.0
    assert refs["LEG030804"].prix_unitaire_ht == 3.3622
    assert refs["LEG030804"].montant_ht == 47.07
    assert refs["LEG030271"].quantite_facturee == 2.0
    assert refs["LEG030271"].montant_ht == 11.12
    assert f.total_ht_affiche == 58.19
    assert round(sum(l.montant_ht for l in f.lignes), 2) == f.total_ht_affiche


def test_parse_facture_coredime_7_ligne_fret_livraison_avion_incluse():
    """Référence "9993" ("LIVRAISON AVION", frais de port) — élargir la
    référence aux formats purement numériques (voir test 5) la fait
    ressortir comme une LigneFacture à part entière (montant 0,10€,
    négligeable) : accepté, elle ne trouvera simplement aucune
    correspondance dans le Suivi et finira "inconnue" au rapprochement —
    jamais un mauvais rattachement, voir bandeau du module."""

    f = _parser("facture_coredime_7_ligne_fret_exclue.pdf")

    assert len(f.lignes) == 2
    assert f.lignes[0].reference_fournisseur == "9993"
    assert f.lignes[0].montant_ht == 0.1
    assert f.lignes[1].reference_fournisseur == "CAEMPSS424120SH10"
    assert f.lignes[1].montant_ht == 13440.0
    assert f.total_ht_affiche == 13440.1
    # Écart de 0,10€ résiduel connu (2e ligne "9993 LIVRAISON AVION" à
    # 0,00€ non extraite, prix_net "-----" ne matche aucun nombre — voir
    # CLAUDE.md, limite acceptée) : PAS un total exact, volontairement.
    assert round(sum(l.montant_ht for l in f.lignes), 2) == 13440.1


def test_parse_facture_coredime_8_remises_multiples_extraction_partielle_honnete():
    """LIMITE CONNUE ACCEPTÉE (voir CLAUDE.md) : quand un bloc contient
    PLUSIEURS lignes incomplètes ET plusieurs lignes "Remise" (ici :
    document réel avec de nombreuses doubles remises), l'appariement 1:1
    ne s'applique plus (ambiguïté réelle, jamais un choix au hasard) — ces
    lignes restent NON extraites. Ce test fige l'extraction PARTIELLE
    actuelle (12 lignes sur un total de 19 réelles) et vérifie que
    l'écart avec le Total HT affiché est bien signalé (jamais un faux
    total exact)."""

    f = _parser("facture_coredime_8_remises_multiples_partiel.pdf")

    assert f.total_ht_affiche == 1132.51
    assert len(f.lignes) == 12

    total_extrait = round(sum(l.montant_ht for l in f.lignes), 2)
    assert total_extrait == 813.83
    assert total_extrait != f.total_ht_affiche  # écart honnête, pas un faux total exact

    # Les lignes propres (sans double remise) restent exactes.
    refs = {l.reference_fournisseur: l for l in f.lignes}
    assert refs["LEG030015"].montant_ht == 184.0
    assert refs["GFO026705"].montant_ht == 148.8
