"""
Parser facture Electric Plus/GMR (moteur.fournisseurs.electricplus.
parse_facture_electricplus_ocr) — session F4 suite (2026-09-02), voir
CLAUDE.md.

GMR n'envoie pas de BL séparé (déjà su côté BL, voir "GABARIT BL" dans
electricplus.py) : sa FACTURE fait déjà office de BL — d'où un parser à
part, "parse_facture_ocr" plutôt que "parse_facture" (ces factures sont des
SCANS, jamais de texte PDF natif, contrairement aux autres fournisseurs
facture déjà couverts).

Nouveauté de ce fournisseur pour l'acheteur : suite à l'exigence de son
service comptable, chaque facture est désormais accompagnée de son propre
"BON DE COMMANDE" (généré par ce projet) et parfois du DEVIS d'origine —
plusieurs documents de nature différente empilés dans le MÊME fichier PDF,
jamais vu ailleurs dans ce projet avant cette session. Sur de VRAIS PDF
(tests/fixtures/, jamais de texte inventé — règle d'or du projet).
"""

from pathlib import Path

from moteur.ocr import mots_document
from moteur.fournisseurs.electricplus import parse_facture_electricplus_ocr

FIXTURES = Path(__file__).parent / "fixtures"


def _parser(nom_fixture):
    return parse_facture_electricplus_ocr(mots_document(FIXTURES / nom_fixture))


def test_parse_facture_electricplus_1_designation_multiligne():
    """BUG RÉEL CORRIGÉ, le plus important de cette session pour ce
    fournisseur : un article peut être imprimé sur PLUSIEURS lignes
    visuelles OCR — référence + début de désignation sur une ligne,
    quantité/prix SANS suffixe PF/PR sur la ligne suivante, puis un
    complément de désignation (taille, ex. "2,4x180") sur une 3e ligne
    ENCORE APRÈS les prix. 3 des 4 articles de cette facture réelle étaient
    ainsi TOTALEMENT perdus avant le correctif (seul le 4e, resté sur une
    ligne unique, ressortait). Voir _regrouper_articles_electricplus dans
    electricplus.py.

    Ce fichier a aussi 2 pages : la facture, puis NOTRE PROPRE "BON DE
    COMMANDE" — un seul document doit ressortir (la page BC, qui ne porte
    aucune donnée de facture, est exclue AVANT le regroupement par
    identifiant ; sans ce filtre, un nombre accessoire de 6-7 chiffres sur
    la page BC démarrait à tort une 2e "facture" fantôme à 0 ligne)."""

    factures = _parser("facture_electricplus_1_designation_multiligne.pdf")

    assert len(factures) == 1
    f = factures[0]

    assert f.fournisseur == "ELECTRIC PLUS"
    assert f.numero_facture == "4205720"
    assert f.date_facture == "16/07/2026"
    assert f.total_ht_affiche == 88.4

    assert len(f.lignes) == 4
    refs = {l.reference_fournisseur: l for l in f.lignes}

    assert refs["LEG031822"].quantite_facturee == 100.0
    assert refs["LEG031822"].prix_unitaire_ht == 0.05
    assert refs["LEG031822"].montant_ht == 5.0
    assert refs["LEG031822"].designation == "COLLIER COLRING INCOLORE 2,4x180"

    assert refs["LEG031824"].quantite_facturee == 400.0
    assert refs["LEG031824"].montant_ht == 36.0

    assert refs["LEG031827"].quantite_facturee == 500.0
    assert refs["LEG031827"].montant_ht == 45.0

    assert refs["MWK4932430854"].quantite_facturee == 1.0
    assert refs["MWK4932430854"].montant_ht == 2.4

    assert round(sum(l.montant_ht for l in f.lignes), 2) == 88.4


def test_parse_facture_electricplus_2_devis_et_bc_agrafes():
    """BUG RÉEL ÉVITÉ avant toute écriture : ce fichier réel a 3 pages —
    facture, NOTRE PROPRE bon de commande, puis le DEVIS d'origine
    (numéroté "4104132", sans rapport avec le n° de facture "4205769")
    pour le MÊME article. Sans le filtre de page (voir
    _est_page_hors_perimetre_electricplus), la page DEVIS aurait démarré
    un groupe à part et produit une 2e "Facture" fantôme avec la MÊME
    ligne d'article DUPLIQUÉE. Une seule facture doit ressortir."""

    factures = _parser("facture_electricplus_2_devis_et_bc_agrafes.pdf")

    assert len(factures) == 1
    f = factures[0]

    assert f.numero_facture == "4205769"
    assert f.date_facture == "20/07/2026"
    assert f.total_ht_affiche == 24.0

    assert len(f.lignes) == 1
    l0 = f.lignes[0]
    assert l0.reference_fournisseur == "MWK4932430854"
    assert l0.quantite_facturee == 10.0
    assert l0.prix_unitaire_ht == 2.4
    assert l0.montant_ht == 24.0


def test_parse_facture_electricplus_3_commande_cde_et_bc():
    """N° de commande lu via l'en-tête "Référence client:CDE130.035"
    (label déjà connu côté BL, MOTIF_COMMANDE_ELECTRICPLUS) — absent sur
    les 2 autres fixtures de ce fournisseur (agences différentes). Prix
    net imprimé avec le suffixe "PR" (ancre PF/PR déjà connue). Fichier à
    2 pages (facture + notre BON DE COMMANDE) : une seule facture doit
    ressortir."""

    factures = _parser("facture_electricplus_3_commande_cde_et_bc.pdf")

    assert len(factures) == 1
    f = factures[0]

    assert f.numero_facture == "1206609"
    assert f.date_facture == "07/07/2026"
    assert f.numeros_commande == ["130.035"]
    assert f.total_ht_affiche == 380.0

    assert len(f.lignes) == 1
    l0 = f.lignes[0]
    assert l0.reference_fournisseur == "CAB013000T500"
    assert l0.quantite_facturee == 500.0
    assert l0.prix_unitaire_ht == 0.76
    assert l0.montant_ht == 380.0


def test_parse_facture_electricplus_4_pages_miroir_multi_factures():
    """BUG RÉEL CORRIGÉ : sur 3 des 6 pages de ce fichier réel (3 factures,
    chacune + son BC), l'ORDRE DES LIGNES ET DES CELLULES est inversé de
    bout en bout — l'en-tête ressort "MONTANT HT | ... | REFERENCES" (la
    référence en DERNIER) et le pied de tableau ("TOTAL HT"/"CODES TVA")
    apparaît AVANT les lignes d'articles plutôt qu'après. Cohérent avec un
    scan de ce lot précis effectué à l'envers (les 3 factures montrent
    exactement la même inversion), pas une erreur OCR ponctuelle. Voir
    _zone_tableau_electricplus et MOTIF_FACTURE_DATE_ELECTRICPLUS_MIROIR
    dans electricplus.py.

    2 des 3 factures retombent exactement sur leur Total HT. La 3e
    (1206681, une seule ligne à 8,90€) reste à 0 ligne — un fragment
    cellule illisible ("680", ni un nombre ni une référence reconnaissable)
    fait échouer à la fois l'ancre PF et le repli positionnel ; honnêtement
    signalé (aucune ligne), jamais deviné — limite acceptée, une seule
    petite ligne sur 21 vraies factures confrontées cette session."""

    factures = _parser("facture_electricplus_4_pages_miroir_multi_factures.pdf")

    assert len(factures) == 3
    par_facture = {f.numero_facture: f for f in factures}

    f1 = par_facture["1206686"]
    assert f1.date_facture == "15/07/2026"
    assert f1.numeros_commande == ["153.002"]
    assert len(f1.lignes) == 5
    assert round(sum(l.montant_ht for l in f1.lignes), 2) == 460.0

    f2 = par_facture["1206661"]
    assert f2.date_facture == "10/07/2026"
    assert f2.numeros_commande == ["124.034"]
    assert len(f2.lignes) == 12
    assert round(sum(l.montant_ht for l in f2.lignes), 2) == 5600.76

    f3 = par_facture["1206681"]
    assert f3.date_facture == "15/07/2026"
    assert f3.numeros_commande == ["113.068"]
    assert len(f3.lignes) == 0
