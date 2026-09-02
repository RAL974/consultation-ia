"""
Parser facture Cominter Mayotte
(moteur.fournisseurs.cominter_mayotte.parse_facture_cominter_mayotte) —
session F4 suite (2026-09-02), voir CLAUDE.md. Sur de VRAIS PDF
(tests/fixtures/, jamais de texte inventé — règle d'or du projet), extraits
des factures transmises par Prisca LEBLÉ (comptable) pour Cominter Mayotte.

Entité distincte de Cominter Réunion (moteur.fournisseurs.cominter) : même
esprit de structure (ancrage sur le montant) mais DEUX différences réelles
— pas de code TVA après le montant (la référence suit directement), pas de
repère "Signature" avant le tableau d'articles.
"""

from pathlib import Path

from moteur.lecture_pdf import lire_pdf
from moteur.fournisseurs.cominter_mayotte import parse_facture_cominter_mayotte

FIXTURES = Path(__file__).parent / "fixtures"


def _parser(nom_fixture):
    texte = lire_pdf(FIXTURES / nom_fixture)
    return parse_facture_cominter_mayotte(texte)


def test_parse_facture_cominter_mayotte_1_simple():
    """Cas de base : 1 seule ligne, remise 30% systématique."""

    f = _parser("facture_cominter_mayotte_1_simple.pdf")

    assert f.fournisseur == "COMINTER MAYOTTE"
    assert f.numero_facture == "MFAC15584"
    assert f.date_facture == "31/07/2026"
    assert f.numeros_commande == ["24.3155"]
    assert f.numeros_bl == ["MBLC05195"]

    assert len(f.lignes) == 1
    l0 = f.lignes[0]
    assert l0.reference_fournisseur == "L81941"
    assert l0.designation == "Batibox beton Universel"
    assert l0.quantite_facturee == 150.0
    assert l0.montant_ht == 430.50

    assert round(sum(l.montant_ht for l in f.lignes), 2) == 430.50


def test_parse_facture_cominter_mayotte_2_commande_espace():
    """Le n° de commande a un SÉPARATEUR ESPACE dans l'étiquette explicite
    ("- N° de Commande : 24  3109", deux espaces) — normalisé en point
    (re.sub sur tout run d'espaces, pas un simple replace(" ", "."), sinon
    "24..3109" avec un double point)."""

    f = _parser("facture_cominter_mayotte_2_commande_espace.pdf")

    assert f.numero_facture == "MFAC15426"
    assert f.date_facture == "15/07/2026"
    assert f.numeros_commande == ["24.3109"]
    assert f.numeros_bl == ["MBLC05115"]

    assert len(f.lignes) == 1
    l0 = f.lignes[0]
    assert l0.reference_fournisseur == "SY0005565"
    assert l0.quantite_facturee == 24.0
    assert l0.montant_ht == 289.80


def test_parse_facture_cominter_mayotte_3_treize_lignes_note_variante():
    """Document riche (13 lignes) : une note parasite ("VARIANTE DISPO
    GTL") traîne entre le Cdt d'un article (L30008) et la Qté du suivant
    (MER9HKT13) — bornée par le Cdt trouvé, jamais incluse dans la
    désignation. Une référence (L77111L) apparaît deux fois, sur deux
    lignes réellement distinctes (quantités/montants différents) — pas un
    doublon à fusionner."""

    f = _parser("facture_cominter_mayotte_3_treize_lignes_note_variante.pdf")

    assert f.numero_facture == "MFAC15576"
    assert f.numeros_commande == ["155.002"]
    assert f.numeros_bl == ["MBLC05197"]

    assert len(f.lignes) == 13
    l30008 = next(l for l in f.lignes if l.reference_fournisseur == "L30008")
    assert l30008.designation == "Moulure 20x12.5  ss cloison  Barre 2,10 MT"
    assert l30008.montant_ht == 247.98

    refs_l77111l = [l for l in f.lignes if l.reference_fournisseur == "L77111L"]
    assert len(refs_l77111l) == 2
    assert {l.quantite_facturee for l in refs_l77111l} == {14.0, 78.0}

    # Total HT affiché : 6 516,33€.
    assert round(sum(l.montant_ht for l in f.lignes), 2) == 6516.33
