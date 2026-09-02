"""
Parser facture Cominter (moteur.fournisseurs.cominter.parse_facture_cominter)
— session F4 suite (2026-09-02), voir CLAUDE.md. Sur de VRAIS PDF
(tests/fixtures/, jamais de texte inventé — règle d'or du projet), extraits
des 132 pièces transmises par Prisca LEBLÉ (comptable) pour Cominter,
juillet-août 2026. Couvre uniquement "COMINTER" (Comptoir Ouest/
Sainte-Clotinde/Saint-Pierre) — "COMINTER MAYOTTE" (format MFACxxxxx) reste
à construire séparément, voir bandeau GABARIT FACTURE du module.
"""

from pathlib import Path

from moteur.lecture_pdf import lire_pdf
from moteur.fournisseurs.cominter import parse_facture_cominter

FIXTURES = Path(__file__).parent / "fixtures"


def _parser(nom_fixture):
    texte = lire_pdf(FIXTURES / nom_fixture)
    return parse_facture_cominter(texte)


def test_parse_facture_cominter_1_derniere_ligne_note_apres_reference():
    """Bug réel corrigé : sur la DERNIÈRE ligne d'articles du document, une
    note de livraison ("Livraison chantier Anzemberg rue Henri Cornu
    CAMBAIE CG CG") traîne APRÈS la vraie référence (CAETM4288), qui
    n'est donc PAS zone[-1] — la référence est recherchée en partant de
    la FIN de la zone jusqu'à trouver un token qui ressemble vraiment à
    une référence (lettres+chiffres), jamais juste "la dernière ligne"."""

    f = _parser("facture_cominter_1_derniere_ligne_note_apres_reference.pdf")

    assert f.fournisseur == "COMINTER"
    assert f.numero_facture == "OFC194316"
    assert f.date_facture == "15/08/2026"
    assert f.numeros_commande == ["135.039"]
    assert f.numeros_bl == ["OBL108110"]

    assert len(f.lignes) == 9
    derniere = f.lignes[-1]
    assert derniere.reference_fournisseur == "CAETM4288"
    assert derniere.designation == "BAIE TECHNIC1000 800X800 42U CAPACITE DE CHARGE 1000 KG"
    assert derniere.quantite_facturee == 1.0
    assert derniere.montant_ht == 901.7

    premiere = f.lignes[0]
    assert premiere.reference_fournisseur == "CAELK2766"
    assert premiere.quantite_facturee == 1.0
    assert premiere.prix_unitaire_ht == 435.94
    assert premiere.montant_ht == 435.94

    # Total HT affiché sur le document (relevé à la main) : 2 040,30€.
    assert round(sum(l.montant_ht for l in f.lignes), 2) == 2040.30


def test_parse_facture_cominter_2_remise_optionnelle_et_ecotaxe_sans_unite():
    """Document riche (19 lignes) : la remise (%) est OPTIONNELLE ligne par
    ligne (présente sur les 2 premiers disjoncteurs, absente ensuite) et
    les lignes d'éco-participation ("ECO-TAXEx") n'ont PAS de champ
    Cdt/Unité du tout (contrairement à un article normal) — un ancrage sur
    Cdt (comme les devis du même fournisseur) les aurait perdues ; l'ancrage
    sur le MONTANT, présent sur CHAQUE ligne, les retrouve toutes."""

    f = _parser("facture_cominter_2_remise_optionnelle_et_ecotaxe_sans_unite.pdf")

    assert f.numero_facture == "OFC193413"
    assert f.date_facture == "31/07/2026"
    assert f.numeros_commande == ["M2.22.082"]
    assert f.numeros_bl == ["OBL107786"]

    assert len(f.lignes) == 19
    refs = {l.reference_fournisseur: l for l in f.lignes}

    # Disjoncteurs : "4.5 KA" dans la désignation (capacité de coupure,
    # déjà connue côté BL du même fournisseur comme source de confusion —
    # ici simplement du texte, jamais pris pour une cellule numérique).
    assert refs["L406773"].quantite_facturee == 20.0
    assert refs["L406773"].montant_ht == 125.40
    assert refs["L406773"].designation == "Disjoncteur DNX3 1P+NG 10A   4.5 KA"

    # Éco-taxes : pas de Cdt/Unité, référence directement après la désignation.
    assert refs["ECO-TAXE3"].quantite_facturee == 15.0
    assert refs["ECO-TAXE3"].montant_ht == 1.80
    assert refs["ECO-TAXE3"].designation == "Eco - contribution 0.12€"
    assert refs["ECO-TAXE21/0"].montant_ht == 9.60

    # Câbles vendus "Cour" (couronne) : Cdt reconnu, désignation propre.
    assert refs["C002B"].designation == "Cable rigide HO7VU 2.5 BLEU C100M"
    assert refs["C002B"].montant_ht == 134.12

    # Total HT affiché : 1 767,36€ (repris aussi dans le BC annexé au même PDF).
    assert round(sum(l.montant_ht for l in f.lignes), 2) == 1767.36


def test_parse_facture_cominter_3_code_tva_colle_montant_et_commande_espace():
    """2 bugs réels corrigés sur ce document : (1) le code TVA est COLLÉ à
    la ligne montant ("17,33 € 1", pas sur sa propre ligne comme les 2
    autres fixtures) — sans repli, tout le document ressortait à 0 ligne ;
    (2) le n° de commande est réimprimé dans le bloc Signature avec un
    SÉPARATEUR ESPACE ("BC N°24 3240"), pas point/tiret — l'en-tête, lui,
    le tronque ("BC: 3240", le "24" perdu), donc SEUL le repli élargi sur
    le bloc Signature retrouve la valeur complète."""

    f = _parser("facture_cominter_3_code_tva_colle_montant_et_commande_espace.pdf")

    assert f.numero_facture == "NFA018127"
    assert f.date_facture == "04/08/2026"
    # Hypothèse à valider en recette réelle (pas de commande "24.3240"
    # encore confrontée au vrai Suivi à ce jour) : le "24" pourrait être un
    # préfixe année/type comme les BdC manuels déjà documentés (Coredime,
    # Cominter Mayotte), à confirmer plutôt qu'à corriger sans preuve.
    assert f.numeros_commande == ["24.3240"]
    assert f.numeros_bl == ["NBL016409"]

    assert len(f.lignes) == 2
    l0, l1 = f.lignes
    assert l0.reference_fournisseur == "AGI437128"
    assert l0.quantite_facturee == 9.0
    assert l0.montant_ht == 17.33
    assert l1.reference_fournisseur == "AGI437129"
    assert l1.quantite_facturee == 12.0
    assert l1.montant_ht == 23.10

    # Total HT affiché : 40,43€.
    assert round(sum(l.montant_ht for l in f.lignes), 2) == 40.43


def test_parse_facture_cominter_4_sans_repere_signature_unite_minuscule():
    """2 bugs réels corrigés sur ce document (agence Saint-Pierre) : (1)
    AUCUN repère "Signature" avant le bloc [date, n° de BL, n° de
    commande] — contrairement aux 3 autres fixtures — repli sur la
    1ère ligne qui ressemble à un n° de BL ; (2) le Cdt/unité est imprimé
    tout en minuscule ("unite", pas "Unité") — comparaison casse-insensible
    désormais. Vérifie aussi que la référence de la DERNIÈRE ligne
    (ECO-TAXE21/0) n'est pas polluée par la note "BC N°24 1581 DU
    07/07/26" qui la suit avant "Article 7" (même garde-fou que la
    fixture 1, ordre différent ici : désignation APRÈS la référence)."""

    f = _parser("facture_cominter_4_sans_repere_signature_unite_minuscule.pdf")

    assert f.numero_facture == "NF155008"
    assert f.date_facture == "15/07/2026"
    assert f.numeros_commande == ["24.1581"]
    assert f.numeros_bl == ["NL132081"]

    assert len(f.lignes) == 5
    refs = {l.reference_fournisseur: l for l in f.lignes}

    assert refs["L69731L"].quantite_facturee == 10.0
    assert refs["L69731L"].montant_ht == 58.30
    assert refs["L69731L"].designation == "Prise 2P+T saillie   Plexo gris"

    derniere = f.lignes[-1]
    assert derniere.reference_fournisseur == "ECO-TAXE21/0"
    assert derniere.designation == "Eco - contrbution 0.24€"
    assert derniere.montant_ht == 2.40

    # Total HT affiché : 178,10€.
    assert round(sum(l.montant_ht for l in f.lignes), 2) == 178.10
