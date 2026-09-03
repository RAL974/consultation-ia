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
    lignes restent NON extraites, l'écart avec le Total HT affiché est
    signalé (jamais un faux total exact).

    Compte de lignes RELEVÉ (12 -> 21) après le correctif fin_zone (voir
    CLAUDE.md, gros lot Coredime H1 2026) : cette pièce s'étale elle aussi
    sur plusieurs folios, la même troncature prématurée qui a motivé ce
    correctif l'affectait déjà — les 9 lignes retrouvées en plus sont
    propres (sans remise multiple), seule la limite remise-multiple
    d'origine reste (124,37€ résiduels, inchangé dans sa nature)."""

    f = _parser("facture_coredime_8_remises_multiples_partiel.pdf")

    assert f.total_ht_affiche == 1132.51
    assert len(f.lignes) == 21

    total_extrait = round(sum(l.montant_ht for l in f.lignes), 2)
    assert total_extrait == 1008.14
    assert total_extrait != f.total_ht_affiche  # écart honnête, pas un faux total exact

    # Les lignes propres (sans double remise) restent exactes.
    refs = {l.reference_fournisseur: l for l in f.lignes}
    assert refs["LEG030015"].montant_ht == 184.0
    assert refs["GFO026705"].montant_ht == 148.8


def test_parse_facture_coredime_9_multi_folio_zone_tronquee():
    """BUG RÉEL CORRIGÉ (gros lot Coredime H1 2026, ~48/328 factures
    touchées, jusqu'à 85% du montant manquant sur les pires cas) : une
    facture à beaucoup de lignes s'étale sur PLUSIEURS folios (pages) DE
    CONTENU, chacun avec son propre repère "##ESIGUID" répété en en-tête —
    l'ancienne borne de zone ("2e occurrence de ##ESIGUID" = fin de la
    facture) prenait à tort le 2e FOLIO pour notre propre bon de commande
    annexe, tronquant tout le reste. Sur cette pièce réelle (24 lignes sur
    2 folios), seules 11 étaient extraites avant le correctif (2984,98€
    manquants sur 8438,47€) ; le nouveau repère ("BON DE COMMANDE", notre
    propre annexe, en-tête isolé sur sa ligne) retombe désormais sur 22.

    Petit résidu honnête restant (18,60€) : la référence HAGMJT702 est
    imprimée DEUX FOIS sur ce document réel (une fois par folio, prix très
    légèrement différent — 12,2340 vs 12,2300 — montant identique 12,23€
    arrondi) — un seul exemple à ce jour, pas de règle de déduplication
    inventée (pourrait être une vraie 2e ligne de commande) ; l'autocontrôle
    le signale honnêtement plutôt que de deviner."""

    f = _parser("facture_coredime_9_multi_folio_zone_tronquee.pdf")

    assert f.numero_facture == "6105181"
    assert f.date_facture == "20/05/2026"
    assert f.numeros_bl == ["B022876"]
    assert f.total_ht_affiche == 8438.47

    assert len(f.lignes) == 22
    total_extrait = round(sum(l.montant_ht for l in f.lignes), 2)
    assert total_extrait == 8457.07

    refs = [l.reference_fournisseur for l in f.lignes]
    assert refs.count("HAGMJT702") == 2  # voir docstring — résidu honnête, pas dédupliqué

    # Lignes du 2e folio, invisibles avant le correctif — preuve directe
    # que la zone s'étend maintenant au-delà du 1er folio.
    refs_set = set(refs)
    assert "HAGXVL122STI" in refs_set
    assert "HAGHMC499" in refs_set
    assert "LEG401333" in refs_set


def test_parse_facture_coredime_10_bon_livraison_annexe_duplique():
    """2e BUG RÉEL CORRIGÉ, trouvé juste après le précédent sur le même
    gros lot : Coredime peut annexer son PROPRE "BON DE LIVRAISON" à la
    suite de la facture, avec un tableau d'articles dans EXACTEMENT le
    même format qu'une vraie ligne de facture — sur cette pièce réelle,
    les 4 lignes de la facture étaient DOUBLÉES (240,00€ extraits au lieu
    de 120,00€), la zone (faute de "BON DE COMMANDE" sur cette pièce)
    s'étendant "jusqu'à la fin du texte" et absorbant le 2e tableau.

    1er essai abandonné : chercher le titre "B O N  D E  L I V R A I S O N"
    (lettres espacées, comme "F A C T U R E") — il s'imprime en PIED de
    son propre bloc, donc APRÈS le tableau à exclure, trop tard pour
    borner quoi que ce soit. Repère retenu : "COR B<num>", la référence
    isolée en TÊTE de ce même bloc annexe (même famille que "COR F<num>"
    en tête de la vraie facture — seule la lettre change, B comme "Bon de
    livraison")."""

    f = _parser("facture_coredime_10_bon_livraison_annexe_duplique.pdf")

    assert f.numero_facture == "6200396"
    assert f.date_facture == "28/01/2026"
    assert f.numeros_bl == ["B010202"]
    assert f.total_ht_affiche == 120.0

    assert len(f.lignes) == 4  # jamais 8 (le tableau dupliqué de l'annexe BL)
    total_extrait = round(sum(l.montant_ht for l in f.lignes), 2)
    assert total_extrait == 120.0

    refs = {l.reference_fournisseur: l for l in f.lignes}
    assert refs["GFO026504"].montant_ht == 36.7
    assert refs["WAG221-2411"].montant_ht == 33.87


def test_parse_facture_coredime_11_telephone_confondu_avec_ligne_incomplete():
    """3e BUG RÉEL CORRIGÉ, trouvé APRÈS le correctif "BON DE COMMANDE"
    lui-même (même gros lot) : ce titre s'imprime en PIED de son propre
    bloc annexe, APRÈS le bloc signature "DATE/ACHETEUR/VISA/téléphone" —
    notre propre numéro ("0693 86 68 03") restait alors DANS la zone
    scannée et matchait accidentellement MOTIF_LIGNE_INCOMPLETE_COREDIME
    (une 2e "ligne incomplète" à côté de la vraie, LEG031955) —
    désamorçant l'appariement 1:1 (remise double) pourtant sans ambiguïté
    sur cette pièce : 0 ligne extraite pour une facture d'1 seule ligne
    réelle (29,16€). Repère retenu : "DESTINATAIRE", la toute PREMIÈRE
    ligne de notre BC, bien avant le tableau et la signature."""

    f = _parser("facture_coredime_11_telephone_confondu_avec_ligne_incomplete.pdf")

    assert f.numero_facture == "6401314"
    assert f.date_facture == "29/06/2026"
    assert f.numeros_commande == ["129.026"]
    assert f.numeros_bl == ["B027782.1"]
    assert f.total_ht_affiche == 29.16

    assert len(f.lignes) == 1
    l0 = f.lignes[0]
    assert l0.reference_fournisseur == "LEG031955"
    assert l0.quantite_facturee == 300.0
    assert l0.prix_unitaire_ht == 0.0972
    assert l0.montant_ht == 29.16
