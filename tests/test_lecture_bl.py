"""
Détection de fournisseur PAR PAGE (moteur/rapprochement/lecture_bl.py) —
cas réel rencontré 2 fois en session (scan groupé mélangeant plusieururs
fournisseurs dans un même fichier, jusqu'ici traité à la main, voir
CLAUDE.md). Fixtures construites en COMBINANT de vraies pages de BL déjà
verrouillées ailleurs (tests/fixtures/), jamais de contenu inventé — les
pages elles-mêmes sont 100% réelles, seul leur assemblage dans un même
fichier est synthétique (reproduit fidèlement le geste de l'acheteur qui
scanne plusieurs BL papier à la suite)."""

import fitz

from conftest import FIXTURES
from moteur.rapprochement.lecture_bl import lire_bl, _parser_groupe_fournisseur


def _combiner_pages(tmp_path, *fixtures, nom="combine.pdf"):
    combine = fitz.open()
    for nom_fixture in fixtures:
        with fitz.open(FIXTURES / nom_fixture) as doc:
            combine.insert_pdf(doc)
    chemin = tmp_path / nom
    combine.save(chemin)
    combine.close()
    return chemin


def test_lire_bl_fichier_mono_fournisseur_comportement_inchange(tmp_path):
    """Un fichier à UN SEUL fournisseur (cas de loin le plus courant) doit
    se comporter exactement comme avant : un seul BonLivraison, aucune
    anomalie."""

    chemin = _combiner_pages(tmp_path, "bl_dist109_1.pdf")

    bons, raisons = lire_bl(chemin)

    assert raisons == []
    [bl] = bons
    assert bl.fournisseur == "109 DISTRIBUTION"
    assert bl.numero_commande == "123.096"
    assert bl.numero_bl == "735136"


def test_lire_bl_fichier_multi_fournisseur_page_par_page(tmp_path):
    """Fichier à 4 pages mélangeant DEUX fournisseurs, chacun apparaissant
    sur des pages NON CONTIGUËS (COREDIME p0 et p2, 109 Distribution p1 et
    p3) — reproduit fidèlement le cas réel rencontré 2 fois en session
    (scan groupé). COREDIME ne sait pas répartir lui-même plusieurs BL sur
    plusieurs pages (contrairement à 109 Distribution, qui le fait via son
    propre numéro de BL) : les 2 pages COREDIME, bien que du même
    fournisseur, portent des commandes RÉELLEMENT différentes et ne
    doivent JAMAIS être fusionnées sous un seul BonLivraison — c'est
    exactement le bug dangereux rencontré en session (voir CLAUDE.md,
    doc07205620260824145119.pdf)."""

    chemin = _combiner_pages(
        tmp_path,
        "bl_coredime_1.pdf",   # p0 : COREDIME, commande 123.097, BL CORB032399
        "bl_dist109_1.pdf",    # p1 : 109 DISTRIBUTION, commande 123.096, BL 735136
        "bl_coredime_3.pdf",   # p2 : COREDIME, commande 131.153, BL CORB032442
        "bl_dist109_2.pdf",    # p3 : 109 DISTRIBUTION, commande 131.156, BL 736366
    )

    bons, raisons = lire_bl(chemin)

    assert raisons == []
    assert len(bons) == 4

    par_page = {bl.pages[0]: bl for bl in bons if bl.pages and len(bl.pages) == 1}
    assert set(par_page) == {0, 1, 2, 3}

    assert par_page[0].fournisseur == "COREDIME"
    assert par_page[0].numero_commande == "123.097"
    assert par_page[0].numero_bl == "CORB032399"

    assert par_page[1].fournisseur == "109 DISTRIBUTION"
    assert par_page[1].numero_commande == "123.096"
    assert par_page[1].numero_bl == "735136"

    # Les 2 pages COREDIME n'ont PAS été fusionnées : chacune garde sa
    # propre commande, distincte de l'autre.
    assert par_page[2].fournisseur == "COREDIME"
    assert par_page[2].numero_commande == "131.153"
    assert par_page[2].numero_bl == "CORB032442"
    assert par_page[2].numero_commande != par_page[0].numero_commande

    assert par_page[3].fournisseur == "109 DISTRIBUTION"
    assert par_page[3].numero_commande == "131.156"
    assert par_page[3].numero_bl == "736366"

    for bl in bons:
        assert bl.fichier == chemin.name


def test_parser_groupe_fournisseur_sans_parser_bl_produit_une_anomalie():
    """"RAVATE PRO" est reconnu par le détecteur de fournisseur mais n'a
    PAS de parser BL dédié (voir moteur/fournisseurs/ravate.py,
    FOURNISSEURS = ["RAVATE"] seulement — limitation réelle documentée
    dans CLAUDE.md). Une page qui lui est attribuée doit produire une
    anomalie claire, jamais une exception ni une perte silencieuse."""

    bons, anomalies = _parser_groupe_fournisseur("RAVATE PRO", [0], [[]])

    assert bons == []
    assert len(anomalies) == 1
    assert "RAVATE PRO" in anomalies[0]
    assert "pas encore de parser" in anomalies[0]
