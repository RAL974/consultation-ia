"""
moteur.rapprochement.lecture_facture — tolérance aux pannes (même principe
que tests/test_lecture_bl.py côté BL) : un fichier illisible, d'un
fournisseur non reconnu, ou d'un fournisseur reconnu mais sans parser
facture, ne doit jamais lever d'exception ni bloquer les autres fichiers."""

from conftest import FIXTURES
from moteur.rapprochement.lecture_facture import analyser_dossier, lire_facture


def test_lire_facture_fichier_reel():
    factures, raisons = lire_facture(FIXTURES / "facture_dist109_1_simple.pdf")

    assert raisons == []
    [f] = factures
    assert f.fournisseur == "109 DISTRIBUTION"
    assert f.numero_facture == "360311"
    assert f.fichier == "facture_dist109_1_simple.pdf"


def test_lire_facture_fournisseur_non_reconnu(tmp_path):
    chemin = tmp_path / "inconnu.pdf"
    import fitz

    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((72, 72), "Ceci n'est le devis/facture d'aucun fournisseur connu.")
    doc.save(chemin)
    doc.close()

    factures, raisons = lire_facture(chemin)

    assert factures == []
    assert raisons == ["Fournisseur non reconnu"]


def test_lire_facture_pdf_illisible(tmp_path):
    chemin = tmp_path / "corrompu.pdf"
    chemin.write_bytes(b"pas un vrai PDF")

    factures, raisons = lire_facture(chemin)

    assert factures == []
    assert len(raisons) == 1
    assert "illisible" in raisons[0]


def test_analyser_dossier_absent(tmp_path):
    factures, anomalies = analyser_dossier(tmp_path / "n_existe_pas")

    assert factures == []
    assert anomalies == []


def test_analyser_dossier_reel(tmp_path):
    import shutil

    shutil.copy2(FIXTURES / "facture_dist109_1_simple.pdf", tmp_path / "facture_dist109_1_simple.pdf")
    shutil.copy2(FIXTURES / "facture_dist109_3_multi_bl_meme_commande.pdf", tmp_path / "autre.pdf")

    factures, anomalies = analyser_dossier(tmp_path)

    assert anomalies == []
    assert len(factures) == 2
    assert {f.numero_facture for f in factures} == {"360311", "360366"}
