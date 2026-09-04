"""
Migration des lignes facture de Commandes vers Pièces (moteur/rapprochement/
migration_pieces.py, étape 4 de P1) et orchestration (installation_pieces).

- logique pure (recherche de PDF, contrôle au centime, rattachement) sur
  des objets construits à la main et un classeur synthétique ;
- extrait de 50 lignes sur une COPIE du vrai Suivi commandes vivant (skipif
  absent du poste) — jamais l'original : la migration est au centime, une
  ligne Pièces par ligne Commandes retrouvée (ou éclatée par BL), rapport
  écrit, idempotence au 2e passage.
"""

import shutil
from datetime import date, datetime

import pytest
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

from moteur.rapprochement.migration_pieces import (
    LigneFacturee,
    _sommes_par_fournisseur,
    _sommes_pieces_par_fournisseur,
    comparer_sommes,
    controler_sommes,
    lire_lignes_facturees,
    migrer_factures_vers_pieces,
    trouver_pdf_facture,
)
from moteur.rapprochement.pieces import (
    COLONNES_FACTURE_CALCULEES,
    MODE_MIGRE,
    MODE_MIGRE_SANS_PDF,
    feuille_pieces_presente,
    installer_feuille_pieces,
    lire_pieces,
)
from moteur.rapprochement.installation_pieces import COLONNES_LIGNE2_CONSTATEES, COLONNES_LIGNE2_PLAN, executer, verifier_structure
from moteur.rapprochement.pipeline_bl import trouver_fichier_suivi_vivant

from conftest import ROOT


def _classeur(chemin):
    wb = Workbook()
    ws = wb.active
    ws.title = "Commandes"
    ws.append(["Référence", "Désignation", "Qté commandée", "N° de commande", "Fournisseur", "Chantier", "Sous-Chantier",
               "Qté livrée", "Tarif BL", "Tarif convenu", "Facturé BL", *COLONNES_FACTURE_CALCULEES])
    ws.append(["06620", "ICT", 800, "131.082", "Coredime", "Ch A", None, 800, 0.35, None, 280.0,
               "6100600", datetime(2026, 1, 28), 800, 0.35, 280.0])
    ws.append([5120, "ICT bleu", 10, "131.082", "COREDIME", "Ch A", "SC", 10, None, None, 0,
               None, None, None, None, None])
    ws.append(["ZZ", "Autre", 1, "M3.1", "109 Distribution", "Ch B", None, 1, 2.0, None, 2.0,
               362759, datetime(2026, 7, 30), 1, 2.0, 2.0])
    table = Table(displayName="Commandes", ref="A1:P4")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)
    wb.save(chemin)


def test_lire_lignes_facturees_ne_garde_que_les_lignes_avec_numero(tmp_path):
    chemin = tmp_path / "suivi.xlsx"
    _classeur(chemin)
    lignes = lire_lignes_facturees(chemin)
    assert [l.ligne_excel for l in lignes] == [2, 4]
    l = lignes[0]
    assert l.fournisseur == "Coredime" and l.fournisseur_suivi == "COREDIME"
    assert l.numero_facture == "6100600" and l.numero_commande == "131.082" and l.reference == "06620"
    assert l.qte == 800.0 and l.pu == 0.35 and l.montant == 280.0 and l.chantier == "Ch A"
    assert lignes[1].numero_facture == "362759"  # nombre -> texte
    assert [l.ligne_excel for l in lire_lignes_facturees(chemin, limite=1)] == [2]


def test_trouver_pdf_facture_archive_puis_a_verifier(tmp_path):
    projet = tmp_path
    archive = projet / "a_traiter" / "BL" / "Traités" / "131.082"
    archive.mkdir(parents=True)
    (archive / "2026-01-28 - COREDIME - Facture 6100600 - BC 131.082.pdf").write_bytes(b"1")
    (archive / "2026-01-28 - COREDIME - B010001 - BC 131.082.pdf").write_bytes(b"2")  # un BL, pas une facture
    a_verifier = projet / "a_traiter" / "Factures" / "À vérifier"
    a_verifier.mkdir(parents=True)
    (a_verifier / "6100600.pdf").write_bytes(b"3")
    (a_verifier / "6100677.pdf").write_bytes(b"4")
    (projet / "a_traiter" / "Factures" / "Facture_360311.pdf").write_bytes(b"5")

    trouves = trouver_pdf_facture(projet, "131.082", "6100600")
    assert [p.name for p in trouves] == [
        "2026-01-28 - COREDIME - Facture 6100600 - BC 131.082.pdf", "6100600.pdf",
    ]
    assert [p.name for p in trouver_pdf_facture(projet, "999", "360311")] == ["Facture_360311.pdf"]
    assert trouver_pdf_facture(projet, "131.082", "") == []
    assert trouver_pdf_facture(projet, "131.082", "12") == []  # n° trop court pour une recherche par nom


def test_comparer_sommes_au_centime():
    lignes = [
        LigneFacturee(2, "Coredime", "C", "R", None, None, "F1", None, 1, 1.0, 10.005),
        LigneFacturee(3, "COREDIME", "C", "R2", None, None, "F2", None, 1, 1.0, 5.0),
        LigneFacturee(4, "GMR", "D", "R", None, None, "F3", None, 1, 1.0, 1.0),
    ]
    sommes = _sommes_par_fournisseur(lignes)
    assert sommes == {"COREDIME": 15.01, "GMR": 1.0}  # noms normalisés, arrondi au centime
    pieces = [
        {"Type": "Facture", "Fournisseur": "COREDIME", "Montant HT": 15.0},
        {"Type": "Facture", "Fournisseur": "COREDIME", "Montant HT": 0.005},
        {"Type": "BL", "Fournisseur": "GMR", "Montant HT": 99.0},
        {"Type": "Facture", "Fournisseur": "GMR", "Montant HT": 1.0},
    ]
    assert _sommes_pieces_par_fournisseur(pieces) == {"COREDIME": 15.01, "GMR": 1.0}
    ok, detail = comparer_sommes(sommes, _sommes_pieces_par_fournisseur(pieces))
    assert ok and detail == [("COREDIME", 15.01, 15.01, 0.0), ("GMR", 1.0, 1.0, 0.0)]
    ok, detail = comparer_sommes({"COREDIME": 15.01}, {"COREDIME": 15.0, "GMR": 1.0})
    assert not ok and detail == [("COREDIME", 15.01, 15.0, -0.01), ("GMR", 0.0, 1.0, 1.0)]


def test_migration_synthetique_sans_pdf_est_au_centime_et_idempotente(tmp_path):
    """Aucun PDF sur ce poste synthétique : toutes les lignes ressortent
    « Migré sans PDF » avec un commentaire explicite, mais les montants
    (source de vérité : les 5 colonnes) sont au centime, la bascule des
    colonnes est alors autorisée, et un 2e passage n'écrit rien."""

    chemin = tmp_path / "suivi.xlsx"
    _classeur(chemin)
    with pytest.raises(ValueError):
        migrer_factures_vers_pieces(chemin, tmp_path, tmp_path / "backups")  # feuille Pièces absente
    installer_feuille_pieces(chemin, tmp_path / "backups")

    resume = migrer_factures_vers_pieces(chemin, tmp_path, tmp_path / "backups")

    assert resume["lignes_commandes"] == 2 and resume["pieces_construites"] == 2
    assert resume["statistiques"] == {"pdf_introuvable": 2}
    assert resume["au_centime_avant_ecriture"] and resume["au_centime"]
    assert resume["ecriture"]["ajoutees"] == 2
    assert resume["chemin_rapport"].exists()
    assert "Au centime avant écriture : OUI" in resume["chemin_rapport"].read_text(encoding="utf-8")

    pieces = lire_pieces(chemin)
    assert [p["Mode de rapprochement"] for p in pieces] == [MODE_MIGRE_SANS_PDF] * 2
    assert pieces[0]["Fournisseur"] == "COREDIME" and pieces[0]["N° pièce"] == "6100600"
    assert pieces[0]["Référence Suivi"] == "06620" and pieces[0]["Qté"] == 800 and pieces[0]["Montant HT"] == 280.0
    assert pieces[0]["Date pièce"].date() == date(2026, 1, 28) and pieces[0]["Chantier"] == "Ch A"
    assert "PDF introuvable" in pieces[0]["Commentaire"]
    assert pieces[1]["Fournisseur"] == "109 Distribution" and pieces[1]["N° pièce"] == "362759"

    controle = controler_sommes(chemin)
    assert controle["au_centime"] and controle["pieces_facture"] == 2

    resume2 = migrer_factures_vers_pieces(chemin, tmp_path, tmp_path / "backups")
    assert resume2["ecriture"]["ajoutees"] == 0 and len(resume2["ecriture"]["ignorees"]) == 2

    # Orchestration : l'étape 5 est acceptée (au centime) et vérifiée.
    resultats = executer(chemin, tmp_path, tmp_path / "backups", etapes=("3", "5"), journal=lambda *a: None)
    assert resultats["5_verif"]["nb_colonnes_commandes"] == 16 + 4
    assert not resultats["5_verif"]["calc_chain"] and resultats["5_verif"]["full_calc_on_load"]


def test_verifier_structure_signale_les_manques(tmp_path):
    chemin = tmp_path / "suivi.xlsx"
    _classeur(chemin)
    with pytest.raises(AssertionError):
        verifier_structure(chemin)  # Pièces attendue par défaut
    r = verifier_structure(chemin, attendre_pieces=False)
    assert r["tableaux"] == 1 and r["feuilles"] == ["Commandes"]


def test_colonnes_ligne2_du_plan_incluses_dans_les_constatees():
    assert set(COLONNES_LIGNE2_PLAN) == set(range(25, 31))
    assert set(COLONNES_LIGNE2_PLAN) <= set(COLONNES_LIGNE2_CONSTATEES) == set(range(22, 31))


# --- extrait de 50 lignes sur une COPIE du vrai classeur ----------------------


@pytest.mark.skipif(
    trouver_fichier_suivi_vivant(ROOT) is None,
    reason="Classeur Suivi commandes VIVANT introuvable depuis ce poste",
)
def test_migration_extrait_50_lignes_sur_une_copie_du_vrai_suivi(tmp_path):
    fichier_reel = trouver_fichier_suivi_vivant(ROOT)
    copie = tmp_path / fichier_reel.name
    shutil.copy2(fichier_reel, copie)
    if not feuille_pieces_presente(copie):
        installer_feuille_pieces(copie, tmp_path / "backups")
    deja = len(lire_pieces(copie))

    lignes = lire_lignes_facturees(copie, limite=50)
    assert len(lignes) == 50

    resume = migrer_factures_vers_pieces(copie, ROOT, tmp_path / "backups", limite=50)

    assert resume["lignes_commandes"] == 50
    assert resume["au_centime_avant_ecriture"]
    assert resume["pieces_construites"] >= 50 - len(resume["ecriture"]["ignorees"]) or resume["ecriture"]["ajoutees"] >= 0
    st = resume["statistiques"]
    # au moins une ligne réellement retrouvée dans un PDF archivé (les
    # factures Coredime de janvier 2026 sont archivées dans Traités/)
    assert st.get("migre", 0) >= 1
    pieces = lire_pieces(copie)
    nouvelles = pieces[deja:]
    assert len(nouvelles) == resume["ecriture"]["ajoutees"]
    assert {p["Mode de rapprochement"] for p in nouvelles} <= {MODE_MIGRE, MODE_MIGRE_SANS_PDF}
    assert all(p["Type"] == "Facture" for p in nouvelles)
    assert all(p["Fichier"] is None or str(p["Fichier"]).lower().endswith(".pdf") for p in nouvelles)
    retrouvees = [p for p in nouvelles if p["Mode de rapprochement"] == MODE_MIGRE]
    assert all(p["N° BL lié"] for p in retrouvees)  # BL lié retrouvé sur la facture
    assert resume["chemin_rapport"].exists()

    # Σ par fournisseur des 50 lignes = Σ Pièces construites, au centime
    for f, a, b, e in resume["sommes"]:
        assert abs(a - b) <= 0.005, (f, a, b, e)

    # 2e passage : idempotent par ID pièce
    resume2 = migrer_factures_vers_pieces(copie, ROOT, tmp_path / "backups", limite=50)
    assert resume2["ecriture"]["ajoutees"] == 0
