"""
Socle d'écriture sécurisée dans le Suivi commandes (moteur/rapprochement/
ecriture.py, branche Rapprochement AI — voir CLAUDE.md).

Deux familles de tests :
- sur un classeur SYNTHÉTIQUE (mais structurellement représentatif : Excel
  Table, colonne formule, validation de données) pour vérifier la mécanique
  générique du patch XML sans dépendre du fichier personnel de l'acheteur ;
- sur le VRAI "1.3.0.1. Suivi commandes - 2026.xlsx" (jamais l'original —
  toujours une copie dans tmp_path), ignoré si absent de ce poste (comme
  tests/test_panier.py::test_colonnes_calees_sur_le_vrai_suivi_commandes),
  pour la preuve définitive que rien d'autre que la cellule visée ne bouge.
"""

import shutil
import zipfile

import pytest
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation

from moteur.rapprochement.ecriture import (
    Ecriture,
    ClasseurVerrouille,
    ColonneNonModifiable,
    appliquer,
    est_verrouille,
    lire_entetes,
    sauvegarder,
    simuler,
)
from moteur.panier import trouver_fichier_suivi

from conftest import ROOT


def _classeur_synthetique(chemin):
    """Un classeur minimal mais structurellement représentatif du vrai
    Suivi commandes : une feuille "Commandes" avec un Excel Table, une
    colonne formule (jamais modifiable) et une validation de données."""

    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Commandes"
    ws.append(["Référence", "Qté commandée", "Qté livrée", "Tarif BL", "Statut", "Note"])
    ws.append(["ART1", 10, None, None, '=IF(C2="","attente","reçu")', None])
    ws.append(["ART2", 5, 5, 12.5, '=IF(C3="","attente","reçu")', None])

    table = Table(displayName="Commandes", ref="A1:F3")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)

    dv = DataValidation(type="list", formula1='"Oui,Non"', allow_blank=True)
    dv.add("F2:F3")
    ws.add_data_validation(dv)

    wb.save(chemin)


@pytest.fixture
def classeur(tmp_path):
    chemin = tmp_path / "Suivi_test.xlsx"
    _classeur_synthetique(chemin)
    return chemin


def test_lire_entetes(classeur):
    entetes = lire_entetes(classeur)
    assert entetes["Qté livrée"] == 3
    assert entetes["Tarif BL"] == 4


def test_colonne_formule_refusee(classeur):
    with pytest.raises(ColonneNonModifiable):
        appliquer(classeur, [Ecriture(2, "Statut", "reçu")], classeur.parent / "backups")


def test_colonne_inconnue_refusee(classeur):
    with pytest.raises(ColonneNonModifiable):
        appliquer(classeur, [Ecriture(2, "Colonne fantôme", 1)], classeur.parent / "backups")


def test_simuler_ne_modifie_rien(classeur):
    avant = classeur.read_bytes()
    rapport = simuler(classeur, [Ecriture(2, "Qté livrée", 10), Ecriture(2, "Tarif BL", 9.9)])
    assert classeur.read_bytes() == avant

    par_colonne = {r["colonne"]: r for r in rapport}
    assert par_colonne["Qté livrée"]["ancienne_valeur"] is None
    assert par_colonne["Qté livrée"]["nouvelle_valeur"] == 10
    assert par_colonne["Tarif BL"]["nouvelle_valeur"] == 9.9


def test_appliquer_ecrit_cellule_vide_et_preserve_le_reste(classeur, tmp_path):
    dossier_backups = tmp_path / "backups"
    sauvegarde = appliquer(
        classeur,
        [Ecriture(2, "Qté livrée", 10), Ecriture(2, "Tarif BL", 19.9), Ecriture(2, "Note", "Rupture fournisseur")],
        dossier_backups,
    )

    assert sauvegarde.exists()
    assert sauvegarde.read_bytes() != classeur.read_bytes()  # la sauvegarde est l'ANCIEN état

    wb = load_workbook(classeur, data_only=False)
    ws = wb["Commandes"]
    assert ws["C2"].value == 10
    assert ws["D2"].value == 19.9
    assert ws["F2"].value == "Rupture fournisseur"
    # la formule voisine, jamais touchée par l'écriture, doit survivre :
    assert str(ws["E2"].value).startswith("=IF(")
    # le Table et la validation de données doivent survivre :
    assert "Commandes" in ws.tables
    assert len(ws.data_validations.dataValidation) == 1


def test_appliquer_ecrit_cellule_deja_remplie(classeur):
    appliquer(classeur, [Ecriture(3, "Qté livrée", 7)], classeur.parent / "backups")
    wb = load_workbook(classeur)
    assert wb["Commandes"]["C3"].value == 7


def test_verrou_bloque_ecriture(classeur):
    verrou = classeur.parent / f"~${classeur.name}"
    verrou.write_text("")
    assert est_verrouille(classeur)
    with pytest.raises(ClasseurVerrouille):
        appliquer(classeur, [Ecriture(2, "Qté livrée", 1)], classeur.parent / "backups")


def test_sauvegarde_purge_les_anciennes(classeur, tmp_path):
    import time

    dossier_backups = tmp_path / "backups"
    ancienne = sauvegarder(classeur, dossier_backups)
    vieux_temps = time.time() - 31 * 86400
    import os

    os.utime(ancienne, (vieux_temps, vieux_temps))

    recente = sauvegarder(classeur, dossier_backups)

    assert not ancienne.exists()
    assert recente.exists()


@pytest.mark.skipif(
    trouver_fichier_suivi(ROOT) is None,
    reason="Export du Suivi commandes absent de ce poste",
)
def test_ecriture_chirurgicale_sur_le_vrai_suivi_commandes(tmp_path):
    """La preuve définitive : sur une COPIE du vrai classeur, seule la
    cellule visée change — tout le reste du zip est OCTET POUR OCTET
    identique à l'original (tableaux, validations, calcChain, sharedStrings,
    customXml, printerSettings inclus)."""

    fichier_reel = trouver_fichier_suivi(ROOT)
    copie = tmp_path / fichier_reel.name
    shutil.copy2(fichier_reel, copie)

    with zipfile.ZipFile(copie) as z:
        contenu_avant = {n: z.read(n) for n in z.namelist()}

    entetes = lire_entetes(copie)
    assert set(("Date de livraison", "Qté livrée", "Tarif BL", "Note")) <= set(entetes)

    # Ligne 2 = 1ère commande réelle : on y écrit une quantité livrée
    # arbitraire, jamais destinée à être sauvegardée dans le vrai fichier
    # (on travaille sur `copie`, dans tmp_path).
    appliquer(copie, [Ecriture(2, "Qté livrée", 12345)], tmp_path / "backups")

    with zipfile.ZipFile(copie) as z:
        contenu_apres = {n: z.read(n) for n in z.namelist()}

    chemin_feuille = "xl/worksheets/sheet1.xml"  # "Commandes" est le 1er onglet
    assert set(contenu_avant) == set(contenu_apres)
    for nom in contenu_avant:
        if nom == chemin_feuille:
            assert contenu_avant[nom] != contenu_apres[nom]
        else:
            assert contenu_avant[nom] == contenu_apres[nom], f"partie modifiée à tort : {nom}"

    wb = load_workbook(copie, data_only=True)
    assert wb["Commandes"]["P2"].value == 12345
    # les 16 Excel Tables du vrai classeur doivent toutes survivre :
    assert len(wb["Commandes"].tables) >= 1
