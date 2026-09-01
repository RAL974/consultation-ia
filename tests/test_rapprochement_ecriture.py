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

import re
import shutil
import zipfile
from datetime import date

import pytest
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation

from moteur.rapprochement.ecriture import (
    COLONNES_MODIFIABLES,
    ENTETES_FACTURE,
    Ecriture,
    ClasseurVerrouille,
    ColonneNonModifiable,
    ajouter_entetes_saisie,
    appliquer,
    est_verrouille,
    lire_entetes,
    sauvegarder,
    simuler,
)
from moteur.panier import trouver_fichier_suivi
from moteur.rapprochement.pipeline_bl import trouver_fichier_suivi_vivant

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


def test_colonnes_facture_modifiables():
    """4 colonnes facture (voir CLAUDE.md, session F2, Volet 1) — même
    statut que les 4 colonnes BL déjà là : saisie brute, jamais une
    formule."""

    assert set(("N° facture", "Date facture", "Qté facturée", "PU facturé")) <= set(COLONNES_MODIFIABLES)


def test_appliquer_ecrit_les_colonnes_facture(tmp_path):
    """Le classeur synthétique existant n'a pas les colonnes facture — un
    2e classeur, dédié, les ajoute pour prouver que le patch chirurgical
    fonctionne aussi pour elles (le vrai Suivi ne les a pas encore créées à
    ce jour, voir CLAUDE.md — ce test tourne donc sur un classeur qui, lui,
    les a)."""

    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Commandes"
    ws.append(["Référence", "Qté livrée", "Tarif BL", "N° facture", "Date facture", "Qté facturée", "PU facturé"])
    ws.append(["ART1", 10, 5.0, None, None, None, None])
    chemin = tmp_path / "suivi_facture.xlsx"
    wb.save(chemin)

    appliquer(
        chemin,
        [
            Ecriture(2, "N° facture", "360311"),
            Ecriture(2, "Qté facturée", 10.0),
            Ecriture(2, "PU facturé", 5.0),
        ],
        tmp_path / "backups",
    )

    wb2 = load_workbook(chemin, data_only=True)
    ws2 = wb2["Commandes"]
    assert ws2["D2"].value == "360311"
    assert ws2["F2"].value == 10.0
    assert ws2["G2"].value == 5.0


def test_ajouter_entetes_saisie_integre_les_colonnes_dans_le_tableau(classeur, tmp_path):
    """Les nouveaux en-têtes vont à la suite de la dernière colonne
    existante ET rejoignent le tableau structuré (ref étendue, autoFilter
    étendu, tableColumns complété) — jamais posés à côté : un tri du
    tableau ne suit pas les colonnes restées hors tableau, les données
    facture se retrouveraient sur les mauvaises lignes (décision explicite
    de l'acheteur, voir CLAUDE.md)."""

    sauvegarde = ajouter_entetes_saisie(
        classeur, ["N° facture", "Date facture"], tmp_path / "backups"
    )
    assert sauvegarde.exists()

    entetes = lire_entetes(classeur)
    assert entetes["N° facture"] == 7  # G, juste après F ("Note")
    assert entetes["Date facture"] == 8  # H

    wb = load_workbook(classeur, data_only=False)
    ws = wb["Commandes"]
    assert ws["G1"].value == "N° facture"
    assert ws["H1"].value == "Date facture"
    assert str(ws["E2"].value).startswith("=IF(")  # formule voisine intacte

    table = ws.tables["Commandes"]
    assert table.ref == "A1:H3"  # le tableau s'étend pour couvrir les 2 colonnes
    noms_colonnes = [c.name for c in table.tableColumns]
    assert len(noms_colonnes) == 8
    assert noms_colonnes[-2:] == ["N° facture", "Date facture"]


def test_ajouter_entetes_saisie_refuse_colonne_deja_dans_le_tableau(classeur, tmp_path):
    """Garde-fou défensif : un nom déjà présent comme colonne du TABLEAU
    (mais pas forcément détecté par lire_entetes(), qui ne lit que la
    ligne 1 de la FEUILLE) doit aussi bloquer — simule une
    désynchronisation entre la feuille et la définition du tableau."""

    with zipfile.ZipFile(classeur) as z:
        contenu = {n: z.read(n) for n in z.namelist()}
    xml = contenu["xl/tables/table1.xml"].decode("utf-8")
    xml = xml.replace(
        "</tableColumns>", '<tableColumn id="7" name="N° facture"/></tableColumns>'
    )
    contenu["xl/tables/table1.xml"] = xml.encode("utf-8")
    with zipfile.ZipFile(classeur, "w", zipfile.ZIP_DEFLATED) as z:
        for nom, donnees in contenu.items():
            z.writestr(nom, donnees)

    avant = classeur.read_bytes()
    with pytest.raises(ValueError):
        ajouter_entetes_saisie(classeur, ["N° facture"], tmp_path / "backups")
    assert classeur.read_bytes() == avant  # refus AVANT toute écriture
    assert not (tmp_path / "backups").exists()


def test_ajouter_entetes_saisie_refuse_nom_deja_present(classeur, tmp_path):
    avant = classeur.read_bytes()
    with pytest.raises(ValueError):
        ajouter_entetes_saisie(classeur, ["Note"], tmp_path / "backups")
    assert classeur.read_bytes() == avant  # refus AVANT toute écriture


def test_ajouter_entetes_saisie_refuse_noms_en_double(classeur, tmp_path):
    with pytest.raises(ValueError):
        ajouter_entetes_saisie(classeur, ["N° facture", "N° facture"], tmp_path / "backups")


def test_ajouter_entetes_saisie_bloque_si_verrouille(classeur, tmp_path):
    verrou = classeur.parent / f"~${classeur.name}"
    verrou.write_text("")
    with pytest.raises(ClasseurVerrouille):
        ajouter_entetes_saisie(classeur, ["N° facture"], tmp_path / "backups")


def test_ajouter_entetes_saisie_refuse_cellule_deja_presente(classeur, tmp_path):
    """Garde-fou défensif : même une cellule présente mais SANS valeur (donc
    invisible de lire_entetes) à l'emplacement visé doit bloquer l'ajout —
    simule un résidu jamais nettoyé à la position G1 (celle que la fonction
    viserait normalement)."""

    with zipfile.ZipFile(classeur) as z:
        contenu = {n: z.read(n) for n in z.namelist()}
    xml = contenu["xl/worksheets/sheet1.xml"].decode("utf-8")
    xml = re.sub(
        r'(<row r="1"[^>]*>.*?)(</row>)',
        r'\1<c r="G1" s="1"/>\2',
        xml,
        count=1,
        flags=re.S,
    )
    contenu["xl/worksheets/sheet1.xml"] = xml.encode("utf-8")
    with zipfile.ZipFile(classeur, "w", zipfile.ZIP_DEFLATED) as z:
        for nom, donnees in contenu.items():
            z.writestr(nom, donnees)

    avant = classeur.read_bytes()
    with pytest.raises(ValueError):
        ajouter_entetes_saisie(classeur, ["N° facture"], tmp_path / "backups")
    assert classeur.read_bytes() == avant  # refus AVANT toute écriture (pas de sauvegarde non plus)
    assert not (tmp_path / "backups").exists()


@pytest.mark.skipif(
    trouver_fichier_suivi_vivant(ROOT) is None,
    reason="Classeur Suivi commandes VIVANT introuvable depuis ce poste (dossier "
    "frère « 1.3.0.1. Commandes courantes » absent ou inaccessible)",
)
def test_ajouter_entetes_saisie_sur_le_vrai_suivi_commandes_vivant(tmp_path):
    """La preuve définitive pour l'étape 1 (voir CLAUDE.md) : sur une COPIE
    du classeur VIVANT réellement utilisé par l'acheteur (refondu — feuilles
    Dashboard/Analyses/Calculs — jamais l'export périmé de la racine du
    dépôt, voir trouver_fichier_suivi_vivant), ajouter des colonnes DANS le
    tableau structuré Commandes (décision explicite de l'acheteur — un tri
    du tableau ne suit pas une colonne restée hors tableau) ne touche
    STRICTEMENT que la feuille (ligne 1 + <dimension>) et la définition du
    tableau (ref, autoFilter, tableColumns) : les formules existantes
    (Statut commande, Facturé BL...) et tout le reste du zip survivent
    identiques.

    L'état "avant" (ref du tableau, nombre de colonnes) est lu dynamiquement
    plutôt que supposé figé : les 5 vraies colonnes facture (ENTETES_FACTURE)
    ont été créées pour de vrai dans le classeur vivant à l'étape 2 de cette
    même session — ce test réutilise donc 2 noms JAMAIS utilisés en pratique,
    uniquement pour continuer à prouver le MÉCANISME sur la vraie structure
    du fichier, indéfiniment (même après l'ajout d'autres colonnes futures)."""

    fichier_reel = trouver_fichier_suivi_vivant(ROOT)
    copie = tmp_path / fichier_reel.name
    shutil.copy2(fichier_reel, copie)
    avant_fichier = copie.read_bytes()

    with zipfile.ZipFile(copie) as z:
        contenu_avant = {n: z.read(n) for n in z.namelist()}

    wb_avant = load_workbook(copie, data_only=False)
    ws_avant = wb_avant["Commandes"]
    table_avant = ws_avant.tables["Commandes"]
    ref_avant = table_avant.ref
    nb_colonnes_avant = len(table_avant.tableColumns)
    derniere_colonne_avant = max(lire_entetes(copie).values())
    formule_statut_avant = ws_avant["U100"].value
    formule_facture_bl_avant = ws_avant["AV100"].value
    wb_avant.close()

    noms_test = ["Colonne test rapprochement A", "Colonne test rapprochement B"]
    sauvegarde = ajouter_entetes_saisie(copie, noms_test, tmp_path / "backups")
    assert sauvegarde.exists()
    assert sauvegarde.read_bytes() == avant_fichier  # sauvegarde = état AVANT patch, intégral
    assert sauvegarde.read_bytes() != copie.read_bytes()

    with zipfile.ZipFile(copie) as z:
        contenu_apres = {n: z.read(n) for n in z.namelist()}

    chemin_feuille = "xl/worksheets/sheet1.xml"
    chemin_table = "xl/tables/table1.xml"  # tableau structuré "Commandes"
    assert set(contenu_avant) == set(contenu_apres)
    parties_modifiees = {chemin_feuille, chemin_table}
    for nom in contenu_avant:
        if nom in parties_modifiees:
            assert contenu_avant[nom] != contenu_apres[nom]
        else:
            assert contenu_avant[nom] == contenu_apres[nom], f"partie modifiée à tort : {nom}"

    # À l'intérieur de sheet1.xml : seules la ligne 1 et <dimension> ont
    # changé, rigoureusement rien d'autre (aucune autre ligne, aucun
    # décalage de calcChain/sharedStrings...).
    def _sans_ligne1_ni_dimension(xml):
        xml = re.sub(r'<row r="1"[^>]*>.*?</row>', "", xml, count=1, flags=re.S)
        xml = re.sub(r'<dimension ref="[^"]*"\s*/>', "", xml, count=1)
        return xml

    xml_feuille_avant = contenu_avant[chemin_feuille].decode("utf-8")
    xml_feuille_apres = contenu_apres[chemin_feuille].decode("utf-8")
    assert _sans_ligne1_ni_dimension(xml_feuille_avant) == _sans_ligne1_ni_dimension(xml_feuille_apres)

    nouvelle_derniere_colonne = get_column_letter(derniere_colonne_avant + len(noms_test))
    m_dim = re.search(r'<dimension ref="[A-Z]+\d+:([A-Z]+)(\d+)"\s*/>', xml_feuille_avant)
    assert m_dim is not None
    ligne_fin = m_dim.group(2)
    assert f'<dimension ref="A1:{nouvelle_derniere_colonne}{ligne_fin}"/>' in xml_feuille_apres

    entetes_apres = lire_entetes(copie)
    for i, nom in enumerate(noms_test):
        assert entetes_apres[nom] == derniere_colonne_avant + 1 + i

    # À l'intérieur de table1.xml : reconstruction INDÉPENDANTE (remplacement
    # de texte brut, pas la même logique regex que _etendre_tableau) de ce
    # que le patch doit produire — seuls le ref du <table>, le ref du
    # <autoFilter> et le count de <tableColumns> changent, plus les
    # <tableColumn> ajoutés en toute fin, rien d'autre.
    xml_table_avant = contenu_avant[chemin_table].decode("utf-8")
    xml_table_apres = contenu_apres[chemin_table].decode("utf-8")
    ancienne_ref_attr = f'ref="{ref_avant}"'
    ancien_count_attr = f'<tableColumns count="{nb_colonnes_avant}">'
    assert ancienne_ref_attr in xml_table_avant
    assert ancien_count_attr in xml_table_avant

    m_ref = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)$", ref_avant)
    debut_col, debut_ligne, fin_col, fin_ligne = m_ref.groups()
    nouvelle_fin_col = get_column_letter(column_index_from_string(fin_col) + len(noms_test))
    nouvelle_ref = f"{debut_col}{debut_ligne}:{nouvelle_fin_col}{fin_ligne}"

    ids_existants = [int(i) for i in re.findall(r'<tableColumn\b[^>]*\bid="(\d+)"', xml_table_avant)]
    id_suivant = max(ids_existants) + 1
    fragments = "".join(
        f'<tableColumn id="{id_suivant + i}" name="{nom}"/>' for i, nom in enumerate(noms_test)
    )
    nouveau_count_attr = f'<tableColumns count="{nb_colonnes_avant + len(noms_test)}">'
    xml_table_attendu = xml_table_avant.replace(ancienne_ref_attr, f'ref="{nouvelle_ref}"')
    xml_table_attendu = xml_table_attendu.replace(ancien_count_attr, nouveau_count_attr)
    xml_table_attendu = xml_table_attendu.replace("</tableColumns>", fragments + "</tableColumns>")
    assert xml_table_apres == xml_table_attendu

    wb_apres = load_workbook(copie, data_only=False)
    ws_apres = wb_apres["Commandes"]
    col1 = get_column_letter(derniere_colonne_avant + 1)
    col2 = get_column_letter(derniere_colonne_avant + 2)
    assert ws_apres[f"{col1}1"].value == noms_test[0]
    assert ws_apres[f"{col2}1"].value == noms_test[1]
    # le tableau structuré Commandes s'étend pour couvrir les nouvelles
    # colonnes (indépendant, via le parseur XML d'openpyxl) :
    table_apres = ws_apres.tables["Commandes"]
    assert table_apres.ref == nouvelle_ref
    noms_colonnes_apres = [c.name for c in table_apres.tableColumns]
    assert len(noms_colonnes_apres) == nb_colonnes_avant + len(noms_test)
    assert noms_colonnes_apres[-len(noms_test):] == noms_test
    # les formules existantes, protégées par COLONNES_MODIFIABLES, sont
    # EXACTEMENT les mêmes qu'avant (aucun calculatedColumnFormula ajouté
    # aux nouvelles colonnes, aucune formule existante retouchée) :
    assert ws_apres["U100"].value == formule_statut_avant
    assert ws_apres["AV100"].value == formule_facture_bl_avant
    wb_apres.close()


@pytest.mark.skipif(
    trouver_fichier_suivi_vivant(ROOT) is None,
    reason="Classeur Suivi commandes VIVANT introuvable depuis ce poste",
)
def test_appliquer_ecrit_les_colonnes_facture_sur_le_vrai_suivi_vivant(tmp_path):
    """Étape 3 (voir CLAUDE.md) : une fois les 5 colonnes facture réellement
    créées dans le classeur vivant par ajouter_entetes_saisie() (étape 2),
    appliquer() doit pouvoir y écrire comme dans n'importe quelle colonne de
    COLONNES_MODIFIABLES — même patch chirurgical (une seule partie du zip
    modifiée), aucune autre partie touchée."""

    fichier_reel = trouver_fichier_suivi_vivant(ROOT)
    copie = tmp_path / fichier_reel.name
    shutil.copy2(fichier_reel, copie)

    entetes = lire_entetes(copie)
    assert set(ENTETES_FACTURE) <= set(entetes)  # les 5 colonnes existent bien désormais

    with zipfile.ZipFile(copie) as z:
        contenu_avant = {n: z.read(n) for n in z.namelist()}

    appliquer(
        copie,
        [
            Ecriture(2, "N° facture", "TEST-360999"),
            Ecriture(2, "Date facture", date(2026, 9, 1)),
            Ecriture(2, "Qté facturée", 7.0),
            Ecriture(2, "PU facturé", 12.5),
            Ecriture(2, "Montant facturé HT", 87.5),
        ],
        tmp_path / "backups",
    )

    with zipfile.ZipFile(copie) as z:
        contenu_apres = {n: z.read(n) for n in z.namelist()}

    chemin_feuille = "xl/worksheets/sheet1.xml"
    assert set(contenu_avant) == set(contenu_apres)
    for nom in contenu_avant:
        if nom == chemin_feuille:
            assert contenu_avant[nom] != contenu_apres[nom]
        else:
            assert contenu_avant[nom] == contenu_apres[nom], f"partie modifiée à tort : {nom}"

    wb = load_workbook(copie, data_only=True)
    ws = wb["Commandes"]
    assert ws.cell(row=2, column=entetes["N° facture"]).value == "TEST-360999"
    assert ws.cell(row=2, column=entetes["Qté facturée"]).value == 7.0
    assert ws.cell(row=2, column=entetes["PU facturé"]).value == 12.5
    assert ws.cell(row=2, column=entetes["Montant facturé HT"]).value == 87.5


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
