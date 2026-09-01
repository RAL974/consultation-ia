"""
Logique de moteur/rapprochement/pipeline_facture.py — même esprit que
tests/test_rapprochement_pipeline_bl.py (BL) : objets Facture/
CorrespondanceFacture construits à la main pour la logique pure, un vrai
classeur xlsx (tmp_path, jamais le vrai Suivi) pour ce qui touche à
l'écriture réelle.

IMPORTANT (voir CLAUDE.md, session F2, "Volet 1") : les colonnes facture
(N° facture / Date facture / Qté facturée / PU facturé) ne sont PAS ENCORE
créées dans le VRAI Suivi commandes à ce jour — les tests d'écriture réelle
ci-dessous utilisent donc un classeur SYNTHÉTIQUE qui, lui, les a, pour
prouver que le mécanisme fonctionne dès qu'elles existeront. Un test dédié
vérifie aussi le comportement attendu en leur ABSENCE (échec propre, pas un
plantage)."""

from datetime import date

import pytest
from openpyxl import Workbook

from moteur.rapprochement.ecriture import ColonneNonModifiable
from moteur.rapprochement.matching_facture import CorrespondanceFacture, LigneSuiviFacture, StatutFacture
from moteur.rapprochement.modele_facture import Facture, LigneFacture
from moteur.rapprochement import pipeline_facture
from moteur.rapprochement.pipeline_facture import (
    RapportRapprochementFacture,
    _est_resolu_facture,
    _resoudre_commandes_facture,
    appliquer_et_archiver_factures,
    archiver_facture,
    compter_lignes_a_facturer,
    ecritures_pour_facture,
    regrouper_par_facture,
)


def _facture(fichier, numero_facture="360311", numeros_commande=None, numeros_bl=None,
             date_facture="15/07/2026", lignes=None, type_document="FACTURE"):
    return Facture(
        fournisseur="109 DISTRIBUTION", fichier=fichier, numero_facture=numero_facture,
        date_facture=date_facture, numeros_commande=numeros_commande or [],
        numeros_bl=numeros_bl or [], lignes=lignes or [], type_document=type_document,
    )


def _ligne_facture(**kwargs):
    defaut = dict(reference_fournisseur="REF1", designation="", quantite_facturee=10.0, prix_unitaire_ht=1.0)
    defaut.update(kwargs)
    return LigneFacture(**defaut)


def _ligne_suivi_facture(ligne_excel, **kwargs):
    defaut = dict(
        reference="REF1", designation="", qte_commandee=10.0, qte_livree=10.0,
        tarif_bl=1.0, tarif_convenu=None,
    )
    defaut.update(kwargs)
    return LigneSuiviFacture(ligne_excel=ligne_excel, **defaut)


def _classeur_avec_colonnes_facture(chemin, lignes=()):
    wb = Workbook()
    ws = wb.active
    ws.title = "Commandes"
    ws.append([
        "Référence", "Désignation", "Qté commandée", "N° de commande", "Fournisseur",
        "Qté livrée", "Tarif BL", "Tarif convenu",
        "N° facture", "Date facture", "Qté facturée", "PU facturé",
    ])
    for ligne in lignes:
        ws.append(ligne)
    wb.save(chemin)


def _classeur_sans_colonnes_facture(chemin, lignes=()):
    wb = Workbook()
    ws = wb.active
    ws.title = "Commandes"
    ws.append([
        "Référence", "Désignation", "Qté commandée", "N° de commande", "Fournisseur",
        "Qté livrée", "Tarif BL", "Tarif convenu",
    ])
    for ligne in lignes:
        ws.append(ligne)
    wb.save(chemin)


# --- regroupement / logique pure --------------------------------------------


def test_regrouper_par_facture_separe_les_statuts():

    f1 = _facture("f1.pdf")
    f2 = _facture("f2.pdf")

    c_sur = CorrespondanceFacture(_ligne_facture(), _ligne_suivi_facture(5), StatutFacture.SUR)
    c_inconnu = CorrespondanceFacture(_ligne_facture(), None, StatutFacture.INCONNU, ["ambigu"])

    rapport = RapportRapprochementFacture(surs=[(f1, c_sur)], inconnus=[(f2, c_inconnu)])

    groupes = regrouper_par_facture(rapport)

    assert set(groupes) == {id(f1), id(f2)}
    assert groupes[id(f1)]["sur"] == [c_sur]
    assert groupes[id(f2)]["inconnu"] == [c_inconnu]


def test_est_resolu_facture_toutes_resolues():
    c = CorrespondanceFacture(_ligne_facture(), _ligne_suivi_facture(5), StatutFacture.SUR)
    g = {"sur": [c], "a_confirmer": [], "deja_a_jour": [], "inconnu": [], "anomalies": []}

    assert _est_resolu_facture(g, {id(c)}) is True
    assert _est_resolu_facture(g, set()) is False


def test_est_resolu_facture_avec_anomalie_jamais_resolue():
    g = {"sur": [], "a_confirmer": [], "deja_a_jour": [], "inconnu": [], "anomalies": ["commande introuvable"]}

    assert _est_resolu_facture(g, set()) is False


def test_ecritures_pour_facture_construit_les_4_champs():

    f = _facture("f1.pdf", numero_facture="360311", date_facture="15/07/2026")
    lf = _ligne_facture(quantite_facturee=200.0, prix_unitaire_ht=0.6)
    c = CorrespondanceFacture(lf, _ligne_suivi_facture(5), StatutFacture.SUR)

    ecritures = ecritures_pour_facture([(f, c)])

    par_colonne = {e.colonne: e.valeur for e in ecritures if e.ligne == 5}
    assert par_colonne["N° facture"] == "360311"
    assert par_colonne["Date facture"] == date(2026, 7, 15)
    assert par_colonne["Qté facturée"] == 200.0
    assert par_colonne["PU facturé"] == 0.6


def test_ecritures_pour_facture_ignore_deja_a_jour():
    f = _facture("f1.pdf")
    c_sur = CorrespondanceFacture(_ligne_facture(), _ligne_suivi_facture(5), StatutFacture.SUR)
    c_deja = CorrespondanceFacture(_ligne_facture(), _ligne_suivi_facture(6), StatutFacture.DEJA_A_JOUR)

    ecritures = ecritures_pour_facture([(f, c_sur), (f, c_deja)])

    assert all(e.ligne != 6 for e in ecritures)
    assert any(e.ligne == 5 for e in ecritures)


# --- résolution de commande par bloc de BL ----------------------------------


def test_resoudre_commandes_facture_entete_direct_applique_a_tous_les_blocs():
    """En-tête clair (N°Réf.Client au format Suivi) -> appliqué à TOUS les
    blocs de BL de la facture, même s'il y en a plusieurs (cas de loin le
    plus courant sur le lot de cadrage réel, voir CLAUDE.md)."""

    f = _facture(
        "f1.pdf", numeros_commande=["129.034"],
        lignes=[
            _ligne_facture(reference_fournisseur="A", numero_bl="BL1"),
            _ligne_facture(reference_fournisseur="B", numero_bl="BL2"),
        ],
    )

    resolutions = _resoudre_commandes_facture(f, fichier_suivi=None)  # jamais lu si l'en-tête suffit

    assert resolutions["BL1"] == ("129.034", False, None)
    assert resolutions["BL2"] == ("129.034", False, None)


def test_resoudre_commandes_facture_deduction_par_contenu_en_repli(tmp_path):
    """N°Réf.Client vide/non exploitable -> repli sur la déduction par
    contenu (même mécanisme que côté BL, voir
    matching.deduire_commande_par_contenu) : au moins 2 références/quantités
    du bloc concordent EXACTEMENT avec une seule commande du Suivi."""

    # Références avec un cœur NUMÉRIQUE (>= 4 chiffres) — deduire_commande_
    # par_contenu s'appuie sur moteur.base.coeur_numerique, qui ignore toute
    # référence sans chiffres exploitables (voir son seuil de fiabilité).
    chemin_suivi = tmp_path / "suivi.xlsx"
    _classeur_avec_colonnes_facture(chemin_suivi, lignes=[
        ["REF4501", "", 5, "129.099", "109 DISTRIBUTION", 5, 1.0, None, None, None, None, None],
        ["REF7802", "", 3, "129.099", "109 DISTRIBUTION", 3, 1.0, None, None, None, None, None],
    ])

    f = _facture(
        "f1.pdf", numeros_commande=[],  # en-tête illisible/non exploitable
        lignes=[
            _ligne_facture(reference_fournisseur="REF4501", quantite_facturee=5.0, numero_bl="BL1"),
            _ligne_facture(reference_fournisseur="REF7802", quantite_facturee=3.0, numero_bl="BL1"),
        ],
    )

    resolutions = _resoudre_commandes_facture(f, chemin_suivi)

    numero_commande, deduit, raison = resolutions["BL1"]
    assert numero_commande == "129.099"
    assert deduit is True
    assert "déduit du contenu" in raison


def test_resoudre_commandes_facture_rien_de_deductible_retourne_none(tmp_path):

    chemin_suivi = tmp_path / "suivi.xlsx"
    _classeur_avec_colonnes_facture(chemin_suivi, lignes=[
        ["AUTREREF", "", 5, "999.999", "109 DISTRIBUTION", 5, 1.0, None, None, None, None, None],
    ])

    f = _facture(
        "f1.pdf", numeros_commande=[],
        lignes=[_ligne_facture(reference_fournisseur="INCONNUE", numero_bl="BL1")],
    )

    resolutions = _resoudre_commandes_facture(f, chemin_suivi)

    assert resolutions["BL1"] == (None, False, None)


# --- archivage (copie multi-commande, jamais un déplacement) ---------------


def test_archiver_facture_copie_dans_chaque_commande_concernee_et_supprime_la_source(tmp_path):
    """Une facture peut couvrir PLUSIEURS commandes (cas réel,
    Facture_365533.pdf) — copiée dans CHAQUE dossier de commande concerné,
    jamais découpée (contrairement à un BL Cominter multi-BL, voir bandeau
    du module) ; la source n'est supprimée qu'une fois toutes les copies
    faites."""

    source = tmp_path / "a_traiter" / "Factures" / "Facture_365533.pdf"
    source.parent.mkdir(parents=True)
    source.write_bytes(b"%PDF-1.4 factice")

    traites = tmp_path / "a_traiter" / "BL" / "Traités"
    f = _facture("Facture_365533.pdf", numero_facture="365533", date_facture="17/08/2026")

    cibles = archiver_facture(source, f, ["132.008", "132.033"], traites)

    assert not source.exists()
    assert len(cibles) == 2
    assert {c.parent.name for c in cibles} == {"132.008", "132.033"}
    for c in cibles:
        assert c.exists()
        assert "365533" in c.name
        assert "109 DISTRIBUTION" in c.name


def test_archiver_facture_evite_lecrasement(tmp_path):

    traites = tmp_path / "Traités"
    f = _facture("f.pdf", numero_facture="1", date_facture="01/01/2026")

    source1 = tmp_path / "f1.pdf"
    source1.write_bytes(b"1")
    [cible1] = archiver_facture(source1, f, ["C1"], traites)

    source2 = tmp_path / "f2.pdf"
    source2.write_bytes(b"2")
    [cible2] = archiver_facture(source2, f, ["C1"], traites)

    assert cible1 != cible2
    assert cible1.exists() and cible2.exists()


# --- résorption --------------------------------------------------------------


def test_compter_lignes_a_facturer(tmp_path):

    chemin_suivi = tmp_path / "suivi.xlsx"
    _classeur_avec_colonnes_facture(chemin_suivi, lignes=[
        ["REF1", "", 10, "C1", "109 DISTRIBUTION", 10, 1.0, None, "F1", None, 10, 1.0],  # déjà facturée
        ["REF2", "", 5, "C2", "109 DISTRIBUTION", 5, 1.0, None, None, None, None, None],  # livrée, pas facturée
        ["REF3", "", 5, "C3", "109 DISTRIBUTION", 0, None, None, None, None, None, None],  # pas livrée du tout
    ])

    r = compter_lignes_a_facturer(chemin_suivi, "109 DISTRIBUTION")

    assert r == {"livrees": 2, "a_facturer": 1, "deja_facturees": 1}


def test_compter_lignes_a_facturer_colonnes_absentes_reste_honnete(tmp_path):
    """Sans les colonnes facture (état actuel du vrai Suivi, voir CLAUDE.md),
    toute ligne livrée ressort "à facturer" — signal honnête, pas un 0
    trompeur qui masquerait leur absence."""

    chemin_suivi = tmp_path / "suivi.xlsx"
    _classeur_sans_colonnes_facture(chemin_suivi, lignes=[
        ["REF1", "", 10, "C1", "109 DISTRIBUTION", 10, 1.0, None],
    ])

    r = compter_lignes_a_facturer(chemin_suivi, "109 DISTRIBUTION")

    assert r == {"livrees": 1, "a_facturer": 1, "deja_facturees": 0}


# --- écriture réelle de bout en bout (classeur synthétique, jamais le vrai) -


def test_appliquer_et_archiver_factures_ecrit_et_archive(tmp_path):
    """Preuve que le mécanisme d'écriture fonctionne de bout en bout — sur
    un classeur SYNTHÉTIQUE qui a les colonnes facture (le vrai Suivi ne
    les a pas encore, voir bandeau du module)."""

    chemin_suivi = tmp_path / "suivi.xlsx"
    _classeur_avec_colonnes_facture(chemin_suivi, lignes=[
        ["REF1", "Article test", 200, "129.034", "109 DISTRIBUTION", 200, 0.6, None, None, None, None, None],
    ])

    dossier_a_traiter = tmp_path / "a_traiter" / "Factures"
    dossier_a_traiter.mkdir(parents=True)
    (dossier_a_traiter / "f1.pdf").write_bytes(b"1")

    lf = _ligne_facture(reference_fournisseur="REF1", quantite_facturee=200.0, prix_unitaire_ht=0.6)
    lf.numero_commande = "129.034"
    f = _facture("f1.pdf", numero_facture="360311", date_facture="15/07/2026", lignes=[lf])

    ls = _ligne_suivi_facture(2, reference="REF1", qte_commandee=200.0, qte_livree=200.0, tarif_bl=0.6)
    c = CorrespondanceFacture(lf, ls, StatutFacture.SUR)

    rapport = RapportRapprochementFacture(surs=[(f, c)], fichier_suivi=chemin_suivi)

    resume = appliquer_et_archiver_factures(tmp_path, dossier_a_traiter, rapport, rapport.surs)

    assert resume["lignes_ecrites"] == 1
    assert resume["sauvegarde"].exists()

    from openpyxl import load_workbook
    wb = load_workbook(chemin_suivi, data_only=True)
    ws = wb["Commandes"]
    assert ws["I2"].value == "360311"       # N° facture
    assert ws["K2"].value == 200.0          # Qté facturée
    assert ws["L2"].value == 0.6            # PU facturé

    [(fichier, cibles)] = resume["factures_archivees"]
    assert fichier == "f1.pdf"
    [cible] = cibles
    assert cible.exists()
    assert cible.parent.name == "129.034"
    assert not (dossier_a_traiter / "f1.pdf").exists()

    assert resume["resorption"] is not None
    assert resume["chemin_rapport"].exists()


def test_appliquer_et_archiver_factures_bloc_non_resolu_va_en_a_verifier(tmp_path):

    chemin_suivi = tmp_path / "suivi.xlsx"
    _classeur_avec_colonnes_facture(chemin_suivi)

    dossier_a_traiter = tmp_path / "a_traiter" / "Factures"
    dossier_a_traiter.mkdir(parents=True)
    (dossier_a_traiter / "f1.pdf").write_bytes(b"1")

    f = _facture("f1.pdf")
    c_inconnu = CorrespondanceFacture(_ligne_facture(), None, StatutFacture.INCONNU, ["référence inconnue"])

    rapport = RapportRapprochementFacture(inconnus=[(f, c_inconnu)], fichier_suivi=chemin_suivi)

    resume = appliquer_et_archiver_factures(tmp_path, dossier_a_traiter, rapport, [])

    assert resume["factures_archivees"] == []
    [(fichier, cible, raisons)] = resume["factures_a_verifier"]
    assert fichier == "f1.pdf"
    assert cible == dossier_a_traiter / "À vérifier" / "f1.pdf"
    assert cible.exists()
    assert "référence inconnue" in raisons[0]


def test_ecriture_echoue_proprement_sans_les_colonnes_facture(tmp_path):
    """État ACTUEL du vrai Suivi commandes (voir CLAUDE.md, Volet 1) : les
    colonnes facture n'existent pas encore -> ColonneNonModifiable, message
    clair, jamais un plantage générique ni une écriture partielle."""

    chemin_suivi = tmp_path / "suivi.xlsx"
    _classeur_sans_colonnes_facture(chemin_suivi, lignes=[
        ["REF1", "", 10, "C1", "109 DISTRIBUTION", 10, 1.0, None],
    ])

    dossier_a_traiter = tmp_path / "a_traiter" / "Factures"
    dossier_a_traiter.mkdir(parents=True)
    (dossier_a_traiter / "f1.pdf").write_bytes(b"1")

    lf = _ligne_facture(reference_fournisseur="REF1", quantite_facturee=10.0, prix_unitaire_ht=1.0)
    f = _facture("f1.pdf", lignes=[lf])
    ls = _ligne_suivi_facture(2, reference="REF1", qte_livree=10.0, tarif_bl=1.0)
    c = CorrespondanceFacture(lf, ls, StatutFacture.SUR)

    rapport = RapportRapprochementFacture(surs=[(f, c)], fichier_suivi=chemin_suivi)

    with pytest.raises(ColonneNonModifiable):
        appliquer_et_archiver_factures(tmp_path, dossier_a_traiter, rapport, rapport.surs)

    # Rien n'a été déplacé : mieux vaut tout laisser en l'état qu'archiver
    # sur la base d'une écriture qui n'a en réalité jamais eu lieu.
    assert (dossier_a_traiter / "f1.pdf").exists()
