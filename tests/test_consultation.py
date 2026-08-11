"""
moteur/consultation.py — résolution du dossier de consultation à traiter.
Remplace l'ancien schéma besoins/ + devis/<chantier>/ + resultats/ partagé
par un dossier autonome et rejouable par consultation, voir CLAUDE.md.
"""

import pytest

from moteur.consultation import (
    ConsultationIntrouvable, lister_consultations, resoudre_dossier_consultation,
)


def _creer_consultation(dossier_projet, nom):
    dossier = dossier_projet / "consultations" / nom
    dossier.mkdir(parents=True)
    return dossier


def test_auto_selection_si_un_seul_dossier(tmp_path):
    attendu = _creer_consultation(tmp_path, "Doujani")
    assert resoudre_dossier_consultation(tmp_path) == attendu


def test_ambigu_leve_une_erreur_avec_les_choix(tmp_path):
    _creer_consultation(tmp_path, "Doujani")
    _creer_consultation(tmp_path, "Atelier")

    with pytest.raises(ConsultationIntrouvable) as exc:
        resoudre_dossier_consultation(tmp_path)

    assert "Atelier" in str(exc.value)
    assert "Doujani" in str(exc.value)


def test_aucun_dossier_leve_une_erreur_explicite(tmp_path):
    with pytest.raises(ConsultationIntrouvable):
        resoudre_dossier_consultation(tmp_path)


def test_cible_par_nom(tmp_path):
    _creer_consultation(tmp_path, "Doujani")
    attendu = _creer_consultation(tmp_path, "Atelier")
    assert resoudre_dossier_consultation(tmp_path, "Atelier") == attendu


def test_cible_par_chemin_direct(tmp_path):
    dossier = tmp_path / "ailleurs"
    dossier.mkdir()
    assert resoudre_dossier_consultation(tmp_path, str(dossier)) == dossier


def test_cible_introuvable_leve_une_erreur(tmp_path):
    with pytest.raises(ConsultationIntrouvable):
        resoudre_dossier_consultation(tmp_path, "PasLa")


def test_lister_consultations_triees(tmp_path):
    _creer_consultation(tmp_path, "Zeta")
    _creer_consultation(tmp_path, "Alpha")
    noms = [d.name for d in lister_consultations(tmp_path)]
    assert noms == ["Alpha", "Zeta"]
