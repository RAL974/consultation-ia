# -*- coding: utf-8 -*-
"""
moteur.fnp_brouillon — uniquement les fonctions PURES (_euro_fr,
_corps_mail_fnp) : creer_brouillon_fnp() elle-même exige Outlook/pywin32
réels, jamais testée automatiquement (même limite déjà acceptée pour
moteur.fnp_brouillon avant cette session — pas de mock COM dans ce projet).
"""

from datetime import date

from moteur.fnp import RapportFNP, ReservesFNP, LigneFNP, DossierTransitaire, AjustementFNP, FactureRecueNonRapprochee
from moteur.fnp_brouillon import _corps_mail_fnp, _euro_fr

# Séparateur de milliers de _euro_fr : espace INSÉCABLE (U+00A0), convention
# typographique française correcte (empêche un nombre de se couper en fin de
# ligne) — écrit ici en   explicite pour ne dépendre d'aucun caractère
# ambigu tapé au clavier.
NBSP = " "


def test_euro_fr_format_francais():
    assert _euro_fr(1234.5) == f"1{NBSP}234,50 €"
    assert _euro_fr(0) == "0,00 €"
    assert _euro_fr(1234567.89) == f"1{NBSP}234{NBSP}567,89 €"


def _rapport_minimal(**kwargs):
    defaut = dict(
        mois="2026-08", fin_de_mois=date(2026, 8, 31), date_generation=None,
        chemin_suivi=None, suivi_modifie_le=None, depuis=None,
    )
    defaut.update(kwargs)
    return RapportFNP(**defaut)


def test_corps_mail_fnp_totaux_de_base():
    ligne_fnp = LigneFNP(
        ligne_excel=2, fournisseur="COREDIME", numero_commande="M1", chantier="C1",
        reference="R1", designation="D1", qte_livree=1.0, montant_ht=1000.0,
        source_prix="Tarif BL", date_livraison=date(2026, 8, 1), anciennete_jours=30,
        numero_facture="", date_facture=None, note="",
    )
    dossier_transit = DossierTransitaire(
        numero_dossier="R1", designation="D", numero_commande="M1", chantier="C1",
        fournisseur="F1", transitaire="T1", ref_transport="", date_depart=None,
        date_arrivee=date(2026, 8, 10), montant_marchandise=5000.0, cout_estime=250.0,
        anciennete_jours=21,
    )
    rapport = _rapport_minimal(lignes_bl=[ligne_fnp], dossiers_transitaires=[dossier_transit])

    corps = _corps_mail_fnp(rapport, "août 2026")

    assert "août 2026" in corps
    assert f"1{NBSP}000,00 €" in corps
    assert "250,00 €" in corps
    assert "(1 ligne(s))" in corps
    assert "(1 dossier(s))" in corps


def test_corps_mail_fnp_ajustements_et_factures_recues_optionnels():
    rapport_sans = _rapport_minimal()
    corps_sans = _corps_mail_fnp(rapport_sans, "août 2026")
    assert "Déclaré par l'acheteur" not in corps_sans
    assert "pas encore rapprochée" not in corps_sans

    ajustement = AjustementFNP(
        type="AUTRE", libelle="Test", fournisseur_ou_transitaire="X", chantier="",
        piece="", date_livraison=None, montant_ht=99.0, source="", commentaire="",
    )
    facture_recue = FactureRecueNonRapprochee(
        ligne_excel=2, fournisseur="COREDIME", numero_commande="M1", reference="R1",
        numero_facture="F1", date_facture=date(2026, 8, 5), montant_facture_bl=42.0,
    )
    rapport_avec = _rapport_minimal(ajustements=[ajustement], factures_recues_non_rapprochees=[facture_recue])
    corps_avec = _corps_mail_fnp(rapport_avec, "août 2026")

    assert "Déclaré par l'acheteur" in corps_avec
    assert "99,00 €" in corps_avec
    assert "1 facture(s) déjà reçue(s)" in corps_avec


def test_corps_mail_fnp_reserves_en_clair():
    rapport = _rapport_minimal(reserves=ReservesFNP(
        n_bdc_manuel_24x=3, n_transitaires_sans_estimation=2, n_dossiers_speciales_total=31,
    ))

    corps = _corps_mail_fnp(rapport, "août 2026")

    assert "3 facture(s) portant sur un bon manuel" in corps
    assert "2 dossier(s) transitaire sans estimation" in corps
    assert "31 dossier(s)" in corps
