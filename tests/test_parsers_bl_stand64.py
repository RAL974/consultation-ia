"""
Tests de non-régression du rapprochement STAND 64 (Rapprochement AI —
nouveau fournisseur, à partir de VRAIES pièces scannées, tests/fixtures/
bl_stand64_*.pdf, copies de a_traiter/BL/).

Les 2 seules pièces réelles disponibles à ce jour partagent la même
structure (tableau simple, colonnes DANS L'ORDRE contrairement au devis de
ce fournisseur, voir bandeau GABARIT BL de moteur/fournisseurs/stand64.py) —
"Eco-part" toujours vide sur les deux, aucune vraie remise rencontrée.
"""

import functools

from moteur.ocr import mots_document, texte_ocr
from moteur.detecteur import detecter_fournisseur
from moteur.fournisseurs.stand64 import parse_bl_stand64
from moteur.rapprochement.modele_bl import LigneBL

from conftest import FIXTURES


@functools.lru_cache(maxsize=None)
def _mots(nom):
    return mots_document(FIXTURES / nom)


def test_detection_fournisseur_bl_stand64():
    assert detecter_fournisseur(texte_ocr(FIXTURES / "bl_stand64_1.pdf")) == "STAND 64"


def test_parse_bl_stand64_1_six_lignes():
    bl = parse_bl_stand64(_mots("bl_stand64_1.pdf"))

    assert bl.fournisseur == "STAND 64"
    assert bl.numero_bl == "45617"
    assert bl.date_bl == "18/08/2026"
    assert bl.numero_commande == "M2.23.058"

    assert bl.lignes == [
        LigneBL(reference_fournisseur='WESTI-73044', designation='COMETIll5PALES132BLANC+ERABLE/BLANC+chainette', quantite_livree=7.0, prix_net=91.0, montant=637.0),
        LigneBL(reference_fournisseur='WESTI-73044', designation='COMET III5PALES132BLANC+ERABLE/BLANC+chainette', quantite_livree=3.0, prix_net=91.0, montant=273.0),
        LigneBL(reference_fournisseur='WESTI-COMET-KITLUM-B', designation='COMETKITLUMIERE2X40WBLANC', quantite_livree=3.0, prix_net=7.0, montant=21.0),
        LigneBL(reference_fournisseur='WESTI-78801', designation='COMMANDEMURALEVENT4VITESSES+LUMIERE', quantite_livree=10.0, prix_net=15.0, montant=150.0),
        LigneBL(reference_fournisseur='WESTI-78800', designation='COMMANDEMURALEVENT4VITESSES', quantite_livree=5.0, prix_net=12.0, montant=60.0),
        LigneBL(reference_fournisseur='WESTI-78095', designation='TELECOMMANDEINFRAROUGE3VITESSES+LUMIEREsanspiles', quantite_livree=8.0, prix_net=35.0, montant=280.0),
    ]

    # Total HT affiché sur la pièce : 1 421,00€ (voir bandeau GABARIT BL,
    # pas encore de champ total_ht_affiche pour ce fournisseur).
    assert round(sum(l.montant for l in bl.lignes), 2) == 1421.0


def test_parse_bl_stand64_2_trois_lignes_commande_courte():
    """Commande sans point ni tiret rallongé ("M4.270") — vérifie que le
    motif de commande ne capture pas un fragment tronqué."""

    bl = parse_bl_stand64(_mots("bl_stand64_2.pdf"))

    assert bl.numero_bl == "45618"
    assert bl.date_bl == "18/08/2026"
    assert bl.numero_commande == "M4.270"

    assert bl.lignes == [
        LigneBL(reference_fournisseur='WESTI-COMET-KITLUM-N', designation='COMETKITLUMIERE2X40WNOIR', quantite_livree=1.0, prix_net=7.0, montant=7.0),
        LigneBL(reference_fournisseur='WESTI-73045', designation='COMETII5PALES132NOIR+ROSEWOOD/NOIR+chainette', quantite_livree=1.0, prix_net=91.0, montant=91.0),
        LigneBL(reference_fournisseur='WESTI-78095', designation='TELECOMMANDEINFRAROUGE3VITESSES+LUMIEREsanspiles', quantite_livree=1.0, prix_net=35.0, montant=35.0),
    ]

    assert round(sum(l.montant for l in bl.lignes), 2) == 133.0


def test_parse_bl_stand64_3_ecopart_renseigne_et_designation_manquante():
    """Cas réel (commande M2.5.126) : l'Eco-part est RENSEIGNÉE (2,88€,
    pas vide comme sur les 2 premières pièces) — décale toute la lecture
    positionnelle d'une cellule si on suppose toujours 6 cellules
    chiffrées. En plus, la 1re ligne de désignation ("ES52-2678 3 PALES
    Ø132 6 VITESSES 18W LED...") n'a PAS été détectée par l'OCR sur ce
    document précis — seules les lignes de désignation SUIVANTES l'ont
    été ("BLANC/BLANC+TELECOMMANDE", "PRIX NETS", "MATERIEL DISPONIBLE CE
    JOUR"), raccordées à la ligne chiffrée par look-ahead."""

    bl = parse_bl_stand64(_mots("bl_stand64_3_ecopart_renseigne_desig_manquante.pdf"))

    assert bl.numero_bl == "45647"
    assert bl.numero_commande == "M2.5.126"

    assert bl.lignes == [
        LigneBL(
            reference_fournisseur='ELIOT-ES52-2678-BLC+BLC',
            designation='BLANC/BLANC+TELECOMMANDE PRIX NETS MATERIELDISPONIBLECEJOUR',
            quantite_livree=18.0, prix_net=95.0, montant=1710.0,
        ),
    ]


def test_parse_bl_stand64_4_coche_collee_a_la_quantite():
    """Cas réel (commande 131.170) : une coche imprimée à côté de la Qté
    est lue par l'OCR comme un "V" collé à la valeur ("40,00V" au lieu de
    "40,00") — sans nettoyage, to_float() échouait et TOUTE la ligne
    disparaissait silencieusement (0 ligne pour un Total HT de 580,00€
    affiché)."""

    bl = parse_bl_stand64(_mots("bl_stand64_4_coche_collee_qte.pdf"))

    assert bl.numero_bl == "45646"
    assert bl.numero_commande == "131.170"

    assert bl.lignes == [
        LigneBL(reference_fournisseur='ELIOT-SUNE27-275-BL', designation='SUN27HUBLOTE27IP65IK1OCL2NOIR', quantite_livree=40.0, prix_net=14.5, montant=580.0),
    ]
