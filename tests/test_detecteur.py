"""Détection du fournisseur — migré depuis test_pipeline.py (racine)."""

from moteur.detecteur import detecter_fournisseur


def test_detection_faux_positifs():
    # "DEMANDE" ne doit pas déclencher le faux positif DEM (\bDEM\b)
    assert detecter_fournisseur("DEMANDE DE PRIX RAVATE") == "RAVATE"
    assert detecter_fournisseur("VOICI LA DEMANDE DU CLIENT") == "INCONNU"
    assert detecter_fournisseur("FACTURE DEM OCEAN INDIEN") == "DEM"
    # COMINTER MAYOTTE doit être reconnu avant le motif générique COMINTER
    assert detecter_fournisseur("COMINTER MAYOTTE SARL") == "COMINTER MAYOTTE"


def test_comminter_mayotte_sans_le_mot_mayotte():
    # Un vrai devis Cominter Mayotte ne contient PAS la chaîne "COMINTER
    # MAYOTTE" (voir tests/fixtures/cominter_mayotte.pdf) : reconnu via
    # l'e-mail @cominter.yt, propre à Cominter.
    assert detecter_fournisseur("E-mail. contact@cominter.yt") == "COMINTER MAYOTTE"


def test_mamoudzou_seul_pas_suffisant_edoi_aussi_base_la_bas():
    # "Mamoudzou" seul ne doit PAS déclencher Cominter Mayotte : EDOI est
    # aussi basé à Mamoudzou (contact.edoi@sonepar.fr) -> faux ami réel,
    # rencontré sur un vrai PDF EDOI.
    assert detecter_fournisseur("B.P 244 ZI KAWENI - 97600 Mamoudzou "
                                 "contact.edoi@sonepar.fr") == "EDOI"


def test_clareo_marque_dans_devis_coredime_pas_confondu():
    # "Clareo" est aussi une marque de luminaires vendue par d'autres
    # fournisseurs (ex. Coredime) : un simple mot ne suffit pas.
    texte_coredime_avec_marque_clareo = (
        "www.coredime.re Panel High Lumen CLAREO 600x600"
    )
    assert detecter_fournisseur(texte_coredime_avec_marque_clareo) == "COREDIME"
    assert detecter_fournisseur("Clareo S.A.S 11 rue Christophe Colomb") == "CLAREO"
    assert detecter_fournisseur("jessica.urso@clareolighting.com") == "CLAREO"
