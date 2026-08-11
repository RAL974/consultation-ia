"""
Normalisation de prix — le mécanisme "à vérifier" (moteur/normalisation.py).

Piège métier visé (CLAUDE.md) : un prix à la boîte de 100 vs à l'unité, un
prix à la barre vs au mètre, ne se comparent qu'après conversion. Quand le
facteur de conversion est incertain, le prix est laissé tel quel et la
ligne doit être marquée "à vérifier" (fiable=False) plutôt que convertie en
silence.
"""

from moteur.normalisation import prix_normalise


def test_boite_sans_quantite_connue_marquee_a_verifier():
    # "Boîte" sans nombre d'unités lisible dans la désignation : le prix
    # est laissé tel quel (pas de division au hasard) et marqué incertain.
    prix, base, fiable = prix_normalise(45.0, "BTE", "Boîte de connecteurs RJ45", est_cable=False)
    assert base == "€/boite"
    assert fiable is False
    assert prix == 45.0


def test_boite_de_n_explicite_convertie_et_fiable():
    prix, base, fiable = prix_normalise(50.0, "BTE", "Boîte de 100 vis", est_cable=False)
    assert base == "€/u"
    assert fiable is True
    assert prix == 0.5


def test_cable_en_metre_deja_comparable():
    prix, base, fiable = prix_normalise(2.5, "MT", "R2V 3G1.5", est_cable=True)
    assert (prix, base, fiable) == (2.5, "€/m", True)


def test_cable_en_barre_avec_longueur_convertie():
    prix, base, fiable = prix_normalise(30.0, "BARRE", "Moulure 2.10m blanche", est_cable=True)
    assert base == "€/m"
    assert fiable is True
    assert round(prix, 4) == round(30.0 / 2.10, 4)


def test_cable_en_barre_sans_longueur_marque_a_verifier():
    # Aucune longueur lisible dans la désignation : jamais de conversion
    # devinée -> prix brut conservé, marqué à vérifier.
    prix, base, fiable = prix_normalise(30.0, "BARRE", "Moulure blanche sans longueur", est_cable=True)
    assert base == "€/barre"
    assert fiable is False
    assert prix == 30.0


def test_unite_simple_toujours_fiable():
    prix, base, fiable = prix_normalise(12.5, "UN", "Disjoncteur DX3 16A", est_cable=False)
    assert (prix, base, fiable) == (12.5, "€/u", True)
