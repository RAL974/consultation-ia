"""
Lecture du fichier besoin (.txt / .xlsx).

Couvre en particulier la régression trouvée sur un vrai fichier de
production (besoins/Besoin Doujani.txt) : une ligne d'en-têtes collée
depuis Excel (tabulations, "Désignation / Référence / Qté") en tête d'un
.txt était lue comme une ligne de besoin fantôme (qté 0, sans référence).
"""

from moteur.besoin import _lire_txt, _est_ligne_entete, lire_fichier_besoin


def test_ligne_entete_detectee():
    assert _est_ligne_entete("Désignation\tRéférence\tQté")
    assert _est_ligne_entete("Demande d'origine;Référence;Quantité")
    assert not _est_ligne_entete("Coffret XL³ 160 classe II 2 rangées;401802;9")
    assert not _est_ligne_entete("10 interrupteurs double allumage")


def test_lire_txt_ignore_entete_collee_avec_tabulations(tmp_path):
    # Reproduit le format réel de besoins/Besoin Doujani.txt : en-tête
    # tabulée, puis des lignes point-virgule dont chaque cellule garde son
    # propre séparateur de fin (collage depuis Excel).
    fichier = tmp_path / "Besoin Test.txt"
    fichier.write_text(
        "Désignation\tRéférence\tQté\n"
        "Coffret XL³ 160 classe II 2 rangées;\t401802;\t9\n"
        "Bornes de raccordement pour peigne universel;\t404905;\t6\n",
        encoding="utf-8",
    )

    besoin = _lire_txt(fichier)

    assert len(besoin) == 2
    assert besoin[0].designation == "Coffret XL³ 160 classe II 2 rangées"
    assert besoin[0].reference == "401802"
    assert besoin[0].quantite == 9

    # lire_fichier_besoin (point d'entrée public) doit voir la même chose
    assert len(lire_fichier_besoin(fichier)) == 2


def test_lire_txt_sans_entete_toujours_ok(tmp_path):
    # Ancien format, sans ligne d'en-têtes : rien ne doit être perdu.
    fichier = tmp_path / "Besoin Test.txt"
    fichier.write_text(
        "100 douilles pour DCL encastrée ; 710100 ; 100\n"
        "15 DX3-ID 2P 63A type A 30mA ; 411651 ; 15\n",
        encoding="utf-8",
    )

    besoin = _lire_txt(fichier)

    assert len(besoin) == 2
    assert besoin[0].reference == "710100"
    assert besoin[0].quantite == 100
