from moteur.modele import Article
from moteur.autocontrole import controler_articles


def _art(qte, prix_net, montant):
    return Article(fournisseur="X", devis="D", reference_fournisseur="REF",
                    reference_distributeur="", designation="D", quantite=qte,
                    unite="UN", prix_brut=prix_net, prix_net=prix_net, montant=montant)


def test_ligne_conforme_pas_signalee():
    assert controler_articles([_art(10, 2.5, 25.0)]) == []


def test_ecart_dans_la_tolerance_pas_signale():
    # 0.02€ pile à la limite : pas signalé (strictement supérieur qui déclenche)
    assert controler_articles([_art(10, 2.5, 25.02)]) == []


def test_ecart_hors_tolerance_signale():
    anomalies = controler_articles([_art(10, 2.5, 30.0)])
    assert len(anomalies) == 1
    assert anomalies[0].attendu == 25.0
    assert anomalies[0].ecart == 5.0


def test_quantite_nulle_ignoree():
    # Rien à contrôler : pas de faux positif sur une ligne sans quantité
    assert controler_articles([_art(0, 2.5, 0.0)]) == []


def test_prix_net_nul_ignore():
    assert controler_articles([_art(10, 0.0, 0.0)]) == []


def test_plusieurs_lignes_seules_les_suspectes_ressortent():
    articles = [_art(10, 2.5, 25.0), _art(4, 1.0, 100.0)]
    anomalies = controler_articles(articles)
    assert len(anomalies) == 1
    assert anomalies[0].article is articles[1]
