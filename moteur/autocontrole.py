"""
Autocontrôle des lignes extraites — le meilleur détecteur d'erreur de
parsing qui existe : pour chaque article, quantité × prix_net doit
correspondre au montant lu sur le devis.

Additif et non intrusif : une ligne suspecte n'est PAS retirée, seulement
signalée (à l'acheteur, PDF ouvert à côté, ou à une prochaine session de
correction du parser). Ne change donc aucune sortie existante des parsers.
"""

from dataclasses import dataclass

from moteur.modele import Article

TOLERANCE_EUROS = 0.02  # quelques centimes


@dataclass
class Anomalie:
    article: Article
    attendu: float
    ecart: float


def controler_articles(articles: list[Article], tolerance: float = TOLERANCE_EUROS) -> list[Anomalie]:
    """
    Retourne les articles pour lesquels quantité × prix_net s'écarte du
    montant lu de plus que `tolerance` (en €). Les lignes sans quantité ou
    sans prix net (0) ne peuvent pas être contrôlées : ignorées.
    """
    anomalies = []

    for article in articles:

        if not article.quantite or not article.prix_net:
            continue

        attendu = round(article.quantite * article.prix_net, 2)
        ecart = round(abs(attendu - article.montant), 2)

        if ecart > tolerance:
            anomalies.append(Anomalie(article=article, attendu=attendu, ecart=ecart))

    return anomalies


def imprimer_anomalies(anomalies: list[Anomalie], log=print) -> None:
    for anomalie in anomalies:
        a = anomalie.article
        log(
            f"!! Ligne suspecte [{a.fournisseur} {a.reference_fournisseur}] : "
            f"qté×prix net = {anomalie.attendu:.2f}€ mais montant lu = "
            f"{a.montant:.2f}€ (écart {anomalie.ecart:.2f}€) — à vérifier sur le PDF."
        )
