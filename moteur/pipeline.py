"""
Pipeline de génération d'un comparatif — commun à main.py (ligne de
commande) et gui.py (interface graphique).

Ne dépend d'aucune bibliothèque d'interface : toute la sortie passe par
print() comme le reste du moteur ; c'est à l'appelant de rediriger
sys.stdout s'il veut capturer le journal ailleurs qu'au terminal (voir
gui.py, qui le redirige vers une file d'attente pour l'afficher dans sa
fenêtre).
"""

from pathlib import Path

from moteur.besoin import lire_fichier_besoin, nom_chantier
from moteur.lecture_pdf import analyser_devis
from moteur.excel import exporter_comparatif
from moteur.base import BaseArticles
from moteur.referentiel import Referentiel
from moteur.journal import capturer_journal


def generer_comparatif(dossier_projet, fichier_besoin, dossier_devis, dossier_resultats=None):
    """
    Génère le comparatif d'un chantier.

    - dossier_projet   : racine du projet (contient base/, moteur/,
                        referentiel/) — jamais le dossier de consultation.
    - fichier_besoin   : chemin du fichier besoin (.xlsx ou .txt), n'importe
                        où sur le disque, ou None si aucun besoin n'est
                        disponible (comparaison de toutes les offres, sans
                        nom de chantier -> "Comparatif.xlsx").
    - dossier_devis    : dossier contenant les PDF de devis à comparer.
    - dossier_resultats : dossier où écrire le Comparatif généré ; par
                        défaut dossier_projet/"resultats" (compatibilité
                        historique). Pour une consultation, passer son
                        sous-dossier resultats/ (voir moteur/consultation.py).

    Le nom du chantier (donc du comparatif généré) est déduit du nom du
    fichier besoin. Retourne le chemin du comparatif généré, ou None si
    aucun article n'a pu être extrait des devis.
    """

    dossier_projet = Path(dossier_projet)
    dossier_resultats = Path(dossier_resultats) if dossier_resultats else dossier_projet / "resultats"

    with capturer_journal(dossier_resultats, prefixe="comparatif"):

        base = BaseArticles(dossier_projet / "base")
        base.importer_bdd(dossier_projet / "base" / "BDD_articles.csv")
        base.importer_familles(dossier_projet / "base" / "familles.csv")

        dossier_referentiel = dossier_projet / "referentiel"
        referentiel = Referentiel(dossier_projet / "moteur")
        referentiel.importer_bdd(dossier_projet / "base" / "BDD_articles.csv")
        referentiel.importer_composes(dossier_referentiel / "composes.csv")
        referentiel.appliquer_confirmations(dossier_referentiel / "A_confirmer.xlsx")

        if fichier_besoin:
            fichier_besoin = Path(fichier_besoin)
            chantier = nom_chantier(fichier_besoin)
            besoin = lire_fichier_besoin(fichier_besoin)
        else:
            chantier = ""
            besoin = []

        besoin_avec_ref = [b for b in besoin if b.reference.strip()]

        if not besoin:
            print("Aucune ligne de besoin lue : comparaison de toutes les lignes des devis.")
        elif not besoin_avec_ref:
            print("Besoin sans références : comparaison de toutes les offres.")
            print("  -> Ajoute une référence par ligne pour un rapprochement ligne à ligne.")

        articles = analyser_devis(dossier_devis)

        comparatif = None

        if not articles:
            print(f"\nAucun article extrait. Vérifier le dossier {dossier_devis}.")
        else:
            base.enrichir_depuis_devis(articles)

            if besoin and not besoin_avec_ref:
                comparatif = exporter_comparatif(
                    articles, dossier_projet, None, base,
                    besoin_brut=besoin, nom_chantier=chantier, referentiel=referentiel,
                    dossier_resultats=dossier_resultats,
                )
            else:
                comparatif = exporter_comparatif(
                    articles, dossier_projet, besoin, base, nom_chantier=chantier,
                    referentiel=referentiel, dossier_resultats=dossier_resultats,
                )

        referentiel.exporter_apprentissage(dossier_referentiel)
        referentiel.ecrire_a_confirmer(dossier_referentiel)
        referentiel.fermer()

        base.fermer()

    return comparatif
