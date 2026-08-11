"""
Lecture des PDF du dossier devis/ et extraction des articles.
"""

import fitz
from pathlib import Path

from moteur.detecteur import detecter_fournisseur
from moteur.parsers import parser_pdf
from moteur.autocontrole import controler_articles, imprimer_anomalies


def lire_pdf(pdf: Path) -> str:

    texte = ""

    with fitz.open(pdf) as doc:

        for page in doc:
            texte += page.get_text()

    return texte


def analyser_devis(dossier) -> list:
    """
    Lit tous les PDF du dossier et retourne la liste des articles.

    Tolérant aux pannes : un PDF illisible (corrompu, protégé par mot de
    passe...) ou une erreur de parser sur UN devis ne doit jamais empêcher
    le traitement des autres. Chaque anomalie est journalisée avec une
    raison en clair, et un rapport de synthèse (X PDF lus, Y lignes
    extraites, Z anomalies) est imprimé en fin de lecture.
    """

    dossier = Path(dossier)

    pdfs = sorted(dossier.glob("*.pdf"))

    print(f"\n{len(pdfs)} PDF trouvé(s)\n")

    tous_les_articles = []
    anomalies = []
    pdfs_lus = 0

    for pdf in pdfs:

        print("-" * 60)
        print(f"Lecture : {pdf.name}")

        try:
            texte = lire_pdf(pdf)
        except Exception as e:
            print(f"!! PDF illisible, ignoré : {e}")
            anomalies.append((pdf.name, f"PDF illisible ({e})"))
            continue

        try:
            fournisseur = detecter_fournisseur(texte)

            print(f"Fournisseur : {fournisseur}")

            if fournisseur == "INCONNU":
                print("!! Fournisseur non reconnu, PDF ignoré.")
                anomalies.append((pdf.name, "Fournisseur non reconnu"))
                continue

            articles = parser_pdf(fournisseur, texte)

            print(f"Articles extraits : {len(articles)}")

            if len(articles) == 0:
                print("!! Aucun article extrait : parser manquant ou format inattendu.")
                anomalies.append((
                    pdf.name,
                    f"Fournisseur {fournisseur} reconnu mais aucun article extrait",
                ))

            imprimer_anomalies(controler_articles(articles))

            tous_les_articles.extend(articles)
            pdfs_lus += 1

        except Exception as e:
            print(f"!! Erreur lors de l'analyse, PDF ignoré : {e}")
            anomalies.append((pdf.name, f"Erreur d'analyse ({e})"))
            continue

    print("=" * 60)
    print("Résumé de la lecture des devis")
    print("=" * 60)
    print(f"{pdfs_lus} PDF lu(s) sur {len(pdfs)}, {len(tous_les_articles)} ligne(s) extraite(s) au total.")
    if anomalies:
        print(f"{len(anomalies)} anomalie(s) :")
        for nom, raison in anomalies:
            print(f"  - {nom} : {raison}")
    else:
        print("Aucune anomalie.")
    print("=" * 60)

    return tous_les_articles
