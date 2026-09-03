"""
Lecture des factures déposées dans a_traiter/Factures/ : texte PDF natif
(moteur/lecture_pdf.lire_pdf, réutilisé tel quel — voir CLAUDE.md, "Volet 2" :
79/79 vraies factures 109 Distribution en texte natif, JAMAIS de scan chez ce
fournisseur, contrairement à ses BL) + détection fournisseur
(moteur/detecteur.py, réutilisé tel quel) + parser facture dédié
(moteur/rapprochement/parsers_facture.py) — même principe que
moteur/rapprochement/lecture_bl.py pour les BL.

**Repli OCR (session F4/Stand64+Electric Plus)** : si le PDF n'a AUCUN texte
natif (scan pur, `lire_pdf()` renvoie une chaîne vide), repli générique sur
l'OCR (moteur/ocr.py, même mécanique que lecture_bl.py) — nécessaire pour
Electric Plus/GMR (sa facture fait aussi office de BL, voir
moteur.fournisseurs.electricplus, déjà scannée côté BL) et potentiellement
tout futur fournisseur qui enverrait une facture scannée plutôt qu'en texte
natif (ex. les 4 factures Stand 64 "FA <numéro>.pdf" identifiées comme des
scans, voir CLAUDE.md — pas encore couvertes, aucun parse_facture_ocr écrit
pour Stand 64 à ce jour). Ce n'est PAS un cas spécial câblé sur un nom de
fournisseur : la détection se fait sur le texte OCR de la même façon que sur
le texte natif, seul le CHOIX du registre (texte vs OCR) dépend de la
présence ou non de texte natif.

Même tolérance aux pannes que lecture_bl.py : une facture illisible, d'un
fournisseur non reconnu, ou d'un fournisseur reconnu mais sans parser
facture (texte NI OCR), ne bloque JAMAIS le traitement des autres fichiers
du lot.

Pas de détection PAR PAGE ici (contrairement à lecture_bl.py) : aucune
facture réelle ne mélange plusieurs fournisseurs sur des pages différentes
à ce jour (règle d'or — à ajouter si un cas réel se présente, sur le même
modèle que lecture_bl.lire_bl)."""

from pathlib import Path

from moteur.detecteur import detecter_fournisseur
from moteur.lecture_pdf import lire_pdf
from moteur.ocr import lignes_ocr, mots_document
from moteur.rapprochement.parsers_facture import parser_facture, parser_facture_ocr

EXTENSIONS_SUPPORTEES = (".pdf",)


def _texte_page(mots_page: list[dict]) -> str:
    return "\n".join(lignes_ocr(mots_page))


def lire_facture(chemin):
    """Lit un seul fichier. Retourne (liste_de_Facture, liste_de_raisons_en_clair)
    — la liste de raisons est vide en cas de succès complet. Une LISTE de
    Facture, pas une seule : un parser fournisseur pourrait, comme côté BL,
    retourner plusieurs documents pour un même fichier (cas réel désormais :
    Electric Plus, voir moteur.fournisseurs.electricplus.parse_facture_ocr,
    même principe que son BL)."""

    chemin = Path(chemin)

    try:
        texte = lire_pdf(chemin)
    except Exception as e:
        return [], [f"PDF illisible ({e})"]

    if texte.strip():
        fournisseur = detecter_fournisseur(texte)

        if fournisseur == "INCONNU":
            return [], ["Fournisseur non reconnu"]

        resultat = parser_facture(fournisseur, texte)

        if resultat is None:
            return [], [f"Fournisseur {fournisseur} reconnu mais pas encore de parser facture"]

    else:
        # PDF sans texte natif (scan) : repli OCR, voir bandeau du module.
        try:
            mots_par_page = mots_document(chemin)
        except Exception as e:
            return [], [f"PDF illisible ({e})"]

        texte_ocr = "\n".join(_texte_page(mots) for mots in mots_par_page)
        fournisseur = detecter_fournisseur(texte_ocr)

        if fournisseur == "INCONNU":
            return [], ["Fournisseur non reconnu (OCR)"]

        resultat = parser_facture_ocr(fournisseur, mots_par_page)

        if resultat is None:
            return [], [f"Fournisseur {fournisseur} reconnu (OCR) mais pas encore de parser facture"]

    factures = resultat if isinstance(resultat, list) else [resultat]

    for f in factures:
        f.fichier = chemin.name

    return factures, []


def analyser_dossier(dossier):
    """Lit toutes les factures de `dossier`. Retourne (factures, anomalies)
    — anomalies : liste de (nom_fichier, raison_en_clair)."""

    dossier = Path(dossier)

    if not dossier.is_dir():
        print(f"\n{dossier} n'existe pas encore (aucune facture déposée) — rien à lire.\n")
        return [], []

    fichiers = sorted(
        f for f in dossier.iterdir()
        if f.is_file() and f.suffix.lower() in EXTENSIONS_SUPPORTEES
    )

    print(f"\n{len(fichiers)} facture(s) trouvée(s) dans {dossier}\n")

    factures = []
    anomalies = []

    for fichier in fichiers:

        print("-" * 60)
        print(f"Lecture : {fichier.name}")

        try:
            lues, raisons = lire_facture(fichier)
        except Exception as e:
            print(f"!! Erreur de lecture, ignorée : {e}")
            anomalies.append((fichier.name, f"Erreur de lecture ({e})"))
            continue

        for raison in raisons:
            print(f"!! {raison}")
            anomalies.append((fichier.name, raison))

        for f in lues:
            print(
                f"Fournisseur : {f.fournisseur} — {len(f.lignes)} ligne(s), "
                f"facture {f.numero_facture or '?'}, BL {', '.join(f.numeros_bl) or '?'}, "
                f"commande {', '.join(f.numeros_commande) or '?'}"
            )
            if not f.lignes:
                anomalies.append((fichier.name, f"Aucune ligne extraite (facture {f.numero_facture or '?'})"))
            factures.append(f)

    print("=" * 60)
    print("Résumé de la lecture des factures")
    print("=" * 60)
    print(f"{len(factures)} facture(s) lue(s) sur {len(fichiers)}.")
    if anomalies:
        print(f"{len(anomalies)} anomalie(s) :")
        for nom, raison in anomalies:
            print(f"  - {nom} : {raison}")
    else:
        print("Aucune anomalie.")
    print("=" * 60)

    return factures, anomalies
