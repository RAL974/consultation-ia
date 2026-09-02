"""
Parser COMINTER MAYOTTE (Cominter, Mamoudzou — contact@cominter.yt).

Entité distincte de Cominter Réunion (SIRET, adresse et e-mail différents) :
la question posée en début de session ("le format est-il identique ?") a une
réponse claire — NON. La structure de devis est différente de celle de
`moteur/fournisseurs/cominter.py` (v1/v2), d'où ce module séparé.

Écrit sur 1 vrai PDF (tests/fixtures/cominter_mayotte.pdf) : à confirmer sur
un 2e devis Mayotte si l'acheteur en dépose un (remise différente,
plusieurs pages...).

LIMITE CONNUE (voir CLAUDE.md) : sur ce PDF, le DERNIER article de la page
("L76565") a ses valeurs numériques (qté, prix, remise, montant) extraites
AVANT sa référence/désignation dans le flux de texte, au lieu d'après comme
partout ailleurs — un bloc orphelin apparaît en tête de document
("38,00"/"16,78"/"30%"/"446,35€") qui lui appartient réellement (38 x
16,78 x 70 % = 446,35€, exact). Cause probable : ordre d'extraction PyMuPDF
différent pour le dernier bloc d'une page. Un seul exemple observé -> pas
de règle générale à en tirer (règle d'or) ; l'article correspondant est
signalé "bloc incomplet" plutôt que rattaché à tort. Total du devis
(9 630,36€) donc supérieur de 446,35€ à la somme des lignes extraites —
visible en comparant au PDF, pas de contrôle automatique fiable possible
sur ce seul exemple.
"""

import re

from moteur.modele import Article
from moteur.outils import to_float, lignes_propres
from moteur.fournisseurs.cominter import (
    MOTIF_BL_FACTURE_COMINTER,
    MOTIF_COMMANDE_FACTURE_COMINTER,
    MOTIF_DATE_FACTURE_COMINTER,
    MOTIF_FIN_TABLEAU_FACTURE_COMINTER,
    MOTIF_MONTANT_LIGNE_FACTURE_COMINTER,
    MOTIF_NUMERO_FACTURE_COMINTER,
    MOTIF_QTE_PX_FACTURE_COMINTER,
    MOTIF_REF_ARTICLE_BL_COMINTER,
    MOTIF_REM_FACTURE_COMINTER,
    _CDT_FACTURE_COMINTER,
)
from moteur.rapprochement.modele_facture import Facture, LigneFacture

# --- GABARIT (Cominter Mayotte) ---------------------------------------------
# Bloc ancré sur le marqueur "Unité" (toujours seul sur sa ligne, une fois
# par article) :
#     Référence           (avant "Unité", motif ^L\d+$ ou "ZZ" — code
#                           générique pour un article non catalogué, la
#                           vraie référence se retrouve alors en tête de
#                           la désignation)
#     Désignation          (1 à 3 lignes, jusqu'à "Unité")
#     Unité                (marqueur)
#     Qté
#     Px unitaire
#     [Remise %]            (facultative : absente sur certaines lignes,
#                            alors Montant = Qté x Px unitaire directement)
#     Montant (€)
MARQUEUR = "Unité"
MOTIF_REF = re.compile(r"^L\d+$|^ZZ$")
MOTIF_MONEY = re.compile(r"^[\d\s]+,\d{2}\s*€$")
MOTIF_DEVIS = r"Devis\s*:\s*([A-Z0-9]+)"
# --- fin GABARIT -------------------------------------------------------------


def parse_cominter_mayotte(texte: str) -> list[Article]:

    articles = []

    m = re.search(MOTIF_DEVIS, texte)
    devis = m.group(1) if m else ""

    lignes = lignes_propres(texte)
    n = len(lignes)

    for i, ligne in enumerate(lignes):

        if ligne != MARQUEUR:
            continue

        # Remonte jusqu'à la référence précédente
        j = i - 1
        while j >= 0 and not MOTIF_REF.match(lignes[j]):
            j -= 1

        if j < 0:
            continue

        ref = lignes[j]
        designation = " ".join(lignes[j + 1:i]).strip()

        # Avance jusqu'au montant (€) : qté, puis 1 (px unitaire seul) ou
        # 2 (px unitaire + remise %) valeurs intermédiaires.
        k = i + 1
        if k >= n:
            continue
        qte_brute = lignes[k]

        valeurs = []
        k += 1
        while k < n and not MOTIF_MONEY.match(lignes[k]) and len(valeurs) < 2:
            valeurs.append(lignes[k])
            k += 1

        if k >= n or not MOTIF_MONEY.match(lignes[k]):
            print(f"Erreur lecture article COMINTER MAYOTTE ({ref}) : bloc incomplet")
            continue

        montant = to_float(lignes[k])
        quantite = to_float(qte_brute)
        prix_unitaire = to_float(valeurs[0]) if valeurs else 0.0

        articles.append(
            Article(
                fournisseur="COMINTER MAYOTTE",
                devis=devis,
                reference_fournisseur=ref,
                reference_distributeur="",
                designation=designation,
                quantite=quantite,
                unite="UN",
                prix_brut=prix_unitaire,
                prix_net=round(montant / quantite, 4) if quantite else 0.0,
                montant=montant,
            )
        )

    return articles


# --- GABARIT FACTURE (Cominter Mayotte) ------------------------------------
# Session F4 suite (2026-09-02), 12 pièces réelles (MFACxxxxx.pdf). Structure
# de facture PROCHE de celle de Cominter Réunion (moteur.fournisseurs.
# cominter, voir son bandeau GABARIT FACTURE — réutilisé au maximum ici :
# détection de fournisseur/date/n° de BL, ancrage sur le montant, référence
# la plupart du temps dans la zone bornée par les deux montants
# consécutifs) mais avec DEUX différences réelles constatées :
# 1. AUCUN code TVA après le montant — la référence suit DIRECTEMENT (pas
#    de ligne "1"/"2" à consommer entre les deux, contrairement à Cominter
#    Réunion). Sur les 3 pièces vues, la référence est TOUJOURS en
#    PREMIÈRE position de la zone (jamais en dernier comme certains cas
#    Réunion) — pas de recherche à rebours nécessaire ici.
# 2. AUCUN repère "Signature" avant le bloc [date, n° de BL, n° de
#    commande] (comme NF155008.pdf côté Réunion, agence Saint-Pierre) — le
#    repli déjà écrit dans _zone_articles_facture_cominter (1re ligne qui
#    ressemble à un n° de BL) sert ici de repère PRINCIPAL, pas seulement
#    de repli. Fin de zone : "NET A PAYER" (pas de paragraphe légal
#    "Article 7" chez cette agence).
# N° de commande : le "BC N°..." de l'en-tête peut être tronqué ou avoir un
# séparateur espace ("BC N° 24  3109") — l'étiquette EXPLICITE "- N° de
# Commande : ..." plus loin dans le document est PRÉFÉRÉE quand présente
# (toujours vue jusqu'ici), avec un repli sur l'en-tête sinon.
# Note parasite possible ENTRE le Cdt d'un article et le Qté du suivant
# (cas réel "VARIANTE DISPO GTL", MFAC15576.pdf) : sans risque, la
# désignation est bornée au Cdt trouvé (jamais au-delà), la note reste
# hors zone comme n'importe quel autre reliquat de fin de bloc.
MOTIF_LABEL_COMMANDE_FACTURE_MAYOTTE = re.compile(r"-\s*N°\s*de\s*Commande\s*:\s*(.+)")
# --- fin GABARIT FACTURE ----------------------------------------------------


def _zone_articles_facture_cominter_mayotte(lignes: list[str]):

    i_bl = next((i for i, l in enumerate(lignes) if MOTIF_BL_FACTURE_COMINTER.match(l.strip())), None)
    if i_bl is None:
        return None

    debut = max(i_bl - 1, 0)
    i_fin = next(
        (i for i in range(debut, len(lignes)) if MOTIF_FIN_TABLEAU_FACTURE_COMINTER.search(lignes[i])),
        len(lignes),
    )

    return debut, i_fin


def _entete_bl_facture_cominter_mayotte(lignes: list[str], debut: int, fin: int):

    numero_bl = ""
    for l in lignes[debut:min(debut + 3, fin)]:
        if MOTIF_BL_FACTURE_COMINTER.match(l.strip()):
            numero_bl = l.strip()
            break

    numero_commande = ""
    for l in lignes[debut:fin]:
        m = MOTIF_LABEL_COMMANDE_FACTURE_MAYOTTE.search(l)
        if m:
            numero_commande = re.sub(r"\s+", ".", m.group(1).strip())
            break

    if not numero_commande:
        for l in lignes[debut:min(debut + 3, fin)]:
            m = MOTIF_COMMANDE_FACTURE_COMINTER.search(l.strip())
            if m:
                numero_commande = m.group(1).upper().replace(" ", ".")
                break

    return numero_bl, numero_commande


def _lignes_facture_cominter_mayotte(lignes: list[str], debut: int, fin: int, numero_bl: str) -> list[LigneFacture]:

    blocs = []
    for i in range(debut, fin):
        m = MOTIF_MONTANT_LIGNE_FACTURE_COMINTER.match(lignes[i].strip())
        if not m:
            continue
        montant = to_float(m.group(1))

        k = i - 1
        if k >= 0 and MOTIF_REM_FACTURE_COMINTER.match(lignes[k].strip()):
            k -= 1
        if k < 1:
            continue
        if not MOTIF_QTE_PX_FACTURE_COMINTER.match(lignes[k].strip()):
            continue
        if not MOTIF_QTE_PX_FACTURE_COMINTER.match(lignes[k - 1].strip()):
            continue
        qte = to_float(lignes[k - 1].strip())

        blocs.append((i, qte, montant))

    articles = []
    for idx, (i_montant, qte, montant) in enumerate(blocs):

        j = i_montant + 1  # pas de code TVA chez Mayotte : la référence suit directement
        fin_zone = blocs[idx + 1][0] - 3 if idx + 1 < len(blocs) else fin
        # -3 : recule avant [Qté, Px unitaire, Remise?] du bloc suivant — au
        # pire une remise absente laisse une ligne de trop dans la zone,
        # sans conséquence (elle ne matche jamais MOTIF_REF_ARTICLE_BL_COMINTER
        # ni un mot-clé Cdt).

        zone = [lignes[k].strip() for k in range(j, max(fin_zone, j)) if lignes[k].strip()]
        if not zone or not qte:
            continue

        if not MOTIF_REF_ARTICLE_BL_COMINTER.match(zone[0]):
            continue
        reference = zone[0]
        reste = zone[1:]

        i_cdt = next((k for k, c in enumerate(reste) if c.lower() in _CDT_FACTURE_COMINTER), None)
        designation = " ".join(reste[:i_cdt] if i_cdt is not None else reste)

        articles.append(LigneFacture(
            reference_fournisseur=reference,
            designation=designation,
            quantite_facturee=qte,
            prix_unitaire_ht=round(montant / qte, 4),
            montant_ht=montant,
            numero_bl=numero_bl,
        ))

    return articles


def parse_facture_cominter_mayotte(texte: str) -> Facture:

    lignes = [l.rstrip() for l in texte.splitlines()]

    m_num = MOTIF_NUMERO_FACTURE_COMINTER.search(texte)
    numero_facture = m_num.group(1) if m_num else ""

    zone = _zone_articles_facture_cominter_mayotte(lignes)
    if zone is None:
        return Facture(
            fournisseur="COMINTER MAYOTTE", fichier="", numero_facture=numero_facture, date_facture="",
        )

    debut, fin = zone

    date_facture = ""
    entete_texte = "\n".join(lignes[:debut])
    m_date = MOTIF_DATE_FACTURE_COMINTER.search(entete_texte)
    if m_date:
        jour, mois, annee = m_date.groups()
        date_facture = f"{int(jour):02d}/{mois}/20{annee}"

    numero_bl, numero_commande = _entete_bl_facture_cominter_mayotte(lignes, debut, fin)

    lignes_facture = _lignes_facture_cominter_mayotte(lignes, debut, fin, numero_bl)

    return Facture(
        fournisseur="COMINTER MAYOTTE",
        fichier="",
        numero_facture=numero_facture,
        date_facture=date_facture,
        numeros_commande=[numero_commande] if numero_commande else [],
        numeros_bl=[numero_bl] if numero_bl else [],
        lignes=lignes_facture,
        total_ht_affiche=None,
    )


# Déclaration pour le chargement automatique
FOURNISSEURS = ['COMINTER MAYOTTE']
parse = parse_cominter_mayotte
parse_facture = parse_facture_cominter_mayotte
