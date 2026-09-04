import re

import fitz

from moteur.modele import Article
from moteur.outils import to_float as _f, chercher_devis
from moteur.fournisseurs._gabarit import scan_regex
from moteur import grille
from moteur.ocr import regrouper_lignes
from moteur.rapprochement.modele_bl import BonLivraison, LigneBL
from moteur.rapprochement.modele_facture import Facture, LigneFacture

# --- GABARIT (Coredime) ---------------------------------------------------
# Ligne type (colonnes à espaces variables, espace possible en tête) :
#     LEG411651   DISPO DX3-ID 2P 63A A 30MA TGA   15  U*            42,0000  630,00 6
#     SCHDZ5CA162 DISPO EMBOUT REP C.16MM2        100  40%  0,7600    0,4560   45,60 6
#
# - DISPO / AEC en tête de désignation = disponibilité
# - Remise éventuelle (40%) suivie du prix de base, avant le prix net
# - Unités : U* M* B* UN (quirk d'origine : seuls U*/M* sont normalisés en
#   UN/MT ; B* reste "B" tel quel — comportement préservé, pas "corrigé")
# - Les lignes ECOTAXE (ECO-...) sont ignorées
# - Garde-fou existant : la ligne est écartée (silencieusement) si
#   qté×prix net s'écarte de plus de 5 % (ou 1€) du montant lu — redondant
#   depuis l'ajout de l'autocontrôle global (moteur/autocontrole.py), mais
#   laissé tel quel : le retirer changerait quelles lignes sont gardées.
MOTIF_DEVIS = r"COR\s+B\d+"
MOTIF_LIGNE = re.compile(
    r"^\s*([A-Z][A-Z0-9]+)\s+"                  # référence
    r"(.+?)\s+"                                  # désignation (avec DISPO/AEC)
    r"(?:\*{3}\s+)?"                              # *** éventuel (ecotaxe)
    r"([\d]+)\s+"                                 # quantité
    r"(?:([UMB]\*|UN|BTE)\s+)?"                   # unité (parfois absente)
    r"(?:(\d+)%\s+([\d,]+)\s+)?"                  # remise % + prix base (facultatif)
    r"([\d,]+)\s+"                                # prix net
    r"([\d\s,]+?)\s+"                             # montant
    r"(\d)\s*$"                                   # code TVA
)
# --- fin GABARIT -----------------------------------------------------------


def parse_coredime(texte: str):

    articles = []

    devis = chercher_devis(texte, MOTIF_DEVIS)

    for _i, m in scan_regex(texte.splitlines(), MOTIF_LIGNE):

        ref = m.group(1)

        if ref.startswith("ECO"):
            continue

        designation = m.group(2).strip()
        quantite = float(m.group(3))
        unite_brute = m.group(4) or "UN"
        unite = unite_brute[0] if unite_brute.endswith("*") else unite_brute
        prix_base = _f(m.group(6)) if m.group(6) else None
        prix_net = _f(m.group(7))
        montant = _f(m.group(8))

        dispo = ""

        if designation.startswith("DISPO "):
            dispo = "DISPO"
            designation = designation[6:]

        elif designation.startswith("AEC "):
            dispo = "AEC"
            designation = designation[4:]

        # Garde-fou : le montant doit correspondre à qté x prix net (à 5 % près)
        if quantite and prix_net and abs(montant - quantite * prix_net) > max(
            0.05 * montant, 1.0
        ):
            continue

        articles.append(
            Article(
                fournisseur="COREDIME",
                devis=devis,
                reference_fournisseur=ref,
                reference_distributeur="",
                designation=designation,
                quantite=quantite,
                unite="UN" if unite == "U" else ("MT" if unite == "M" else unite),
                prix_brut=prix_base if prix_base else prix_net,
                prix_net=prix_net,
                montant=montant,
                disponibilite=dispo,
            )
        )

    return articles


# --- GABARIT BL (Coredime) --------------------------------------------------
# BL scanné (image pure, comme 109 Distribution — voir moteur/ocr.py). La
# plupart des BL réels vus n'affichent AUCUN prix (réglé à la facture, voir
# CLAUDE.md "Ref à livrer directement") : pas de garde-fou "Total HT"
# possible ici contrairement à 109 Distribution. La quantité livrée est
# donc le seul champ vraiment critique à bien extraire ; l'OCR abîme
# parfois la cellule Quantité elle-même (coche "livré" imprimée dedans,
# ex. "300" lu "3007" — cas réel), mais le document réimprime souvent la
# même quantité un peu plus loin sur la ligne dans une mention "<qté> X 1
# unite" (ex. "39% 300x1unite"), utilisée en repli/confirmation quand
# présente et plus fiable que la cellule Quantité elle-même dans ce cas.
#
# IMPORTANT (signalé par l'acheteur — PAS encore rencontré sur un vrai PDF
# à ce jour, implémenté par prudence sur sa description directe des BL
# papier, À VALIDER dès qu'un cas réel se présente) : Coredime livre
# souvent en plusieurs fois. Un même BL peut lister à la fois les articles
# LIVRÉS et les articles NON livrés, ces derniers sous une mention
# "*Reste à livrer" — tout ce qui suit cette mention est EXCLU de la
# quantité livrée (jamais compté comme livré). À ne pas confondre avec
# "Ref à livrer directement" (déjà vu sur un vrai PDF,
# tests/fixtures/bl_coredime_4.pdf) : ces articles-là SONT bien livrés,
# seul le prix est différé à la facture.
MOTIF_COMMANDE_COREDIME = re.compile(
    r"(?:\bBC|COMMANDE\s*N)\s*[°o]?\s*([A-Z]?\d{1,4}(?:[.\-]\d{1,4}){1,2})",
    re.IGNORECASE,
)
# Repli si le séparateur (point) a été perdu par l'OCR, ex. "BC123097" au
# lieu de "BC 123.097" (1 seul cas réel vu) — découpage 3+3, convention
# observée sur toutes les autres références de commande de ce projet.
MOTIF_COMMANDE_COREDIME_SANS_SEPARATEUR = re.compile(r"\bBC\s*[°o]?\s*(\d{6})\b", re.IGNORECASE)
MOTIF_BL_COREDIME = re.compile(r"COR\s*B\d+(?:\.\d+)?", re.IGNORECASE)
MOTIF_DATE_LIVRAISON_COREDIME = re.compile(r"^(\d{2})\.(\d{2})\.(\d{2})$")
MOTIF_QTE_CONFIRMATION_COREDIME = re.compile(r"([0-9oO]+)\s*[xX]?\s*1\s*unite", re.IGNORECASE)

# BUG RÉEL CORRIGÉ (nouveau lot de BL réels) : l'OCR a lu "Code aricle"
# (le "T" de "article" disparu) sur un document par ailleurs propre (une
# seule ligne, "10 X 1 unite" bien confirmée) — l'en-tête de tableau
# n'était alors JAMAIS trouvé, faisant disparaître TOUTE la ligne (0
# extraite pour un article pourtant simple à lire). "T" rendu optionnel,
# seule lettre concernée par cette corruption sur le document vu.
MOTIF_ENTETE_TABLEAU_BL_COREDIME = re.compile(r"CODEART?ICLE")
MOTIF_PIED_TABLEAU_BL_COREDIME = re.compile(r"^COLIS|BASETVA|PENALITESDERETARD")
MOTIF_RESTE_A_LIVRER_COREDIME = re.compile(r"RESTE\s*A\s*LIVRER")
# Code alphanumérique 5-15 caractères avec au moins un chiffre (ex.
# "LEG031919", "227059360", "LBCLASTD02") — pas de ponctuation, ce qui
# exclut naturellement les codes ECO-taxe ("ECO-23", tiret) et les bouts de
# texte de désignation qui débordent sur une 2e ligne (contiennent un
# point, ex. "Autoris.OMA2500081").
MOTIF_REF_ARTICLE_COREDIME = re.compile(r"^(?=[A-Z0-9]*\d)[A-Z0-9]{5,15}$")
# --- fin GABARIT BL -----------------------------------------------------------


def _sans_espaces_coredime(s: str) -> str:
    return re.sub(r"\s+", "", s.upper())


def _numero_commande_coredime(texte: str) -> str:
    """"COMMANDE N° 123.077" / "BC154027" / "BC 241728" -> "123.077" —
    factorisé ici (utilisé par le BL ET la facture, voir GABARIT FACTURE
    plus bas) pour ne pas dupliquer la même logique deux fois."""

    m = MOTIF_COMMANDE_COREDIME.search(texte)
    if m:
        return m.group(1).upper().replace(" ", ".")

    m = MOTIF_COMMANDE_COREDIME_SANS_SEPARATEUR.search(texte)
    if m:
        brut = m.group(1)
        return f"{brut[:3]}.{brut[3:]}"

    return ""


# Valeur BRUTE de "Réf.:" (voir GABARIT FACTURE) — capturée QUE
# _numero_commande_coredime() ait réussi ou non à la convertir en n° de
# commande exploitable ; sert uniquement à reconnaître un bon manuel
# "BCN 241461" en aval (moteur.rapprochement.matching_facture.
# est_bdc_manuel_24x), jamais à deviner une commande.
MOTIF_REF_BRUTE_COREDIME = re.compile(r"R[ée]f\.?\s*:\s*(.+)$", re.MULTILINE)


def _ref_brute_coredime(texte_bloc: str) -> str:
    m = MOTIF_REF_BRUTE_COREDIME.search(texte_bloc)
    return m.group(1).strip() if m else ""


def _normaliser_date_coredime(brut: str) -> str:
    """"06.08.26" -> "06/08/2026" (même format que les autres fournisseurs
    BL, voir moteur/rapprochement/pipeline_bl._parser_date_bl)."""

    m = MOTIF_DATE_LIVRAISON_COREDIME.match(brut.strip())
    if not m:
        return ""
    jour, mois, an2 = m.groups()
    return f"{jour}/{mois}/20{an2}"


def _zone_tableau_bl_coredime(lignes_groupees: list[list[dict]]) -> list[list[dict]]:
    """Lignes (groupées par ligne visuelle, mots OCR positionnés) entre
    l'en-tête de colonnes et le premier repère de pied de tableau."""

    i_entete = next(
        (i for i, ligne in enumerate(lignes_groupees)
         if any(MOTIF_ENTETE_TABLEAU_BL_COREDIME.search(_sans_espaces_coredime(m["texte"])) for m in ligne)),
        None,
    )
    if i_entete is None:
        return []

    i_pied = next(
        (i for i, ligne in enumerate(lignes_groupees)
         if i > i_entete and any(
             MOTIF_PIED_TABLEAU_BL_COREDIME.search(_sans_espaces_coredime(m["texte"])) for m in ligne
         )),
        None,
    )

    return lignes_groupees[i_entete + 1:(i_pied if i_pied is not None else len(lignes_groupees))]


def _quantite_ligne_bl_coredime(cellules: list[str]) -> float:
    """Mention "<qté> X 1 unite" en priorité (voir bandeau), sinon la
    cellule Quantité (index 1, juste après la désignation) en repli."""

    # "1" mal lu par l'OCR en tant que lettre isolée ("l", "I", ou même
    # "i" minuscule — cas réel : "i x 1 unite") : remplacé seulement en mot
    # entier pour ne pas abîmer "unite" (le "i" qu'il contient n'est PAS
    # un mot isolé, \b ne matche pas dedans). Même famille, cas réel
    # supplémentaire : le "1" DE TÊTE d'un nombre à 2 chiffres est parfois
    # lu "l" collé au chiffre suivant ("lo x 1 unite" pour "10 x 1 unite")
    # — remplacé seulement en DÉBUT de mot, juste avant un chiffre ou un
    # "o"/"O" (lui-même déjà toléré comme "0" juste après), pour ne pas
    # abîmer un mot commençant réellement par "l" (ex. "livraison").
    texte_confirmation = " ".join(cellules[2:])
    texte_confirmation = re.sub(r"\b[lIi]\b", "1", texte_confirmation)
    texte_confirmation = re.sub(r"\bl(?=[oO0-9])", "1", texte_confirmation)
    m = MOTIF_QTE_CONFIRMATION_COREDIME.search(texte_confirmation)
    if m:
        brut = m.group(1).replace("o", "0").replace("O", "0")
        if brut.isdigit():
            return float(brut)

    if len(cellules) > 2:
        brut = cellules[2].replace("o", "0").replace("O", "0")
        m2 = re.match(r"^(\d+)", brut)
        if m2:
            return float(m2.group(1))

    return 0.0


def _ligne_bl_vers_article_coredime(cellules: list[str]) -> LigneBL | None:

    if len(cellules) < 2:
        return None

    reference = cellules[0].strip()

    if not MOTIF_REF_ARTICLE_COREDIME.match(reference):
        return None

    quantite = _quantite_ligne_bl_coredime(cellules)

    if not quantite:
        return None

    return LigneBL(
        reference_fournisseur=reference,
        designation=cellules[1].strip(),
        quantite_livree=quantite,
        prix_net=None,
        montant=None,
    )


def parse_bl_coredime(mots_par_page: list[list[dict]]) -> BonLivraison:

    lignes_plates = [
        mot["texte"]
        for mots in mots_par_page
        for ligne in regrouper_lignes(mots)
        for mot in ligne
    ]
    texte = "\n".join(lignes_plates)

    # _numero_commande_coredime() cherche sur le texte ORIGINAL (pas
    # compacté) : _sans_espaces_coredime collerait "Référence" et "BC"
    # ensemble et détruirait la frontière de mot ("\bBC") dont son motif
    # de repli dépend (bug réel rencontré ici).
    numero_commande = _numero_commande_coredime(texte)

    numero_bl = ""
    m = MOTIF_BL_COREDIME.search(texte)
    if m:
        numero_bl = m.group(0).upper().replace(" ", "")

    date_bl = ""
    for i, ligne in enumerate(lignes_plates):
        if _sans_espaces_coredime(ligne).startswith("LIVRAISONDU") and i + 1 < len(lignes_plates):
            date_bl = _normaliser_date_coredime(lignes_plates[i + 1])
            break

    articles = []
    exclues_reste_a_livrer = 0

    for mots in mots_par_page:

        # BUG RÉEL CORRIGÉ (1er cas réel rencontré, commande M3.23.043,
        # BL sur 2 pages) : `reste_a_livrer` était déclaré HORS de cette
        # boucle (une seule fois pour tout le document) — une fois activé
        # par un "*Reste à livrer" en page 1, il restait actif pour
        # TOUJOURS, excluant à tort des lignes de la page 2 pourtant
        # livrées (chacune avec sa propre confirmation "<qté> x 1 unite").
        # Structure réelle observée : Coredime réimprime en page 2, avec
        # confirmation de livraison, exactement les articles que la page 1
        # avait listés comme "reste à livrer" — chaque PAGE a son propre
        # statut, réinitialisé ici pour ne plus jamais contaminer les
        # pages suivantes.
        reste_a_livrer = False

        lignes_zone = _zone_tableau_bl_coredime(regrouper_lignes(mots))
        i = 0

        while i < len(lignes_zone):

            cellules = [m["texte"] for m in lignes_zone[i]]
            i += 1

            if MOTIF_RESTE_A_LIVRER_COREDIME.search(_sans_espaces_coredime(" ".join(cellules))):
                reste_a_livrer = True
                continue

            # Désignation sur 2 lignes (cas réel : "DL MOSAIC 2 CD FLASH
            # ROUGE" puis "Autoris. OMA2500081" avant la quantité) : si la
            # ligne courante n'a pas de quantité exploitable et que la
            # ligne suivante ne commence pas par une référence article,
            # c'est la suite de la même ligne — on raccorde ses cellules
            # utiles (tout sauf sa 1ère, qui n'est que du texte).
            if (
                _quantite_ligne_bl_coredime(cellules) == 0
                and i < len(lignes_zone)
            ):
                cellules_suivantes = [m["texte"] for m in lignes_zone[i]]
                premiere_suivante = cellules_suivantes[0].strip() if cellules_suivantes else ""
                if (
                    premiere_suivante
                    and not MOTIF_REF_ARTICLE_COREDIME.match(premiere_suivante)
                    and not MOTIF_RESTE_A_LIVRER_COREDIME.search(_sans_espaces_coredime(premiere_suivante))
                ):
                    cellules = cellules + cellules_suivantes[1:]
                    i += 1

            article = _ligne_bl_vers_article_coredime(cellules)

            if article is None:
                continue

            if reste_a_livrer:
                exclues_reste_a_livrer += 1
                continue

            articles.append(article)

    if exclues_reste_a_livrer:
        print(
            f"!! COREDIME (BL) : {exclues_reste_a_livrer} ligne(s) sous "
            f"« Reste à livrer » exclue(s) de la quantité livrée (voir bandeau GABARIT BL)."
        )

    return BonLivraison(
        fournisseur="COREDIME",
        fichier="",
        numero_bl=numero_bl,
        date_bl=date_bl,
        numero_commande=numero_commande,
        lignes=articles,
        total_ht_affiche=None,
    )


# --- GABARIT FACTURE (Coredime) ---------------------------------------------
# Facture en texte PDF natif (jamais de scan, comme le devis du même
# fournisseur) — mais moteur.lecture_pdf.lire_pdf() concatène TOUTES les
# pages du PDF, qui assemble souvent PLUSIEURS documents à la suite dans
# le MÊME fichier (Facture, puis CGV, puis parfois un Bon de Livraison
# RE-SCANNÉ en image et une copie de notre propre Bon de Commande, vus
# sur les pièces réelles de Prisca LEBLÉ, session F4) : l'extraction est
# bornée à la PAGE 0 (la Facture elle-même) — jamais un scan du texte
# entier, même principe que 109 Distribution.
#
# BUG RÉEL CORRIGÉ (trouvé en confrontant le parser aux 70 pièces
# réelles) : un 1er essai bornait sur "----- IMPORTANT -----" (repère de
# CONTENU, pas de page) — mais l'ordre d'extraction PyMuPDF de ce gabarit
# est si scramblé qu'une VRAIE ligne d'article peut apparaître APRÈS ce
# repère, toujours au sein de la page 0 (ex. "LEG040580 ... 64,92 1" sur
# 6107304.pdf, listée après "----- IMPORTANT -----" mais avant la fin
# réelle de la page 0) — cette borne perdait silencieusement des lignes
# entières sur 38/69 factures réelles testées. Corrigé en bornant sur la
# VRAIE limite de page : chaque page recommence par un bloc identifiable
# "##ESIGUID;...\n##NM#FR;...\n#####DEMAT-FJ;...;<Facture|Avoir>;...\n"
# (métadonnées machine-lisible du système de dématérialisation, présentes
# sur CHAQUE page, y compris la page scannée en image) — la zone
# d'extraction va du 1er marqueur "* BON D'EXPEDITION" jusqu'au début de
# la 2e occurrence de ce bloc (= début de la page 1, la CGV), jamais un
# repère de contenu qui peut apparaître n'importe où dans le flux.
#
# Type de document (Facture/Avoir) lu sur la ligne de métadonnées
# machine-lisible imprimée en tête de CHAQUE page par le système de
# dématérialisation ("#####DEMAT-FJ;...;Facture;..." ou "...;Avoir;...")
# — plus fiable que le texte visible, confirmé par le préfixe visible
# "COR F<num>"/"COR A<num>" en repli. AUCUNE ligne d'un AVOIR n'est
# extraite en détail (format numérique différent : quantité et montant
# avec un "-" COLLÉ en fin de nombre, ex. "1,00-U*"/"98,00-6", jamais
# rencontré ailleurs dans ce projet) : sans intérêt réel, un AVOIR n'est
# de toute façon jamais rapproché automatiquement (voir
# moteur.rapprochement.pipeline_facture) — 1er exemple réel de ce projet.
#
# Une facture peut citer PLUSIEURS "BON D'EXPEDITION" (= plusieurs BL),
# chacun potentiellement sur une commande différente (même principe que
# 109 Distribution) — zone bornée PAR BLOC, entre un marqueur et le
# suivant (ou la fin de la zone facture). Facture.numeros_commande
# regroupe les références DISTINCTES trouvées bloc par bloc : si un seul
# candidat distinct, moteur.rapprochement.pipeline_facture l'applique à
# tous les blocs (cas le plus courant, vérifié sur les pièces réelles) ;
# si plusieurs, chaque bloc retombe sur la déduction par contenu déjà
# généraliste — aucune modification du pipeline nécessaire ici.
#
# Total HT affiché retrouvé via un repère fiable trouvé sur les pièces
# réelles : la ligne "COR F<num>"/"COR A<num>" ISOLÉE (rien d'autre sur
# sa propre ligne, contrairement à celle d'en-tête qui porte aussi le
# code client) est TOUJOURS précédée de 3 valeurs non vides, dans cet
# ordre : Total HT, Total TVA, Total TTC (la case "Base TVA" intercalée
# entre les deux premières ressort vide à cette position précise, sa
# vraie valeur — quand elle existe — s'imprime ailleurs dans le flux).
#
# Quantité tantôt entière ("500"), tantôt à 2 décimales ("10,00") sur le
# MÊME gabarit — les deux formes acceptées. Colonne "Un" (U*/M*) ABSENTE
# quand une remise est appliquée (remise% + prix de base à la place,
# vérifié sur plusieurs documents réels) — les deux formes traitées
# comme des variantes mutuellement exclusives de la même ligne, jamais
# les deux à la fois (jamais observé).
#
# LIMITE CONNUE, NON corrigée (un seul exemple à ce jour, règle d'or) :
# une "remise double" ("Remise 35,00+31,00%") peut être imprimée sur une
# ligne totalement DÉCONNECTÉE de sa référence/désignation dans le flux
# PyMuPDF linéaire (positionnées à des endroits complètement différents
# du texte, pas seulement une ligne d'écart) — cette ligne n'est alors
# PAS extraite. L'autocontrôle Total HT ci-dessous le signale
# honnêtement plutôt que de deviner un rattachement (vu sur
# tests/fixtures/facture_coredime_5_remise_double_dispersee.pdf).
MOTIF_TYPE_DOCUMENT_COREDIME_FACTURE = re.compile(
    r"#####DEMAT-FJ;[^;]*;[^;]*;[^;]*;[^;]*;(Facture|Avoir);", re.IGNORECASE,
)
MOTIF_NUMERO_FACTURE_COREDIME = re.compile(r"^COR\s+[FA](\d+)\s*$", re.MULTILINE)
MOTIF_DATE_FACTURE_COREDIME = re.compile(
    r"#####DEMAT-FJ;[^;]*;[^;]*;[^;]*;[^;]*;(?:Facture|Avoir);(\d{4})-(\d{2})-(\d{2});",
)
MOTIF_ANCRE_TOTAUX_COREDIME = re.compile(r"^COR\s+[FA]\d+\s*$")
MOTIF_BLOC_FACTURE_COREDIME = re.compile(
    r"\*\s*(?:BON D'EXPEDITION|Bon de Facturation)\s*N?°?\s*(B\d+(?:\.\d+)?)", re.IGNORECASE,
)
# Repère de fin de zone facture (BUG RÉEL CORRIGÉ, gros lot Coredime H1
# 2026, ~48/328 factures avec un écart de Total HT — parfois plus de 80%
# du montant manquant) : remplace l'ancienne heuristique "2e occurrence du
# marqueur de page ##ESIGUID", qui supposait à tort qu'une facture
# Coredime tient sur UNE SEULE page de contenu suivie d'UNE SEULE page
# annexe — une facture à beaucoup de lignes s'étale en réalité sur
# PLUSIEURS folios (pages) DE CONTENU, chacun avec son propre repère
# "##ESIGUID" répété en en-tête, faisant conclure à tort que le 2e FOLIO
# marquait la fin de la facture, tronquant silencieusement tout le reste
# (cas réel 6105181.pdf : 24 lignes réelles sur 2 folios, seules 11
# extraites, 2984,98€ manquants).
#
# Deux annexes possibles marquent la VRAIE fin de zone (exigence du
# service comptable, déjà documentée pour 109 Distribution/Electric
# Plus/EDOI — présentes sur la quasi-totalité des pièces récentes) :
# notre propre bon de commande, ou le bon de livraison de Coredime
# lui-même (celui-ci en plus, jamais à la place — cas réel 6200396.pdf).
# Absentes -> la zone s'étend jusqu'à la fin du texte, comportement sûr
# par défaut : MOTIF_LIGNE_FACTURE_COREDIME est trop strict (référence +
# désignation + quantité + prix + montant + TVA sur UNE SEULE ligne
# visuelle) pour matcher un faux positif dans les CGV ou dans une annexe
# (champs éclatés sur des lignes séparées, jamais tous réunis sur une
# seule ligne) — SAUF le tableau d'articles du bon de livraison, identique
# en tout point à celui d'une facture (d'où le besoin de l'exclure
# explicitement).
#
# LEÇON COMMUNE AUX DEUX ANNEXES, trouvée en 2 temps sur des cas réels
# différents : le premier réflexe (chercher le TITRE de l'annexe, "BON DE
# COMMANDE"/"B O N  D E  L I V R A I S O N") échoue dans les DEUX cas, car
# Coredime imprime ces titres en PIED de leur propre bloc (comme
# "F A C T U R E" en pied de la vraie facture) — donc APRÈS le contenu
# qu'il fallait justement exclure :
# - Annexe BON DE LIVRAISON (cas réel 6200396.pdf) : chercher son titre ne
#   borne rien d'utile, le tableau dupliqué (identique à la facture,
#   DOUBLANT le total extrait : 240,00€ au lieu de 120,00€) le précède.
#   Repère retenu : "COR B<num>", la référence isolée en TÊTE de ce même
#   bloc annexe (même famille que "COR F<num>" en tête de la vraie
#   facture, déjà utilisé par MOTIF_ANCRE_TOTAUX_COREDIME — seule la
#   lettre change, "B" comme "Bon de livraison" plutôt que "F" comme
#   "Facture").
# - Annexe BON DE COMMANDE (cas réel 6401314.pdf, trouvé APRÈS le
#   correctif précédent) : le titre "BON DE COMMANDE" s'imprime après le
#   tableau ET après le bloc signature "DATE/ACHETEUR/VISA/téléphone" —
#   notre propre numéro de téléphone ("0693 86 68 03") restait alors DANS
#   la zone scannée et matchait accidentellement
#   MOTIF_LIGNE_INCOMPLETE_COREDIME, comme un 2e faux candidat "remise
#   double" à côté du vrai — désamorçant l'appariement 1:1 pourtant sans
#   ambiguïté (voir _lignes_remise_double_coredime, qui exige EXACTEMENT
#   1 ligne incomplète). Repère retenu : "DESTINATAIRE", la toute
#   PREMIÈRE ligne de contenu de notre BC (avant le fournisseur, le
#   chantier, le tableau et la signature).
MOTIF_DEBUT_ANNEXE_BL_COREDIME = re.compile(r"^\s*COR\s+B\d+", re.MULTILINE)
MOTIF_DEBUT_ANNEXE_BC_COREDIME = re.compile(r"^\s*DESTINATAIRE\s*$", re.MULTILINE | re.IGNORECASE)
MOTIF_LIGNE_FACTURE_COREDIME = re.compile(
    r"^\s*([A-Z0-9][A-Z0-9\-]+)\s+"                  # référence (tiret parfois, ex. "WAG221-425" ; parfois purement numérique, ex. "227060133")
    r"(.+?)\s+"                                       # désignation
    r"(\d+(?:,\d+)?)\s+"                              # quantité (entier ou décimal)
    r"(?:([UMB]\*|UN|BTE)\s+|(?:\d+)%\s+[\d,]+\s+)?"    # unité OU remise%+prix base (mutuellement exclusifs)
    r"([\d,]+)\s+"                                    # prix net HT
    r"([\d\s,]+?)\s+"                                 # montant
    r"(\d)\s*$"                                       # code TVA
)
# Repli quand une double remise ("Remise 35,00+26,00%") est appliquée :
# tout le bloc [référence + désignation + quantité + prix base] peut être
# imprimé SANS sa fin (prix net/montant/TVA), la fin apparaissant sur une
# ligne "Remise ..." totalement disjointe ailleurs dans le flux PyMuPDF
# scramblé (vu sur plusieurs pièces réelles, ex. 6107462.pdf, 6108474.pdf,
# 6108047.pdf) — jamais rattachée par un simple regex multi-ligne.
#
# RATTACHEMENT PAR COORDONNÉES (voir moteur.grille, _apparier_par_position_
# coredime) : chaque ligne "Remise" est rattachée à la ligne incomplète
# dont le Y est immédiatement AU-DESSUS d'elle dans le document — mêmes
# rangées visuelles que "<qté> X 1 unite" (elle-même juste au-dessus),
# vérifié sur un document réel à 3 doubles remises
# (facture_coredime_6108846_remise_double_x3.pdf) : reconstitue
# EXACTEMENT les 3 bonnes paires (confirmées par cohérence arithmétique
# indépendante, qté × prix net = montant sur les 3). Remplace l'ancien
# appariement 1:1 (qui n'osait un rattachement que si le bloc ne
# contenait QU'UNE SEULE ligne incomplète ET QU'UNE SEULE ligne "Remise" —
# gardé en repli SANS coordonnées, ex. `chemin` non fourni à
# parse_facture_coredime : jamais un choix au hasard entre plusieurs).
MOTIF_LIGNE_INCOMPLETE_COREDIME = re.compile(
    r"^\s*([A-Z0-9][A-Z0-9\-]+)\s+"                  # référence (voir MOTIF_LIGNE_FACTURE_COREDIME)
    r"(.+?)\s+"                                       # désignation
    r"(\d+(?:,\d+)?)\s+"                              # quantité
    r"(?:[UMB]\*|UN|BTE)?\s*"                          # unité optionnelle
    r"([\d,]+)\s*$"                                    # prix base HT — RIEN après (sinon ce serait une ligne complète)
)
MOTIF_LIGNE_REMISE_SEULE_COREDIME = re.compile(
    r"^\s*Remise\s+([\d,]+\+[\d,]+%)\s+"              # taux double (ancre de position, groupe 1)
    r"([\d,]+)\s+"                                    # prix net HT
    r"([\d\s,]+?)\s+"                                 # montant
    r"(\d)\s*$"                                        # code TVA
)
# BUG RÉEL CORRIGÉ (session S0, 2 pièces réelles, 6107800.pdf et
# 6100226.pdf) : la ligne d'éco-taxe a DEUX prix consécutifs (brut et net,
# toujours identiques — aucune remise vue sur une éco-taxe à ce jour) là où
# une ligne d'article normale n'en affiche qu'UN SEUL, ex. réel :
#   " ECO-23               ECOTAXE                       *** 10   UN      0,08      0,08       0,80 1"
# MOTIF_LIGNE_FACTURE_COREDIME matchait quand même (design non-greedy de la
# désignation qui absorbe le "***"), mais son groupe MONTANT
# (`[\d\s,]+?`, qui tolère des espaces internes) fusionnait alors les DEUX
# derniers nombres ("0,08       0,80") en une seule chaîne que `_f()` ne
# sait pas convertir correctement -> montant=0,0 au lieu de 0,80 (10 x
# 0,08). Silencieux : le garde-fou qté x prix vs montant
# (`abs(0 - 0.8) > max(0.05*0, 1.0)` = `0.8 > 1.0` = Faux) ne se déclenche
# pas sur un si petit écart absolu. Repère fiable, jamais vu sur une ligne
# d'article normale : le "***" imprimé à la place d'une vraie quantité de
# vente précède TOUJOURS cette structure à 2 prix. Traité en amont du
# passage normal (lignes isolées avant que MOTIF_LIGNE_FACTURE_COREDIME ne
# les voie, voir parse_facture_coredime) plutôt qu'en repli après coup —
# jamais une double extraction de la même ligne.
MOTIF_LIGNE_ECOTAXE_COREDIME = re.compile(
    r"^\s*([A-Z0-9][A-Z0-9\-]+)\s+"                  # référence (ex. "ECO-23")
    r"(.+?)\s+"                                       # désignation (ex. "ECOTAXE")
    r"\*\*\*\s+"                                       # placeholder de quantité de vente
    r"(\d+(?:,\d+)?)\s+"                              # quantité réelle
    r"(?:([UMB]\*|UN|BTE)\s+)?"                        # unité
    r"[\d,]+\s+"                                       # prix brut (identique au net, ignoré)
    r"([\d,]+)\s+"                                    # prix net
    r"([\d,]+)\s+"                                    # montant
    r"(\d)\s*$"                                        # code TVA
)
# --- fin GABARIT FACTURE -----------------------------------------------------


def _position_page_coredime(lignes_avec_page: list, sous_texte: str):
    """Comme moteur.grille.position_y, mais retourne (page, y) plutôt que
    le seul Y — indispensable pour comparer des lignes de PAGES
    DIFFÉRENTES : chaque page a son propre système de coordonnées (Y
    redémarre près de 0 en haut de CHAQUE page) — un tri par Y brut
    mélangerait à tort le bas d'une page avec le haut de la suivante si
    leurs valeurs se chevauchent par coïncidence (bug réel évité en
    confrontant un document réel à 2 pages : deux lignes de pages
    différentes partageaient exactement le même Y). Comparer des tuples
    (page, y) trie naturellement PAGE D'ABORD, Y ENSUITE."""

    for page, ligne in lignes_avec_page:
        texte_ligne = " ".join(m["texte"] for m in ligne)
        if sous_texte in texte_ligne:
            return (page, sum(m["y0"] for m in ligne) / len(ligne))
    return None


def _apparier_par_position_coredime(incompletes: list, remises: list, lignes_grille_bloc: list) -> list | None:
    """Rattache chaque ligne "Remise" à la ligne d'article incomplète dont
    la position (page, Y) est immédiatement AU-DESSUS d'elle (voir
    _position_page_coredime) — gère n'importe quel nombre de remises/
    lignes incomplètes, contrairement à l'ancien appariement 1:1. Retourne
    None si UNE SEULE position est introuvable (jamais un rattachement à
    l'aveugle sur une coordonnée manquante) — la ligne concernée reste
    alors non extraite, comme avant.

    `lignes_grille_bloc` : liste de (page, ligne) — voir
    _position_page_coredime, jamais une liste de lignes nues (perdrait la
    page, seule façon fiable de les ordonner entre pages différentes)."""

    positions_articles = []
    for m in incompletes:
        pos = _position_page_coredime(lignes_grille_bloc, m.group(1))
        if pos is None:
            return None
        positions_articles.append((pos, m))

    positions_remises = []
    for m in remises:
        pos = _position_page_coredime(lignes_grille_bloc, m.group(1))  # "35,00+31,00%", unique par ligne
        if pos is None:
            return None
        positions_remises.append((pos, m))

    paires = []
    for pos_remise, m_remise in sorted(positions_remises, key=lambda t: t[0]):
        candidats = [t for t in positions_articles if t[0] < pos_remise]
        if not candidats:
            return None
        choisi = max(candidats, key=lambda t: t[0])
        paires.append((choisi[1], m_remise))
        positions_articles.remove(choisi)

    return paires


def _lignes_remise_double_coredime(texte_bloc: str, refs_deja_extraites: set,
                                    lignes_grille_bloc: list | None = None) -> list:
    """Voir MOTIF_LIGNE_INCOMPLETE_COREDIME : identifie les lignes
    incomplètes et les lignes "Remise" du bloc (en excluant les références
    déjà capturées par le passage normal, pour ne jamais compter une même
    ligne deux fois), puis les apparie :

    - `lignes_grille_bloc` fourni (voir moteur.grille, `chemin` passé à
      parse_facture_coredime) : appariement PAR POSITION
      (_apparier_par_position_coredime), gère n'importe quel nombre de
      lignes de chaque sorte.
    - `lignes_grille_bloc` absent (aucune coordonnée disponible) : repli
      sur l'ANCIEN comportement, n'ose une correspondance que dans le cas
      sans ambiguïté (1 seule ligne de chaque), jamais un choix au hasard
      entre plusieurs."""

    lignes = texte_bloc.splitlines()

    incompletes = [
        m for l in lignes
        if (m := MOTIF_LIGNE_INCOMPLETE_COREDIME.match(l)) and m.group(1) not in refs_deja_extraites
    ]
    remises = [m for l in lignes if (m := MOTIF_LIGNE_REMISE_SEULE_COREDIME.match(l))]

    if not incompletes or not remises:
        return []

    if lignes_grille_bloc is not None:
        paires = _apparier_par_position_coredime(incompletes, remises, lignes_grille_bloc)
        if paires is None:
            return []
    else:
        if len(incompletes) != 1 or len(remises) != 1:
            return []
        paires = [(incompletes[0], remises[0])]

    resultat = []
    for m_ref, m_remise in paires:

        prix_net = _f(m_remise.group(2))
        montant = _f(m_remise.group(3))
        quantite = _f(m_ref.group(3))

        if quantite and prix_net and abs(montant - quantite * prix_net) > max(0.05 * montant, 1.0):
            continue

        resultat.append(LigneFacture(
            reference_fournisseur=m_ref.group(1),
            designation=m_ref.group(2).strip(),
            quantite_facturee=quantite,
            prix_unitaire_ht=prix_net,
            montant_ht=montant,
            numero_bl="",  # renseigné par l'appelant (numero_bl_bloc)
        ))

    return resultat


def _limites_pages_coredime(doc) -> list:
    """Longueur CUMULÉE de page.get_text() pour chaque page — permet de
    retrouver, pour un intervalle de caractères dans `texte` (= la simple
    concaténation de ces mêmes textes, voir moteur.lecture_pdf.lire_pdf),
    quelle(s) page(s) physique(s) il recouvre."""

    limites = []
    cumul = 0
    for page in doc:
        cumul += len(page.get_text())
        limites.append(cumul)
    return limites


def _pages_pour_intervalle_coredime(limites_pages: list, debut: int, fin: int) -> list:
    resultat = []
    debut_page = 0
    for i, fin_page in enumerate(limites_pages):
        if fin_page > debut and debut_page < fin:
            resultat.append(i)
        debut_page = fin_page
    return resultat


def _type_document_facture_coredime(texte: str) -> str:
    m = MOTIF_TYPE_DOCUMENT_COREDIME_FACTURE.search(texte)
    if m:
        return "AVOIR" if m.group(1).upper() == "AVOIR" else "FACTURE"
    # Repli sur le préfixe visible si la ligne de métadonnées est absente
    # (jamais rencontré sur les pièces réelles, prudence défensive).
    return "AVOIR" if re.search(r"\bCOR\s+A\d+", texte) else "FACTURE"


def _date_facture_coredime(texte: str) -> str:
    m = MOTIF_DATE_FACTURE_COREDIME.search(texte)
    if not m:
        return ""
    an, mois, jour = m.groups()
    return f"{jour}/{mois}/{an}"


def _total_ht_facture_coredime(texte: str):
    """Voir bandeau GABARIT FACTURE : la ligne "COR F<num>"/"COR A<num>"
    ISOLÉE est précédée de 3 valeurs non vides (Total HT, Total TVA,
    Total TTC, dans cet ordre) — la 1ère est le Total HT recherché."""

    lignes = texte.splitlines()
    for i, ligne in enumerate(lignes):
        if MOTIF_ANCRE_TOTAUX_COREDIME.match(ligne.strip()):
            valeurs = [l.strip() for l in lignes[max(0, i - 4):i] if l.strip()]
            if len(valeurs) >= 3:
                return _f(valeurs[0])
            return None
    return None


def parse_facture_coredime(texte: str, chemin=None) -> Facture:
    """`chemin` (optionnel, chemin du PDF source) : quand fourni, active le
    rattachement PAR COORDONNÉES des lignes "Remise" multiples (voir
    _apparier_par_position_coredime, moteur.grille) — en son absence,
    repli sur l'ancien appariement 1:1 (jamais un choix au hasard, voir
    _lignes_remise_double_coredime). `lecture_facture.lire_facture()`
    fournit toujours ce chemin en production ; seuls d'éventuels appels
    directs avec juste `texte` (tests unitaires sur chaîne synthétique)
    perdent cette amélioration, sans jamais planter."""

    m_num = MOTIF_NUMERO_FACTURE_COREDIME.search(texte)
    numero_facture = m_num.group(1) if m_num else ""

    type_document = _type_document_facture_coredime(texte)
    date_facture = _date_facture_coredime(texte)

    facture_vide = lambda **kw: Facture(
        fournisseur="COREDIME", fichier="", numero_facture=numero_facture,
        date_facture=date_facture, type_document=type_document, **kw,
    )

    if type_document == "AVOIR":
        # Jamais rapprochée automatiquement (voir bandeau) — pas la peine
        # d'extraire ses lignes, format numérique différent (voir plus haut).
        return facture_vide()

    marqueurs = list(MOTIF_BLOC_FACTURE_COREDIME.finditer(texte))
    if not marqueurs:
        return facture_vide()

    # Fin de la zone facture = début de la PREMIÈRE annexe rencontrée
    # (notre propre bon de commande, ou le bon de livraison de Coredime
    # lui-même — voir bandeau), sinon la fin du texte.
    positions_fin = [
        m.start() for m in (
            MOTIF_DEBUT_ANNEXE_BC_COREDIME.search(texte),
            MOTIF_DEBUT_ANNEXE_BL_COREDIME.search(texte),
        ) if m
    ]
    fin_zone = min(positions_fin) if positions_fin else len(texte)

    # Lignes-grille par page (voir moteur.grille), calculées UNE SEULE FOIS
    # pour tout le document si `chemin` est fourni — jamais si absent
    # (aucune coordonnée disponible, comportement historique préservé).
    pages_lignes_grille = None
    limites_pages = None
    if chemin is not None:
        try:
            with fitz.open(chemin) as doc:
                pages_lignes_grille = [grille.lignes(grille.mots(p)) for p in doc]
                limites_pages = _limites_pages_coredime(doc)
        except Exception:
            pages_lignes_grille = None
            limites_pages = None

    lignes_facture = []
    commandes_vues = []
    commandes_brutes_vues = []

    for i, m_bloc in enumerate(marqueurs):

        numero_bl_bloc = m_bloc.group(1).upper()
        debut_bloc = m_bloc.end()
        fin_bloc = marqueurs[i + 1].start() if i + 1 < len(marqueurs) else fin_zone
        texte_bloc = texte[debut_bloc:fin_bloc]

        numero_commande_bloc = _numero_commande_coredime(texte_bloc)
        if numero_commande_bloc and numero_commande_bloc not in commandes_vues:
            commandes_vues.append(numero_commande_bloc)
        elif not numero_commande_bloc:
            ref_brute = _ref_brute_coredime(texte_bloc)
            if ref_brute and ref_brute not in commandes_brutes_vues:
                commandes_brutes_vues.append(ref_brute)

        refs_extraites_bloc = set()

        lignes_bloc_brutes = texte_bloc.splitlines()

        # Isole les lignes d'éco-taxe AVANT le passage normal (voir
        # MOTIF_LIGNE_ECOTAXE_COREDIME) : la ligne est retirée du texte
        # soumis à MOTIF_LIGNE_FACTURE_COREDIME (qui la matcherait aussi,
        # avec un montant faux) pour ne jamais la compter deux fois.
        indices_ecotaxe = set()
        for i2, ligne_brute in enumerate(lignes_bloc_brutes):
            m_eco = MOTIF_LIGNE_ECOTAXE_COREDIME.match(ligne_brute)
            if not m_eco:
                continue
            ref = m_eco.group(1)
            quantite = _f(m_eco.group(3))
            prix_net = _f(m_eco.group(5))
            montant = _f(m_eco.group(6))
            if quantite and prix_net and abs(montant - quantite * prix_net) > max(0.05 * montant, 1.0):
                continue
            indices_ecotaxe.add(i2)
            refs_extraites_bloc.add(ref)
            lignes_facture.append(LigneFacture(
                reference_fournisseur=ref,
                designation=m_eco.group(2).strip(),
                quantite_facturee=quantite,
                prix_unitaire_ht=prix_net,
                montant_ht=montant,
                numero_bl=numero_bl_bloc,
            ))

        lignes_bloc_sans_ecotaxe = [
            "" if i2 in indices_ecotaxe else l for i2, l in enumerate(lignes_bloc_brutes)
        ]

        for _i2, m in scan_regex(lignes_bloc_sans_ecotaxe, MOTIF_LIGNE_FACTURE_COREDIME):

            ref = m.group(1)
            designation = m.group(2).strip()
            quantite = _f(m.group(3))
            # group(4) = unité de vente (U*/M*...), absente quand une remise
            # est appliquée (voir bandeau) — jamais utilisée pour le
            # rapprochement, uniquement là pour border correctement la regex.
            prix_net = _f(m.group(5))
            montant = _f(m.group(6))

            # Même garde-fou que le devis/BL du même fournisseur : la
            # ligne est écartée si qté x prix net s'écarte trop du
            # montant lu (repère mal accroché, ex. remise double
            # dispersée — voir bandeau).
            if quantite and prix_net and abs(montant - quantite * prix_net) > max(0.05 * montant, 1.0):
                continue

            refs_extraites_bloc.add(ref)
            lignes_facture.append(LigneFacture(
                reference_fournisseur=ref,
                designation=designation,
                quantite_facturee=quantite,
                prix_unitaire_ht=prix_net,
                montant_ht=montant,
                numero_bl=numero_bl_bloc,
            ))

        lignes_grille_bloc = None
        if pages_lignes_grille is not None:
            pages_concernees = _pages_pour_intervalle_coredime(limites_pages, debut_bloc, fin_bloc)
            lignes_grille_bloc = [
                (p, ligne) for p in pages_concernees for ligne in pages_lignes_grille[p]
            ]

        for ligne_remise in _lignes_remise_double_coredime(texte_bloc, refs_extraites_bloc, lignes_grille_bloc):
            ligne_remise.numero_bl = numero_bl_bloc
            lignes_facture.append(ligne_remise)

    total_ht_affiche = _total_ht_facture_coredime(texte)

    if total_ht_affiche is not None:
        total_extrait = round(sum(l.montant_ht for l in lignes_facture), 2)
        if abs(total_ht_affiche - total_extrait) > 0.02:
            print(
                f"!! COREDIME (Facture) : Total HT du PDF ({total_ht_affiche:.2f}€) "
                f"!= somme des lignes extraites ({total_extrait:.2f}€) "
                f"— une ligne a peut-être été oubliée ou mal lue (voir bandeau GABARIT FACTURE)."
            )

    return facture_vide(
        numeros_commande=commandes_vues,
        numeros_commande_bruts=commandes_brutes_vues,
        numeros_bl=[m.group(1).upper() for m in marqueurs],
        lignes=lignes_facture,
        total_ht_affiche=total_ht_affiche,
    )


# Déclaration pour le chargement automatique
FOURNISSEURS = ['COREDIME']
parse = parse_coredime
parse_bl = parse_bl_coredime
parse_facture = parse_facture_coredime
