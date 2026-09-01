import re

from moteur.modele import Article
from moteur.outils import to_float as _f, chercher_devis
from moteur.fournisseurs._gabarit import scan_regex
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
MOTIF_DEBUT_PAGE_COREDIME = re.compile(r"##ESIGUID;")
MOTIF_LIGNE_FACTURE_COREDIME = re.compile(
    r"^\s*([A-Z0-9][A-Z0-9\-]+)\s+"                  # référence (tiret parfois, ex. "WAG221-425" ; parfois purement numérique, ex. "227060133")
    r"(.+?)\s+"                                       # désignation
    r"(\d+(?:,\d+)?)\s+"                              # quantité (entier ou décimal)
    r"(?:([UMB]\*|UN|BTE)\s+|(?:\d+)%\s+[\d,]+\s+)?"    # unité OU remise%+prix base (mutuellement exclusifs)
    r"([\d,]+)\s+"                                    # prix net HT
    r"([\d\s,]+?)\s+"                                 # montant
    r"(\d)\s*$"                                       # code TVA
)
# Repli SÛR uniquement (voir _lignes_remise_double_coredime) : quand une
# double remise ("Remise 35,00+26,00%") est appliquée, tout le bloc
# [référence + désignation + quantité + prix base] peut être imprimé SANS
# sa fin (prix net/montant/TVA), la fin apparaissant sur une ligne "Remise
# ..." totalement disjointe ailleurs dans le flux PyMuPDF scramblé (vu sur
# plusieurs pièces réelles, ex. 6107462.pdf, 6108474.pdf, 6108047.pdf) —
# jamais rattachée par un simple regex multi-ligne, JAMAIS reconstruite
# sauf quand il n'existe QU'UNE SEULE ligne incomplète ET QU'UNE SEULE
# ligne "Remise" dans le même bloc (aucune ambiguïté possible sur qui va
# avec qui) ; sinon laissée non extraite, l'autocontrôle Total HT le
# signale honnêtement plutôt que de deviner un appariement.
MOTIF_LIGNE_INCOMPLETE_COREDIME = re.compile(
    r"^\s*([A-Z0-9][A-Z0-9\-]+)\s+"                  # référence (voir MOTIF_LIGNE_FACTURE_COREDIME)
    r"(.+?)\s+"                                       # désignation
    r"(\d+(?:,\d+)?)\s+"                              # quantité
    r"(?:[UMB]\*|UN|BTE)?\s*"                          # unité optionnelle
    r"([\d,]+)\s*$"                                    # prix base HT — RIEN après (sinon ce serait une ligne complète)
)
MOTIF_LIGNE_REMISE_SEULE_COREDIME = re.compile(
    r"^\s*Remise\s+[\d,]+\+[\d,]+%\s+"
    r"([\d,]+)\s+"                                    # prix net HT
    r"([\d\s,]+?)\s+"                                 # montant
    r"(\d)\s*$"                                        # code TVA
)
# --- fin GABARIT FACTURE -----------------------------------------------------


def _lignes_remise_double_coredime(texte_bloc: str, refs_deja_extraites: set) -> list:
    """Voir MOTIF_LIGNE_INCOMPLETE_COREDIME : n'apparie une ligne
    incomplète à une ligne "Remise" que si le bloc n'en contient QU'UNE
    SEULE de chaque, en excluant les références déjà capturées par le
    passage normal (MOTIF_LIGNE_FACTURE_COREDIME) pour ne jamais compter
    une même ligne deux fois."""

    lignes = texte_bloc.splitlines()

    incompletes = [
        m for l in lignes
        if (m := MOTIF_LIGNE_INCOMPLETE_COREDIME.match(l)) and m.group(1) not in refs_deja_extraites
    ]
    remises = [m for l in lignes if (m := MOTIF_LIGNE_REMISE_SEULE_COREDIME.match(l))]

    if len(incompletes) != 1 or len(remises) != 1:
        return []

    m_ref, m_remise = incompletes[0], remises[0]
    prix_net = _f(m_remise.group(1))
    montant = _f(m_remise.group(2))
    quantite = _f(m_ref.group(3))

    if quantite and prix_net and abs(montant - quantite * prix_net) > max(0.05 * montant, 1.0):
        return []

    return [LigneFacture(
        reference_fournisseur=m_ref.group(1),
        designation=m_ref.group(2).strip(),
        quantite_facturee=quantite,
        prix_unitaire_ht=prix_net,
        montant_ht=montant,
        numero_bl="",  # renseigné par l'appelant (numero_bl_bloc)
    )]


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


def parse_facture_coredime(texte: str) -> Facture:

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

    # Fin de la page 0 (la Facture) = début de la 2e occurrence du bloc de
    # métadonnées de page (voir bandeau) — jamais un repère de CONTENU.
    debuts_page = list(MOTIF_DEBUT_PAGE_COREDIME.finditer(texte))
    fin_zone = debuts_page[1].start() if len(debuts_page) > 1 else len(texte)

    lignes_facture = []
    commandes_vues = []

    for i, m_bloc in enumerate(marqueurs):

        numero_bl_bloc = m_bloc.group(1).upper()
        debut_bloc = m_bloc.end()
        fin_bloc = marqueurs[i + 1].start() if i + 1 < len(marqueurs) else fin_zone
        texte_bloc = texte[debut_bloc:fin_bloc]

        numero_commande_bloc = _numero_commande_coredime(texte_bloc)
        if numero_commande_bloc and numero_commande_bloc not in commandes_vues:
            commandes_vues.append(numero_commande_bloc)

        refs_extraites_bloc = set()

        for _i2, m in scan_regex(texte_bloc.splitlines(), MOTIF_LIGNE_FACTURE_COREDIME):

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

        for ligne_remise in _lignes_remise_double_coredime(texte_bloc, refs_extraites_bloc):
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
        numeros_bl=[m.group(1).upper() for m in marqueurs],
        lignes=lignes_facture,
        total_ht_affiche=total_ht_affiche,
    )


# Déclaration pour le chargement automatique
FOURNISSEURS = ['COREDIME']
parse = parse_coredime
parse_bl = parse_bl_coredime
parse_facture = parse_facture_coredime
