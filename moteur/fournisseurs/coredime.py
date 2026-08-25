import re

from moteur.modele import Article
from moteur.outils import to_float as _f, chercher_devis
from moteur.fournisseurs._gabarit import scan_regex
from moteur.ocr import regrouper_lignes
from moteur.rapprochement.modele_bl import BonLivraison, LigneBL

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

    numero_commande = ""
    m = MOTIF_COMMANDE_COREDIME.search(texte)
    if m:
        numero_commande = m.group(1).upper().replace(" ", ".")
    else:
        # Repli sur le texte ORIGINAL (pas compacté) : _sans_espaces_coredime
        # collerait "Référence" et "BC" ensemble et détruirait la frontière
        # de mot ("\bBC") dont ce motif dépend (bug réel rencontré ici).
        m = MOTIF_COMMANDE_COREDIME_SANS_SEPARATEUR.search(texte)
        if m:
            brut = m.group(1)
            numero_commande = f"{brut[:3]}.{brut[3:]}"

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


# Déclaration pour le chargement automatique
FOURNISSEURS = ['COREDIME']
parse = parse_coredime
parse_bl = parse_bl_coredime
