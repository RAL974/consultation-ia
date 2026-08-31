"""
Parser 109 DISTRIBUTION.

RÉÉCRIT sur un vrai PDF (`tests/fixtures/109_distribution.pdf`) : l'ancien
gabarit supposait un ordre de champs jamais observé en réalité (0 article
extrait — voir CLAUDE.md, "points fragiles" de la session précédente).

DEUX VARIANTES réelles coexistent chez ce fournisseur (même marque, mais
109 Est/Sud/Ouest/Nord partagent apparemment plusieurs gabarits de devis) :
- "commande" (ex. tests/fixtures/109_distribution.pdf) : la référence vient
  AVANT le bloc chiffré, motif "Commande client n° ... du ...".
- "devis_bpu" (ex. devis "e.s BPU Région...") : la référence vient APRÈS,
  motif "Devis n° ... du ...". Découverte via l'autocontrôle Total HT
  (voir plus bas) : la variante "commande" y donnait 0 article, l'écart
  entre Total HT du PDF et somme extraite l'a immédiatement signalé.
Les deux sont essayées sur chaque bloc ancré sur le code TVA ; le premier
dont la référence "a la forme attendue" est retenu.
"""

import re
import unicodedata

from moteur.modele import Article
from moteur.outils import to_float, lignes_propres
from moteur.fournisseurs._gabarit import scan_ancre
from moteur.ocr import pages_par_identifiant, regrouper_lignes
from moteur.rapprochement.modele_bl import BonLivraison, LigneBL

# --- GABARIT (109 Distribution) --------------------------------------------
# Bloc ancré sur le code TVA (C0 à C9, un seul mot sur sa ligne).
MARQUEURS_TVA = tuple(f"C{d}" for d in range(10))

# Variante "commande" : Référence, Désignation, Total, Qté, TVA, P.U.Net, [code]
OFFSETS_COMMANDE = {
    "reference_fournisseur": -4,
    "designation": -3,
    "total": -2,
    "qte": -1,
    "pu_net_affiche": 1,
}
MOTIF_REF_COMMANDE = re.compile(r"^\d{6,}-")  # ex. "45001001-PRYSM-H07VU1.5B"

# Variante "devis_bpu" : Désignation, Qté, TVA, Total, P.U.Net, Référence
OFFSETS_DEVIS_BPU = {
    "designation": -2,
    "qte": -1,
    "total": 1,
    "pu_net_affiche": 2,
    "reference_fournisseur": 3,
}
# ex. "LK4288", "CR1/3G2.5", "M-160101" (tiret), "GOUJON 8X70" (espace) —
# ces deux derniers manquaient silencieusement (0 anomalie levée à tort)
# tant que le contrôle Total HT du PDF n'a pas signalé l'écart (cas réel,
# chantier Kanopée CDC). Borne haute élargie à 20 (cas réel,
# "FRN1X6G3-3G1.5 T", 16 caractères — devis ISHOP Saint-Denis, 321051) :
# ce n'était pas une 3e variante de gabarit, juste une référence plus
# longue que celles vues jusque-là, écartée à tort par la borne à 15.
# "+" ajouté à la classe de caractères (cas réel, "BTSOUT3X150+70"/
# "BTSOUT3X95+50", devis BT - Floe 321273) : 2 lignes sur 5 manquaient
# silencieusement (1 539,00€ sur 1 626,40€), signalé par l'autocontrôle
# Total HT.
MOTIF_REF_DEVIS_BPU = re.compile(r"^[A-Z0-9./+\-\s]{5,20}$")

# Variante "devis_bpu" avec colonne Rem% RENSEIGNÉE (cas réel, devis ISHOP
# 321106/Réglettes - Rico Carpaye) : la colonne "Rem%" (remise, ex.
# "2,94") n'apparaît dans le texte extrait QUE si elle est non nulle sur
# CE document — quand elle est présente, elle s'intercale entre Total et
# P.U.Net, décalant tout le reste d'un cran (même famille que la cellule
# Eco-part optionnelle déjà rencontrée côté BL Stand 64). Sans cette
# variante, ce devis ressortait à 0 article malgré un Total HT affiché ;
# la référence attendue (ex. "302758", 6 chiffres) tombait sur la cellule
# P.U.Net ("45,00000", avec une virgule — rejetée par MOTIF_REF_DEVIS_BPU,
# ce qui évite justement toute confusion silencieuse avec la variante
# sans Rem%, voir parse_109 ci-dessous).
OFFSETS_DEVIS_BPU_REMISE = {
    "designation": -2,
    "qte": -1,
    "total": 1,
    "rem_pct": 2,
    "pu_net_affiche": 3,
    "reference_fournisseur": 4,
}

MOTIF_DEVIS = r"(?:Commande client|Devis) n[°o]\s*([\d\s]+?)\s+du"
MOTIF_TOTAL_HT = r"Total HT\s*\n\s*([\d\s]+,\d{2})\s*(?:EUR|€)"
# --- fin GABARIT -------------------------------------------------------------


def _autocontrole_total_ht(texte: str, articles: list[Article]) -> None:
    """Contrôle additif propre à 109 Distribution : le PDF affiche un
    Total HT global -> le comparer à la somme des lignes extraites est le
    meilleur signal possible qu'aucune ligne n'a été oubliée ou mal lue,
    ou que la variante de gabarit essayée ne correspond pas à ce PDF."""

    m = re.search(MOTIF_TOTAL_HT, texte)
    if not m:
        return

    total_pdf = to_float(m.group(1))
    total_extrait = round(sum(a.montant for a in articles), 2)

    if abs(total_pdf - total_extrait) > 0.02:
        print(
            f"!! 109 DISTRIBUTION : Total HT du PDF ({total_pdf:.2f}€) "
            f"!= somme des lignes extraites ({total_extrait:.2f}€) "
            f"— une ligne a peut-être été oubliée ou mal lue, ou ce PDF "
            f"correspond à une 3e variante de gabarit jamais vue."
        )


def _bloc_vers_article(bloc, devis, ref) -> Article | None:

    if bloc["qte"] is None or bloc["total"] is None or bloc["pu_net_affiche"] is None:
        return None

    quantite = to_float(bloc["qte"])
    total = to_float(bloc["total"])
    pu_net_affiche = to_float(bloc["pu_net_affiche"])

    if not quantite:
        return None

    return Article(
        fournisseur="109 DISTRIBUTION",
        devis=devis,
        reference_fournisseur=ref,
        reference_distributeur="",
        designation=(bloc["designation"] or "").strip(),
        quantite=quantite,
        unite="UN",
        prix_brut=pu_net_affiche,
        prix_net=round(total / quantite, 4),
        montant=total,
    )


def parse_109(texte: str) -> list[Article]:

    articles = []

    m = re.search(MOTIF_DEVIS, texte)
    devis = m.group(1).replace(" ", "") if m else ""

    lignes = lignes_propres(texte)

    blocs_commande = scan_ancre(lignes, MARQUEURS_TVA, OFFSETS_COMMANDE)
    blocs_devis_bpu = scan_ancre(lignes, MARQUEURS_TVA, OFFSETS_DEVIS_BPU)
    blocs_devis_bpu_remise = scan_ancre(lignes, MARQUEURS_TVA, OFFSETS_DEVIS_BPU_REMISE)

    for bloc_c, bloc_d, bloc_dr in zip(blocs_commande, blocs_devis_bpu, blocs_devis_bpu_remise):

        ref_c = (bloc_c["reference_fournisseur"] or "").strip()
        ref_d = (bloc_d["reference_fournisseur"] or "").strip()
        ref_dr = (bloc_dr["reference_fournisseur"] or "").strip()

        if MOTIF_REF_COMMANDE.match(ref_c):
            article = _bloc_vers_article(bloc_c, devis, ref_c)
        elif MOTIF_REF_DEVIS_BPU.match(ref_d):
            article = _bloc_vers_article(bloc_d, devis, ref_d)
        elif MOTIF_REF_DEVIS_BPU.match(ref_dr):
            article = _bloc_vers_article(bloc_dr, devis, ref_dr)
        else:
            article = None

        if article:
            articles.append(article)

    _autocontrole_total_ht(texte, articles)

    return articles


# --- GABARIT BL (109 Distribution) ------------------------------------------
# BL scanné (image pure, voir moteur/ocr.py) : impossible de réutiliser
# scan_ancre() à décalage fixe comme pour le devis ci-dessus — l'OCR "rate"
# parfois une cellule entière (ex. la quantité) sans que la ligne bascule
# pour autant, ce qui décale tous les offsets suivants (constaté sur les 4
# vrais BL de session R2 : tests/fixtures/bl_dist109_*.pdf). Seules DEUX
# positions sont fiables sur CHAQUE ligne observée, relatives au code TVA
# (C0 à C9, "CO" ajouté car l'OCR confond régulièrement 0/O) : le P.U.Net
# juste avant, le Total juste après. La quantité livrée en est déduite
# (Total / P.U.Net) plutôt que lue directement — la cellule Qté peut être
# tronquée par une coche de "livré" imprimée dedans (ex. "32" lu "327").
# Un même fichier peut aussi contenir PLUSIEURS BL scannés à la suite
# ("scans en masse") — voir parse_bl_109() plus bas, même principe que
# Cominter Ouest (moteur.ocr.pages_par_identifiant).
MOTIF_TVA_BL = re.compile(r"^C[0-9O]$")
# Le P.U.Net a TOUJOURS 5 décimales sur ce gabarit (ex. "22,00000"),
# contrairement à l'éco-part et au Total qui n'en ont que 2 — signal fiable
# pour le distinguer d'une cellule "Eco-part" intercalée avant le code TVA,
# ou pour le retrouver directement quand l'OCR n'a lu AUCUN code TVA sur la
# ligne (voir _ligne_bl_vers_article, cas réels bl_dist109_7.pdf).
MOTIF_PU_NET_BL = re.compile(r"^\d+[,.]\d{5}$")
# N°Réf.Client : "123.096" (2 groupes) le plus courant, mais aussi des
# codes chantier "M3.14.353" (préfixe lettre + 3 groupes) — cas réel trouvé
# en recette (voir CLAUDE.md, session R2 suite). Capture tout le motif tel
# quel plutôt que 2 groupes fixes : plus robuste aux variantes.
MOTIF_COMMANDE_BL = re.compile(r"^([A-Z]?\d{1,4}(?:[.\-]\d{1,4}){1,2})$", re.IGNORECASE)
MOTIF_ENTETE_TABLEAU_BL = re.compile(r"REFERENCEARTICLE")
# "Total Eco-part HT" (1er repère essayé) est parfois lu de travers par
# l'OCR ("Tatal Eco-part HT", voire pire) — "Total HT" juste après, lui,
# est resté fiable sur TOUS les documents vus jusqu'ici. Repli sur ce 2e
# repère quand le 1er ne matche pas (cas réel bl_dist109_8, pages 6-7 :
# sans ce repli, la fin du tableau n'était jamais trouvée -> 0 ligne
# extraite malgré un Total HT bien affiché).
MOTIF_PIED_TABLEAU_BL = re.compile(r"TOTALECO-?PARTHT|TOTALHT")
MOTIF_BL_NUMERO_DATE = re.compile(
    r"ivrais\w*n?\D*?(\d{5,7})\D*?du\D*?(\d{1,2}\D?\d{2}\D?\d{4})", re.IGNORECASE
)
# BUG RÉEL CORRIGÉ (recette réelle, cas réel signalé par l'acheteur) : ce
# fournisseur envoie aussi des "Retour n° X du date" — un document qui
# ANNULE une ligne d'un BL précédent (ex. article listé mais pas coché
# "livré" à réception), jamais une livraison en soi. Le corps d'un retour
# référence TOUJOURS le BL qu'il annule ("Bon de livraison n° Y du date"),
# ce qui faisait matcher MOTIF_BL_NUMERO_DATE à tort sur CETTE référence
# (au lieu de l'identité propre du retour) — le document se retrouvait
# fusionné à tort comme une "page supplémentaire" du BL Y (voir
# moteur.ocr.pages_par_identifiant), avec un risque réel de compter deux
# fois la même quantité si le retour avait été traité comme une livraison
# normale. MOTIF_IDENTIFIANT_PAGE_BL sert désormais à regrouper les pages
# (reconnaît "Retour n°" AVANT "livraison n°" dans le texte d'une page de
# retour, puisqu'il apparaît en premier dans le document) ; MOTIF_RETOUR_
# NUMERO_DATE sert à détecter qu'UN groupe de pages est un retour plutôt
# qu'un BL (voir parse_bl_109()).
MOTIF_RETOUR_NUMERO_DATE = re.compile(
    r"Retour\w*\D*?n[°o]?\D*?(\d{4,7})\D*?du\D*?(\d{1,2}\D?\d{2}\D?\d{4})", re.IGNORECASE
)
MOTIF_IDENTIFIANT_PAGE_BL = re.compile(
    r"(?:ivrais\w*n?|Retour\w*\D*?n[°o]?)\D*?(\d{4,7})\D*?du\D*?(\d{1,2}\D?\d{2}\D?\d{4})",
    re.IGNORECASE,
)
MOTIF_TOTAL_HT_BL = re.compile(r"TOTALHT\D*?(\d[\d\s]*[,.]\d{2})EUR")
# --- fin GABARIT BL ----------------------------------------------------------


def _sans_espaces(s: str) -> str:
    # BUG RÉEL CORRIGÉ (recette réelle, bl_dist109_6.pdf) : l'OCR lit
    # parfois l'en-tête "Référence article" AVEC son accent ("Reférence",
    # cf. "É" majuscule après .upper()) au lieu de la forme sans accent
    # habituellement lue — MOTIF_ENTETE_TABLEAU_BL ("REFERENCEARTICLE") ne
    # matchait alors jamais, faisant disparaître TOUT le tableau (0 ligne
    # extraite alors qu'un Total HT était bien affiché). Les accents sont
    # désormais retirés avant comparaison, pour les deux graphies.
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", "", s.upper())


def _zone_tableau_bl(lignes_groupees: list[list[dict]]) -> list[list[dict]]:
    """Lignes (groupées par ligne visuelle, mots OCR positionnés) strictement
    entre l'en-tête de colonnes et le premier total — exclut le tableau de
    répartition TVA du pied de page, qui contient lui aussi des codes
    "C1"/"C0" isolés."""

    i_entete = next(
        (i for i, ligne in enumerate(lignes_groupees)
         if any(MOTIF_ENTETE_TABLEAU_BL.search(_sans_espaces(m["texte"])) for m in ligne)),
        None,
    )
    i_pied = next(
        (i for i, ligne in enumerate(lignes_groupees)
         if any(MOTIF_PIED_TABLEAU_BL.search(_sans_espaces(m["texte"])) for m in ligne)),
        None,
    )

    if i_entete is None or i_pied is None or i_pied <= i_entete:
        return []

    return lignes_groupees[i_entete + 1:i_pied]


def _ligne_bl_vers_article(cellules: list[str]) -> LigneBL | None:

    if len(cellules) < 2:
        return None

    indices_tva = [
        i for i, c in enumerate(cellules) if MOTIF_TVA_BL.match(c.strip().upper())
    ]

    if indices_tva:
        i_tva = indices_tva[0]

        if i_tva == 0 or i_tva + 1 >= len(cellules):
            return None

        i_pu_net = i_tva - 1
        if (
            not MOTIF_PU_NET_BL.match(cellules[i_pu_net].strip())
            and i_tva - 2 >= 0
            and MOTIF_PU_NET_BL.match(cellules[i_tva - 2].strip())
        ):
            # BUG RÉEL CORRIGÉ (recette réelle, bl_dist109_7.pdf) : une
            # cellule "Eco-part" (2 décimales) s'intercale parfois entre le
            # vrai P.U.Net (toujours 5 décimales) et le code TVA — sans ce
            # repli, le code prenait l'éco-part pour le P.U.Net, donnant
            # une quantité totalement fausse (104,76 au lieu de 3 réels,
            # ligne 302304).
            i_pu_net = i_tva - 2

        total = to_float(cellules[i_tva + 1])
    else:
        # BUG RÉEL CORRIGÉ (recette réelle, bl_dist109_7.pdf) : certaines
        # lignes n'ont AUCUNE cellule de code TVA lue par l'OCR (3 lignes
        # sur 7 sur ce document, silencieusement ignorées avant ce
        # correctif) — repli sur le format du P.U.Net lui-même (5
        # décimales) pour le retrouver sans dépendre du code TVA ; le
        # Total reste alors la DERNIÈRE cellule de la ligne (vérifié exact
        # sur les 3 lignes concernées).
        candidats_pu_net = [i for i, c in enumerate(cellules) if MOTIF_PU_NET_BL.match(c.strip())]
        if len(candidats_pu_net) != 1:
            return None
        i_pu_net = candidats_pu_net[0]
        total = to_float(cellules[-1])

    if i_pu_net <= 0:
        return None

    pu_net = to_float(cellules[i_pu_net])

    if not pu_net or not total:
        return None

    return LigneBL(
        reference_fournisseur=cellules[0].strip(),
        designation=cellules[1].strip() if i_pu_net > 1 else "",
        quantite_livree=round(total / pu_net, 4),
        prix_net=pu_net,
        montant=total,
    )


def _parse_un_bl_109(mots_par_page: list[list[dict]]) -> BonLivraison:

    lignes_plates = [
        mot["texte"]
        for mots in mots_par_page
        for ligne in regrouper_lignes(mots)
        for mot in ligne
    ]
    texte = "\n".join(lignes_plates)

    numero_commande = ""
    for ligne in lignes_plates:
        m = MOTIF_COMMANDE_BL.match(ligne.strip())
        if m:
            numero_commande = m.group(1).upper().replace("-", ".")
            break

    texte_aplati = texte.replace("\n", "")

    # Un "Retour n° X du date" (voir bandeau GABARIT BL) référence TOUJOURS
    # le BL qu'il annule ("Bon de livraison n° Y du date") dans son corps —
    # vérifier d'ABORD s'il s'agit d'un retour, sinon MOTIF_BL_NUMERO_DATE
    # matcherait à tort cette référence comme si c'était l'identité propre
    # du document.
    type_document = "BL"
    numero_bl, date_bl, numero_bl_origine = "", "", ""

    m_retour = MOTIF_RETOUR_NUMERO_DATE.search(texte_aplati)
    if m_retour:
        type_document = "RETOUR"
        numero_bl, date_bl = m_retour.group(1), m_retour.group(2)
        m_origine = MOTIF_BL_NUMERO_DATE.search(texte_aplati)
        if m_origine:
            numero_bl_origine = m_origine.group(1)
    else:
        m = MOTIF_BL_NUMERO_DATE.search(texte_aplati)
        if m:
            numero_bl, date_bl = m.group(1), m.group(2)

    total_ht_affiche = None
    m = MOTIF_TOTAL_HT_BL.search(_sans_espaces(texte))
    if m:
        total_ht_affiche = to_float(m.group(1))

    articles = []
    for mots in mots_par_page:
        for ligne_mots in _zone_tableau_bl(regrouper_lignes(mots)):
            cellules = [m["texte"] for m in ligne_mots]
            article = _ligne_bl_vers_article(cellules)
            if article:
                articles.append(article)

    bl = BonLivraison(
        fournisseur="109 DISTRIBUTION",
        fichier="",
        numero_bl=numero_bl,
        date_bl=date_bl,
        numero_commande=numero_commande,
        lignes=articles,
        total_ht_affiche=total_ht_affiche,
        type_document=type_document,
        numero_bl_origine=numero_bl_origine,
    )

    if total_ht_affiche is not None:
        total_extrait = round(sum(a.montant for a in articles), 2)
        if abs(total_ht_affiche - total_extrait) > 0.02:
            print(
                f"!! 109 DISTRIBUTION (BL) : Total HT du PDF ({total_ht_affiche:.2f}€) "
                f"!= somme des lignes extraites ({total_extrait:.2f}€) "
                f"— une ligne a peut-être été oubliée ou mal lue par l'OCR."
            )

    return bl


def parse_bl_109(mots_par_page: list[list[dict]]) -> list[BonLivraison]:
    """Un même fichier peut contenir PLUSIEURS BL (et/ou RETOURS, voir
    bandeau GABARIT BL) scannés à la suite (cas réel, "scans en masse"
    déposés par l'acheteur — même principe que Cominter Ouest, voir
    moteur.ocr.pages_par_identifiant) : retourne une liste, une entrée par
    document détecté. MOTIF_IDENTIFIANT_PAGE_BL (pas MOTIF_BL_NUMERO_DATE
    seul) sert d'identifiant de regroupement : reconnaît "Retour n°" en
    plus de "livraison n°", pour qu'une page de retour ne soit jamais
    fusionnée à tort avec le BL qu'elle référence dans son corps (bug réel
    corrigé — voir MOTIF_RETOUR_NUMERO_DATE). Chaque BonLivraison porte
    aussi les indices de page qu'il occupe dans le fichier source
    (bl.pages), pour l'archivage individuel (voir
    moteur.rapprochement.pipeline_bl, "archivage par BL individuel")."""

    groupes_indices = pages_par_identifiant(mots_par_page, MOTIF_IDENTIFIANT_PAGE_BL)

    resultat = []
    for indices in groupes_indices:
        bl = _parse_un_bl_109([mots_par_page[i] for i in indices])
        bl.pages = indices
        resultat.append(bl)

    return resultat


# Déclaration pour le chargement automatique
FOURNISSEURS = ['109 DISTRIBUTION']
parse = parse_109
parse_bl = parse_bl_109
