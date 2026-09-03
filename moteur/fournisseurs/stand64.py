"""
Parser Stand 64 (STAND 64, luminaires - prescription).

Format difficile : la référence est souvent un placeholder (ZARTICLENP =
article non prédéfini par le commercial, absent du catalogue fournisseur ;
ZARTICLETVA = éco-contribution). Le produit réel n'est décrit que dans le
texte. Colonnes en ordre inversé, nombre de colonnes numériques variable
(éco-part présente ou non).

Ancre fiable : après la description viennent, dans l'ordre,
    Qté, code TVA (C0/C1/C4/C7), Total HT, P.U Net, [autres nombres], Référence

**Placeholder ZARTICLENP/"Alternative:ZARTICLENP" : la vraie référence
fabricant, quand elle existe, est alors donnée SEULEMENT entre parenthèses
en fin de désignation** (ex. "...SD-WOOD RING SUSPENSION ... IP20 BOIS
(KUBIA-ART00031180)"), pas dans le champ référence — cas réel, chantier
Cosinus (`tests/fixtures/stand64_cosinus.pdf`). Sans extraire ce code, TOUTES
les lignes personnalisées d'un même devis partagent la même "référence" et
s'écrasent entre elles au comparateur (34 lignes -> 5 conservées, constaté
sur ce PDF) : `_reference_reelle()` la récupère quand le champ référence
imprimé est un tel placeholder ; sinon la référence imprimée fait foi, comme
avant.

PARSER MEILLEUR EFFORT : à contrôler visuellement. Les lignes
"ECO CONTRIBUTION" (ZARTICLETVA) sont ignorées. Un garde-fou
qté x P.U Net ≈ Total HT écarte les lectures douteuses.
"""

import re
import unicodedata

from moteur.modele import Article
from moteur.outils import to_float, lignes_propres
from moteur.ocr import regrouper_lignes
from moteur.rapprochement.modele_bl import BonLivraison, LigneBL
from moteur.rapprochement.modele_facture import Facture, LigneFacture

_NUM = re.compile(r"^\s*\d[\d\s]*,\d{2}\s*$")
_TVA = re.compile(r"^C\d$")
_MARQUEUR = re.compile(r"^(REPERE|REPRE|VARIANTE|DETECTEUR|PRIX NETS)", re.I)
_CODE_INTEGRE = re.compile(r"\(([A-Z0-9][A-Z0-9\-\.]{2,})\)\s*$")


def _reference_reelle(ref: str, designation: str) -> str:
    """Récupère le code fabricant intégré en fin de désignation quand la
    référence imprimée est un placeholder générique ("ZARTICLENP",
    "Alternative:ZARTICLENP"...) — sinon la référence imprimée telle quelle."""

    if "ZARTICLE" not in ref.upper():
        return ref

    m = _CODE_INTEGRE.search(designation)
    return m.group(1) if m else ref


def parse_stand64(texte: str) -> list[Article]:

    articles = []

    m = re.search(r"Devis n°\s*([\d\s]+?)\s+du", texte)
    devis = m.group(1).replace(" ", "") if m else ""

    lignes = lignes_propres(texte)
    n = len(lignes)

    # Début après l'en-tête "Description"
    debut = 0
    for k, l in enumerate(lignes):
        if l == "Description":
            debut = k + 1
            break

    i = debut
    desc_buffer = []

    while i < n:
        l = lignes[i]

        # Pied de page légal répété à chaque page : sauter jusqu'à la
        # reprise du tableau ("Description") sur la page suivante, ou finir.
        if l.startswith(("Le transfert", "Total Eco-part", "Total HT",
                         "Mode de paiement", "Signature")):
            desc_buffer = []
            p = i + 1
            while p < n and lignes[p] != "Description":
                p += 1
            if p >= n:
                break
            i = p + 1
            continue

        # Marqueurs de repère / variante : réinitialisent la description
        if _MARQUEUR.match(l):
            desc_buffer = []
            i += 1
            continue

        # Début d'un bloc chiffré : Qté (nombre) suivi d'un code TVA
        if _NUM.match(l) and i + 1 < n and _TVA.match(lignes[i + 1].strip()):
            try:
                quantite = to_float(l)
                total = to_float(lignes[i + 2])
                prix_net = to_float(lignes[i + 3])

                # Avancer jusqu'à la référence (1re ligne non numérique)
                j = i + 4
                while j < n and _NUM.match(lignes[j]):
                    j += 1
                ref = lignes[j] if j < n else ""
            except Exception:
                desc_buffer = []
                i += 1
                continue

            designation = " ".join(desc_buffer).strip()

            # Ignorer l'éco-contribution
            est_eco = ref.upper() == "ZARTICLETVA" or "ECO CONTRIBUTION" in designation.upper()

            # Garde-fou cohérence
            coherent = quantite and prix_net and abs(
                total - quantite * prix_net
            ) <= max(0.02 * total, 1)

            if not est_eco and designation and coherent:
                articles.append(
                    Article(
                        fournisseur="STAND 64",
                        devis=devis,
                        reference_fournisseur=_reference_reelle(ref, designation),
                        reference_distributeur="",
                        designation=designation,
                        quantite=quantite,
                        unite="UN",
                        prix_brut=prix_net,
                        prix_net=prix_net,
                        montant=total,
                    )
                )

            desc_buffer = []
            i = j + 1
            continue

        # Sinon : ligne de description
        desc_buffer.append(l)
        i += 1

    return articles


# --- GABARIT BL (Stand 64) --------------------------------------------------
# BL scanné (image nette, comme les autres — voir moteur/ocr.py), TOUJOURS
# 2 vrais PDF vus à ce jour (M2.23.058, M4.270). Tableau simple, une ligne
# visuelle = un article, colonnes DANS CET ORDRE (contrairement au devis,
# qui les a en ordre inversé) :
#     Référence article | Description | Qté | P.U | Rem% | P.U Net |
#     Eco-part | Total HT | TVA
# "Eco-part" est TOUJOURS vide sur les 2 PDF vus (aucune cellule imprimée
# dans ce cas, pas juste "0,00") -> la ligne chiffrée a alors exactement
# 6 cellules après la désignation (Qté, P.U, Rem%, P.U Net, Total HT, TVA).
# Ancre fiable : le code TVA (C0..C9) en toute fin de ligne. Repéré par la
# POSITION relative à cette ancre (comme 109 Distribution/Comptoir du
# Cabling), pas par un compte de cellules fixe pour la désignation (elle
# peut déborder sur plusieurs cellules OCR selon la longueur du libellé).
# Montant = Total HT (affiché), jamais recalculé : qté x P.U Net retombe
# exactement dessus sur les 2 PDF vus (aucune remise réelle rencontrée à
# ce jour, Rem% toujours à 0,00) -> pas encore de cas réel avec une vraie
# remise pour valider l'ordre eco-part/remise en cas de conflit.
#
# BUG RÉEL CORRIGÉ (nouveau lot de BL réels) : l'hypothèse "Eco-part
# toujours vide" ne tient PAS sur tous les documents — un vrai BL
# (ELIOT-ES52-2678-BLC+BLC, commande M2.5.126) a l'Eco-part RENSEIGNÉE
# (2,88€), ajoutant UNE cellule numérique de plus avant Total HT et
# décalant tout le compte fixe de 6 cellules (la Qté "18,00" se
# retrouvait alors happée dans la désignation, et le P.U "95,00" pris à
# tort pour la Qté). `_ligne_bl_vers_article_stand64()` essaie donc les
# DEUX hypothèses (6 cellules sans éco-part, 7 avec) et retient celle dont
# Qté x P.U Net retombe sur le Total HT affiché — même principe que les
# replis positionnels validés par cohérence arithmétique déjà utilisés
# chez 109 Distribution/Electric Plus.
#
# BUG RÉEL CORRIGÉ (même lot) : une coche/checkmark imprimée à côté de la
# Qté est parfois lue par l'OCR comme une lettre COLLÉE à la valeur (ex.
# "40,00V" au lieu de "40,00", le "V" ressemblant au symbole ✓) — sans
# nettoyage, `to_float()` levait une exception et TOUTE la ligne
# disparaissait silencieusement (0 ligne extraite alors qu'un Total HT
# était bien affiché). `_nombre_bl_stand64()` retire une lettre isolée en
# fin de cellule avant conversion.
MOTIF_TVA_BL_STAND64 = re.compile(r"^C\d$")
MOTIF_ENTETE_TABLEAU_BL_STAND64 = re.compile(r"REFERENCEARTICLE")
MOTIF_PIED_TABLEAU_BL_STAND64 = re.compile(r"TOTALECO-PARTHT|MODEDEPAIEMENT")
MOTIF_BL_STAND64 = re.compile(r"LIVRAISON\D{0,3}(\d{3,6})\s*DU\s*(\d{2}/\d{2}/\d{4})", re.IGNORECASE)
MOTIF_COMMANDE_BL_STAND64 = re.compile(r"BC\s*N[^A-Z0-9]{0,2}([A-Z]?\d[\dA-Z.\-]*)", re.IGNORECASE)
# --- fin GABARIT BL -----------------------------------------------------------


def _sans_espaces_stand64(s: str) -> str:
    # Même repli que 109 Distribution (moteur/fournisseurs/dist109.py,
    # _sans_espaces) : l'en-tête "Référence article" peut ressortir de
    # l'OCR AVEC son accent — les accents sont retirés avant comparaison.
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", "", s.upper())


def _zone_tableau_bl_stand64(lignes_groupees: list[list[dict]]) -> list[list[dict]]:

    i_entete = next(
        (i for i, ligne in enumerate(lignes_groupees)
         if any(MOTIF_ENTETE_TABLEAU_BL_STAND64.search(_sans_espaces_stand64(m["texte"])) for m in ligne)),
        None,
    )
    if i_entete is None:
        return []

    i_pied = next(
        (i for i, ligne in enumerate(lignes_groupees)
         if i > i_entete and any(
             MOTIF_PIED_TABLEAU_BL_STAND64.search(_sans_espaces_stand64(m["texte"])) for m in ligne
         )),
        None,
    )

    return lignes_groupees[i_entete + 1:(i_pied if i_pied is not None else len(lignes_groupees))]


def _nombre_bl_stand64(cellule: str) -> float:
    # Coche/checkmark imprimée collée à la valeur, lue "V" par l'OCR (voir
    # bandeau GABARIT BL) — retirée si elle traîne en fin de cellule.
    return to_float(re.sub(r"[A-Za-z]$", "", cellule.strip()))


def _hypothese_stand64(cellules: list[str], decalage: int) -> tuple | None:
    """Tente une position de Qté/P.U Net à `decalage` cellules de la fin
    (6 = pas d'éco-part, 7 = éco-part renseignée) -> (quantite, pu_net,
    total_ht) si les 3 valeurs se lisent ET que qté x P.U Net retombe sur
    le Total HT affiché, None sinon."""

    if len(cellules) < decalage + 1:
        return None
    try:
        total_ht = _nombre_bl_stand64(cellules[-2])
        pu_net = _nombre_bl_stand64(cellules[-4] if decalage == 7 else cellules[-3])
        quantite = _nombre_bl_stand64(cellules[-decalage])
    except Exception:
        return None

    if not quantite or not pu_net:
        return None
    if abs(quantite * pu_net - total_ht) > max(0.02 * total_ht, 0.5):
        return None

    return quantite, pu_net, total_ht


def _ligne_bl_vers_article_stand64(cellules: list[str]) -> LigneBL | None:

    if len(cellules) < 8 or not MOTIF_TVA_BL_STAND64.match(cellules[-1].strip()):
        return None

    for decalage in (6, 7):
        hypothese = _hypothese_stand64(cellules, decalage)
        if hypothese is not None:
            quantite, pu_net, total_ht = hypothese
            break
    else:
        return None

    reference = cellules[0].strip()
    # Désignation parfois VIDE ici (voir bandeau GABARIT BL — 1re ligne de
    # désignation non détectée par l'OCR sur un vrai document, seule la
    # ligne chiffrée l'a été) : raccordée depuis les lignes suivantes par
    # l'appelant (parse_bl_stand64), jamais rejetée ici (qté/prix restent
    # exploitables sans désignation, celle-ci ne sert pas au rapprochement).
    designation = " ".join(cellules[1:-decalage]).strip()

    if not reference:
        return None

    return LigneBL(
        reference_fournisseur=reference,
        designation=designation,
        quantite_livree=quantite,
        prix_net=pu_net,
        montant=total_ht,
    )


def parse_bl_stand64(mots_par_page: list[list[dict]]) -> BonLivraison:

    lignes_plates = [
        mot["texte"]
        for mots in mots_par_page
        for ligne in regrouper_lignes(mots)
        for mot in ligne
    ]
    texte = "\n".join(lignes_plates)
    texte_compact = _sans_espaces_stand64(texte)

    numero_bl = ""
    date_bl = ""
    m = MOTIF_BL_STAND64.search(texte_compact)
    if m:
        numero_bl = m.group(1)
        date_bl = m.group(2)

    numero_commande = ""
    m = MOTIF_COMMANDE_BL_STAND64.search(texte)
    if m:
        numero_commande = m.group(1).upper()

    articles = []
    for mots in mots_par_page:
        lignes_zone = _zone_tableau_bl_stand64(regrouper_lignes(mots))
        i = 0
        while i < len(lignes_zone):
            cellules = [m["texte"] for m in lignes_zone[i]]
            i += 1
            article = _ligne_bl_vers_article_stand64(cellules)
            if article is None:
                continue
            if not article.designation:
                # Désignation manquante (voir bandeau GABARIT BL) :
                # raccorde les lignes suivantes tant qu'elles ne sont
                # PAS elles-mêmes une ligne chiffrée reconnue (texte de
                # désignation pur, ex. "BLANC/BLANC+TELECOMMANDE",
                # "PRIX NETS", "MATERIEL DISPONIBLE CE JOUR").
                morceaux = []
                while i < len(lignes_zone):
                    cellules_suivantes = [m["texte"] for m in lignes_zone[i]]
                    if _ligne_bl_vers_article_stand64(cellules_suivantes) is not None:
                        break
                    morceaux.extend(cellules_suivantes)
                    i += 1
                article.designation = " ".join(morceaux).strip()
            articles.append(article)

    return BonLivraison(
        fournisseur="STAND 64",
        fichier="",
        numero_bl=numero_bl,
        date_bl=date_bl,
        numero_commande=numero_commande,
        lignes=articles,
        total_ht_affiche=None,
    )


# --- GABARIT FACTURE (Stand 64) --------------------------------------------
# Session F4 suite (2026-09-02+), ~25 pièces réelles déposées par
# l'acheteur (Facture_XXXXX.pdf), texte NATIF — jamais de scan (contrairement
# au BL du même fournisseur, qui lui en a besoin malgré l'apparence nette).
#
# Structure par article, ancrée sur le MONTANT — mais ici le Total HT de
# chaque ligne est imprimé AVANT sa désignation (pas après, contrairement à
# la plupart des autres fournisseurs) : vérifié par cohérence arithmétique
# EXACTE sur une pièce riche à 6 articles (facture_33707) — la somme des 6
# "Total HT" retombe pile sur le Total HT affiché en pied de facture
# (3 615,00€), et pour CHAQUE ligne, Total HT = Qté x P.U Net, et
# P.U Net = P.U x (1 - Rem%/100) à l'euro près. Ordre après le code TVA :
# Qté, P.U Net, P.U (brut), Rem%.
#
# N° de commande : préfère notre propre "BON DE COMMANDE" (reproduit en
# pied de certaines factures, pas systématique) au "BC N°..." de l'en-tête,
# qui peut souffrir d'un décalage du point réel constaté (cas réel
# facture_33707 : en-tête "BC N°M2.220.78" vs BON DE COMMANDE "M2.22.078" —
# mêmes chiffres "22078", point décalé d'une position) — pas de correctif
# général tenté (un seul exemple, règle d'or), juste une préférence de
# source quand les deux sont disponibles.
#
# Eco-part : colonne listée dans l'en-tête du tableau mais jamais rencontrée
# non nulle sur les pièces vues à ce jour (toujours 0,00€ en pied de
# facture) — pas de position confirmée dans le bloc de 4 valeurs, à
# éclaircir dès qu'un exemple réel avec Eco-part renseignée se présentera.
MOTIF_NUMERO_FACTURE_STAND64 = re.compile(r"Facture client n°\s*([\d\s]+?)\s+du\s+(\d{2}/\d{2}/\d{4})")
MOTIF_ECHEANCE_STAND64 = re.compile(r"Date d'échéance\s*:\s*(\d{2}/\d{2}/\d{4})")
MOTIF_BL_FACTURE_STAND64 = re.compile(r"Bon de livraison n°\s*(\d+)\s+du\s+(\d{2}/\d{2}/\d{4})")
MOTIF_COMMANDE_ENTETE_STAND64 = re.compile(r"BC N°\s*([A-Z0-9][A-Z0-9.\-]*)")
MOTIF_TOTAL_HT_AFFICHE_STAND64 = re.compile(r"Total HT\s*\n\s*([\d\s]+,\d{2})\s*EUR")
MOTIF_NUM_FACTURE_STAND64 = re.compile(r"^\d[\d\s]*,\d{2}$")
MOTIF_TVA_FACTURE_STAND64 = re.compile(r"^C\d$")

# BUG RÉEL CORRIGÉ : un mot de désignation tout en MAJUSCULES sans aucun
# chiffre ni tiret (ex. "CHAINETTE", fin d'une désignation qui déborde sur
# 2 lignes) matchait à tort un motif référence trop permissif — TOUTES les
# vraies références vues à ce jour contiennent au moins UN TIRET (même
# "WESTI-COMET-KITLUM-N", sans aucun chiffre) : signal fiable, contrairement
# à "au moins un chiffre" qui aurait exclu cette référence légitime. "+"
# toléré dans les segments (référence réelle "ELIOT-ES52-2678-BLC+BLC",
# déjà connue côté BL du même fournisseur, commande M2.5.126).
MOTIF_REF_FACTURE_STAND64 = re.compile(r"^[A-Z][A-Z0-9+]*(?:-[A-Z0-9+]+)+$")
MOTIF_FIN_TABLEAU_FACTURE_STAND64 = re.compile(r"^Total$")
# --- fin GABARIT FACTURE -----------------------------------------------------


def _entete_facture_stand64(lignes: list[str]) -> str:
    """N° de commande : préfère le "BON DE COMMANDE" (notre propre BC, voir
    bandeau) ; repli sur le "BC N°..." de l'en-tête sinon."""

    for i, l in enumerate(lignes):
        if l.strip().upper() == "BON DE COMMANDE" and i + 1 < len(lignes):
            candidat = lignes[i + 1].strip()
            if candidat:
                return candidat

    for l in lignes:
        m = MOTIF_COMMANDE_ENTETE_STAND64.search(l)
        if m:
            return m.group(1).upper()

    return ""


def _zone_articles_facture_stand64(lignes: list[str]):
    """(indice_debut, indice_fin) — entre le "Bon de livraison n°..." (le
    tableau d'articles suit directement) et le "Total" bare qui introduit
    le tableau de ventilation TVA."""

    i_bl = next((i for i, l in enumerate(lignes) if MOTIF_BL_FACTURE_STAND64.search(l)), None)
    if i_bl is None:
        return None

    debut = i_bl + 1
    i_fin = next(
        (i for i in range(debut, len(lignes)) if MOTIF_FIN_TABLEAU_FACTURE_STAND64.match(lignes[i].strip())),
        len(lignes),
    )

    return debut, i_fin


def _lignes_facture_stand64(lignes: list[str], debut: int, fin: int, numero_bl: str) -> list[LigneFacture]:
    """Ancrée sur la RÉFÉRENCE (fiable dans les deux sens, voir
    MOTIF_REF_FACTURE_STAND64), pas sur un compte de valeurs numériques —
    celui-ci varie réellement d'une ligne à l'autre au sein d'un MÊME
    document (Eco-part présente ou non par ligne, cas réel facture_34184 :
    2 lignes sur 3 avec Eco-part renseignée, 1 sans, un compte fixe aurait
    décalé une ligne sur deux). Qté et P.U Net sont à position FIXE juste
    après le code TVA (fiable, quel que soit le nombre de valeurs
    restantes) ; Total HT et désignation sont retrouvés en remontant DEPUIS
    la référence : tout ce qui n'est pas numérique juste avant elle est
    désignation, le premier nombre rencontré au-delà est le Total HT — ne
    dépend jamais du nombre de valeurs consommées par la ligne PRÉCÉDENTE."""

    indices_ref = [i for i in range(debut, fin) if MOTIF_REF_FACTURE_STAND64.match(lignes[i].strip())]

    articles = []
    for idx_ref in indices_ref:

        i_tva = idx_ref + 1
        if i_tva >= fin or not MOTIF_TVA_FACTURE_STAND64.match(lignes[i_tva].strip()):
            continue

        i_qte = i_tva + 1
        if i_qte + 1 >= fin:
            continue
        try:
            qte = to_float(lignes[i_qte].strip())
            pu_net = to_float(lignes[i_qte + 1].strip())
        except ValueError:
            continue

        k = idx_ref - 1
        designation_lignes = []
        while k >= debut and not MOTIF_NUM_FACTURE_STAND64.match(lignes[k].strip()):
            if lignes[k].strip():
                designation_lignes.insert(0, lignes[k].strip())
            k -= 1
        if k < debut:
            continue
        total_ht = to_float(lignes[k].strip())

        if not qte:
            continue

        articles.append(LigneFacture(
            reference_fournisseur=lignes[idx_ref].strip(),
            designation=" ".join(designation_lignes),
            quantite_facturee=qte,
            prix_unitaire_ht=pu_net,
            montant_ht=total_ht,
            numero_bl=numero_bl,
        ))

    return articles


def parse_facture_stand64(texte: str) -> Facture:

    lignes = lignes_propres(texte)

    numero_facture = ""
    date_facture = ""
    m_num = MOTIF_NUMERO_FACTURE_STAND64.search(texte)
    if m_num:
        numero_facture = m_num.group(1).replace(" ", "")
        date_facture = m_num.group(2)

    date_echeance = ""
    m_ech = MOTIF_ECHEANCE_STAND64.search(texte)
    if m_ech:
        date_echeance = m_ech.group(1)

    zone = _zone_articles_facture_stand64(lignes)
    if zone is None:
        return Facture(
            fournisseur="STAND 64", fichier="", numero_facture=numero_facture,
            date_facture=date_facture, date_echeance=date_echeance,
        )

    debut, fin = zone

    m_bl = MOTIF_BL_FACTURE_STAND64.search(lignes[debut - 1])
    numero_bl = m_bl.group(1) if m_bl else ""

    numero_commande = _entete_facture_stand64(lignes)

    lignes_facture = _lignes_facture_stand64(lignes, debut, fin, numero_bl)

    total_ht_affiche = None
    m_total = MOTIF_TOTAL_HT_AFFICHE_STAND64.search(texte)
    if m_total:
        total_ht_affiche = to_float(m_total.group(1))
        somme = round(sum(l.montant_ht for l in lignes_facture), 2)
        if abs(somme - total_ht_affiche) > 0.02:
            print(
                f"!! STAND 64 (Facture) : Total HT du PDF ({total_ht_affiche:.2f}€) != somme "
                f"des lignes extraites ({somme:.2f}€) — une ligne a peut-être été oubliée ou mal lue "
                "(voir bandeau GABARIT FACTURE)."
            )

    return Facture(
        fournisseur="STAND 64",
        fichier="",
        numero_facture=numero_facture,
        date_facture=date_facture,
        date_echeance=date_echeance,
        numeros_commande=[numero_commande] if numero_commande else [],
        numeros_bl=[numero_bl] if numero_bl else [],
        lignes=lignes_facture,
        total_ht_affiche=total_ht_affiche,
    )


FOURNISSEURS = ["STAND 64"]
parse = parse_stand64
parse_bl = parse_bl_stand64
parse_facture = parse_facture_stand64
