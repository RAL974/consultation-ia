import re

from moteur.modele import Article
from moteur.outils import to_float as _to_float, lignes_propres, chercher_devis
from moteur.fournisseurs._gabarit import scan_ancre, disponibilite_apres, diviser_qte_unite
from moteur.ocr import regrouper_lignes
from moteur.rapprochement.modele_bl import BonLivraison, LigneBL

# --- GABARIT (Ravate Elec / Ravate Pro) ---------------------------------
# Structure ancrée sur "Réf. FNR :" : Ravate facture toujours par la Réf.
# FNR, jamais par le code article interne (règle métier — voir CLAUDE.md).
MARQUEUR = ("RÉF. FNR :", "REF. FNR :")
MOTIF_DEVIS = r"\d{8}DV\d{4}"
OFFSETS = {
    "reference_fournisseur": 1,
    "reference_distributeur": 2,
    "designation": 3,
    "qte_unite": 4,
    "prix_brut": 5,
    "prix_net": 6,
    "montant": 7,
}
OFFSET_DISPONIBILITE = 8  # zone libre scannée jusqu'au marqueur suivant
# --- fin GABARIT ---------------------------------------------------------


def parse_ravate_core(texte: str, fournisseur: str) -> list[Article]:
    """Cœur commun aux devis Ravate (Elec et Pro)."""

    articles = []

    devis = chercher_devis(texte, MOTIF_DEVIS)

    lignes = lignes_propres(texte)

    for bloc in scan_ancre(lignes, MARQUEUR, OFFSETS):

        try:
            if any(bloc[champ] is None for champ in OFFSETS):
                raise IndexError("bloc incomplet")

            quantite, unite = diviser_qte_unite(bloc["qte_unite"])

            articles.append(
                Article(
                    fournisseur=fournisseur,
                    devis=devis,
                    reference_fournisseur=bloc["reference_fournisseur"],
                    reference_distributeur=bloc["reference_distributeur"],
                    designation=bloc["designation"],
                    quantite=quantite,
                    unite=unite,
                    prix_brut=_to_float(bloc["prix_brut"]),
                    prix_net=_to_float(bloc["prix_net"]),
                    montant=_to_float(bloc["montant"]),
                    disponibilite=disponibilite_apres(
                        lignes, bloc["_i"] + OFFSET_DISPONIBILITE, MARQUEUR
                    ),
                )
            )
        except Exception as e:
            print(f"Erreur lecture article {fournisseur} : {e}")

    return articles


def parse_ravate(texte: str) -> list[Article]:
    return parse_ravate_core(texte, "RAVATE")


# --- GABARIT BL (Ravate Elec) ------------------------------------------------
# BL scanné (image pure, voir moteur/ocr.py). Les 6 vrais BL Ravate
# disponibles à ce jour (règle d'or) sont TOUS des "ravatelec" (branding
# visible sur le document : "SITE EXPEDITION ... RAVATELEC LE PORT") — RAVATE
# PRO reste NON couvert côté BL (aucune pièce réelle disponible), contrairement
# au devis où l'acheteur a confirmé une structure identique (voir
# parse_ravate_core ci-dessus, PAS supposé ici sans confirmation équivalente
# côté BL).
#
# Chaque article est réparti sur PLUSIEURS lignes visuelles OCR : la
# désignation sur une ligne, les chiffres (Px Brut, Px Net, Remises, Montant)
# sur une autre. DEUX codes différents cohabitent — le "Code Art" (interne
# Ravate, TOUJOURS numérique pur, préfixe "100" [9 chiffres] ou "44200"
# [8 chiffres] sur tous les exemples vus) et la "Référence fournisseur"
# (le vrai code métier, alphanumérique ou numérique court, ex. "VK16BT",
# "R2V5G10", "404926", "069864") — leur position L'UN PAR RAPPORT À L'AUTRE
# VARIE selon le scan (Code Art tantôt sur la ligne désignation, tantôt en
# tête de la ligne chiffrée ; Référence fournisseur tantôt isolée sur sa
# propre ligne, tantôt en tête de la ligne chiffrée à la place du Code Art).
# Comme pour le devis de ce fournisseur (voir parse_ravate_core, "toujours
# la Réf. FNR, jamais la réf interne" — règle métier confirmée), LigneBL
# utilise TOUJOURS la Référence fournisseur, jamais le Code Art — repérée en
# EXCLUANT explicitement la forme du Code Art (MOTIF_CODE_ART_BL_RAVATE),
# jamais l'inverse (la Réf. fournisseur n'a pas une forme unique fiable).
#
# La ligne CHIFFRÉE est repérée par ses 4 DERNIÈRES cellules — Px Brut |
# Px Net | Remises (montant €, PAS un pourcentage) | Montant HT, TOUJOURS
# dans cet ordre — plutôt que par un nombre de cellules fixe (5 ou 7 selon
# que la quantité/unité soit inline ou non) : vérifié par cohérence
# arithmétique EXACTE sur les 6 pièces réelles disponibles, ex. :
#   - Câble    : Px Brut 2,78 / Px Net 1,15 / Remises 163,00 / Montant 115,00
#     -> (2,78-1,15) x 100 = 163,00 (qté 100)
#   - Poussoir : Px Brut 52,00 / Px Net 19,19 / Remises 98,43 / Montant 57,57
#     -> (52,00-19,19) x 3 = 98,43 (qté 3)
# La quantité livrée est donc TOUJOURS déduite de Montant / Px net (comme
# 109 Distribution/Cominter/Electric Plus), jamais lue dans une cellule Qté
# dont la position varie. La colonne "Remises" (jamais conservée) sert
# uniquement à cette vérification de cohérence.
#
# Cas réel d'OCR : une cellule monétaire peut avoir sa virgule lue comme un
# astérisque ("99*91" au lieu de "99,91", vérifié par cohérence
# arithmétique — 130 x 6,75 = 877,50, cohérent avec Px Brut ≈ 99,91 avant
# remise) — toléré comme séparateur décimal au même titre que "," et ".".
MOTIF_CODE_ART_BL_RAVATE = re.compile(r"^(100\d{6}|44200\d{3})$")
MOTIF_REF_FOURNISSEUR_BL_RAVATE = re.compile(r"^[A-Z0-9][A-Z0-9./]{2,10}$")
# BUG RÉEL CORRIGÉ (2e lot du jour) : "Reference" lu "Reterence" par l'OCR
# (F confondu avec T) sur un document par ailleurs propre — l'ancre ne
# matchait alors plus DU TOUT, la zone du tableau ressortait vide (0 ligne
# extraite pour un Total HT de 204,86€ affiché). "F" rendu tolérant au T.
MOTIF_ENTETE_TABLEAU_BL_RAVATE = re.compile(r"RE[FT]ERENCEFOURNISSEUR")
MOTIF_PIED_TABLEAU_BL_RAVATE = re.compile(r"\d*LIGNE")
# "BC 00312608CC0056 AU 04/08/2026 M3.23.033" (espaces, tirets ou rien du
# tout selon le scan, y compris entre "AU" et la date, et entre la date et
# la commande) -> capture le dernier jeton de la ligne. Plus simple et plus
# fiable qu'une déduction par contenu (voir Cominter) : ce fournisseur
# imprime le n° de commande en clair sur chaque bloc de commande. Le jour
# et le mois de la date ([\dB]{1,2} / \S{2}) tolèrent une confusion OCR
# chiffre/lettre (ex. "1B/0B/2026" au lieu de "18/08/2026", vu réellement —
# le même "B" à la place de "8" apparaît aussi DANS la commande elle-même,
# ex. "M3.10.1B2" pour "M3.10.182" — cas réel, nom de fichier du BL confirme
# la vraie valeur). Groupe captant : [\dB] au lieu de \d, "B" remplacé par
# "8" après capture (voir plus bas). "-" toléré aussi DANS la classe de
# caractères de la commande (pas seulement en sortie) : cas réel
# "2026131-162" (année et commande collées sans espace) — sans "-" dans la
# classe, la capture s'arrêtait juste avant le tiret et le \s*$ final ne
# matchait plus DU TOUT, faisant échouer toute la ligne (numéro de
# commande ET exclusion de cette ligne-repère hors de la désignation, qui
# réutilise ce même motif).
# BUG RÉEL CORRIGÉ (2e lot, fichier multi-fournisseurs) : "AU" collé
# directement au tiret introducteur de la date ("AU-25/08/2026") et la
# date collée au tiret introducteur de la commande ("2026--135-049") —
# les deux "\s*" d'origine n'acceptaient QUE des espaces, jamais un tiret
# à cet endroit précis (distinct du tiret déjà toléré DANS la commande
# elle-même) : la commande ressortait introuvable alors que le BL était
# par ailleurs propre. "\s*" élargi en "[\s-]*" aux deux endroits.
# BUG RÉEL CORRIGÉ (2e lot du jour, 2 documents du même lot) : le "/"
# séparateur de date lu "7" par l'OCR ("AU16706/2026...", voire les DEUX
# "/" lus "7" — "AU1670672026." — confirmé par la cellule date isolée du
# même document, "1670672026", même corruption). "/" élargi en "[/.7]"
# aux deux séparateurs de la date (sans risque de confusion : jour et mois
# ont une longueur fixe dans le motif, "7" ne peut donc jamais être pris
# pour un chiffre du jour/mois lui-même). Le séparateur avant la commande
# tolère aussi "." en plus de espace/tiret (vu collé par un point dans le
# même document, "...2026.M3.18.223").
# BUG RÉEL CORRIGÉ (BL 143.194 GYSM.jpg) : un préfixe de nom de chantier
# ("GYSM-143.194", collé sans espace après la date) n'était pas reconnu —
# le groupe capturant ([A-Z]?...) ne tolère qu'UNE SEULE lettre de préfixe
# (le "M"/"i"/"o" habituel DEVANT la commande elle-même), pas tout un mot
# comme "GYSM". Sans ce cas, `numero_commande` ressortait vide et le BL
# s'est archivé dans "Commande inconnue" au lieu de "143.194" — corrigé en
# tolérant un mot de lettres majuscules suivi d'un tiret AVANT la commande
# elle-même, optionnel pour ne rien changer aux cas déjà couverts (un seul
# "M"/"i"/"o" isolé ne peut jamais matcher "LETTRES-", faute du tiret qui
# suit immédiatement dans la commande réelle).
MOTIF_BC_COMMANDE_BL_RAVATE = re.compile(r"AU[\s-]*[\dB]{1,2}[/.7]\S{2}[/.7]\d{2,4}[\s.\-]*(?:[A-Z]+-)?([A-Z]?[\dB][\dB.-]*)\s*$", re.IGNORECASE)
# N° de BL (ex. "00312608LV0027") : motif stable sur les 6 pièces réelles
# (préfixe variable + "LV" + 3-5 chiffres), contrairement à une position de
# cellule fixe (la date de livraison se trouve tantôt seule sur sa propre
# ligne, tantôt dans la même ligne que ce code — les deux cas réels
# coexistent).
MOTIF_NUMERO_BL_RAVATE = re.compile(r"(\d{6,10}LV\d{3,5})")
# BUG RÉEL CORRIGÉ (2e lot, fichier multi-fournisseurs) : le 2e séparateur
# lu "." au lieu de "/" par l'OCR (ex. "25/08.2026" au lieu de
# "25/08/2026", sur une pièce par ailleurs propre) — les deux séparateurs
# acceptent désormais "/" OU "." indifféremment, même famille de tolérance
# que les confusions de ponctuation déjà rencontrées chez ce fournisseur.
MOTIF_DATE_BL_RAVATE = re.compile(r"(\d{1,2})[/.](\d{2})[/.](\d{4})")
# Repli si le code "...LV####" est trop abîmé par l'OCR pour être reconnu
# (cas réel : "QQ312608LV0Q48", Q à la place de 0 à plusieurs endroits) :
# une cellule ISOLÉE (ligne entière) au format JJ/MM/AAAA, préfixe non-chiffre
# optionnel toléré (le ":" de "Date :" parfois collé devant, cf. "：06/08/2026").
# Exclut naturellement la ligne "BC n°... AU <date>..." (jamais isolée).
MOTIF_DATE_ISOLEE_BL_RAVATE = re.compile(r"^\D{0,2}(\d{1,2})[/.](\d{2})[/.](\d{4})$")
MOTIF_TOTAL_HT_BL_RAVATE = re.compile(r"TOTALHT(\d[\d\s]*[,.]\d{2})")
# BUG RÉEL CORRIGÉ (2e lot du jour) : la virgule décimale peut aussi être
# lue ":" par l'OCR ("0:00" au lieu de "0,00", sur une ligne par ailleurs
# propre) — même famille que la confusion "*" déjà tolérée. Sans risque de
# confondre avec le cas COMBO ci-dessous (2 montants séparés par ":") : la
# partie après les 2 décimales ne peut alors JAMAIS tenir dans les 3
# caractères de fin autorisés ici, le motif échoue proprement et le repli
# sur MOTIF_ARGENT_COMBO_BL_RAVATE prend le relais.
MOTIF_ARGENT_BL_RAVATE = re.compile(r"^(\d[\d\s]*[,.*:]\d{2})[:.\d]{0,3}$")
# Cas réel (commande M3.10.182) : Px Net et Remises collés dans UNE SEULE
# cellule séparés par ":" et SANS espace ("6,75:1288,30") — même famille que
# le bug remise/Px net Cominter (voir moteur/fournisseurs/cominter.py) mais
# ici les DEUX valeurs sont déjà des montants complets, pas un %. Extrait la
# 1ère (Px Net, la seule utile) ; la 2e (Remises) n'est jamais conservée.
MOTIF_ARGENT_COMBO_BL_RAVATE = re.compile(r"^(\d[\d\s]*[,.*]\d{2})\s*:\s*\d[\d\s]*[,.*]\d{2}\s*$")
# --- fin GABARIT BL -----------------------------------------------------------


def _sans_espaces_bl_ravate(s: str) -> str:
    return re.sub(r"\s+", "", s.upper())


def _argent_bl_ravate(cellule: str):
    cellule = cellule.strip()
    m = MOTIF_ARGENT_BL_RAVATE.match(cellule)
    if m:
        return _to_float(m.group(1).replace("*", ",").replace(":", ","))
    m = MOTIF_ARGENT_COMBO_BL_RAVATE.match(cellule)
    if m:
        return _to_float(m.group(1).replace("*", ","))
    return None


def _zone_tableau_bl_ravate(lignes_groupees: list[list[dict]]) -> list[list[dict]]:
    """Bornée par en-tête/pied recherchés sur la ligne VISUELLE ENTIÈRE
    (cellules jointes), pas cellule par cellule : "Reference fournisseur"
    est parfois lu en une seule cellule, parfois éclaté en plusieurs
    cellules adjacentes ("Reference" | "fournisseur") selon le scan — une
    recherche cellule par cellule ratait ce 2e cas (bug réel trouvé en
    recette, 0 ligne extraite sur 4 pièces fraîches)."""

    i_entete = next(
        (i for i, ligne in enumerate(lignes_groupees)
         if MOTIF_ENTETE_TABLEAU_BL_RAVATE.search(_sans_espaces_bl_ravate(" ".join(m["texte"] for m in ligne)))),
        None,
    )
    if i_entete is None:
        return []

    i_pied = next(
        (i for i, ligne in enumerate(lignes_groupees)
         if i > i_entete and MOTIF_PIED_TABLEAU_BL_RAVATE.search(
             _sans_espaces_bl_ravate(" ".join(m["texte"] for m in ligne))
         )),
        None,
    )

    return lignes_groupees[i_entete + 1:(i_pied if i_pied is not None else len(lignes_groupees))]


def _ligne_chiffree_bl_ravate(cellules: list[str]):
    """Une ligne CHIFFRÉE a ses 4 dernières cellules qui parsent toutes
    comme un montant — repère indépendant du nombre total de cellules
    (5 ou 7 selon que qté+unité soient inline, voir bandeau GABARIT BL)."""

    if len(cellules) < 4:
        return None

    quatre = [_argent_bl_ravate(c) for c in cellules[-4:]]
    if not any(v is None for v in quatre):
        pu_brut, pu_net, remise, montant = quatre
        return pu_brut, pu_net, remise, montant

    # Repli (cas réel, commande M3.10.182) : qté+unité collées dans UNE
    # SEULE cellule juste avant Px Brut (ex. ":MT:130,00:") — décale la
    # fenêtre des 4 dernières cellules d'un cran, la 4e en partant de la
    # fin n'est alors PLUS Px Brut mais cette cellule qté+unité (jamais un
    # montant valide). Les 3 DERNIÈRES cellules (Px Brut, Px Net, Montant)
    # suffisent dans ce cas — Remises non calculable, jamais utilisée que
    # pour la vérification de cohérence documentée plus haut.
    if len(cellules) >= 3:
        trois = [_argent_bl_ravate(c) for c in cellules[-3:]]
        if not any(v is None for v in trois):
            pu_brut, pu_net, montant = trois
            return pu_brut, pu_net, None, montant

    return None


def _ligne_bl_vers_article_ravate(cellules: list[str], designation: str, ref_fournisseur_isolee: str) -> LigneBL | None:

    chiffres = _ligne_chiffree_bl_ravate(cellules)
    if chiffres is None:
        return None

    pu_brut, prix_net, remise, montant = chiffres
    if not montant or not prix_net:
        return None

    tete = cellules[0].strip()
    if MOTIF_CODE_ART_BL_RAVATE.match(tete):
        # Code Art (interne) en tête de la ligne chiffrée -> la vraie
        # Référence fournisseur vient d'une ligne isolée voisine (voir
        # bandeau GABARIT BL) ; à défaut, mieux vaut garder le Code Art
        # que de perdre la ligne (règle d'or : jamais deviner, mais jamais
        # perdre une ligne silencieusement non plus).
        reference = ref_fournisseur_isolee or tete
    else:
        reference = tete

    return LigneBL(
        reference_fournisseur=reference,
        designation=designation.strip(),
        quantite_livree=round(montant / prix_net, 4),
        prix_net=prix_net,
        montant=montant,
    )


def parse_bl_ravate(mots_par_page: list[list[dict]]) -> BonLivraison:

    lignes_plates = [
        " ".join(m["texte"] for m in ligne)
        for mots in mots_par_page
        for ligne in regrouper_lignes(mots)
    ]
    texte = "\n".join(lignes_plates)

    numero_commande = ""
    for ligne in lignes_plates:
        m = MOTIF_BC_COMMANDE_BL_RAVATE.search(ligne.strip())
        if m:
            # BUG RÉEL CORRIGÉ : le tiret ("131-162" au lieu de "131.162",
            # cas réel) doit être normalisé en point comme chez les autres
            # fournisseurs — jamais fait avant, la commande restait
            # introuvable dans le Suivi. "B" -> "8" pour les OCR corrompus
            # au même endroit (voir MOTIF_BC_COMMANDE_BL_RAVATE).
            numero_commande = m.group(1).upper().replace("B", "8").replace("-", ".")
            break

    # N° de BL et date de livraison : recherchés dans une fenêtre de lignes
    # autour du code "...LV####" (voir bandeau GABARIT BL) — la date peut
    # être sur la MÊME ligne visuelle que ce code, ou sur celle juste
    # avant, selon le scan (2 cas réels distincts observés).
    numero_bl, date_bl = "", ""
    for i, ligne in enumerate(lignes_plates):
        m = MOTIF_NUMERO_BL_RAVATE.search(ligne)
        if not m:
            continue
        numero_bl = m.group(1).upper()
        for candidate in lignes_plates[max(0, i - 1):i + 1]:
            m_date = MOTIF_DATE_BL_RAVATE.search(candidate)
            if m_date:
                date_bl = f"{int(m_date.group(1)):02d}/{m_date.group(2)}/{m_date.group(3)}"
                break
        break

    if not date_bl:
        for ligne in lignes_plates:
            m_date = MOTIF_DATE_ISOLEE_BL_RAVATE.match(ligne.strip())
            if m_date:
                date_bl = f"{int(m_date.group(1)):02d}/{m_date.group(2)}/{m_date.group(3)}"
                break

    total_ht_affiche = None
    m = MOTIF_TOTAL_HT_BL_RAVATE.search(_sans_espaces_bl_ravate(texte))
    if m:
        total_ht_affiche = _to_float(m.group(1))

    articles = []
    for mots in mots_par_page:

        lignes_zone = _zone_tableau_bl_ravate(regrouper_lignes(mots))

        designation_en_attente = ""
        ref_fournisseur_isolee = ""

        for i, ligne_mots in enumerate(lignes_zone):

            cellules = [m["texte"] for m in ligne_mots]
            ligne_jointe = " ".join(cellules)

            article = _ligne_bl_vers_article_ravate(cellules, designation_en_attente, ref_fournisseur_isolee)
            if article:
                if MOTIF_CODE_ART_BL_RAVATE.match(article.reference_fournisseur):
                    # BUG RÉEL CORRIGÉ (recette, commande 123.095) : à
                    # défaut de Référence fournisseur isolée PRÉCÉDENTE
                    # (voir _ligne_bl_vers_article_ravate), elle peut aussi
                    # se trouver sur la ligne SUIVANTE (ex. "R2V3G2.5T1 |
                    # :MT:100.00", cas réel) — jamais vue avant celle-ci
                    # sur ce document précis. Cherche 1 seule ligne en
                    # avant, jamais plus loin (pas de "presque").
                    if i + 1 < len(lignes_zone):
                        cellules_suivantes = [m["texte"] for m in lignes_zone[i + 1]]
                        if (
                            len(cellules_suivantes) >= 1
                            and MOTIF_REF_FOURNISSEUR_BL_RAVATE.match(cellules_suivantes[0].strip())
                            and not MOTIF_CODE_ART_BL_RAVATE.match(cellules_suivantes[0].strip())
                        ):
                            article.reference_fournisseur = cellules_suivantes[0].strip()
                articles.append(article)
                designation_en_attente = ""
                ref_fournisseur_isolee = ""
            elif MOTIF_BC_COMMANDE_BL_RAVATE.search(ligne_jointe.strip()):
                # Ligne repère "BC n°... AU <date> <commande>" — jamais une
                # désignation (bug réel trouvé en recette : accumulée à
                # tort en tête de la désignation du 1er article).
                designation_en_attente = ""
                ref_fournisseur_isolee = ""
            elif (
                len(cellules) == 1
                and MOTIF_REF_FOURNISSEUR_BL_RAVATE.match(cellules[0].strip())
                and not MOTIF_CODE_ART_BL_RAVATE.match(cellules[0].strip())
            ):
                # Référence fournisseur isolée sur sa propre ligne (voir
                # bandeau GABARIT BL) — mémorisée pour la prochaine ligne
                # chiffrée dont la tête serait un Code Art.
                ref_fournisseur_isolee = cellules[0].strip()
            else:
                designation_en_attente = (designation_en_attente + " " + ligne_jointe).strip()

    bl = BonLivraison(
        fournisseur="RAVATE",
        fichier="",
        numero_bl=numero_bl,
        date_bl=date_bl,
        numero_commande=numero_commande,
        lignes=articles,
        total_ht_affiche=total_ht_affiche,
    )

    if total_ht_affiche is not None:
        total_extrait = round(sum(a.montant for a in articles), 2)
        if abs(total_ht_affiche - total_extrait) > 0.02:
            print(
                f"!! RAVATE (BL) : Total HT du PDF ({total_ht_affiche:.2f}€) "
                f"!= somme des lignes extraites ({total_extrait:.2f}€) "
                f"— une ligne a peut-être été oubliée ou mal lue par l'OCR."
            )

    return bl


FOURNISSEURS = ["RAVATE"]
parse = parse_ravate
parse_bl = parse_bl_ravate
