"""
Parser COMINTER (et COMINTER MAYOTTE).

TROIS formats réels coexistent chez ce fournisseur (le texte extrait du
PDF éclate chaque champ sur sa propre ligne, mais pas dans le même ordre) :

v1 : Référence, Désignation (1-2 lignes), Montant HT €, Quantité,
     Prix Unit € [remise %], Code TVA (chiffre), Prix Net €,
     Conditionnement (Unité / Boîte de 1 / Sachet de / Barre).

v2 : Qté, Px unitaire, Montant HT€, [code TVA], Désignation (1-2 lignes),
     Cdt (Unité/Barre...), Référence — la référence vient APRÈS le Cdt.

v3 (chantier Kanopée CDC, devis DV121328/DV124395) : Qté, Px unitaire,
     Montant HT€, Rem% (1 chiffre), Référence, Désignation (1-2 lignes),
     Cdt — la référence vient AVANT la désignation, contrairement à v2.
     parse_cominter() essaie les trois et garde le plus complet ; sur ces
     2 PDF réels, le total extrait retombe exactement sur le Total HT
     affiché (6 010,49€ et 9 515,39€).
"""

import re

from moteur.modele import Article
from moteur.ocr import pages_par_identifiant, regrouper_lignes
from moteur.outils import to_float
from moteur.rapprochement.modele_bl import BonLivraison, LigneBL

# --- GABARIT (Cominter / Cominter Mayotte) ---------------------------------
# Deux formats de PDF coexistent chez Cominter (v1, v2 ci-dessous) : une
# logique procédurale réelle de "quel format essayer", pas une donnée — elle
# reste en Python explicite (voir CLAUDE.md). Constantes alignées sur le
# même bandeau que les autres fournisseurs, par cohérence de lecture.
_REF = re.compile(r"^[A-Z][A-Z0-9/\-]{3,}$")
_MONEY = re.compile(r"^\d[\d\s]*,\d{2}\s*€$")
_MONEY_REM = re.compile(r"^(\d[\d\s]*,\d{2})\s*€(?:\s+(\d+)%)?$")
_QTE = re.compile(r"^\d[\d\s]*,\d{2}$")

_UNITES = {
    "UNITÉ": "UN",
    "UNITE": "UN",
    "BOÎTE": "BTE",
    "BOITE": "BTE",
    "SACHET": "SCH",
    "BARRE": "BARRE",
}
# --- fin GABARIT (suite : _CDT et les deux stratégies plus bas) ------------


def _f(v: str) -> float:
    return float(v.replace(" ", "").replace("\u202f", "").replace("€", "").replace(",", "."))


def _unite(ligne: str) -> str:

    mot = ligne.strip().upper().split()

    if not mot:
        return ""

    return _UNITES.get(mot[0], mot[0])


def _parse_cominter_v1(texte: str) -> list[Article]:

    articles = []

    devis = ""

    # 3 lettres habituellement ("ODE270211"), mais 2 sur les PDF v3
    # ("DV121328") — cas réel, chantier Kanopée CDC.
    m = re.search(r"Devis\s*:\s*([A-Z]{2,4}\d+)", texte)

    if m:
        devis = m.group(1)

    lignes = [l.rstrip() for l in texte.splitlines()]

    i = 0
    n = len(lignes)

    while i < n:

        ligne = lignes[i].strip()

        # Une référence candidate, hors éco-taxes
        if not _REF.match(ligne) or ligne.startswith("ECO"):
            i += 1
            continue

        ref = ligne

        # Désignation : lignes jusqu'au montant (max 4 lignes)
        j = i + 1
        designation = []

        while j < n and j <= i + 4 and not _MONEY.match(lignes[j].strip()):
            designation.append(lignes[j].strip())
            j += 1

        if j >= n or not _MONEY.match(lignes[j].strip()):
            i += 1
            continue

        try:

            montant = _f(lignes[j].strip())

            quantite = _f(lignes[j + 1].strip())

            m_prix = _MONEY_REM.match(lignes[j + 2].strip())
            prix_brut = _f(m_prix.group(1)) if m_prix else 0.0

            # lignes[j+3] = code TVA (1 chiffre)
            prix_net = _f(lignes[j + 4].strip())

            unite = _unite(lignes[j + 5]) if j + 5 < n else ""

        except (ValueError, AttributeError, IndexError):
            i += 1
            continue

        articles.append(
            Article(
                fournisseur="COMINTER",
                devis=devis,
                reference_fournisseur=ref,
                reference_distributeur="",
                designation=" ".join(designation),
                quantite=quantite,
                unite=unite,
                prix_brut=prix_brut,
                prix_net=prix_net,
                montant=montant,
            )
        )

        i = j + 5

    return articles



_CDT = ("Unité", "Barre", "Boîte de 1", "Sachet de", "Rouleau", "Mètre")


def _parse_cominter_v2(texte: str) -> list[Article]:
    """
    Nouveau format Cominter : par article, ordre
        Qté, Px unitaire, Montant HT€, [code TVA], Désignation(1-2 l),
        Cdt (Unité/Barre...), Référence
    On s'ancre sur la ligne Cdt : la référence suit, la désignation et le
    bloc prix précèdent.
    """
    articles = []

    m = re.search(r"Devis\s*:\s*([A-Z]{2,4}\d+)", texte)
    devis = m.group(1) if m else ""

    lignes = [l.rstrip() for l in texte.splitlines()]
    n = len(lignes)

    for i, l in enumerate(lignes):

        if l.strip() not in _CDT and not l.strip().startswith(("Boîte", "Sachet", "Barre", "Unité", "Rouleau")):
            continue

        if i + 1 >= n:
            continue

        ref = lignes[i + 1].strip()
        if not _REF.match(ref) or ref.startswith("ECO"):
            continue

        # Remonter : désignation jusqu'au montant €
        j = i - 1
        design = []
        montant = None
        while j >= 0 and j > i - 6:
            lg = lignes[j].strip()
            if _MONEY.match(lg):
                montant = _f(lg)
                break
            design.append(lg)
            j -= 1

        if montant is None or j < 2:
            continue

        # Une éventuelle ligne "code TVA" (1 chiffre) juste après le montant
        if design and re.fullmatch(r"\d", design[-1]):
            design.pop()

        try:
            prix = _f(lignes[j - 1].strip())
            quantite = _f(lignes[j - 2].strip())
        except (ValueError, IndexError):
            continue

        articles.append(
            Article(
                fournisseur="COMINTER",
                devis=devis,
                reference_fournisseur=ref,
                reference_distributeur="",
                designation=" ".join(reversed(design)),
                quantite=quantite,
                unite=_unite(lignes[i]),
                prix_brut=prix,
                prix_net=prix,
                montant=montant,
            )
        )

    return articles


def _parse_cominter_v3(texte: str) -> list[Article]:
    """
    3e format Cominter réel (chantier Kanopée CDC, devis DV121328/DV124395) :
    par article, ordre
        Qté, Px unitaire, Montant HT€, Rem% (1 chiffre), Référence,
        Désignation (1-2 lignes), Cdt (Unité/Barre...)
    — la référence vient AVANT la désignation (contrairement à v2, où elle
    vient après le Cdt). On s'ancre sur la ligne Cdt et on remonte : la
    désignation (1 à 2 lignes) jusqu'à la référence, puis le bloc prix.
    """
    articles = []

    m = re.search(r"Devis\s*:\s*([A-Z]{2,4}\d+)", texte)
    devis = m.group(1) if m else ""

    lignes = [l.rstrip() for l in texte.splitlines()]
    n = len(lignes)

    for i, l in enumerate(lignes):

        if l.strip() not in _CDT and not l.strip().startswith(("Boîte", "Sachet", "Barre", "Unité", "Rouleau")):
            continue

        # Remonter jusqu'à la référence (1 à 2 lignes de désignation)
        j = i - 1
        design = []
        while j >= 0 and j > i - 4 and not _REF.match(lignes[j].strip()):
            design.insert(0, lignes[j].strip())
            j -= 1

        if j < 0 or not _REF.match(lignes[j].strip()) or lignes[j].strip().startswith("ECO"):
            continue

        ref = lignes[j].strip()

        # Avant la référence : Rem% (1 chiffre), Montant€, Px unitaire, Qté
        if j < 4:
            continue

        code_rem = lignes[j - 1].strip()
        if not re.fullmatch(r"\d+%?", code_rem):
            continue

        try:
            if not _MONEY.match(lignes[j - 2].strip()):
                continue
            montant = _f(lignes[j - 2].strip())
            prix = _f(lignes[j - 3].strip())
            quantite = _f(lignes[j - 4].strip())
        except (ValueError, IndexError):
            continue

        articles.append(
            Article(
                fournisseur="COMINTER",
                devis=devis,
                reference_fournisseur=ref,
                reference_distributeur="",
                designation=" ".join(design),
                quantite=quantite,
                unite=_unite(lignes[i]),
                prix_brut=prix,
                prix_net=prix,
                montant=montant,
            )
        )

    return articles


def parse_cominter(texte: str) -> list[Article]:
    """Essaie les trois formats et retourne le plus complet."""
    v1 = _parse_cominter_v1(texte)
    v2 = _parse_cominter_v2(texte)
    v3 = _parse_cominter_v3(texte)
    return max((v1, v2, v3), key=len)



# --- GABARIT BL (Cominter Ouest) --------------------------------------------
# BL scanné (image pure, comme les autres fournisseurs — voir moteur/ocr.py).
#
# PIÈGE réel signalé par l'acheteur AVANT tout code, décisif ici : un même
# fichier PDF peut contenir PLUSIEURS bons de livraison scannés à la suite
# (jusqu'à 8 vus en session R2 suite), chacun avec son propre n° "OBL......"
# — mais un même BL peut aussi déborder sur 2 pages (le tableau d'articles
# tient sur la 1ère, le pied de page des totaux sur la 2e, MÊME n° OBL sur
# les deux). `grouper_pages_par_identifiant()` (moteur/ocr.py, générique)
# regroupe donc les pages par n° OBL AVANT tout parsing : une page sans OBL
# détecté (le pied de page qui déborde) rejoint le groupe précédent plutôt
# que d'ouvrir un nouveau groupe à tort.
#
# Structure de chaque BL : "Bon de livraison : OBL......" (n° de document),
# un champ "Référence" (n° de commande acheteur, ex. "M3.23.030", parfois
# précédé d'un mot de chantier genre "LAGOURGUE 131.155" — cherché par
# recherche libre, pas ancrage strict, pour ignorer ce préfixe), puis un
# tableau Référence/Désignation/Qté+Cdt (souvent COLLÉS en une seule
# cellule OCR, ex. "22,00 Unité")/Px unitaire/Rem%(optionnel)/Px net/
# Montant HT/[code TVA isolé optionnel]. Comme Coredime, le nombre de
# cellules par ligne varie (Rem% absent si 0%) : ancrage sur le MONTANT
# (dernière cellule "argent", ou avant-dernière si un code TVA isolé
# traîne derrière), Px net juste avant — la quantité livrée en est déduite
# (Montant / Px net) plutôt que lue dans la cellule Qté+Cdt collée.
#
# Pas d'autocontrôle Total HT ici (contrairement à 109 Distribution/
# Electric Plus) : le tableau de répartition TVA en pied de page a une
# structure trop irrégulière (colonnes qui se décalent selon le nombre de
# taux de TVA présents) pour extraire la valeur de façon fiable sans plus
# d'exemples réels — mieux vaut l'omettre que de calculer un total faux.
MOTIF_BL_COMINTER = re.compile(r"[0O]BL\s*(\d{5,7})", re.IGNORECASE)
MOTIF_ENTETE_TABLEAU_BL_COMINTER = re.compile(r"DESIGNATION")
MOTIF_PIED_TABLEAU_BL_COMINTER = re.compile(r"ARTICLE7")
# BUG RÉEL CORRIGÉ (session R2 suite, recette réelle) : l'OCR colle parfois
# le mot de chantier au n° de commande SANS aucun espace, ex.
# "LAGOURGUE131.155" (un seul token). L'ancien motif ("[A-Z]?\d..." avec
# préfixe libre) captait alors à tort la DERNIÈRE lettre du mot précédent
# comme si c'était le préfixe du n° de commande ("E131.155" au lieu de
# "131.155") — la commande "131.155" en devenait introuvable dans le
# Suivi, tout le BL correspondant perdu. Le préfixe lettre (chantiers du
# style "M3.23.030") n'est accepté que s'il n'est PAS précédé d'une autre
# lettre (véritable début de mot) ; sinon on ne capture QUE les chiffres.
MOTIF_COMMANDE_BL_COMINTER = re.compile(
    r"((?:(?<![A-Za-z])[A-Z]\d{1,4}|(?<!\d)\d{1,4})(?:[.\-]\d{1,4}){1,2})"
)
MOTIF_MONTANT_BL_COMINTER = re.compile(r"^(\d[\d\s]*[,.]\d{2})\s*E?\s*\d?$", re.IGNORECASE)
MOTIF_REF_ARTICLE_BL_COMINTER = re.compile(r"^(?=[A-Z0-9./\-]*\d)[A-Z0-9./\-]{4,15}$")
# Date au format JJ/MM/AA OU JJ/MM/AAAA — l'année est tantôt sur 2
# chiffres tantôt sur 4 SELON LE SCAN (cas réel : 2 documents du même lot,
# "06/08/26" et "04/08/2026") — cellule "Date" de l'en-tête (Numero/Date/
# Fin de Validité), juste après "Numero".
MOTIF_DATE_BL_COMINTER = re.compile(r"^(\d{1,2})/(\d{2})/(\d{2}|\d{4})$")
# --- fin GABARIT BL -----------------------------------------------------------


def _sans_espaces_bl_cominter(s: str) -> str:
    return re.sub(r"\s+", "", s.upper())


def _argent_bl_cominter(cellule: str):
    m = MOTIF_MONTANT_BL_COMINTER.match(cellule.strip())
    return to_float(m.group(1)) if m else None


def _zone_tableau_bl_cominter(lignes_groupees: list[list[dict]]) -> list[list[dict]]:

    i_entete = next(
        (i for i, ligne in enumerate(lignes_groupees)
         if any(MOTIF_ENTETE_TABLEAU_BL_COMINTER.search(_sans_espaces_bl_cominter(m["texte"])) for m in ligne)),
        None,
    )
    if i_entete is None:
        return []

    i_pied = next(
        (i for i, ligne in enumerate(lignes_groupees)
         if i > i_entete and any(
             MOTIF_PIED_TABLEAU_BL_COMINTER.search(_sans_espaces_bl_cominter(m["texte"])) for m in ligne
         )),
        None,
    )

    return lignes_groupees[i_entete + 1:(i_pied if i_pied is not None else len(lignes_groupees))]


def _prix_net_bl_cominter(cellules: list[str], fin: int):
    """Px net = normalement la cellule juste avant Montant. Deux cas réels
    où elle n'est pas directement lisible :
    - remise % et Px net collés dans la MÊME cellule OCR, AVEC ou SANS
      espace entre le "%" et le prix (ex. "30% 110,67" mais aussi "30%
      435,94" — cas réel, BL ANZEMBERG.pdf/OBL108110, ligne CAELK2766,
      signalé par l'acheteur : sans l'espace optionnel, le motif ne
      matchait jamais et la boucle retombait à tort sur le Px UNITAIRE de
      la cellule précédente, 622,77€ au lieu du vrai Px net 435,94€ — bug
      silencieux, aucune anomalie levée car montant/qté restaient
      cohérents avec le mauvais prix) — le Px net est là, juste après le
      taux ;
    - Px net absent, seule la remise % subsiste (ex. "L600321 ...
      3,00/Unite 26,78 30% 56,24") : reconstruit un prix net effectif à
      partir du Px unitaire (juste avant le taux) et du taux, vérifié
      exact sur un vrai cas (26,78 x 0,70 = 18,746 ; 3 x 18,746 = 56,24,
      montant affiché)."""

    k = fin - 1

    while k >= 1:

        cellule = cellules[k].strip()

        m_combo = re.fullmatch(r"\d+\s*%\s*(\d[\d\s]*[,.]\d{2})", cellule)
        if m_combo:
            return to_float(m_combo.group(1))

        m_pct = re.fullmatch(r"(\d+)\s*%", cellule)
        if m_pct:
            if k - 1 >= 1:
                prix_unitaire = _argent_bl_cominter(cellules[k - 1])
                if prix_unitaire is not None:
                    return round(prix_unitaire * (1 - to_float(m_pct.group(1)) / 100), 4)
            return None

        val = _argent_bl_cominter(cellule)
        if val is not None:
            return val

        k -= 1

    return None


def _quantite_bl_cominter(cellules: list[str], j: int, montant: float, prix_net: float):
    """Quantité livrée : PRIORITÉ à la cellule Qté+Unité imprimée (ex.
    "30,00 Unite" -> 30.0) — c'est la vraie valeur, un compte d'unités
    entières pour la plupart des références (interrupteurs, prises...).

    BUG RÉEL CORRIGÉ (session R2 suite, recette réelle) : cette fonction
    calculait AUPARAVANT la quantité via Montant / Px net (comme pour 109
    Distribution/Electric Plus, où le Px net affiché est fiable) — mais
    chez Cominter, avec une remise, le Px net affiché est ARRONDI à la
    ligne alors que le Montant semble calculé à partir d'une valeur plus
    précise (ex. 30 x 4,165 = 124,95, mais Px net affiché "4,17" ->
    124,95 / 4,17 = 29,9640...). Résultat concret écrit dans le Suivi
    commandes vivant avant correction : "29,96 interrupteur(s)" au lieu de
    30, "149,89" au lieu de 150 — une quantité non entière n'a aucun sens
    pour ce genre d'article et a été repérée par l'acheteur. Montant / Px
    net ne sert plus qu'en dernier repli, si la cellule Qté est absente."""

    if j < len(cellules):
        m = re.match(r"^(\d+(?:[,.]\d+)?)", cellules[j].strip())
        if m:
            try:
                return float(m.group(1).replace(",", "."))
            except ValueError:
                pass

    if prix_net:
        return round(montant / prix_net, 2)

    return None


def _ligne_bl_vers_article_cominter(cellules: list[str]) -> LigneBL | None:

    if not cellules:
        return None

    reference = cellules[0].strip()

    if reference.upper().startswith("ECO") or not MOTIF_REF_ARTICLE_BL_COMINTER.match(reference):
        return None

    fin = len(cellules) - 1
    if fin >= 1 and re.fullmatch(r"\d", cellules[fin].strip()):
        fin -= 1  # code TVA isolé en bout de ligne, pas le montant

    if fin < 2:
        return None

    montant = _argent_bl_cominter(cellules[fin])
    prix_net = _prix_net_bl_cominter(cellules, fin)

    if not montant or not prix_net:
        return None

    # BUG RÉEL CORRIGÉ (recette réelle, M2.17.006/OBL108540,
    # M3.23.042/OBL108537 et M4.272/OBL108653) : la capacité de coupure du
    # disjoncteur DNX³ ("4.5KA" ou "4.5 KA", partie de la désignation) se
    # retrouve parfois dans SA PROPRE cellule OCR au lieu de rester collée
    # au reste de la désignation — `re.match(r"^\d", ...)` (qui accepte
    # tout ce qui COMMENCE par un chiffre) la prenait alors à tort pour la
    # cellule Qté(+Unité), décalant tout le reste d'une cellule. Résultat
    # concret déjà écrit dans le Suivi avant correction : "4,5"
    # disjoncteurs livrés au lieu de 10/20 (la vraie quantité, toujours à
    # la cellule suivante) — repéré par l'acheteur ("il ne peut y avoir
    # que des entiers en quantité"). DEUX tentatives précédentes cassées,
    # gardées en mémoire pour ne pas y retomber :
    # - `fullmatch` sur 2 décimales : rejette aussi le format M4.272, où
    #   Qté+Unité sont DANS LA MÊME cellule ("7,00 Unite" — du texte suit
    #   les 2 décimales, un fullmatch échoue).
    # - Exiger une VIRGULE (pas un point) : rejette à tort une vraie
    #   quantité écrite avec un point ("3.00 Unite", bl_cominter_3.pdf) —
    #   virgule/point ne distinguent PAS de façon fiable une vraie
    #   quantité de "4.5KA" (les deux séparateurs existent des deux
    #   côtés selon les documents).
    # Signal qui tient sur TOUS les cas réels observés (2 fournisseurs
    # confondus, virgule ET point) : une vraie cellule Qté a TOUJOURS
    # exactement 2 chiffres après le séparateur ("10,00", "3.00", "542,00"
    # — jamais "10,0" ni "10,000"), alors que "4.5KA"/"4.5 KA" n'en a
    # qu'UN SEUL ("5"). `re.match` (préfixe, pas `fullmatch`) pour
    # continuer à accepter tout ce qui suit collé ou espacé (unité,
    # "/Unite", TVA glissée...).
    j = 1
    while j < fin - 1 and not re.match(r"^\d+[,.]\d{2}", cellules[j].strip()):
        j += 1
    designation = " ".join(c.strip() for c in cellules[1:j]).strip()

    quantite = _quantite_bl_cominter(cellules, j, montant, prix_net)
    if not quantite:
        return None

    return LigneBL(
        reference_fournisseur=reference,
        designation=designation,
        quantite_livree=quantite,
        prix_net=prix_net,
        montant=montant,
    )


def _parse_un_bl_cominter(mots_par_page_groupe: list[list[dict]]) -> BonLivraison:

    lignes_plates = [
        mot["texte"]
        for mots in mots_par_page_groupe
        for ligne in regrouper_lignes(mots)
        for mot in ligne
    ]
    texte = "\n".join(lignes_plates)

    numero_bl = ""
    m = MOTIF_BL_COMINTER.search(texte)
    if m:
        numero_bl = f"OBL{m.group(1)}"

    # N° de commande : cherché entre l'en-tête "Numero/Date/Fin de
    # Validité" et l'en-tête du tableau d'articles ("Designation") —
    # recherche libre (pas ancrage strict) pour ignorer un éventuel mot de
    # chantier devant (ex. "LAGOURGUE 131.155").
    i_numero = next(
        (i for i, l in enumerate(lignes_plates) if _sans_espaces_bl_cominter(l).startswith("NUMERO")),
        None,
    )
    i_designation = next(
        (i for i, l in enumerate(lignes_plates) if "DESIGNATION" in _sans_espaces_bl_cominter(l)),
        None,
    )

    numero_commande = ""
    date_bl = ""
    if i_numero is not None and i_designation is not None:
        for l in lignes_plates[i_numero + 1:i_designation]:
            if not numero_commande:
                m = MOTIF_COMMANDE_BL_COMINTER.search(l)
                if m:
                    numero_commande = m.group(1).upper().replace(" ", ".")
            if not date_bl:
                m_date = MOTIF_DATE_BL_COMINTER.match(l.strip())
                if m_date:
                    jour, mois, annee = m_date.groups()
                    annee = annee if len(annee) == 4 else f"20{annee}"
                    date_bl = f"{int(jour):02d}/{mois}/{annee}"

    articles = []
    for mots in mots_par_page_groupe:

        lignes_zone = _zone_tableau_bl_cominter(regrouper_lignes(mots))
        i = 0

        while i < len(lignes_zone):

            cellules = [m["texte"] for m in lignes_zone[i]]
            i += 1

            reference_candidate = cellules[0].strip() if cellules else ""
            reference_valide = bool(
                cellules
                and not reference_candidate.upper().startswith("ECO")
                and MOTIF_REF_ARTICLE_BL_COMINTER.match(reference_candidate)
            )

            # BUG RÉEL CORRIGÉ (session R2 suite, recette réelle) : sur une
            # désignation longue, l'OCR renvoie parfois la référence sur
            # sa PROPRE ligne, APRÈS la ligne désignation+qté+prix (au lieu
            # d'être en 1ère cellule de cette même ligne) — cas réel
            # "PLW11643" : la ligne courante n'a pas de référence valide,
            # mais la ligne suivante est une référence isolée (1 seule
            # cellule). Sans ce raccord, la ligne (et sa quantité livrée)
            # disparaissait silencieusement — un article réellement livré
            # n'était alors jamais écrit dans le Suivi.
            if not reference_valide and i < len(lignes_zone):
                cellules_suivantes = [m["texte"] for m in lignes_zone[i]]
                if len(cellules_suivantes) == 1:
                    ref_suivante = cellules_suivantes[0].strip()
                    if (
                        not ref_suivante.upper().startswith("ECO")
                        and MOTIF_REF_ARTICLE_BL_COMINTER.match(ref_suivante)
                    ):
                        cellules = [ref_suivante] + cellules
                        i += 1

            article = _ligne_bl_vers_article_cominter(cellules)
            if article:
                articles.append(article)

    return BonLivraison(
        fournisseur="COMINTER",
        fichier="",
        numero_bl=numero_bl,
        date_bl=date_bl,
        numero_commande=numero_commande,
        lignes=articles,
        total_ht_affiche=None,
    )


def parse_bl_cominter(mots_par_page: list[list[dict]]) -> list[BonLivraison]:
    """Un même fichier peut contenir PLUSIEURS BL scannés à la suite (voir
    bandeau GABARIT BL) : retourne une liste, une entrée par BL détecté.
    Chaque BonLivraison porte aussi les indices de page (0-based) qu'il
    occupe dans le fichier source (bl.pages) — utilisé par
    moteur.rapprochement.pipeline_bl pour archiver chaque BL
    individuellement (découpage PDF), sans attendre que TOUS les BL du
    même fichier soient résolus."""

    groupes_indices = pages_par_identifiant(mots_par_page, MOTIF_BL_COMINTER)

    resultat = []
    for indices in groupes_indices:
        bl = _parse_un_bl_cominter([mots_par_page[i] for i in indices])
        bl.pages = indices
        resultat.append(bl)

    return resultat


# "COMINTER MAYOTTE" est géré par moteur/fournisseurs/cominter_mayotte.py :
# entité distincte (SIRET, adresse, structure de devis différents), voir
# le bandeau GABARIT de ce module.
FOURNISSEURS = ['COMINTER']
parse = parse_cominter
parse_bl = parse_bl_cominter
