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
from moteur.rapprochement.modele_facture import Facture, LigneFacture

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


# --- GABARIT FACTURE (Cominter Ouest / Sainte-Clotille / Saint-Pierre) -----
# Session F4 suite (2026-09-01), 132 pièces réelles reçues par mail de
# Prisca LEBLÉ (comptable), même montage que Coredime : plusieurs .msg
# contenant chacun des e-mails imbriqués, un par facture. Sur ce lot : ~80
# vraies factures natif-texte (JAMAIS de scan chez ce fournisseur pour la
# facture, contrairement à son BL), 12 "COMINTER MAYOTTE" (format
# différent, PAS couvert ici — voir cominter_mayotte.py, à étendre
# séparément), et ~40 pièces jointes ANNEXES (scans purs, 0 caractère :
# "xxxxxx_BL_OBLxxxxxx.pdf"/"xxxxxx_MAN.pdf" — BL/manifeste papier
# scannés) — HORS PÉRIMÈTRE de ce parser. Piège réel repéré AVANT tout
# code : certains "xxxxxx_MAN.pdf" ont en fait du texte natif MAIS ne
# contiennent QUE notre propre "DETAIL DE LA COMMANDE"/"BON DE COMMANDE"
# (aucun champ Cominter : pas de "Facture : XXX", pas de "Signature", pas
# de bloc prix) — détectés "COMINTER" (le mot apparaît dans "DESTINATAIRE
# COMINTER") mais parse_facture_cominter() n'y trouve honnêtement aucune
# ligne (pas de traitement spécial nécessaire, l'architecture tolérante
# aux pannes gère déjà ce cas comme n'importe quel autre 0-ligne).
#
# Structure réelle (texte natif PyMuPDF, ordre de lecture scramblé comme
# les devis du même fournisseur — PAS le même ordre que le BL scanné/OCR,
# mais un contenu de ligne d'article structurellement identique : Qté, Px
# unitaire, Remise% optionnelle, Montant net, Code TVA, Désignation,
# Cdt/Unité optionnel, Référence). Zone bornée entre "Signature" (juste
# avant le bloc [date livraison, n° de BL, n° de commande] puis le tableau
# d'articles) et "Article 7. PROCEDURE DE REMBOURSEMENT" (texte légal fixe,
# présent sur toutes les pièces vues) — jamais au-delà : ~55% des factures
# ont EN PLUS notre propre "DETAIL DE LA COMMANDE"/"BON DE COMMANDE"
# ajouté à la suite dans le même PDF (même piège déjà documenté chez 109
# Distribution) et il ne faut jamais le lire comme des lignes Cominter.
#
# Ancrage sur le MONTANT (seule cellule à la fois "argent" ET présente sur
# CHAQUE ligne, avec ou sans remise) plutôt que sur un Cdt/Unité — celui-ci
# est ABSENT sur les lignes d'éco-participation (ex. "Eco - contribution
# 0.12€" suivi directement de "ECO-TAXE3", sans "Unité" entre les deux,
# alors qu'un article normal a "Désignation / Unité / Référence") : ancrer
# sur Cdt (comme les devis v2/v3 du même fournisseur) aurait perdu ces
# lignes silencieusement. La désignation+référence de chaque article est
# donc bornée par les DEUX montants consécutifs (tout ce qui suit le code
# TVA de l'article N jusqu'au début du bloc prix de l'article N+1) ; la
# DERNIÈRE ligne de cette zone est toujours la référence, l'avant-dernière
# est le Cdt SI elle correspond à un mot-clé connu (Unité/Boîte/Barre/...).
#
# N° de BL : préfixe variable selon la pièce/agence — "OBL" (déjà connu
# côté BL, Comptoir Ouest) MAIS AUSSI "NBL" (Sainte-Clotinde/Saint-Pierre,
# jamais vu côté BL jusqu'ici) — motif élargi à toute forme
# LETTRES(2-5)+CHIFFRES(5-7), pas de préfixe figé en dur.
#
# PAS d'autocontrôle Total HT (même choix que le BL du même fournisseur,
# voir bandeau GABARIT BL ci-dessus) : le tableau de ventilation TVA en
# pied de page a une structure trop irrégulière pour une extraction fiable
# avec les exemples actuels — mieux vaut l'omettre que calculer un total
# faux.
#
# INCERTITUDE ASSUMÉE sur date_facture : DEUX dates apparaissent dans
# l'en-tête scramblé — une tôt dans le texte (juste après le mot "Numéro",
# avant le n° de commande client) et une plus tard, accompagnée d'une
# heure précise (HH:MM:SS) juste avant "Facture : XXX", qui ressemble
# plutôt à un horodatage d'IMPRESSION du PDF qu'à la date commerciale de
# la facture. La PREMIÈRE (sans heure associée) est retenue comme
# date_facture — à valider sur davantage de recettes réelles avant de la
# considérer acquise à 100%.
MOTIF_NUMERO_FACTURE_COMINTER = re.compile(r"Facture\s*:\s*([A-Z]+\d+)")
MOTIF_DATE_FACTURE_COMINTER = re.compile(r"\b(\d{1,2})/(\d{2})/(\d{2})\b")
MOTIF_SIGNATURE_COMINTER = re.compile(r"^Signature$")
MOTIF_FIN_TABLEAU_FACTURE_COMINTER = re.compile(r"Article\s*7\.\s*PROCEDURE|NET A PAYER")
MOTIF_BL_FACTURE_COMINTER = re.compile(r"^[A-Z]{2,5}\d{5,7}$")
# Élargi par rapport à MOTIF_COMMANDE_BL_COMINTER (séparateur espace en
# plus de point/tiret) : cas réel NFA018127.pdf, "BC N°24 3240" (le
# séparateur entre les deux groupes est un espace, pas un point comme
# "M2.22.082" — même famille que le format Cominter Mayotte "24  3155").
MOTIF_COMMANDE_FACTURE_COMINTER = re.compile(
    r"((?:(?<![A-Za-z])[A-Z]\d{1,4}|(?<!\d)\d{1,4})(?:[.\-\s]\d{1,4}){1,2})"
)
MOTIF_LABEL_DATE_COMINTER = re.compile(r"^Date$")
MOTIF_MONTANT_LIGNE_FACTURE_COMINTER = re.compile(r"^(\d[\d\s]*,\d{2})\s*€(?:\s+(\d))?$")
MOTIF_REM_FACTURE_COMINTER = re.compile(r"^(\d+)\s*%$")
MOTIF_QTE_PX_FACTURE_COMINTER = re.compile(r"^(\d[\d\s]*,\d{2})$")
MOTIF_CODE_TVA_FACTURE_COMINTER = re.compile(r"^\d$")
_CDT_FACTURE_COMINTER = {
    "unité", "unite", "boîte", "boite", "barre", "cour", "rouleau", "mètre", "metre", "sachet",
}  # comparé en minuscule (voir NF155008.pdf, agence Saint-Pierre : "unite" tout en minuscule)
# --- fin GABARIT FACTURE ----------------------------------------------------


def _zone_articles_facture_cominter(lignes: list[str]):
    """(indice_debut, indice_fin) de la zone à scanner pour les lignes
    d'articles — entre "Signature" et "Article 7. PROCEDURE...", jamais
    au-delà (voir bandeau : "DETAIL DE LA COMMANDE" = notre propre BC,
    pas des lignes Cominter).

    Repli si "Signature" est ABSENT (cas réel NF155008.pdf, agence
    Saint-Pierre — le bloc [date, n° de BL, n° de commande] démarre
    directement, sans "Signature" devant) : ancrage direct sur la 1ère
    ligne qui ressemble à un n° de BL (MOTIF_BL_FACTURE_COMINTER) — sans
    risque de faux positif, ce motif (2-5 lettres + 5-7 chiffres, rien
    d'autre sur la ligne) ne matche aucune ligne du bandeau légal/adresse
    qui précède."""

    i_signature = next((i for i, l in enumerate(lignes) if MOTIF_SIGNATURE_COMINTER.match(l.strip())), None)

    if i_signature is not None:
        debut = i_signature + 1
    else:
        i_bl = next((i for i, l in enumerate(lignes) if MOTIF_BL_FACTURE_COMINTER.match(l.strip())), None)
        if i_bl is None:
            return None
        debut = max(i_bl - 1, 0)

    i_fin = next(
        (i for i in range(debut, len(lignes)) if MOTIF_FIN_TABLEAU_FACTURE_COMINTER.search(lignes[i])),
        len(lignes),
    )

    return debut, i_fin


def _entete_bl_facture_cominter(lignes: list[str], debut: int, fin: int):
    """(numero_bl, numero_commande) lus juste après "Signature" — même
    esprit que _parse_un_bl_cominter côté BL : recherche libre (pas
    ancrage strict) pour tolérer un mot de chantier ou un préfixe "BC"/"BC
    N°" devant le n° de commande.

    Repli sur l'EN-TÊTE (ligne juste avant le label "Date", tout au début
    du document, AVANT "Signature") si rien trouvé ici — cas réel
    OFC194316.pdf : aucun n° de commande réimprimé dans le bloc Signature
    (juste [date, n° BL], contrairement à OFC193413.pdf qui réimprime la
    commande à cet endroit) ; le n° de commande n'existe alors QUE dans
    l'en-tête. Le bloc Signature reste préféré quand il en fournit un : cas
    réel NFA018127.pdf où l'en-tête tronque la commande ("BC: 3240", un
    seul groupe, "24" perdu) alors que le bloc Signature la réimprime en
    entier ("BC N°24 3240")."""

    numero_bl = ""
    numero_commande = ""

    for l in lignes[debut:min(debut + 6, fin)]:
        l = l.strip()
        if not numero_bl and MOTIF_BL_FACTURE_COMINTER.match(l):
            numero_bl = l
            continue
        if not numero_commande and numero_bl:
            m = MOTIF_COMMANDE_FACTURE_COMINTER.search(l)
            if m:
                numero_commande = m.group(1).upper().replace(" ", ".")

    if not numero_commande:
        i_date = next(
            (i for i in range(debut) if MOTIF_LABEL_DATE_COMINTER.match(lignes[i].strip())),
            None,
        )
        if i_date is not None and i_date > 0:
            m = MOTIF_COMMANDE_FACTURE_COMINTER.search(lignes[i_date - 1].strip())
            if m:
                numero_commande = m.group(1).upper().replace(" ", ".")

    return numero_bl, numero_commande


def _bloc_prix_facture_cominter(lignes: list[str], i_montant: int):
    """Depuis la ligne montant à l'indice i_montant, remonte pour lire
    [Qté, Px unitaire, Remise% optionnelle] — retourne
    (indice_debut_bloc, qte, montant, code_tva_inline) ou None si le motif
    ne correspond pas (voir bandeau : ancrage sur le montant, seule
    cellule fiable sur TOUTE ligne, avec ou sans remise, avec ou sans
    Cdt). `code_tva_inline` est le code TVA capturé sur la MÊME ligne que
    le montant (ex. "17,33 € 1") quand présent — cas réel NFA018127.pdf,
    distinct du cas plus courant où le code TVA est sur sa PROPRE ligne
    suivante (ex. OFC194316.pdf/OFC193413.pdf) : sans cette variante, le
    motif montant (ancré `$`) ne matchait jamais du tout et TOUTE la
    facture ressortait à 0 ligne."""

    m = MOTIF_MONTANT_LIGNE_FACTURE_COMINTER.match(lignes[i_montant].strip())
    if not m:
        return None
    montant = to_float(m.group(1))
    code_tva_inline = m.group(2)

    k = i_montant - 1
    if k >= 0 and MOTIF_REM_FACTURE_COMINTER.match(lignes[k].strip()):
        k -= 1

    if k < 1:
        return None
    if not MOTIF_QTE_PX_FACTURE_COMINTER.match(lignes[k].strip()):
        return None
    if not MOTIF_QTE_PX_FACTURE_COMINTER.match(lignes[k - 1].strip()):
        return None

    qte = to_float(lignes[k - 1].strip())

    return k - 1, qte, montant, code_tva_inline


def _lignes_facture_cominter(lignes: list[str], debut: int, fin: int, numero_bl: str) -> list[LigneFacture]:

    blocs = []
    for i in range(debut, fin):
        b = _bloc_prix_facture_cominter(lignes, i)
        if b is not None:
            blocs.append((i, *b))  # (i_montant, debut_bloc, qte, montant, code_tva_inline)

    articles = []
    for idx, (i_montant, debut_bloc, qte, montant, code_tva_inline) in enumerate(blocs):

        if code_tva_inline is not None:
            j = i_montant + 1
        else:
            j = i_montant + 1
            if j >= fin or not MOTIF_CODE_TVA_FACTURE_COMINTER.match(lignes[j].strip()):
                continue
            j += 1

        fin_zone = blocs[idx + 1][1] if idx + 1 < len(blocs) else fin

        zone = [lignes[k].strip() for k in range(j, fin_zone) if lignes[k].strip()]
        if not zone:
            continue

        # Référence = DERNIÈRE ligne de la zone qui ressemble vraiment à
        # une référence (lettres+chiffres, voir MOTIF_REF_ARTICLE_BL_COMINTER
        # déjà éprouvé côté BL du même fournisseur) — PAS systématiquement
        # zone[-1] : un mot de note/adresse peut traîner APRÈS la vraie
        # référence sur la DERNIÈRE ligne d'articles du document (cas réel
        # OFC194316.pdf, "Livraison chantier Anzemberg... CAMBAIE CG CG"
        # après la référence CAETM4288 — sans ce garde-fou, "CG" ou
        # "CAMBAIE" auraient été pris à tort pour la référence). Tout ce
        # qui suit la référence trouvée est simplement ignoré (note, pas
        # un champ de cette ligne).
        i_ref = next(
            (k for k in range(len(zone) - 1, -1, -1) if MOTIF_REF_ARTICLE_BL_COMINTER.match(zone[k])),
            None,
        )
        if i_ref is None:
            continue
        reference = zone[i_ref]

        if i_ref == 0:
            # Variante réelle (NF155008.pdf, agence Saint-Pierre) : la
            # référence vient AVANT la désignation ("L69731L / Prise 2P+T
            # saillie Plexo gris / unite"), pas après comme les 3 autres
            # fixtures. Bornée par le PROCHAIN Cdt (comme la variante
            # habituelle), ou à défaut (lignes d'éco-participation, jamais
            # de Cdt) à UNE seule ligne — pour ne jamais avaler une note de
            # pied de document du genre "BC N°24 1581 DU 07/07/26" qui
            # peut traîner juste après sur la DERNIÈRE ligne d'articles.
            reste = zone[1:]
            i_cdt = next((k for k, c in enumerate(reste) if c.lower() in _CDT_FACTURE_COMINTER), None)
            designation = " ".join(reste[:i_cdt] if i_cdt is not None else reste[:1])
        else:
            reste = zone[:i_ref]
            if reste and reste[-1].lower() in _CDT_FACTURE_COMINTER:
                designation = " ".join(reste[:-1])
            else:
                designation = " ".join(reste)

        if not qte:
            continue

        articles.append(LigneFacture(
            reference_fournisseur=reference,
            designation=designation,
            quantite_facturee=qte,
            prix_unitaire_ht=round(montant / qte, 4),
            montant_ht=montant,
            numero_bl=numero_bl,
        ))

    return articles


def parse_facture_cominter(texte: str) -> Facture:

    lignes = [l.rstrip() for l in texte.splitlines()]

    m_num = MOTIF_NUMERO_FACTURE_COMINTER.search(texte)
    numero_facture = m_num.group(1) if m_num else ""

    zone = _zone_articles_facture_cominter(lignes)
    if zone is None:
        return Facture(fournisseur="COMINTER", fichier="", numero_facture=numero_facture, date_facture="")

    debut, fin = zone

    date_facture = ""
    m_date = MOTIF_DATE_FACTURE_COMINTER.search(texte[:texte.find("Signature") if "Signature" in texte else len(texte)])
    if m_date:
        jour, mois, annee = m_date.groups()
        date_facture = f"{int(jour):02d}/{mois}/20{annee}"

    numero_bl, numero_commande = _entete_bl_facture_cominter(lignes, debut, fin)

    lignes_facture = _lignes_facture_cominter(lignes, debut, fin, numero_bl)
    for l in lignes_facture:
        l.numero_bl = numero_bl

    return Facture(
        fournisseur="COMINTER",
        fichier="",
        numero_facture=numero_facture,
        date_facture=date_facture,
        numeros_commande=[numero_commande] if numero_commande else [],
        numeros_bl=[numero_bl] if numero_bl else [],
        lignes=lignes_facture,
        total_ht_affiche=None,
    )


# "COMINTER MAYOTTE" est géré par moteur/fournisseurs/cominter_mayotte.py :
# entité distincte (SIRET, adresse, structure de devis différents), voir
# le bandeau GABARIT de ce module. Sa FACTURE (format MFACxxxxx, "N° de
# Commande :"/"Nom du Chantier :" en labels explicites, structure de
# remise différente) n'est PAS couverte ici — reste à construire
# séparément (12 pièces réelles déjà disponibles pour ça, voir a_traiter/
# Factures/ au moment de cette session).
FOURNISSEURS = ['COMINTER']
parse = parse_cominter
parse_bl = parse_bl_cominter
parse_facture = parse_facture_cominter
