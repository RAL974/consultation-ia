"""
Parser BL YESSS ÉLECTRIQUE (marque du groupe CEF SAS, RCS Lyon) — nouveau
fournisseur, 1 seul vrai BL vu à ce jour (`BL M4.273 GENDARMERIE.pdf`,
agence YESSS CAMBAIE). Aucun devis connu pour ce fournisseur — module
dédié BL uniquement.

Structure RÉELLEMENT inhabituelle : chaque champ (Montant, Prix net,
Désignation, Catalogue...) est imprimé comme un LABEL suivi de sa VALEUR,
mais le TEXTE LUI-MÊME est pivoté à 90° sur la page (confirmé en zoomant
sur le PDF rendu — les boîtes OCR de ces mots sont hautes et étroites,
~30px de large pour 120-180px de haut, signature d'un mot tourné). Chaque
"colonne" de la grille d'origine (non tournée) devient ainsi une bande
verticale étroite (~30-40px de large) sur la page rendue : la bande la
plus à gauche porte les LABELS empilés (Montant, Prix net, Désignation,
Catalogue, Fournisseur, Qté, chacun à une hauteur Y différente), la bande
juste à droite porte la valeur RÉELLE de la ligne d'article, et une 3e
bande encore plus à droite porte une valeur "0.00"/vide — un emplacement
de 2e ligne d'article TOUJOURS présent dans le gabarit mais non rempli
sur ce document (une seule ligne d'article vue à ce jour).

Chaque valeur est donc repérée par PROXIMITÉ à son label : le mot le plus
proche (X ET Y) du label, PAS par un ordre de lecture haut/bas classique
(voir _valeur_proche). Une tolérance X ÉTROITE (~40px) est nécessaire :
sans elle, "dispo sous 48h" (dans la bande "2e ligne vide", juste à côté
de la désignation réelle) se retrouve PLUS PROCHE en Y du label
"Désignation" que la vraie désignation "DX3-ID 2P 63AA 30MA TGA" — d'où
la tolérance X qui filtre cette bande voisine.

Quantité TOUJOURS déduite de Montant / Prix net (comme 109 Distribution/
Cominter/Electric Plus/Ravate) : la cellule Qté elle-même est ambiguë sur
ce document (deux valeurs empilées "2"/"0", ligne réelle + emplacement
vide comme les autres champs numériques).

Un seul PDF réel à ce jour (règle d'or) : structure à revalider dès qu'un
2e vrai BL YESSS sera disponible, notamment pour une commande à PLUSIEURS
lignes d'article (jamais rencontrée).
"""

import re

from moteur.ocr import mots_document
from moteur.outils import to_float
from moteur.rapprochement.modele_bl import BonLivraison, LigneBL

# --- GABARIT BL (YESSS) ------------------------------------------------------
MOTIF_NOMBRE_BL_YESSS = re.compile(r"^\d+[.,]\d{1,2}$")
MOTIF_MONTANT_BL_YESSS = re.compile(r"^Montant$", re.IGNORECASE)
MOTIF_PRIX_NET_BL_YESSS = re.compile(r"^Prix\s*net$", re.IGNORECASE)
MOTIF_DESIGNATION_BL_YESSS = re.compile(r"^D.signation$", re.IGNORECASE)
MOTIF_CATALOGUE_BL_YESSS = re.compile(r"^Catalogue$", re.IGNORECASE)
MOTIF_COMMANDE_BL_YESSS = re.compile(r"N.\s*commande\s+([A-Z]?\d[\w.\-]*)", re.IGNORECASE)
MOTIF_NUMERO_BL_YESSS = re.compile(r"\b([A-Z]{2,4}/\d{5,7})\b")
MOIS_FR_YESSS = {
    "janvier": 1, "février": 2, "mars": 3, "avril": 4, "mai": 5, "juin": 6,
    "juillet": 7, "août": 8, "aout": 8, "septembre": 9, "octobre": 10,
    "novembre": 11, "décembre": 12,
}
MOTIF_MOIS_ANNEE_BL_YESSS = re.compile(
    r"(" + "|".join(MOIS_FR_YESSS) + r")\s+(\d{4})", re.IGNORECASE,
)
MOTIF_JOUR_SEUL_BL_YESSS = re.compile(r"^\d{1,2}$")
MOTIF_LABEL_DATE_BL_YESSS = re.compile(r"^Date$", re.IGNORECASE)
# --- fin GABARIT BL -----------------------------------------------------------


def _valeur_proche(mots: list[dict], motif_label, filtre_texte, tolerance_x: float = 45.0):
    """Le mot dont le texte satisfait `filtre_texte`, le plus proche (X ET
    Y) du label matchant `motif_label` — voir bandeau du module. Retourne
    None si le label est introuvable ou qu'aucun candidat ne passe la
    tolérance X."""

    label = next((m for m in mots if motif_label.search(m["texte"].strip())), None)
    if label is None:
        return None

    xc_label = (label["x0"] + label["x1"]) / 2
    yc_label = (label["y0"] + label["y1"]) / 2

    candidats = []
    for m in mots:
        if m is label:
            continue
        texte = m["texte"].strip()
        if not filtre_texte(texte):
            continue
        xc = (m["x0"] + m["x1"]) / 2
        yc = (m["y0"] + m["y1"]) / 2
        if abs(xc - xc_label) <= tolerance_x:
            candidats.append((abs(yc - yc_label), m))

    if not candidats:
        return None

    candidats.sort(key=lambda c: c[0])
    return candidats[0][1]


def _nombre_non_nul_proche(mots: list[dict], motif_label) -> float | None:
    """Comme _valeur_proche, mais parmi les candidats numériques les plus
    proches, retient le premier NON NUL (l'emplacement de 2e ligne
    d'article, toujours "0.00" sur le seul document vu à ce jour, ne doit
    jamais être pris pour une vraie valeur)."""

    label = next((m for m in mots if motif_label.search(m["texte"].strip())), None)
    if label is None:
        return None

    xc_label = (label["x0"] + label["x1"]) / 2
    yc_label = (label["y0"] + label["y1"]) / 2

    candidats = []
    for m in mots:
        if m is label:
            continue
        texte = m["texte"].strip()
        if not MOTIF_NOMBRE_BL_YESSS.fullmatch(texte):
            continue
        xc = (m["x0"] + m["x1"]) / 2
        yc = (m["y0"] + m["y1"]) / 2
        if abs(xc - xc_label) <= 45.0:
            candidats.append((abs(yc - yc_label), to_float(texte)))

    candidats.sort(key=lambda c: c[0])
    for _, valeur in candidats:
        if valeur:
            return valeur
    return None


def _date_bl_yesss(mots: list[dict]) -> str:
    """"24 août 2026" : le mois+année ressort comme un seul mot OCR
    ("aout 2026" ou "août 2026"), mais le jour ("24") ressort SÉPARÉMENT,
    ailleurs sur la page (même repli que d'autres champs de ce gabarit —
    voir bandeau). Le document contient D'AUTRES mentions mois+année SANS
    RAPPORT (le pavé légal cite une loi du "25 janvier 1985") ainsi que
    d'autres mots à 1-2 chiffres (quantité, bouts de code postal...) : le
    bon mot mois+année est retrouvé par proximité au label "Date" (comme
    le reste du gabarit, voir bandeau du module), PUIS le jour par
    proximité à CE mot mois+année précisément — jamais le premier trouvé
    dans un ordre arbitraire."""

    mot_mois_annee = _valeur_proche(
        mots, MOTIF_LABEL_DATE_BL_YESSS,
        lambda t: bool(MOTIF_MOIS_ANNEE_BL_YESSS.search(t)),
    )
    if mot_mois_annee is None:
        return ""

    match = MOTIF_MOIS_ANNEE_BL_YESSS.search(mot_mois_annee["texte"].strip())
    mois = MOIS_FR_YESSS.get(match.group(1).lower())
    annee = match.group(2)
    yc_ref = (mot_mois_annee["y0"] + mot_mois_annee["y1"]) / 2

    candidats = sorted(
        (
            (abs((m["y0"] + m["y1"]) / 2 - yc_ref), m["texte"].strip())
            for m in mots
            if MOTIF_JOUR_SEUL_BL_YESSS.fullmatch(m["texte"].strip())
        ),
        key=lambda c: c[0],
    )
    if not candidats or candidats[0][0] > 200:
        return ""

    jour = int(candidats[0][1])
    return f"{jour:02d}/{mois:02d}/{annee}"


def parse_bl_yesss(mots_par_page: list[list[dict]]) -> BonLivraison:

    mots = [m for page in mots_par_page for m in page]

    numero_commande = ""
    numero_bl = ""
    for m in mots:
        if not numero_commande:
            match = MOTIF_COMMANDE_BL_YESSS.search(m["texte"])
            if match:
                numero_commande = match.group(1)
        if not numero_bl:
            match = MOTIF_NUMERO_BL_YESSS.search(m["texte"])
            if match:
                numero_bl = match.group(1)

    bl = BonLivraison(
        fournisseur="YESSS", fichier="", numero_bl=numero_bl,
        date_bl=_date_bl_yesss(mots), numero_commande=numero_commande,
    )

    reference_mot = _valeur_proche(
        mots, MOTIF_CATALOGUE_BL_YESSS,
        lambda t: bool(re.fullmatch(r"[A-Z0-9][A-Z0-9.\-]{3,14}", t)),
    )
    designation_mot = _valeur_proche(
        mots, MOTIF_DESIGNATION_BL_YESSS,
        lambda t: bool(t) and not t.isdigit(),
    )
    prix_net = _nombre_non_nul_proche(mots, MOTIF_PRIX_NET_BL_YESSS)
    montant = _nombre_non_nul_proche(mots, MOTIF_MONTANT_BL_YESSS)

    if reference_mot is None or prix_net is None or montant is None:
        return bl

    quantite = round(montant / prix_net, 2) if prix_net else 0.0

    bl.lignes.append(LigneBL(
        reference_fournisseur=reference_mot["texte"].strip(),
        designation=designation_mot["texte"].strip() if designation_mot else "",
        quantite_livree=quantite,
        prix_net=prix_net,
        montant=montant,
    ))
    bl.total_ht_affiche = montant

    return bl


FOURNISSEURS = ["YESSS"]

parse_bl = parse_bl_yesss
