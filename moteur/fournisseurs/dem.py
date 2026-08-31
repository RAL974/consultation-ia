import re

from moteur.modele import Article
from moteur.fournisseurs._gabarit import scan_regex

# --- GABARIT (DEM) ---------------------------------------------------------
# Structure du texte extrait :
#     LEG411651        15.00 U    8058.85/C    1208.83 1
#     DX3-ID 2P 63A A 30MA TGA          <- désignation sur la ligne SUIVANTE
#
# ATTENTION : les prix DEM sont exprimés AU CENT (/C).
# Le prix unitaire réel est donc montant / quantité.
# Les nombres utilisent le point décimal (128.52).
MOTIF_DEVIS = r"N°\s*(\d+)\s+du"
# La référence peut contenir des POINTS et des TIRETS (cas réels :
# "FILFR1N3G2.5T500" devis 821409, "BOOMAL.1200.213.5" devis 821416) —
# la classe d'origine ([A-Z][A-Z0-9]+) s'arrêtait au point, le \s+ suivant
# tombait dessus et la ligne entière échouait : 0 article extrait, sans
# aucune anomalie levée. Voir tests/test_parsers.py.
MOTIF_LIGNE = re.compile(
    r"^\s*([A-Z][A-Z0-9.\-]+)\s+"   # référence
    r"([\d.]+)\s+"                   # quantité
    r"([A-Z]{1,3})\s+"               # unité (U...)
    r"([\d.]+)/C\s+"                 # prix au cent
    r"([\d.]+)\s+"                   # total HT
    r"(\d)\s*$"                      # code TVA
)
# --- fin GABARIT -------------------------------------------------------------


def parse_dem(texte: str) -> list[Article]:

    articles = []

    devis = ""
    m = re.search(MOTIF_DEVIS, texte)
    if m:
        devis = m.group(1)

    lignes = texte.splitlines()

    for i, m in scan_regex(lignes, MOTIF_LIGNE):

        ref = m.group(1)
        quantite = float(m.group(2))
        unite = m.group(3)
        prix_cent = float(m.group(4))
        montant = float(m.group(5))

        # Prix unitaire réel : montant / quantité (les prix sont au cent)
        prix_net = round(montant / quantite, 4) if quantite else 0.0
        prix_brut = round(prix_cent / 100, 4)

        # Désignation sur la ligne suivante
        designation = lignes[i + 1].strip() if i + 1 < len(lignes) else ""

        articles.append(
            Article(
                fournisseur="DEM",
                devis=devis,
                reference_fournisseur=ref,
                reference_distributeur="",
                designation=designation,
                quantite=quantite,
                unite=unite,
                prix_brut=prix_brut,
                prix_net=prix_net,
                montant=montant,
            )
        )

    return articles


# --- GABARIT BL (DEM) --------------------------------------------------------
# BL scanné (image pure, comme les autres — voir moteur/ocr.py), 2 vrais BL
# vus à ce jour (M3.14.363 — 2 pages, chacune un bon de livraison DEM
# DIFFÉRENT — voir plus bas ; M3.23.046 — 1 page, 6 lignes). Structure du
# tableau IDENTIQUE au devis (voir bandeau GABARIT ci-dessus) : Référence,
# puis sur la MÊME ligne visuelle Quantité ("250.00 U"), Prix AU CENT
# ("95.00/C"), Montant HT ("237.50") — désignation sur la ligne SUIVANTE.
# Ancre fiable : "QUANTITE" dans l'en-tête du tableau (repéré même très
# déformé par l'OCR, ex. "ＡT ON No QUANTITE P.U. HT TOTAL HT" — seul
# "QUANTITE" reste lisible à chaque fois).
#
# Chaque PAGE est traitée comme un bon de livraison INDÉPENDANT (pas de
# fusion inter-pages) : le n° de BL et sa date changent bel et bien d'une
# page à l'autre sur le fixture 2 pages (706992 le 24/08, 706990 le
# 20/08) — ce sont deux vraies livraisons SÉPARÉES pour la même commande
# (M3.14.363), pas un même BL étalé sur 2 pages.
#
# "Reste à livrer" (cas réel, page 2 du fixture 2 pages) : une ligne SANS
# prix ni montant ("FILSYT15P0.9T SYT15P0.9T500M 250.00 U", désignation
# glued sur la MÊME ligne visuelle faute de place prise par les colonnes
# de prix normalement vides) — jamais livrée, exclue. Confirmé par le fait
# que CETTE MÊME référence (même quantité) réapparaît, cette fois PRICED,
# sur l'AUTRE page du même fichier (la livraison suivante qui la solde).
# L'absence de prix/montant est en elle-même le signal fiable ici (pas
# besoin de mémoriser un drapeau "reste à livrer" par page comme chez
# Coredime : chez DEM, une ligne non livrée n'affiche jamais de prix,
# quelle que soit sa position sur la page).
MOTIF_ENTETE_TABLEAU_BL_DEM = re.compile(r"QUANTITE")
MOTIF_PIED_TABLEAU_BL_DEM = re.compile(r"FRAISDEPORT|TOTALH\.?T", re.IGNORECASE)
MOTIF_QUANTITE_BL_DEM = re.compile(r"^(\d[\d\s]*[.,]\d{2})\s*U$", re.IGNORECASE)
MOTIF_PRIX_AU_CENT_BL_DEM = re.compile(r"^(\d[\d\s]*[.,]\d{2})\s*/\s*C$", re.IGNORECASE)
MOTIF_MONTANT_BL_DEM = re.compile(r"^\d[\d\s]*[.,]\d{2}$")
MOTIF_COMMANDE_BL_DEM = re.compile(r"votre\s*commande\s+(\S+)", re.IGNORECASE)
MOTIF_BL_NUMERO_DATE_DEM = re.compile(r"N.\s*(\d{5,8})\s*du\s*(\d{2}/\d{2}/\d{2,4})", re.IGNORECASE)
MOTIF_TOTAL_HT_BL_DEM = re.compile(r"TOTALH\.?T\.?NET\s*(\d[\d\s]*[.,]\d{2})", re.IGNORECASE)
# --- fin GABARIT BL -----------------------------------------------------------


def _nombre_bl_dem(valeur: str) -> float:
    return float(valeur.replace(" ", "").replace(",", "."))


def _sans_espaces_bl_dem(s: str) -> str:
    return re.sub(r"\s+", "", s.upper())


def _zone_tableau_bl_dem(lignes_groupees):

    i_entete = next(
        (i for i, ligne in enumerate(lignes_groupees)
         if any(MOTIF_ENTETE_TABLEAU_BL_DEM.search(_sans_espaces_bl_dem(m["texte"])) for m in ligne)),
        None,
    )
    if i_entete is None:
        return []

    i_pied = next(
        (i for i, ligne in enumerate(lignes_groupees)
         if i > i_entete and any(
             MOTIF_PIED_TABLEAU_BL_DEM.search(_sans_espaces_bl_dem(m["texte"])) for m in ligne
         )),
        None,
    )

    return lignes_groupees[i_entete + 1:(i_pied if i_pied is not None else len(lignes_groupees))]


def _parse_une_page_bl_dem(mots):

    from moteur.ocr import regrouper_lignes
    from moteur.rapprochement.modele_bl import BonLivraison, LigneBL

    lignes_plates = [" ".join(m["texte"] for m in ligne) for ligne in regrouper_lignes(mots)]
    texte = "\n".join(lignes_plates)

    numero_commande = ""
    m = MOTIF_COMMANDE_BL_DEM.search(texte)
    if m:
        numero_commande = m.group(1).upper()

    numero_bl, date_bl = "", ""
    m = MOTIF_BL_NUMERO_DATE_DEM.search(_sans_espaces_bl_dem(texte))
    if m:
        numero_bl = m.group(1)
        jour, mois, an = m.group(2).split("/")
        an = an if len(an) == 4 else f"20{an}"
        date_bl = f"{jour}/{mois}/{an}"

    total_ht_affiche = None
    m = MOTIF_TOTAL_HT_BL_DEM.search(_sans_espaces_bl_dem(texte))
    if m:
        total_ht_affiche = _nombre_bl_dem(m.group(1))

    zone = _zone_tableau_bl_dem(regrouper_lignes(mots))

    articles = []
    i = 0
    while i < len(zone):
        cellules = [m["texte"] for m in zone[i]]

        if not cellules:
            i += 1
            continue

        reference = cellules[0].strip()

        i_qte = next(
            (j for j, c in enumerate(cellules) if MOTIF_QUANTITE_BL_DEM.match(c.strip())),
            None,
        )
        if i_qte is None or i_qte == 0:
            i += 1
            continue

        quantite = _nombre_bl_dem(MOTIF_QUANTITE_BL_DEM.match(cellules[i_qte].strip()).group(1))

        i_prix = next(
            (j for j, c in enumerate(cellules) if j > i_qte and MOTIF_PRIX_AU_CENT_BL_DEM.match(c.strip())),
            None,
        )
        i_montant = next(
            (j for j, c in enumerate(cellules) if j > i_qte and MOTIF_MONTANT_BL_DEM.match(c.strip())
             and not MOTIF_PRIX_AU_CENT_BL_DEM.match(c.strip())),
            None,
        )

        if i_prix is None or i_montant is None:
            # Voir bandeau GABARIT BL : ligne "reste à livrer" (aucun
            # prix/montant imprimé) — jamais livrée, exclue.
            i += 1
            continue

        montant = _nombre_bl_dem(cellules[i_montant].strip())
        prix_net = round(montant / quantite, 4) if quantite else 0.0

        designation = " ".join(c.strip() for c in cellules[1:i_qte]).strip()
        if not designation and i + 1 < len(zone):
            suivante = [m["texte"] for m in zone[i + 1]]
            if not any(MOTIF_QUANTITE_BL_DEM.match(c.strip()) for c in suivante):
                designation = " ".join(c.strip() for c in suivante).strip()
                i += 1

        articles.append(LigneBL(
            reference_fournisseur=reference, designation=designation,
            quantite_livree=quantite, prix_net=prix_net, montant=montant,
        ))
        i += 1

    return BonLivraison(
        fournisseur="DEM", fichier="", numero_bl=numero_bl, date_bl=date_bl,
        numero_commande=numero_commande, lignes=articles, total_ht_affiche=total_ht_affiche,
    )


def parse_bl_dem(mots_par_page: list[list[dict]]) -> list:
    """Chaque page est un bon de livraison DEM indépendant (voir bandeau
    GABARIT BL) — jamais de fusion inter-pages pour ce fournisseur."""

    return [_parse_une_page_bl_dem(mots) for mots in mots_par_page]


# Déclaration pour le chargement automatique
FOURNISSEURS = ['DEM']
parse = parse_dem
parse_bl = parse_bl_dem
