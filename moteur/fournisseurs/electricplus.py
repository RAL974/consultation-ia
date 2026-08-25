import re

from moteur.modele import Article
from moteur.fournisseurs._gabarit import scan_ancre
from moteur.ocr import pages_par_identifiant, regrouper_lignes
from moteur.outils import to_float
from moteur.rapprochement.modele_bl import BonLivraison, LigneBL

# --- GABARIT (Electric Plus Réunion — marque publique du canal GMR) --------
# Structure du texte extrait, un champ par ligne, ancrée sur "PF" :
#     Référence          (i-5)
#     Désignation        (i-4)
#     Quantité           (i-3)
#     Prix unit. HT      (i-2)
#     P.U. net HT        (i-1)
#     PF                 (i)
#     Montant HT         (i+1)
#     Code TVA           (i+2)
MARQUEUR = "PF"
OFFSETS = {
    "reference_fournisseur": -5,
    "designation": -4,
    "quantite": -3,
    "prix_brut": -2,
    "prix_net": -1,
    "montant": 1,
}
MOTIF_DEVIS = r"^\s*(\d{7})\s*$"
MOTIF_REF = re.compile(r"^[A-Z0-9]{4,}$")
# --- fin GABARIT -------------------------------------------------------------


def _f(v: str) -> float:
    return float(v.replace(" ", "").replace(" ", "").replace(",", "."))


def parse_electricplus(texte: str) -> list[Article]:

    articles = []

    devis = ""
    m = re.search(MOTIF_DEVIS, texte, re.MULTILINE)
    if m:
        devis = m.group(1)

    lignes = [l.rstrip() for l in texte.splitlines()]

    for bloc in scan_ancre(lignes, MARQUEUR, OFFSETS):

        try:
            ref = bloc["reference_fournisseur"].strip()
            designation = bloc["designation"].strip()

            quantite = _f(bloc["quantite"])
            prix_brut = _f(bloc["prix_brut"])
            prix_net = _f(bloc["prix_net"])
            montant = _f(bloc["montant"])

        except (ValueError, IndexError, AttributeError):
            continue

        if not MOTIF_REF.match(ref):
            continue

        articles.append(
            Article(
                fournisseur="ELECTRIC PLUS",
                devis=devis,
                reference_fournisseur=ref,
                reference_distributeur="",
                designation=designation,
                quantite=quantite,
                unite="UN",
                prix_brut=prix_brut,
                prix_net=prix_net,
                montant=montant,
            )
        )

    return articles


# --- GABARIT BL (Electric Plus / GMR) ---------------------------------------
# GMR n'envoie PAS de bon de livraison séparé (confirmé par l'acheteur,
# session R2 suite) : le rapprochement se fait directement à partir de ses
# FACTURES (structure similaire aux BL des autres fournisseurs :
# "Référence client : CDE <commande>", tableau réf/désignation/qté/prix,
# "Total HT" en pied de page comme 109 Distribution -> autocontrôle
# possible ici, contrairement à Coredime qui n'affiche aucun prix).
#
# Ligne d'article OCR : la désignation est parfois éclatée en PLUSIEURS
# cellules, un mot par cellule (contrairement aux BL 109 Distribution/
# Coredime où elle tient sur une seule) — la référence et le prix
# encadrent un nombre variable de cellules de désignation. Ancre fiable :
# la cellule "<prix> PF" (prix fournisseur, même convention que le devis
# de ce fournisseur ci-dessus), toujours suivie du P.U. net puis du
# Montant. La quantité livrée est déduite de Montant / P.U. net (même
# logique que 109 Distribution) plutôt que lue dans la cellule Qté, qui
# peut être absente de l'OCR selon les fichiers.
#
# Cas réel rencontré : la référence "LEG004107" a été lue coupée en deux
# cellules ("LE" + "4107", ou "LEL" + "4107") sur certains scans du MÊME
# document (comparé à un autre scan de la même facture où elle est lue
# entière) — reconstituée si la 1ère cellule est purement alphabétique et
# courte (<=3 caractères) et la 2e purement numérique.
# P[FR] plutôt que PF strict : cas réel (bl_electricplus_8.pdf) où l'OCR a
# lu "PR" au lieu de "PF" (F confondu avec R) -> 0 ligne extraite sur toute
# la facture avant ce correctif, même famille que la tolérance O/0 déjà en
# place sur cette même ancre.
MOTIF_PF_ELECTRICPLUS = re.compile(r"^([\doO]+[,.][\doO]+)\s*P[FR]$", re.IGNORECASE)
MOTIF_COMMANDE_ELECTRICPLUS = re.compile(
    r"CDE\s*[°o]?\s*([A-Z]?\d{1,4}(?:[.\-]\d{1,4}){1,2})", re.IGNORECASE
)
# [^\d]{0,3} plutôt que \D+ ou [^\d\n]+ : date/date_echeance/facture sont
# des cellules OCR CONSÉCUTIVES, séparées d'un simple "\n" une fois
# aplaties (aucun autre caractère entre elles) — \D+/[^\d\n]+ exigeaient
# À TORT au moins un caractère non-numérique NON-saut de ligne, ce qui ne
# matchait jamais (bug réel : numero_bl/date_bl toujours vides). Borné à 3
# caractères pour ne jamais dériver vers un nombre lointain sans rapport
# (même prudence que le bug de séparateur \s trouvé chez Coredime).
MOTIF_FACTURE_DATE_ELECTRICPLUS = re.compile(
    r"(\d{1,2}/\d{2}/\d{2})[^\d]{0,3}\d{1,2}/\d{2}/\d{2}[^\d]{0,3}(\d{6,7})\b"
)
# BUG RÉEL CORRIGÉ (recette réelle, doc07149220260814105422.pdf, 2 pages) :
# un même fichier peut contenir PLUSIEURS FACTURES Electric Plus distinctes
# (commande/n° facture/date différents par page — même principe que les
# scans en masse Cominter Ouest / 109 Distribution, voir moteur.ocr.
# pages_par_identifiant). Avant ce correctif, tout le fichier était traité
# comme UN seul document : le n° de commande et le total venaient de la
# 1ère page trouvée, mais les LIGNES d'articles étaient accumulées sur
# TOUTES les pages sans distinction — un vrai risque de mélanger deux
# livraisons de commandes différentes sous un seul numéro.
#
# Identifiant de regroupement dédié — PAS MOTIF_FACTURE_DATE_ELECTRICPLUS :
# grouper_pages_par_identifiant()/pages_par_identifiant() cherchent le motif
# dans le texte BRUT de la page (mots OCR simplement joints par un espace,
# dans leur ORDRE DE LECTURE OCR, jamais réordonnés en lignes visuelles
# contrairement à regrouper_lignes() utilisé pour l'extraction des champs
# ci-dessous) — sur ce document réel, les cellules "date / date échéance /
# n° facture" ne sont PAS lues dans cet ordre par l'OCR, donc le motif à 3
# cellules adjacentes ne matchait JAMAIS (bug trouvé en recette : les 2
# pages fusionnaient toujours en un seul groupe). Seul le n° de facture
# LUI-MÊME apparaît comme un token isolé fiable dans le flux brut (borné
# par \b des deux côtés — jamais une coïncidence avec un siret/téléphone/
# code postal, tous plus longs ou collés à des lettres sans transition
# \w/\W, vérifié sur ce document réel).
MOTIF_IDENTIFIANT_PAGE_ELECTRICPLUS = re.compile(r"\b(\d{6,7})\b")
MOTIF_ENTETE_TABLEAU_ELECTRICPLUS = re.compile(r"REFERENCES")
MOTIF_PIED_TABLEAU_ELECTRICPLUS = re.compile(r"TOTALHT|CODESTVA")
MOTIF_TOTAL_HT_ELECTRICPLUS = re.compile(r"TOTALHT(\d[\d\s]*[,.]\d{2})")
MOTIF_REF_INCOMPLETE_ELECTRICPLUS = re.compile(r"^[A-Z]{1,3}$")
MOTIF_QTE_OU_NOMBRE_ELECTRICPLUS = re.compile(r"^\d+[.,]\d+(?:MTR|MTS|M)?$", re.IGNORECASE)
# --- fin GABARIT BL -----------------------------------------------------------


def _sans_espaces_electricplus(s: str) -> str:
    return re.sub(r"\s+", "", s.upper())


def _normaliser_date_electricplus(brut: str) -> str:
    """"7/08/26" -> "07/08/2026" (même format que les autres fournisseurs
    BL, voir moteur/rapprochement/pipeline_bl._parser_date_bl)."""

    m = re.match(r"^(\d{1,2})/(\d{2})/(\d{2})$", brut.strip())
    if not m:
        return ""
    jour, mois, an2 = m.groups()
    return f"{int(jour):02d}/{mois}/20{an2}"


def _zone_tableau_electricplus(lignes_groupees: list[list[dict]]) -> list[list[dict]]:

    i_entete = next(
        (i for i, ligne in enumerate(lignes_groupees)
         if any(MOTIF_ENTETE_TABLEAU_ELECTRICPLUS.search(_sans_espaces_electricplus(m["texte"])) for m in ligne)),
        None,
    )
    if i_entete is None:
        return []

    i_pied = next(
        (i for i, ligne in enumerate(lignes_groupees)
         if i > i_entete and any(
             MOTIF_PIED_TABLEAU_ELECTRICPLUS.search(_sans_espaces_electricplus(m["texte"])) for m in ligne
         )),
        None,
    )

    return lignes_groupees[i_entete + 1:(i_pied if i_pied is not None else len(lignes_groupees))]


def _ligne_vers_article_electricplus(cellules: list[str]) -> LigneBL | None:

    i_pf = next(
        (i for i, c in enumerate(cellules) if MOTIF_PF_ELECTRICPLUS.match(c.strip())),
        None,
    )

    if i_pf is not None and i_pf >= 1 and i_pf + 2 < len(cellules):
        brut_prix_net = MOTIF_PF_ELECTRICPLUS.match(cellules[i_pf].strip()).group(1)
        prix_net = to_float(brut_prix_net.replace("o", "0").replace("O", "0"))
        montant = to_float(cellules[i_pf + 2])
        fin_designation = i_pf
        appliquer_trim_generique = True
    else:
        # BUG RÉEL CORRIGÉ (recette réelle, facture 1207019/commande
        # M4.263) : certaines lignes n'affichent AUCUN suffixe "PF" sur
        # leur cellule Prix unit. HT (repéré ici par cohérence
        # arithmétique — qté × P.U.net = Montant, et Montant = Total HT
        # affiché de la page pour ce document à ligne unique). Repli
        # POSITIONNEL depuis la FIN de la ligne (colonnes REFERENCES |
        # DESIGNATION | QTE | PRIX UNIT.HT | P.U.NET HT | MONTANT HT),
        # jamais depuis le début — le nombre de cellules de désignation
        # varie. Un seul exemple réel à ce jour (règle d'or) ; validé
        # uniquement par la cohérence arithmétique ci-dessus.
        if len(cellules) < 5:
            return None
        prix_net = to_float(cellules[-2])
        montant = to_float(cellules[-1])
        fin_designation = len(cellules) - 4
        appliquer_trim_generique = False

    if not prix_net or not montant:
        return None

    reference = cellules[0].strip()
    debut_designation = 1

    if (
        MOTIF_REF_INCOMPLETE_ELECTRICPLUS.match(reference)
        and len(cellules) > 1
        and re.match(r"^\d+$", cellules[1].strip())
    ):
        reference = reference + cellules[1].strip()
        debut_designation = 2

    # La cellule juste avant l'ancre "PF" est souvent la quantité (nombre
    # nu, parfois suivi d'une unité) plutôt que de la désignation — sans
    # objet pour le repli positionnel ci-dessus, déjà correctement placé.
    if appliquer_trim_generique and fin_designation > debut_designation and MOTIF_QTE_OU_NOMBRE_ELECTRICPLUS.match(
        cellules[fin_designation - 1].strip()
    ):
        fin_designation -= 1

    designation = " ".join(c.strip() for c in cellules[debut_designation:fin_designation]).strip()

    return LigneBL(
        reference_fournisseur=reference,
        designation=designation,
        quantite_livree=round(montant / prix_net, 4),
        prix_net=prix_net,
        montant=montant,
    )


def _parse_une_facture_electricplus(mots_par_page: list[list[dict]]) -> BonLivraison:

    lignes_plates = [
        mot["texte"]
        for mots in mots_par_page
        for ligne in regrouper_lignes(mots)
        for mot in ligne
    ]
    texte = "\n".join(lignes_plates)

    numero_commande = ""
    m = MOTIF_COMMANDE_ELECTRICPLUS.search(texte)
    if m:
        numero_commande = m.group(1).upper().replace(" ", ".")

    numero_bl, date_bl = "", ""
    m = MOTIF_FACTURE_DATE_ELECTRICPLUS.search(texte)
    if m:
        date_bl = _normaliser_date_electricplus(m.group(1))
        numero_bl = m.group(2)

    total_ht_affiche = None
    m = MOTIF_TOTAL_HT_ELECTRICPLUS.search(_sans_espaces_electricplus(texte))
    if m:
        total_ht_affiche = to_float(m.group(1))

    articles = []
    for mots in mots_par_page:
        for ligne_mots in _zone_tableau_electricplus(regrouper_lignes(mots)):
            cellules = [m["texte"] for m in ligne_mots]
            article = _ligne_vers_article_electricplus(cellules)
            if article:
                articles.append(article)

    bl = BonLivraison(
        fournisseur="ELECTRIC PLUS",
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
                f"!! ELECTRIC PLUS (facture) : Total HT affiché ({total_ht_affiche:.2f}€) "
                f"!= somme des lignes extraites ({total_extrait:.2f}€) "
                f"— une ligne a peut-être été oubliée ou mal lue par l'OCR."
            )

    return bl


def parse_bl_electricplus(mots_par_page: list[list[dict]]) -> list[BonLivraison]:
    """Un même fichier peut contenir PLUSIEURS FACTURES Electric Plus
    distinctes (cas réel, doc07149220260814105422.pdf : 2 pages, 2
    commandes/n° facture/dates différents) — retourne une liste, une
    entrée par facture détectée, même principe que parse_bl_109/
    parse_bl_cominter (voir moteur.ocr.pages_par_identifiant). Chaque
    BonLivraison porte aussi les indices de page qu'il occupe dans le
    fichier source (bl.pages), pour l'archivage individuel (voir
    moteur.rapprochement.pipeline_bl, "archivage par BL individuel")."""

    groupes_indices = pages_par_identifiant(mots_par_page, MOTIF_IDENTIFIANT_PAGE_ELECTRICPLUS)

    resultat = []
    for indices in groupes_indices:
        bl = _parse_une_facture_electricplus([mots_par_page[i] for i in indices])
        bl.pages = indices
        resultat.append(bl)

    return resultat


# Déclaration pour le chargement automatique
FOURNISSEURS = ['ELECTRIC PLUS', 'GMR']
parse = parse_electricplus
parse_bl = parse_bl_electricplus
