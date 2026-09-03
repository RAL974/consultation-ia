import re

from moteur.modele import Article
from moteur.fournisseurs._gabarit import scan_ancre
from moteur.ocr import pages_par_identifiant, regrouper_lignes
from moteur.outils import to_float
from moteur.rapprochement.modele_bl import BonLivraison, LigneBL
from moteur.rapprochement.modele_facture import Facture, LigneFacture

# --- GABARIT (Electric Plus Réunion — marque publique du canal GMR) --------
# Structure du texte extrait, un champ par ligne, ancrée sur "PF" (ou "PR") :
#     Référence          (i-5)
#     Désignation        (i-4)
#     Quantité           (i-3)
#     Prix unit. HT      (i-2)
#     P.U. net HT        (i-1)
#     PF ou PR           (i)
#     Montant HT         (i+1)
#     Code TVA           (i+2)
# "PR" au lieu de "PF" : constaté sur 2 devis réels (D1109436/BT-Floe,
# D1109369/R2V 3G1.5 - Rico Carpaye), texte PDF natif (pas un artefact
# OCR comme le repli P[FR] déjà en place côté BL) — une 2e valeur réelle
# pour cette même colonne, même position, mêmes calculs (qté × prix_net =
# montant vérifié exact sur les 2 PDF). Sans ce marqueur en plus, ces
# devis ressortaient à 0 article malgré le fournisseur bien reconnu.
# Élargir ce marqueur a aussi révélé (fixture electric_plus_gmr.pdf,
# déjà en place) que 7 lignes "PR" y étaient silencieusement perdues
# depuis le début, sans qu'aucune anomalie ne soit levée (pas de contrôle
# de Total HT document pour ce parser devis, contrairement au BL) — voir
# tests/test_parsers.py::test_parse_electricplus.
# Limite résiduelle NON corrigée (règle d'or, un seul exemple à ce jour) :
# sur ce même document, la ligne WAG2273205 n'a AUCUN marqueur PF/PR du
# tout (montant 120,00€ imprimé directement après le prix net) — reste
# non extraite, voir "Points fragiles" dans CLAUDE.md.
MARQUEUR = ["PF", "PR"]
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
# Repli MIROIR (voir bandeau _zone_tableau_electricplus : même page inversée,
# ce bloc d'en-tête suit alors l'ordre [n° facture, date échéance, date]
# plutôt que [date, date échéance, n° facture]) — \b devant \d{6,7} pour ne
# jamais matcher un nombre à l'intérieur d'un token collé comme "C400002".
MOTIF_FACTURE_DATE_ELECTRICPLUS_MIROIR = re.compile(
    r"\b(\d{6,7})[^\d]{0,3}\d{1,2}/\d{2}/\d{2}[^\d]{0,3}(\d{1,2}/\d{2}/\d{2})\b"
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
# BUG RÉEL CORRIGÉ (2e lot, fichier multi-fournisseurs) : sur ce document,
# l'en-tête de colonnes ("DESIGNATION QTE PRIX UNIT.HT...") s'est retrouvé
# GROUPÉ PAR L'OCR (tolérance Y de regrouper_lignes) sur la MÊME ligne
# visuelle que la référence+désignation du 1er article ("PLA11525
# EMBTMOULUREKEVA32MMX12MM DESIGNATION QTE..."). Cette ligne échouait
# alors totalement (cellules[-2]/cellules[-1] = "MONTANT HT"/"TVA", pas
# des nombres) et sa référence/désignation était perdue — le 1er article
# ressortait avec la référence de la ligne SUIVANTE (le morceau de
# désignation qui y avait débordé) prise à tort pour une référence. Voir
# _zone_tableau_electricplus, qui détache désormais les cellules
# précédant ce motif et les reporte sur la ligne suivante.
MOTIF_ENTETE_COLONNES_ELECTRICPLUS = re.compile(r"^(DESIGNATION|QTE|PRIX|P\.?U\.?|MONTANT|TVA)", re.IGNORECASE)
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

    if i_pied is not None:
        zone = list(lignes_groupees[i_entete + 1:i_pied])
    else:
        # Repli MIROIR (voir plus bas) : sur une page où l'ORDRE DES LIGNES
        # est lui aussi inversé (pas seulement l'ordre des cellules), le
        # repère de PIED apparaît AVANT l'en-tête plutôt qu'après — cas
        # structurellement impossible sur une page normale (un pied de
        # tableau ne précède jamais son propre en-tête), donc un signal sûr.
        i_pied_avant = next(
            (i for i in range(i_entete - 1, -1, -1)
             if any(MOTIF_PIED_TABLEAU_ELECTRICPLUS.search(_sans_espaces_electricplus(m["texte"]))
                    for m in lignes_groupees[i])),
            None,
        )
        zone = list(lignes_groupees[i_pied_avant + 1:i_entete]) if i_pied_avant is not None \
            else list(lignes_groupees[i_entete + 1:])

    # BUG RÉEL CORRIGÉ (session F4, 1 fichier réel sur 21, 3 factures
    # bundlées) : sur certaines pages, l'ORDRE DES LIGNES ET DES CELLULES
    # est inversé de bout en bout (cohérent avec un scan de ce lot précis
    # effectué à l'envers, pas une erreur OCR ponctuelle — les 3 factures de
    # ce fichier montrent exactement la même inversion) : le pied de tableau
    # ("TOTAL HT"/"CODES TVA") apparaît AVANT les lignes d'articles, l'en-tête
    # lui-même ressort "MONTANT HT | P.U.NET HT | PRIX UNIT.HT | QTE |
    # DESIGNATION | REFERENCES" (la référence en DERNIER) au lieu de
    # "REFERENCES | DESIGNATION | ... MONTANT HT" habituel, et chaque ligne
    # d'article suit la même inversion de cellules. Le repli ci-dessus gère
    # l'inversion de LIGNES (cherche le pied AVANT l'en-tête si rien n'est
    # trouvé après) ; la détection ci-dessous (position de la cellule
    # "REFERENCES" dans la ligne d'en-tête, 2e moitié plutôt qu'en tête)
    # réinverse en plus CHAQUE ligne de la zone — le reste du code (repli
    # positionnel, ancre PF, regroupement multi-lignes...) suppose partout
    # l'ordre normal et n'a pas besoin d'être dupliqué pour ce cas.
    entete = lignes_groupees[i_entete]
    pos_ref = next(
        (i for i, m in enumerate(entete) if MOTIF_ENTETE_TABLEAU_ELECTRICPLUS.search(_sans_espaces_electricplus(m["texte"]))),
        None,
    )
    if pos_ref is not None and pos_ref > len(entete) / 2:
        zone = [list(reversed(ligne)) for ligne in zone]

    # Voir bandeau GABARIT BL (MOTIF_ENTETE_COLONNES_ELECTRICPLUS) : si la
    # 1re ligne de la zone mélange des cellules d'en-tête ("DESIGNATION",
    # "QTE"...) AVEC des cellules qui les précèdent (la référence/
    # désignation du 1er article, entraînées dans le même groupement Y),
    # ces cellules précédentes sont détachées et reportées en tête de la
    # ligne suivante plutôt que perdues avec le reste de la ligne d'en-tête.
    if zone:
        i_entete_colonne = next(
            (i for i, m in enumerate(zone[0]) if MOTIF_ENTETE_COLONNES_ELECTRICPLUS.match(m["texte"].strip())),
            None,
        )
        if i_entete_colonne is not None and i_entete_colonne > 0 and len(zone) > 1:
            cellules_reportees = zone[0][:i_entete_colonne]
            zone[1] = cellules_reportees + zone[1]
            zone = zone[1:]

    return zone


# BUG RÉEL CORRIGÉ (session F4, nouveau lot de factures GMR directement
# déposées par l'acheteur — pas encore vu sur les fixtures BL plus
# anciennes) : un article peut être imprimé sur PLUSIEURS lignes visuelles
# — référence + début de désignation sur une ligne, quantité/prix SANS
# suffixe PF/PR sur la ligne suivante, puis un complément de désignation
# (taille, ex. "2,4x180") sur une 3e ligne ENCORE APRÈS les prix (cas réel :
# 3 articles sur 4 d'une même facture, désignations longues qui débordent).
# Chaque ligne prise isolément échouait totalement (ni ancre PF, ni assez
# de cellules pour le repli positionnel) — ces 3 articles disparaissaient
# purement et simplement (0 ligne sur 3, seul le 4e article de la facture,
# resté sur une ligne unique, était extrait).
# _regrouper_articles_electricplus() reconstitue UNE liste de cellules par
# article, quel que soit son étalement d'origine : une ligne dont la 1re
# cellule ressemble à une vraie référence (MOTIF_DEBUT_REFERENCE_ELECTRICPLUS
# — lettres majuscules suivies d'au moins un chiffre, jamais vrai pour un
# nombre nu type "100,00" qui commence par un chiffre, ni pour un mot pur
# type "BLIST"/"2,4x180" qui n'a pas cette forme) démarre un NOUVEL article ;
# toute ligne suivante est absorbée dans l'article courant, CHAQUE cellule
# reclassée individuellement en désignation ou en nombre
# (_cellule_ressemble_a_nombre_electricplus) plutôt que simplement
# concaténée dans l'ordre de lecture — indispensable ici, car la cellule de
# désignation-complément arrive APRÈS les nombres dans l'ordre de lecture
# (2,4x180 est imprimé sous la ligne de prix, pas au-dessus) : une simple
# concaténation aurait décalé le repli positionnel "4 dernières cellules"
# et pris la désignation-complément pour le Montant.
# Sur un article déjà tenant sur une SEULE ligne (cas le plus courant,
# fixtures BL existantes comprises), ce regroupement reproduit exactement
# les mêmes cellules dans le même ordre — comportement STRICTEMENT
# inchangé, vérifié par les 14 tests BL déjà existants (aucune régression).
MOTIF_DEBUT_REFERENCE_ELECTRICPLUS = re.compile(r"^[A-Z]{2,}[A-Z0-9]*\d[A-Z0-9]*$")


def _cellule_ressemble_a_nombre_electricplus(cellule: str) -> bool:
    c = cellule.strip()
    if MOTIF_PF_ELECTRICPLUS.match(c):
        return True
    if MOTIF_QTE_OU_NOMBRE_ELECTRICPLUS.match(c):
        return True
    return bool(re.match(r"^\d[\d\s]*[.,]\d+$", c))


def _regrouper_articles_electricplus(lignes_cellules: list[list[str]]) -> list[list[str]]:

    articles: list[dict] = []

    for ligne in lignes_cellules:

        if ligne and MOTIF_DEBUT_REFERENCE_ELECTRICPLUS.match(ligne[0].strip()):
            articles.append({"reference": ligne[0], "designation": [], "nombres": []})
            reste = ligne[1:]
        elif articles:
            reste = ligne
        else:
            continue  # bruit avant toute référence détectée dans la zone

        for c in reste:
            cle = "nombres" if _cellule_ressemble_a_nombre_electricplus(c) else "designation"
            articles[-1][cle].append(c)

    return [[a["reference"]] + a["designation"] + a["nombres"] for a in articles]


def _champs_ligne_electricplus(cellules: list[str]) -> dict | None:
    """Extraction commune BL/Facture (voir bandeau GABARIT BL : GMR n'envoie
    pas de BL séparé, sa facture EST le BL — un seul jeu de règles
    d'extraction, deux constructeurs d'objet différents en aval, voir
    _ligne_vers_article_electricplus/_ligne_vers_ligne_facture_electricplus).
    Retourne {"reference", "designation", "quantite", "prix_net", "montant"}
    ou None."""

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

    # BUG RÉEL CORRIGÉ (2e lot, fichier multi-fournisseurs) : une ligne
    # chiffrée peut se retrouver SANS aucune référence/désignation
    # adjacente (le regroupement Y de l'OCR les a rattachées à une AUTRE
    # ligne, ou elles manquent purement et simplement sur ce document) —
    # cellules[0] est alors la QUANTITÉ elle-même ("70,00MTR"), jamais une
    # vraie référence. Mieux vaut ne PAS produire de ligne du tout (le
    # contrôle du Total HT du document signalera honnêtement qu'une ligne
    # manque) que d'écrire une "référence" qui n'en est pas une.
    if MOTIF_QTE_OU_NOMBRE_ELECTRICPLUS.match(reference):
        return None

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

    return {
        "reference": reference,
        "designation": designation,
        "quantite": round(montant / prix_net, 4),
        "prix_net": prix_net,
        "montant": montant,
    }


def _ligne_vers_article_electricplus(cellules: list[str]) -> LigneBL | None:

    champs = _champs_ligne_electricplus(cellules)
    if champs is None:
        return None

    return LigneBL(
        reference_fournisseur=champs["reference"],
        designation=champs["designation"],
        quantite_livree=champs["quantite"],
        prix_net=champs["prix_net"],
        montant=champs["montant"],
    )


def _ligne_vers_ligne_facture_electricplus(cellules: list[str]) -> LigneFacture | None:

    champs = _champs_ligne_electricplus(cellules)
    if champs is None:
        return None

    return LigneFacture(
        reference_fournisseur=champs["reference"],
        designation=champs["designation"],
        quantite_facturee=champs["quantite"],
        prix_unitaire_ht=champs["prix_net"],
        montant_ht=champs["montant"],
    )


def _entete_et_lignes_electricplus(mots_par_page: list[list[dict]]):
    """Extraction commune BL/Facture : n° de commande, n° + date du document
    (numero_bl côté BL = numero_facture côté Facture, c'est le MÊME champ
    imprimé — voir bandeau GABARIT BL), Total HT affiché, et les lignes du
    tableau déjà groupées en cellules (pas encore converties en LigneBL/
    LigneFacture, voir les deux appelants ci-dessous)."""

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

    numero_document, date_document = "", ""
    m = MOTIF_FACTURE_DATE_ELECTRICPLUS.search(texte)
    if m:
        date_document = _normaliser_date_electricplus(m.group(1))
        numero_document = m.group(2)
    else:
        m = MOTIF_FACTURE_DATE_ELECTRICPLUS_MIROIR.search(texte)
        if m:
            numero_document = m.group(1)
            date_document = _normaliser_date_electricplus(m.group(2))

    total_ht_affiche = None
    m = MOTIF_TOTAL_HT_ELECTRICPLUS.search(_sans_espaces_electricplus(texte))
    if m:
        total_ht_affiche = to_float(m.group(1))

    lignes_cellules = [
        [m["texte"] for m in ligne_mots]
        for mots in mots_par_page
        for ligne_mots in _zone_tableau_electricplus(regrouper_lignes(mots))
    ]

    return numero_commande, numero_document, date_document, total_ht_affiche, lignes_cellules


def _parse_une_facture_electricplus(mots_par_page: list[list[dict]]) -> BonLivraison:

    numero_commande, numero_bl, date_bl, total_ht_affiche, lignes_cellules = (
        _entete_et_lignes_electricplus(mots_par_page)
    )

    articles = []
    for cellules in lignes_cellules:
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


# --- GABARIT FACTURE (Electric Plus / GMR) ----------------------------------
# GMR n'envoie pas de BL séparé (voir bandeau GABARIT BL ci-dessus) : sa
# FACTURE fait déjà office de BL, donc du flux Facture (F4) comme du flux
# BL — même document, même OCR, même extraction (_entete_et_lignes_
# electricplus/_champs_ligne_electricplus, partagés). "numero_bl" côté BL
# et "numero_facture" côté Facture sont le MÊME champ imprimé sur le
# document (pas deux numéros différents à concilier).
#
# Ces factures sont des SCANS (comme le BL du même fournisseur) — jamais de
# texte PDF natif — d'où parse_facture_ocr (mots_par_page), pas parse_facture
# (texte) : voir moteur/rapprochement/lecture_facture.py, repli OCR générique
# quand lire_pdf() ne renvoie aucun texte.
#
# BUG RÉEL ÉVITÉ avant toute écriture (session F4) : suite à l'exigence du
# service comptable de l'acheteur (facture + NOTRE PROPRE bon de commande +
# éventuellement le DEVIS d'origine, agrafés ensemble — voir CLAUDE.md),
# un même fichier peut contenir des pages qui ne sont PAS des factures GMR :
# - une page DEVIS, avec SA PROPRE numérotation (6-7 chiffres, ex.
#   "4104132") pour le MÊME article que la facture ;
# - notre propre "BON DE COMMANDE" (généré par ce projet), qui peut porter
#   un nombre de 6-7 chiffres SANS RAPPORT (ex. une date collée "120720")
#   qui ressemble par coïncidence à un identifiant de regroupement.
# Sans filtrage, CHACUNE de ces pages peut démarrer un groupe à part entière
# (pages_par_identifiant) — une page DEVIS produirait une 2e "Facture"
# fantôme avec la ligne d'article DUPLIQUÉE (silencieux, potentiellement
# grave) ; une page de BON DE COMMANDE produit une "Facture" fantôme à 0
# ligne (inoffensif pour les montants mais fait basculer le fichier ENTIER
# vers "à vérifier" — même une facture par ailleurs parfaitement exacte,
# repéré en recette réelle sur 4205720.pdf : Total HT extrait = Total HT
# affiché au centime près, mais 2e "facture" fantôme à 0 ligne quand même
# produite par la page BC). _est_page_hors_perimetre_electricplus() exclut
# les DEUX AVANT tout regroupement par identifiant — repéré par "BON DE
# COMMANDE" (marqueur constant sur les 3 vraies pages BC confrontées à ce
# jour, malgré des libellés de détail différents ensuite : "DETAILSCOMMANDE"
# vs "DETAILDELACOMMANDE") ou "DEVIS". Exclure une page par erreur (ex. une
# page BC qui ne porte AUCUNE de ces mentions) est sans gravité : elle ne
# porte de toute façon aucune donnée de facture exploitable.
MOTIF_PAGE_HORS_PERIMETRE_ELECTRICPLUS = re.compile(r"DEVIS|BONDECOMMANDE", re.IGNORECASE)


def _est_page_hors_perimetre_electricplus(mots_page: list[dict]) -> bool:
    texte = _sans_espaces_electricplus(" ".join(m["texte"] for m in mots_page))
    return bool(MOTIF_PAGE_HORS_PERIMETRE_ELECTRICPLUS.search(texte))
# --- fin GABARIT FACTURE -----------------------------------------------------


def _construire_facture_electricplus(mots_par_page: list[list[dict]]) -> Facture:

    numero_commande, numero_facture, date_facture, total_ht_affiche, lignes_cellules_brutes = (
        _entete_et_lignes_electricplus(mots_par_page)
    )

    # Regroupement multi-lignes (voir bandeau ci-dessus) : ICI SEULEMENT,
    # jamais côté BL (_parse_une_facture_electricplus) — appliquer ce
    # regroupement là-bas a fait régresser 8 tests déjà verrouillés (une
    # ligne de bruit, auparavant isolée et donc silencieusement ignorée,
    # se retrouvait absorbée dans la désignation d'un article réel). Aucun
    # cas réel à ce jour ne montre ce besoin côté BL — seulement côté
    # Facture (nouveau lot de factures GMR déposées directement).
    lignes_cellules = _regrouper_articles_electricplus(lignes_cellules_brutes)

    lignes = []
    for cellules in lignes_cellules:
        ligne = _ligne_vers_ligne_facture_electricplus(cellules)
        if ligne:
            ligne.numero_bl = numero_facture  # un seul "bloc" : la facture entière (voir bandeau)
            lignes.append(ligne)

    facture = Facture(
        fournisseur="ELECTRIC PLUS",
        fichier="",
        numero_facture=numero_facture,
        date_facture=date_facture,
        numeros_commande=[numero_commande] if numero_commande else [],
        numeros_bl=[numero_facture] if numero_facture else [],
        lignes=lignes,
        total_ht_affiche=total_ht_affiche,
    )

    if total_ht_affiche is not None:
        total_extrait = round(sum((l.montant_ht or 0) for l in lignes), 2)
        if abs(total_ht_affiche - total_extrait) > 0.02:
            print(
                f"!! ELECTRIC PLUS (facture) : Total HT affiché ({total_ht_affiche:.2f}€) "
                f"!= somme des lignes extraites ({total_extrait:.2f}€) "
                f"— une ligne a peut-être été oubliée ou mal lue par l'OCR."
            )

    return facture


def parse_facture_electricplus_ocr(mots_par_page: list[list[dict]]) -> list[Facture]:
    """Un même fichier peut contenir PLUSIEURS factures Electric Plus
    distinctes (même découpage par page que parse_bl_electricplus, voir
    moteur.ocr.pages_par_identifiant — cas réel déjà rencontré côté BL,
    doc07149220260814105422.pdf). Les pages DEVIS/BON DE COMMANDE (voir
    bandeau GABARIT FACTURE) sont exclues AVANT le regroupement, jamais
    après."""

    pages_utiles = [mots for mots in mots_par_page if not _est_page_hors_perimetre_electricplus(mots)]

    if not pages_utiles:
        return []

    groupes_indices = pages_par_identifiant(pages_utiles, MOTIF_IDENTIFIANT_PAGE_ELECTRICPLUS)

    return [
        _construire_facture_electricplus([pages_utiles[i] for i in indices])
        for indices in groupes_indices
    ]


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
parse_facture_ocr = parse_facture_electricplus_ocr
