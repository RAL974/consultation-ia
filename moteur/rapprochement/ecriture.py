"""
Écriture sécurisée dans le classeur Suivi commandes (feuille "Commandes") —
socle de Rapprochement AI (voir CLAUDE.md, branche "Rapprochement AI").

Le classeur est VIVANT : ouvert dans Excel une bonne partie de la journée,
potentiellement par plusieurs personnes (Xavier compris). Trois garde-fous,
dans cet ordre :

1. Verrou Excel (fichier "~$..." à côté du classeur) : jamais d'écriture en
   force, `appliquer()` lève ClasseurVerrouille et laisse réessayer plus
   tard (voir est_verrouille()).
2. Sauvegarde horodatée AVANT toute écriture, avec rotation (voir
   sauvegarder()).
3. Écriture CHIRURGICALE : seule la valeur des cellules demandées est
   modifiée, en patchant directement la partie XML de la feuille visée dans
   le zip OOXML (un .xlsx est un zip) — jamais de désérialisation/
   réécriture complète du classeur.

   openpyxl (load_workbook() + save()) a été essayé en premier sur une
   COPIE du vrai Suivi commandes et écarté pour cet usage : même sans
   toucher à rien d'autre qu'une seule cellule, la réécriture complète fait
   disparaître xl/calcChain.xml, xl/metadata.xml et customXml/* (item de
   propriétés du document), et dégrade des validations de données
   ("Data Validation extension is not supported and will be removed",
   avertissement obtenu sur le classeur réel à l'ouverture) — inacceptable
   sur un fichier vivant riche en formules, tableaux structurés (16 Excel
   Tables) et validations. Le patch XML ciblé ne touche STRICTEMENT rien
   d'autre que le texte des cellules demandées ; chaque autre partie du zip
   (styles, tableaux, validations, calcChain, sharedStrings, customXml,
   printerSettings...) est recopiée octet pour octet.

Seules 4 colonnes de la feuille "Commandes" sont de VRAIES données saisies
au clavier (vérifié cellule par cellule sur le classeur réel) :
"Date de livraison", "Qté livrée", "Tarif BL", et "Note" (texte libre déjà
utilisé comme valeur "magique" par plusieurs formules : "Rupture
fournisseur", "Reliquat soldé", "Commande annulée"). TOUT le reste de la
feuille — Statut commande, Reliquat, RAL, Soldé, Reste à facturer, Facturé
BL, Potentiel factu... — est calculé par une formule Excel à partir de ces
4 colonnes (XLOOKUP/IF en cascade côté "Statut commande" notamment) :
écrire une valeur figée dedans romprait le calcul à la prochaine ouverture
Excel, silencieusement. `appliquer()` refuse toute colonne hors de
COLONNES_MODIFIABLES.

5 colonnes SUPPLÉMENTAIRES, propres au rapprochement des FACTURES (voir
CLAUDE.md, session de création des colonnes facture) : "N° facture",
"Date facture", "Qté facturée", "PU facturé", "Montant facturé HT" — même
statut que les 4 premières (saisie brute). Contrairement à la 1ère
proposition (session F2, "Volet 1", jamais appliquée dans Excel) où
l'acheteur devait créer ces colonnes à la main, elles sont ajoutées par
l'OUTIL lui-même (`ajouter_entetes_saisie()`, ci-dessous) — jamais de
manipulation Excel demandée à l'acheteur. Elles sont intégrées DANS le
tableau structuré Commandes (pas posées à côté, contrairement à un 1er
essai abandonné en cours de session) : décision explicite de l'acheteur —
une colonne hors tableau ne suit pas un tri du tableau, les données
facture se retrouveraient sur les mauvaises lignes.
`ajouter_entetes_saisie()` patche donc DEUX parties du zip (la feuille ET
la définition XML du tableau : ref, autoFilter, tableColumns), jamais
plus — aucun calculatedColumnFormula, aucune totalsRow n'accompagne ces
colonnes : "Montant facturé HT" est ici une simple colonne de saisie, pas
(encore) un calcul ; "Écart facture" et "Statut facture" restent HORS
PÉRIMÈTRE, reportés à une session dédiée où l'outil écrira lui-même les
formules dans le XML (jamais collées à la main par l'acheteur). "Statut
commande" (la formule existante, dont dépend le Dashboard) n'est et ne
sera jamais modifié par ce module.
"""

import re
import shutil
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.datetime import to_excel as _date_vers_excel

FEUILLE_COMMANDES = "Commandes"

# Les seules colonnes que Rapprochement AI a le droit d'écrire (voir
# bandeau ci-dessus) — toute autre colonne de "Commandes" est une formule.
COLONNES_MODIFIABLES = (
    "Date de livraison", "Qté livrée", "Tarif BL", "Note",
    "N° facture", "Date facture", "Qté facturée", "PU facturé",
    "Montant facturé HT",
)

# En-têtes ajoutés DANS le tableau structuré Commandes par
# ajouter_entetes_saisie(), dans cet ordre exact (voir CLAUDE.md) — la
# prochaine session (rapprochement des factures) les utilise tel quel.
ENTETES_FACTURE = (
    "N° facture", "Date facture", "Qté facturée", "PU facturé",
    "Montant facturé HT",
)

RETENTION_BACKUPS_JOURS = 30

_NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_NS_PKGREL = "{http://schemas.openxmlformats.org/package/2006/relationships}"


class ClasseurVerrouille(Exception):
    """Le Suivi commandes est actuellement ouvert dans Excel (fichier ~$ présent)."""


class ColonneNonModifiable(ValueError):
    """Tentative d'écriture dans une colonne calculée par formule (ou
    absente des en-têtes) — jamais autorisée, voir COLONNES_MODIFIABLES."""


@dataclass(frozen=True)
class Ecriture:
    """Une écriture demandée : ligne Excel 1-based (la ligne 1 est
    l'en-tête, donc la 1ère commande est en ligne 2), nom d'en-tête (doit
    être dans COLONNES_MODIFIABLES), valeur Python (str / int / float /
    date/datetime)."""

    ligne: int
    colonne: str
    valeur: object


def est_verrouille(fichier) -> bool:
    """True si `fichier` est actuellement ouvert dans Excel (présence du
    fichier de verrou "~$<nom>" dans le même dossier)."""

    fichier = Path(fichier)
    return (fichier.parent / f"~${fichier.name}").exists()


def sauvegarder(fichier, dossier_backups) -> Path:
    """Copie horodatée de `fichier` dans `dossier_backups`, puis purge les
    sauvegardes de plus de RETENTION_BACKUPS_JOURS jours. Retourne le
    chemin de la copie créée."""

    fichier = Path(fichier)
    dossier_backups = Path(dossier_backups)
    dossier_backups.mkdir(parents=True, exist_ok=True)

    horodatage = datetime.now().strftime("%Y%m%d_%H%M%S")
    cible = dossier_backups / f"{fichier.stem}_{horodatage}{fichier.suffix}"
    compteur = 1
    while cible.exists():  # deux sauvegardes dans la même seconde (tests, double-clic...)
        compteur += 1
        cible = dossier_backups / f"{fichier.stem}_{horodatage}_{compteur}{fichier.suffix}"
    shutil.copy2(fichier, cible)

    _purger_backups_anciens(dossier_backups)
    return cible


def _purger_backups_anciens(dossier_backups, jours=RETENTION_BACKUPS_JOURS):
    seuil = datetime.now().timestamp() - jours * 86400
    for f in Path(dossier_backups).glob("*.xlsx"):
        if f.stat().st_mtime < seuil:
            f.unlink()


def lire_entetes(fichier, feuille=FEUILLE_COMMANDES) -> dict:
    """{nom d'en-tête: index de colonne 1-based} de la 1ère ligne de
    `feuille`."""

    wb = load_workbook(fichier, read_only=True, data_only=True)
    try:
        ws = wb[feuille]
        premiere = next(ws.iter_rows(min_row=1, max_row=1))
        return {
            str(c.value).strip(): i + 1
            for i, c in enumerate(premiere)
            if c.value is not None
        }
    finally:
        wb.close()


def _verifier_colonnes(ecritures, entetes):
    for e in ecritures:
        if e.colonne not in COLONNES_MODIFIABLES:
            raise ColonneNonModifiable(
                f"« {e.colonne} » n'est pas une colonne saisissable (c'est une "
                f"formule dans le Suivi commandes) — colonnes autorisées : "
                f"{', '.join(COLONNES_MODIFIABLES)}"
            )
        if e.colonne not in entetes:
            raise ColonneNonModifiable(
                f"« {e.colonne} » introuvable dans les en-têtes de la feuille "
                f"« {FEUILLE_COMMANDES} » — export du Suivi commandes différent "
                f"de celui attendu ?"
            )


def simuler(fichier, ecritures, feuille=FEUILLE_COMMANDES) -> list:
    """Rapport « ce qui sera écrit », sans rien modifier dans `fichier`.
    Une entrée par Ecriture : ligne, colonne, ancienne_valeur,
    nouvelle_valeur. C'est ce rapport que l'acheteur doit voir et confirmer
    avant tout appel à appliquer() — mode simulation par défaut."""

    entetes = lire_entetes(fichier, feuille)
    _verifier_colonnes(ecritures, entetes)

    # Une seule passe séquentielle sur la feuille : en mode read_only,
    # ws.cell(row=, column=) ré-analyse le flux XML depuis le début à
    # CHAQUE appel (accès aléatoire non prévu par openpyxl en read_only) —
    # sur ~5 900 lignes, des dizaines d'écritures rendaient ça minutes,
    # voire jamais fini. On collecte donc d'abord {ligne -> {colonne: idx}}
    # demandées, puis on ne lit qu'une fois chaque ligne concernée.
    lignes_demandees = {}
    for e in ecritures:
        lignes_demandees.setdefault(e.ligne, set()).add(entetes[e.colonne])

    derniere_ligne = max(lignes_demandees)
    valeurs_lues = {}  # (ligne, idx_colonne) -> valeur actuelle
    wb = load_workbook(fichier, read_only=True, data_only=True)
    try:
        ws = wb[feuille]
        for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if i in lignes_demandees:
                for idx in lignes_demandees[i]:
                    valeurs_lues[(i, idx)] = row[idx - 1] if idx - 1 < len(row) else None
            if i >= derniere_ligne:
                break
    finally:
        wb.close()

    rapport = []
    for e in ecritures:
        idx = entetes[e.colonne]
        rapport.append(
            {
                "ligne": e.ligne,
                "colonne": e.colonne,
                "ancienne_valeur": valeurs_lues.get((e.ligne, idx)),
                "nouvelle_valeur": e.valeur,
            }
        )
    return rapport


def appliquer(fichier, ecritures, dossier_backups, feuille=FEUILLE_COMMANDES) -> Path:
    """Écrit réellement `ecritures` dans `fichier` (feuille `feuille`) :
    verrou -> sauvegarde -> patch XML chirurgical. Lève ClasseurVerrouille
    si le fichier est ouvert dans Excel (jamais d'écriture en force) et
    ColonneNonModifiable si une écriture vise une colonne hors de
    COLONNES_MODIFIABLES (jamais de compromis là-dessus, voir bandeau).
    Retourne le chemin de la sauvegarde créée avant écriture."""

    fichier = Path(fichier)
    if est_verrouille(fichier):
        raise ClasseurVerrouille(
            f"« {fichier.name} » est actuellement ouvert dans Excel — "
            "ferme-le puis réessaie (rien n'a été écrit)."
        )

    entetes = lire_entetes(fichier, feuille)
    _verifier_colonnes(ecritures, entetes)

    sauvegarde = sauvegarder(fichier, dossier_backups)

    cellules = {}
    for e in ecritures:
        idx = entetes[e.colonne]
        ref = f"{get_column_letter(idx)}{e.ligne}"
        cellules[ref] = e.valeur

    _patcher_cellules_xlsx(fichier, feuille, cellules)
    return sauvegarde


def ajouter_entetes_saisie(fichier, noms, dossier_backups, feuille=FEUILLE_COMMANDES) -> Path:
    """Ajoute `noms` comme nouvelles colonnes DANS le tableau structuré de
    `feuille` (même nom que la feuille, ex. "Commandes"), à sa suite : une
    cellule d'en-tête (inlineStr) en ligne 1 pour chaque nom, ET les
    <tableColumn> correspondants ajoutés à la définition du tableau (avec
    extension de sa `ref` et de son `<autoFilter>`). Sciemment DANS le
    tableau, pas à côté : une colonne hors tableau ne suit pas un tri du
    tableau, les données se retrouveraient sur les mauvaises lignes (voir
    CLAUDE.md).

    Patch exactement DEUX parties du zip, rien d'autre : la feuille (ligne 1
    + <dimension>) et la définition XML du tableau (ref, autoFilter,
    tableColumns — AUCUN calculatedColumnFormula, AUCUNE totalsRow ajoutée).
    Même esprit chirurgical que appliquer() (voir bandeau) : tout le reste
    du classeur (styles, calcChain, sharedStrings, customXml, les 15 autres
    tableaux, les autres feuilles...) est recopié octet pour octet.

    Mêmes garde-fous que appliquer() : ClasseurVerrouille si le fichier est
    ouvert dans Excel (jamais d'écriture en force), sauvegarde horodatée
    AVANT toute écriture. Lève ValueError si un `nom` existe déjà (en-tête
    de la feuille OU colonne du tableau), ou si une cellule cible n'est pas
    vide (jamais d'écrasement d'une donnée existante). Retourne le chemin
    de la sauvegarde créée avant écriture."""

    fichier = Path(fichier)
    if not noms:
        raise ValueError("Aucun nom d'en-tête fourni.")
    if len(set(noms)) != len(noms):
        raise ValueError(f"Noms d'en-tête en double dans {noms!r}.")

    if est_verrouille(fichier):
        raise ClasseurVerrouille(
            f"« {fichier.name} » est actuellement ouvert dans Excel — "
            "ferme-le puis réessaie (rien n'a été écrit)."
        )

    entetes = lire_entetes(fichier, feuille)
    for nom in noms:
        if nom in entetes:
            raise ValueError(
                f"« {nom} » existe déjà dans les en-têtes de « {feuille} » "
                f"(colonne {get_column_letter(entetes[nom])}) — rien n'a été écrit."
            )

    derniere_colonne = max(entetes.values()) if entetes else 0
    refs = [
        f"{get_column_letter(derniere_colonne + 1 + i)}1" for i in range(len(noms))
    ]

    with zipfile.ZipFile(fichier, "r") as zin:
        chemin_feuille = _chemin_feuille(zin, feuille)
        xml_feuille = zin.read(chemin_feuille).decode("utf-8")
        chemin_table = _chemin_table_par_nom(zin, feuille)
        xml_table = zin.read(chemin_table).decode("utf-8")

    _verifier_cellules_vides(xml_feuille, 1, refs)
    _verifier_colonnes_table_absentes(xml_table, noms)

    colonne_min_requise = derniere_colonne + len(noms)
    nouvelle_dimension = _dimension_etendue(xml_feuille, colonne_min_requise)
    xml_table = _etendre_tableau(xml_table, noms, colonne_min_requise)

    sauvegarde = sauvegarder(fichier, dossier_backups)

    xml_feuille = _remplacer_dans_ligne(xml_feuille, 1, dict(zip(refs, noms)))
    xml_feuille = re.sub(
        r'<dimension ref="[^"]*"\s*/>',
        f'<dimension ref="{nouvelle_dimension}"/>',
        xml_feuille,
        count=1,
    )

    _patcher_parties_xlsx(fichier, {chemin_feuille: xml_feuille, chemin_table: xml_table})
    return sauvegarde


# --- Patch XML chirurgical -------------------------------------------------


def _chemin_feuille(zin: zipfile.ZipFile, nom_feuille: str) -> str:
    """xl/worksheets/sheetN.xml correspondant au nom d'onglet `nom_feuille`,
    en suivant xl/workbook.xml -> xl/_rels/workbook.xml.rels (lecture seule,
    ces deux fichiers de contrôle ne sont jamais réécrits)."""

    workbook_xml = ET.fromstring(zin.read("xl/workbook.xml"))
    rid = None
    for sheet in workbook_xml.iter(f"{_NS_MAIN}sheet"):
        if sheet.get("name") == nom_feuille:
            rid = sheet.get(f"{_NS_REL}id")
            break
    if rid is None:
        raise ValueError(f"Feuille « {nom_feuille} » introuvable dans xl/workbook.xml")

    rels_xml = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
    for rel in rels_xml.iter(f"{_NS_PKGREL}Relationship"):
        if rel.get("Id") == rid:
            cible = rel.get("Target")
            # Target est soit relatif à xl/ ("worksheets/sheet1.xml", forme
            # écrite par Excel), soit relatif à la racine du package
            # ("/xl/worksheets/sheet1.xml", forme écrite par openpyxl).
            return cible.lstrip("/") if cible.startswith("/") else "xl/" + cible
    raise ValueError(f"Relation « {rid} » introuvable dans xl/_rels/workbook.xml.rels")


def _verifier_cellules_vides(xml_feuille: str, num_ligne: int, refs) -> None:
    """Lève ValueError si l'une de `refs` (ex. "AY1") existe déjà comme
    cellule dans la ligne `num_ligne` du XML — jamais d'écrasement d'une
    cellule déjà présente, même vide de valeur."""

    motif_ligne = re.compile(rf'<row r="{num_ligne}"[^>]*>(.*?)</row>', re.S)
    m = motif_ligne.search(xml_feuille)
    contenu = m.group(1) if m else ""
    for ref in refs:
        if re.search(rf'<c r="{ref}"', contenu):
            raise ValueError(
                f"La cellule {ref} existe déjà dans la feuille — écriture "
                "refusée (rien n'a été écrit)."
            )


def _etendre_plage(ref: str, colonne_min_requise: int) -> str:
    """"A1:AX6420" -> "A1:BC6420" si `colonne_min_requise` (1-based) dépasse
    la colonne de fin actuelle — ne réduit jamais, ne touche jamais les
    lignes ni la colonne de début. Partagé par <dimension>, <table ref=...>
    et <autoFilter ref=...> (même forme de plage partout)."""

    m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)$", ref)
    if not m:
        raise ValueError(f"Plage de cellules de forme inattendue : {ref!r}")
    debut_col, debut_ligne, fin_col, fin_ligne = m.groups()
    nouvelle_fin_col = get_column_letter(
        max(column_index_from_string(fin_col), colonne_min_requise)
    )
    return f"{debut_col}{debut_ligne}:{nouvelle_fin_col}{fin_ligne}"


def _dimension_etendue(xml_feuille: str, colonne_min_requise: int) -> str:
    """Nouvelle valeur de <dimension ref="..."/> couvrant au moins la
    colonne `colonne_min_requise` (1-based), même ligne de fin qu'avant."""

    m = re.search(r'<dimension ref="([^"]*)"\s*/>', xml_feuille)
    if not m:
        raise ValueError("<dimension> introuvable ou de forme inattendue dans la feuille.")
    return _etendre_plage(m.group(1), colonne_min_requise)


def _chemin_table_par_nom(zin: zipfile.ZipFile, nom_table: str) -> str:
    """xl/tables/tableN.xml dont l'attribut du <table> racine name="..."
    vaut `nom_table` — recherche directe dans xl/tables/ (pas besoin de
    suivre feuille -> _rels -> table, un seul tableau porte ce nom dans
    tout le classeur)."""

    for nom in zin.namelist():
        if re.match(r"xl/tables/table\d+\.xml$", nom):
            data = zin.read(nom).decode("utf-8")
            if re.search(rf'<table\b[^>]*\bname="{re.escape(nom_table)}"', data):
                return nom
    raise ValueError(f"Tableau structuré « {nom_table} » introuvable dans xl/tables/")


def _verifier_colonnes_table_absentes(xml_table: str, noms) -> None:
    """Lève ValueError si l'un de `noms` est déjà le name= d'un
    <tableColumn> existant — vérification indépendante de lire_entetes()
    (qui lit la ligne 1 de la FEUILLE) pour ne jamais dépendre d'une
    hypothèse de synchronisation entre la feuille et le tableau."""

    m = re.search(r"<tableColumns\b[^>]*>(.*)</tableColumns>", xml_table, re.S)
    contenu = m.group(1) if m else ""
    existants = set(re.findall(r'<tableColumn\b[^>]*\bname="([^"]*)"', contenu))
    for nom in noms:
        if nom in existants:
            raise ValueError(
                f"« {nom} » existe déjà comme colonne du tableau structuré — "
                "rien n'a été écrit."
            )


def _etendre_tableau(xml_table: str, noms, colonne_min_requise: int) -> str:
    """Étend la définition XML d'un tableau structuré (xl/tables/tableN.xml)
    pour y ajouter `noms` comme nouvelles colonnes, à la fin (l'ORDRE des
    <tableColumn> dans le XML est l'ordre gauche->droite affiché par Excel
    — vérifié sur le vrai classeur, où l'attribut id="" n'est PAS
    séquentiel par position) :

    - `ref` du <table> et, s'il existe, du <autoFilter> : étendus pour
      couvrir la nouvelle dernière colonne (jamais les lignes) ;
    - <tableColumns count="..."> : compte mis à jour ;
    - un <tableColumn id="…" name="…"/> par nom — SEULEMENT id et name
      (jamais de calculatedColumnFormula : ces colonnes sont de la saisie
      pure, pas un calcul hérité par une future ligne du tableau) — id =
      max(id existants) + 1, +2, ... pour rester unique."""

    m_table = re.search(r'(<table\b[^>]*\sref=")([^"]*)(")', xml_table)
    if not m_table:
        raise ValueError('<table ref="..."> introuvable ou de forme inattendue.')
    nouvelle_ref = _etendre_plage(m_table.group(2), colonne_min_requise)
    xml_table = (
        xml_table[: m_table.start()]
        + m_table.group(1) + nouvelle_ref + m_table.group(3)
        + xml_table[m_table.end() :]
    )

    m_filtre = re.search(r'(<autoFilter\b[^>]*\sref=")([^"]*)(")', xml_table)
    if m_filtre:
        nouvelle_ref_filtre = _etendre_plage(m_filtre.group(2), colonne_min_requise)
        xml_table = (
            xml_table[: m_filtre.start()]
            + m_filtre.group(1) + nouvelle_ref_filtre + m_filtre.group(3)
            + xml_table[m_filtre.end() :]
        )

    m_cols = re.search(r'<tableColumns\b[^>]*\bcount="(\d+)"', xml_table)
    if not m_cols:
        raise ValueError('<tableColumns count="..."> introuvable ou de forme inattendue.')
    nouveau_compte = int(m_cols.group(1)) + len(noms)
    xml_table = xml_table[: m_cols.start(1)] + str(nouveau_compte) + xml_table[m_cols.end(1) :]

    ids_existants = [int(i) for i in re.findall(r'<tableColumn\b[^>]*\bid="(\d+)"', xml_table)]
    id_suivant = max(ids_existants, default=0) + 1

    fragments = []
    for i, nom in enumerate(noms):
        nom_xml = nom
        for a, b in (("&", "&amp;"), ("<", "&lt;"), (">", "&gt;"), ('"', "&quot;")):
            nom_xml = nom_xml.replace(a, b)
        fragments.append(f'<tableColumn id="{id_suivant + i}" name="{nom_xml}"/>')

    position = xml_table.rindex("</tableColumns>")
    return xml_table[:position] + "".join(fragments) + xml_table[position:]


def _contenu_cellule(valeur):
    """(xml_interne, attribut_t) pour la valeur Python donnée. Jamais de
    sharedStrings.xml pour du texte : chaîne "inline" (t="inlineStr"), pour
    ne pas toucher une table partagée par des milliers d'autres cellules."""

    if isinstance(valeur, bool):
        return (f"<v>{1 if valeur else 0}</v>", ' t="b"')
    if isinstance(valeur, (int, float)):
        texte = repr(float(valeur)) if isinstance(valeur, float) else str(valeur)
        return (f"<v>{texte}</v>", "")
    if isinstance(valeur, (datetime, date)):
        return (f"<v>{repr(float(_date_vers_excel(valeur)))}</v>", "")

    texte = str(valeur)
    for a, b in (("&", "&amp;"), ("<", "&lt;"), (">", "&gt;")):
        texte = texte.replace(a, b)
    return (f'<is><t xml:space="preserve">{texte}</t></is>', ' t="inlineStr"')


def _ecrire_cellule(contenu_ligne: str, ref: str, valeur) -> str:
    """Remplace (ou insère, triée par colonne) la cellule `ref` dans le
    contenu XML d'une <row>...</row> (sans les balises <row> elles-mêmes)."""

    contenu_valeur, attr_t = _contenu_cellule(valeur)
    motif_cellule = re.compile(rf'<c r="{ref}"([^>]*?)(?:/>|>.*?</c>)', re.S)
    m = motif_cellule.search(contenu_ligne)
    if m:
        attrs = re.sub(r'\s+t="[^"]*"', "", m.group(1))
        nouvelle_cellule = f'<c r="{ref}"{attrs}{attr_t}>{contenu_valeur}</c>'
        return contenu_ligne[: m.start()] + nouvelle_cellule + contenu_ligne[m.end() :]

    # Cellule absente du XML (jamais écrite depuis la création de la ligne) :
    # insertion triée par colonne — un ordre non croissant fait "réparer" le
    # fichier par Excel à l'ouverture.
    nouvelle_cellule = f'<c r="{ref}"{attr_t}>{contenu_valeur}</c>'
    col_cible = column_index_from_string(re.match(r"[A-Z]+", ref).group())
    position = len(contenu_ligne)
    for m2 in re.finditer(r'<c r="([A-Z]+)\d+"', contenu_ligne):
        if column_index_from_string(m2.group(1)) > col_cible:
            position = m2.start()
            break
    return contenu_ligne[:position] + nouvelle_cellule + contenu_ligne[position:]


def _remplacer_dans_ligne(xml_feuille: str, num_ligne: int, remplacements: dict) -> str:
    motif_ligne = re.compile(rf'(<row r="{num_ligne}"[^>]*>)(.*?)(</row>)', re.S)
    m = motif_ligne.search(xml_feuille)
    if not m:
        raise ValueError(
            f"Ligne {num_ligne} introuvable dans la feuille (ligne vide ou "
            "hors de la plage de données ?)"
        )
    ouverture, contenu, fermeture = m.group(1), m.group(2), m.group(3)

    for ref, valeur in remplacements.items():
        contenu = _ecrire_cellule(contenu, ref, valeur)

    return xml_feuille[: m.start()] + ouverture + contenu + fermeture + xml_feuille[m.end() :]


def _patcher_parties_xlsx(fichier, parties: dict) -> None:
    """Réécrit `fichier` en place : chaque partie nommée dans `parties`
    (chemin dans le zip -> nouveau contenu XML str) est remplacée en une
    seule passe atomique ; TOUTES les autres parties sont recopiées octet
    pour octet depuis l'original. Utilisé par ajouter_entetes_saisie(), qui
    doit patcher DEUX parties (la feuille et la définition du tableau) sans
    laisser le zip dans un état intermédiaire — _patcher_cellules_xlsx()
    reste inchangée, dédiée à appliquer() (une seule partie, la feuille)."""

    fichier = Path(fichier)
    fichier_tmp = fichier.with_suffix(fichier.suffix + ".tmp")

    with zipfile.ZipFile(fichier, "r") as zin:
        try:
            with zipfile.ZipFile(fichier_tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    donnees = (
                        parties[item.filename].encode("utf-8")
                        if item.filename in parties
                        else zin.read(item.filename)
                    )
                    zout.writestr(item, donnees)
        except Exception:
            fichier_tmp.unlink(missing_ok=True)
            raise

    fichier_tmp.replace(fichier)


def _patcher_cellules_xlsx(fichier, feuille: str, cellules_par_ref: dict):
    """Réécrit `fichier` en place : seule la partie XML de `feuille` est
    modifiée dans le zip, toutes les autres parties (styles, tableaux,
    validations, calcChain, sharedStrings, customXml, printerSettings...)
    sont recopiées octet pour octet depuis l'original."""

    fichier = Path(fichier)
    fichier_tmp = fichier.with_suffix(fichier.suffix + ".tmp")

    with zipfile.ZipFile(fichier, "r") as zin:
        chemin_feuille = _chemin_feuille(zin, feuille)
        xml_feuille = zin.read(chemin_feuille).decode("utf-8")

        cellules_par_ligne = {}
        for ref, valeur in cellules_par_ref.items():
            num_ligne = int(re.match(r"[A-Z]+(\d+)", ref).group(1))
            cellules_par_ligne.setdefault(num_ligne, {})[ref] = valeur

        for num_ligne, remplacements in cellules_par_ligne.items():
            xml_feuille = _remplacer_dans_ligne(xml_feuille, num_ligne, remplacements)

        try:
            with zipfile.ZipFile(fichier_tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    donnees = (
                        xml_feuille.encode("utf-8")
                        if item.filename == chemin_feuille
                        else zin.read(item.filename)
                    )
                    zout.writestr(item, donnees)
        except Exception:
            fichier_tmp.unlink(missing_ok=True)
            raise

    fichier_tmp.replace(fichier)
