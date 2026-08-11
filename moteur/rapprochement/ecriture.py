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
COLONNES_MODIFIABLES = ("Date de livraison", "Qté livrée", "Tarif BL", "Note")

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

    wb = load_workbook(fichier, read_only=True, data_only=True)
    try:
        ws = wb[feuille]
        rapport = []
        for e in ecritures:
            idx = entetes[e.colonne]
            ancienne = ws.cell(row=e.ligne, column=idx).value
            rapport.append(
                {
                    "ligne": e.ligne,
                    "colonne": e.colonne,
                    "ancienne_valeur": ancienne,
                    "nouvelle_valeur": e.valeur,
                }
            )
        return rapport
    finally:
        wb.close()


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
