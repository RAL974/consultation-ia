"""
OCR pour les documents scannés (bons de livraison...) qui n'ont AUCUN texte
extractible par PyMuPDF — vérifié sur les 22 vrais BL déposés en session R1
de Rapprochement AI (0 caractère à chaque fois, voir CLAUDE.md) : ce sont
tous des scans image purs, même quand ils "ont l'air" d'un PDF net.

RapidOCR : pip pur (pas de binaire externe genre tesseract.exe à installer
à la main sur chaque poste), cohérent avec l'auto-installation du projet
(voir moteur/dependances.py). ~7s/page en régime établi sur ce poste
(~25s au tout premier appel, le temps de charger les modèles).

`lignes_ocr()` reconstruit, à partir des mots OCR positionnés, une liste de
lignes au MÊME FORMAT que `moteur.outils.lignes_propres()` sur du texte PDF
natif (une cellule d'un tableau = une ligne, dans l'ordre de lecture
haut->bas puis gauche->droite) : les gabarits fournisseurs peuvent ainsi
réutiliser telles quelles les primitives de moteur/fournisseurs/_gabarit.py
(scan_ancre notamment), déjà éprouvées sur les devis.
"""

from pathlib import Path

import fitz

_EXTENSIONS_IMAGE = (".jpg", ".jpeg", ".png")

_moteur = None


def _get_moteur():
    """Charge RapidOCR à la demande (import + init ~1-2s, coûteux à répéter)."""
    global _moteur
    if _moteur is None:
        from rapidocr_onnxruntime import RapidOCR
        _moteur = RapidOCR()
    return _moteur


def _images_pages(chemin: Path, dpi: int):
    """Rend chaque page d'un document en PNG (une image = une seule 'page')."""
    chemin = Path(chemin)

    if chemin.suffix.lower() in _EXTENSIONS_IMAGE:
        yield chemin.read_bytes()
        return

    with fitz.open(chemin) as doc:
        for page in doc:
            yield page.get_pixmap(dpi=dpi).tobytes("png")


def mots_document(chemin: Path, dpi: int = 200) -> list[list[dict]]:
    """
    OCR de chaque page d'un document (PDF ou image) -> une liste de mots par
    page, chaque mot étant {texte, score, x0, y0, x1, y1} (coordonnées en
    pixels de l'image rendue à `dpi`).
    """
    moteur = _get_moteur()
    pages = []

    for image in _images_pages(chemin, dpi):

        resultat, _ = moteur(image)

        mots = []
        for boite, texte, score in (resultat or []):
            xs = [point[0] for point in boite]
            ys = [point[1] for point in boite]
            mots.append({
                "texte": texte,
                "score": float(score),
                "x0": min(xs), "y0": min(ys),
                "x1": max(xs), "y1": max(ys),
            })

        pages.append(mots)

    return pages


def regrouper_lignes(mots: list[dict], tolerance_y: float = 12.0) -> list[list[dict]]:
    """
    Regroupe les mots d'UNE page en lignes visuelles (tolérance verticale de
    `tolerance_y` pixels, absorbe le léger tremblement des boîtes OCR).
    Retourne une liste de lignes, chacune la liste de ses mots triés de
    gauche à droite — la structure EN LIGNES (une ligne = plusieurs
    cellules) est conservée, contrairement à `lignes_ocr()` qui l'aplatit.
    Utile pour les gabarits BL qui ont besoin de savoir quelles cellules
    appartiennent à la MÊME ligne du tableau (ex. retrouver le total juste
    après le code TVA sur la bonne ligne, pas une autre).
    """
    mots_tries = sorted(mots, key=lambda m: (m["y0"] + m["y1"]) / 2)

    lignes = []
    ligne_courante = []
    y_ref = None

    for mot in mots_tries:
        yc = (mot["y0"] + mot["y1"]) / 2
        if y_ref is None or abs(yc - y_ref) <= tolerance_y:
            ligne_courante.append(mot)
            y_ref = yc if y_ref is None else (y_ref + yc) / 2
        else:
            lignes.append(sorted(ligne_courante, key=lambda m: m["x0"]))
            ligne_courante = [mot]
            y_ref = yc

    if ligne_courante:
        lignes.append(sorted(ligne_courante, key=lambda m: m["x0"]))

    return lignes


def lignes_ocr(mots: list[dict], tolerance_y: float = 12.0) -> list[str]:
    """
    Comme `regrouper_lignes()`, mais aplati en une liste de textes (un par
    mot), ligne par ligne de haut en bas, mot par mot de gauche à droite au
    sein d'une même ligne — même format que `moteur.outils.lignes_propres()`
    sur du texte PDF natif.
    """
    return [
        mot["texte"]
        for ligne in regrouper_lignes(mots, tolerance_y)
        for mot in ligne
    ]


def texte_ocr(chemin: Path, dpi: int = 200) -> str:
    """OCR complet d'un document -> texte façon PDF natif (une ligne par
    mot/cellule), toutes pages concaténées. Pratique pour la détection de
    fournisseur (moteur.detecteur) qui travaille sur du texte plein."""
    pages = mots_document(chemin, dpi)
    return "\n".join(
        ligne
        for mots in pages
        for ligne in lignes_ocr(mots)
    )


def _grouper_indices_par_identifiant(mots_par_page: list[list[dict]], motif_identifiant) -> list[list[int]]:
    """Même logique que grouper_pages_par_identifiant() (voir sa
    docstring), mais retourne les INDICES de page (0-based) de chaque
    groupe plutôt que leurs mots — factorisé pour servir à la fois
    grouper_pages_par_identifiant() (mots) et pages_par_identifiant()
    (indices, pour le découpage PDF par BL)."""

    groupes = []
    identifiant_courant = None

    for i, mots in enumerate(mots_par_page):

        texte_page = " ".join(mot["texte"] for mot in mots)
        m = motif_identifiant.search(texte_page)
        identifiant = (m.group(1) if m.lastindex else m.group(0)) if m else None

        if identifiant and identifiant != identifiant_courant:
            groupes.append([])
            identifiant_courant = identifiant

        if not groupes:
            groupes.append([])

        groupes[-1].append(i)

    return groupes


def grouper_pages_par_identifiant(mots_par_page: list[list[dict]], motif_identifiant) -> list[list[list[dict]]]:
    """
    Découpe `mots_par_page` (une entrée par page, voir mots_document) en
    groupes de pages consécutives partageant le MÊME identifiant de
    document — pour les fournisseurs qui scannent parfois PLUSIEURS BL à
    la suite dans un seul fichier (cas réel Cominter Ouest : jusqu'à 8 BL
    différents dans un même PDF).

    `motif_identifiant` est cherché (re.search) dans le texte aplati de
    CHAQUE page ; dès qu'une page révèle un identifiant DIFFÉRENT du
    groupe en cours, un nouveau groupe démarre. Une page sans identifiant
    trouvé (ex. un pied de page de total qui déborde sur la page
    suivante — cas réel, même BL sur 2 pages) rejoint le groupe précédent
    plutôt que d'en ouvrir un nouveau à tort.

    Si `motif_identifiant` a un groupe captant (1er groupe), c'est LUI qui
    sert de clé de comparaison (pas le texte entier du match) — pour
    ignorer un préfixe bruité par l'OCR (ex. "OBL108056" vs "0BL108056",
    O confondu avec 0 sur des scans différents de la MÊME page) tout en
    gardant les vrais chiffres identifiants.

    Retourne une liste de groupes, chacun au même format que
    `mots_par_page` (directement réutilisable par un parser BL existant,
    qui boucle déjà sur "mots_par_page" sans supposer qu'il s'agit d'un
    seul document)."""

    return [
        [mots_par_page[i] for i in groupe]
        for groupe in _grouper_indices_par_identifiant(mots_par_page, motif_identifiant)
    ]


def pages_par_identifiant(mots_par_page: list[list[dict]], motif_identifiant) -> list[list[int]]:
    """Comme grouper_pages_par_identifiant() (même regroupement, voir sa
    docstring), mais retourne les INDICES de page (0-based) de chaque
    groupe plutôt que leurs mots — pour découper le PDF source par BL
    individuel (voir moteur.rapprochement.pipeline_bl, archivage par BL
    individuel)."""

    return _grouper_indices_par_identifiant(mots_par_page, motif_identifiant)
