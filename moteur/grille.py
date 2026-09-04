"""
Regroupement de mots PDF natifs (PyMuPDF `page.get_text("words")`) en
lignes visuelles triées par position Y — même principe que
`moteur.ocr.regrouper_lignes` (mots OCR), mais pour du texte PDF natif :
les mots sortent de PyMuPDF dans un ordre de LECTURE (souvent scramblé sur
les gabarits à colonnes/tableaux superposés, voir moteur.fournisseurs.
coredime, GABARIT FACTURE), jamais un ordre VISUEL fiable — `lignes()`
reconstruit cet ordre visuel en regroupant les mots par Y.

Contrairement à `moteur.ocr.regrouper_lignes` (tolérance FIXE 12px, pensée
pour une image rendue à un DPI connu), la tolérance ici est dérivée de la
demi-hauteur MÉDIANE des mots de la page : un PDF natif est en points
(taille de police réelle du document, jamais supposée fixe d'un
fournisseur à l'autre).

Sert quand la POSITION RELATIVE de deux lignes (laquelle est au-dessus de
l'autre) est le seul moyen fiable de les rattacher l'une à l'autre — ex.
Coredime : une ligne "Remise x+y%" imprimée à un endroit du flux PyMuPDF
totalement déconnecté de sa ligne d'article, mais toujours juste EN
DESSOUS d'elle visuellement (voir moteur.fournisseurs.coredime,
_apparier_par_position_coredime).
"""

import statistics


def mots(page) -> list[dict]:
    """`page.get_text("words")` (PyMuPDF) -> liste de mots
    {texte, x0, y0, x1, y1} — même format que moteur.ocr.mots_document,
    pour rester compatible avec le même style de consommateurs (ex.
    regrouper par ligne, chercher un mot par sous-texte)."""

    return [
        {"texte": t, "x0": x0, "y0": y0, "x1": x1, "y1": y1}
        for x0, y0, x1, y1, t, *_ in page.get_text("words")
    ]


def lignes(mots_page: list[dict]) -> list[list[dict]]:
    """Regroupe `mots_page` en lignes visuelles (triées par Y puis, au
    sein d'une ligne, par X) — tolérance verticale = demi-hauteur MÉDIANE
    des mots de la page (pas une constante : la taille de police varie
    d'un document à l'autre, contrairement à une image OCR rendue à un
    DPI fixe)."""

    if not mots_page:
        return []

    demi_hauteurs = [(m["y1"] - m["y0"]) / 2 for m in mots_page if m["y1"] > m["y0"]]
    tolerance = statistics.median(demi_hauteurs) if demi_hauteurs else 5.0

    mots_tries = sorted(mots_page, key=lambda m: (m["y0"] + m["y1"]) / 2)

    resultat = []
    ligne_courante = []
    y_ref = None

    for mot in mots_tries:
        yc = (mot["y0"] + mot["y1"]) / 2
        if y_ref is None or abs(yc - y_ref) <= tolerance:
            ligne_courante.append(mot)
            y_ref = yc if y_ref is None else (y_ref + yc) / 2
        else:
            resultat.append(sorted(ligne_courante, key=lambda m: m["x0"]))
            ligne_courante = [mot]
            y_ref = yc

    if ligne_courante:
        resultat.append(sorted(ligne_courante, key=lambda m: m["x0"]))

    return resultat
