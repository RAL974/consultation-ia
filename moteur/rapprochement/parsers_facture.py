"""
Registre des parsers facture fournisseurs — même principe que
moteur/rapprochement/parsers_bl.py (BL) et moteur/parsers.py (devis).

Pour ajouter le support facture d'un fournisseur dont les factures sont en
texte PDF natif (majorité des cas à ce jour — 109 Distribution, Coredime,
Cominter, Stand 64...), ajouter à moteur/fournisseurs/<fournisseur>.py :

    def parse_facture(texte: str):    # texte PDF natif — voir
        ...                            # moteur/rapprochement/lecture_facture.py
        return Facture(...) | None

Pour un fournisseur dont les factures sont des SCANS (aucun texte natif —
cas réel : Electric Plus/GMR, dont la facture fait aussi office de BL, voir
moteur/fournisseurs/electricplus.py), exposer plutôt :

    def parse_facture_ocr(mots_par_page: list[list[dict]]):  # voir moteur/ocr.py
        ...
        return Facture(...) | list[Facture] | None

`lecture_facture.lire_facture()` n'essaie l'OCR QUE si le texte natif est
vide (règle générique, pas un cas spécial câblé sur un nom de fournisseur —
profite à tout futur fournisseur scanné). Un fournisseur peut exposer les
DEUX (texte natif la plupart du temps, repli OCR sur un scan occasionnel) ;
aucun cas réel de ce genre à ce jour.

`FOURNISSEURS` (déjà déclaré pour les devis/BL) sert aussi de clé ici. Un
fournisseur sans parse_facture/parse_facture_ocr reste utilisable pour les
devis/BL normalement.
"""

import importlib
import inspect
import pkgutil

import moteur.fournisseurs as paquet


def _charger_parsers(attribut: str) -> dict:
    parsers = {}

    for info in pkgutil.iter_modules(paquet.__path__):

        if info.name.startswith("_"):
            continue

        module = importlib.import_module(f"moteur.fournisseurs.{info.name}")

        fn = getattr(module, attribut, None)
        noms = getattr(module, "FOURNISSEURS", None)

        if fn is None or not noms:
            continue

        for nom in noms:
            parsers[nom.upper()] = fn

    return parsers


PARSERS_FACTURE = _charger_parsers("parse_facture")
PARSERS_FACTURE_OCR = _charger_parsers("parse_facture_ocr")


def parser_facture(fournisseur: str, texte: str, chemin=None):
    """`chemin` (optionnel, chemin du PDF source) : transmis au parser
    UNIQUEMENT s'il déclare explicitement un paramètre `chemin` dans sa
    signature (ex. moteur.fournisseurs.coredime.parse_facture_coredime,
    voir son bandeau — rattachement par coordonnées des lignes "Remise"
    multiples) — les autres parsers, qui n'en ont pas besoin, continuent
    de recevoir juste `texte`, sans qu'aucun d'eux n'ait à déclarer un
    paramètre inutile."""

    parser = PARSERS_FACTURE.get(fournisseur.upper())

    if parser is None:
        print(f"!! Pas encore de parser facture pour {fournisseur}.")
        return None

    if chemin is not None and "chemin" in inspect.signature(parser).parameters:
        return parser(texte, chemin=chemin)

    return parser(texte)


def parser_facture_ocr(fournisseur: str, mots_par_page: list):

    parser = PARSERS_FACTURE_OCR.get(fournisseur.upper())

    if parser is None:
        print(f"!! Pas encore de parser facture (scan/OCR) pour {fournisseur}.")
        return None

    return parser(mots_par_page)
