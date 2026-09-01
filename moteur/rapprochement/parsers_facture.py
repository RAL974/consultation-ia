"""
Registre des parsers facture fournisseurs — même principe que
moteur/rapprochement/parsers_bl.py (BL) et moteur/parsers.py (devis).

Pour ajouter le support facture d'un fournisseur, ajouter à
moteur/fournisseurs/<fournisseur>.py :

    def parse_facture(texte: str):    # texte PDF natif — voir
        ...                            # moteur/rapprochement/lecture_facture.py
        return Facture(...) | None

`FOURNISSEURS` (déjà déclaré pour les devis/BL) sert aussi de clé ici. Un
fournisseur sans parse_facture reste utilisable pour les devis/BL normalement.
"""

import importlib
import pkgutil

import moteur.fournisseurs as paquet


def _charger_parsers_facture() -> dict:
    parsers = {}

    for info in pkgutil.iter_modules(paquet.__path__):

        if info.name.startswith("_"):
            continue

        module = importlib.import_module(f"moteur.fournisseurs.{info.name}")

        fn = getattr(module, "parse_facture", None)
        noms = getattr(module, "FOURNISSEURS", None)

        if fn is None or not noms:
            continue

        for nom in noms:
            parsers[nom.upper()] = fn

    return parsers


PARSERS_FACTURE = _charger_parsers_facture()


def parser_facture(fournisseur: str, texte: str):

    parser = PARSERS_FACTURE.get(fournisseur.upper())

    if parser is None:
        print(f"!! Pas encore de parser facture pour {fournisseur}.")
        return None

    return parser(texte)
