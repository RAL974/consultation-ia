"""
Registre des parsers BL fournisseurs — chargement automatique, même principe
que moteur/parsers.py (devis).

Pour ajouter le support BL d'un fournisseur, ajouter à
moteur/fournisseurs/<fournisseur>.py :

    def parse_bl(mots_par_page):     # ou parse_bl_<x>(...) affecté à parse_bl
        ...
        return BonLivraison(...) | None

`FOURNISSEURS` (déjà déclaré pour les devis) sert aussi de clé ici. Un
fournisseur sans parse_bl reste utilisable pour les devis normalement.
"""

import importlib
import pkgutil

import moteur.fournisseurs as paquet


def _charger_parsers_bl() -> dict:
    parsers = {}

    for info in pkgutil.iter_modules(paquet.__path__):

        if info.name.startswith("_"):
            continue

        module = importlib.import_module(f"moteur.fournisseurs.{info.name}")

        fn = getattr(module, "parse_bl", None)
        noms = getattr(module, "FOURNISSEURS", None)

        if fn is None or not noms:
            continue

        for nom in noms:
            parsers[nom.upper()] = fn

    return parsers


PARSERS_BL = _charger_parsers_bl()


def parser_bl(fournisseur: str, mots_par_page: list[list[dict]]):

    parser = PARSERS_BL.get(fournisseur.upper())

    if parser is None:
        print(f"!! Pas encore de parser BL pour {fournisseur}.")
        return None

    return parser(mots_par_page)
