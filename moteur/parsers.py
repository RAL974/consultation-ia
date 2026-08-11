"""
Registre des parsers fournisseurs — chargement automatique.

Pour ajouter un fournisseur, il suffit de déposer un fichier
moteur/fournisseurs/nouveau.py qui déclare :

    def parse(texte):        # ou parse_nouveau(texte) affecté à `parse`
        ...
        return [Article, ...]

    FOURNISSEURS = ["NOM_DETECTE"]   # nom(s) renvoyé(s) par le détecteur
    parse = parse_nouveau

Aucune modification de ce fichier n'est nécessaire : les modules du
dossier fournisseurs/ sont découverts et enregistrés automatiquement.
"""

import importlib
import pkgutil

import moteur.fournisseurs as paquet


def _charger_parsers() -> dict:
    parsers = {}

    for info in pkgutil.iter_modules(paquet.__path__):

        if info.name.startswith("_"):
            continue

        module = importlib.import_module(f"moteur.fournisseurs.{info.name}")

        fn = getattr(module, "parse", None)
        noms = getattr(module, "FOURNISSEURS", None)

        if fn is None or not noms:
            continue

        for nom in noms:
            parsers[nom.upper()] = fn

    return parsers


PARSERS = _charger_parsers()


def parser_pdf(fournisseur: str, texte: str):

    parser = PARSERS.get(fournisseur.upper())

    if parser is None:
        print(f"!! Pas encore de parser pour {fournisseur}.")
        return []

    return parser(texte)
