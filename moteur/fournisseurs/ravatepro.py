"""Ravate Pro (ravatepro / ravpro@ravate.com) — distinct de Ravate Elec.
Même structure Réf. FNR, seul le libellé fournisseur change."""

from moteur.fournisseurs.ravate import parse_ravate_core


def parse_ravate_pro(texte: str):
    return parse_ravate_core(texte, "RAVATE PRO")


FOURNISSEURS = ["RAVATE PRO"]
parse = parse_ravate_pro
