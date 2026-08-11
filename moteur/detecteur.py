"""
Détection du fournisseur à partir du texte du PDF.

Règles importantes :
- "COMINTER MAYOTTE" doit être testé AVANT "COMINTER". Un vrai devis Cominter
  Mayotte ne contient PAS la chaîne "COMINTER MAYOTTE" (l'entité s'appelle
  juste "Cominter", basée à Mamoudzou) : on la reconnaît par son e-mail
  (@cominter.yt). "Mamoudzou" seul est TROP LARGE : EDOI est aussi basé à
  Mamoudzou (contact.edoi@sonepar.fr) — utiliser seulement le domaine
  cominter.yt, propre à Cominter.
- "DEM" est testé en mot entier (\bDEM\b) pour éviter les faux positifs
  ("DEMANDE", "DEMI", "ACADEMIE"...).
- "CLAREO" est testé via l'identité de l'entreprise (Clareo S.A.S /
  clareolighting.com), PAS le mot seul : "Clareo" est aussi le nom d'une
  marque de luminaires que revendent d'autres fournisseurs (ex. Coredime) —
  un simple `re.search(r"CLAREO", ...)` prenait alors le mauvais parser sur
  un vrai devis Coredime (voir tests/fixtures).
"""

import re


# Ordre = priorité. Les motifs sont des expressions régulières.
_FOURNISSEURS = [
    ("LEGRAND", r"LEGRAND\s+INTERNAL"),
    ("RAVATE PRO", r"RAVATEPRO|RAVPRO@|RAVATE\s+PRO"),
    ("RAVATE", r"RAVATELEC|RAVATE"),
    ("IMPORELEC", r"IMPORELEC"),
    ("IMPORTER", r"IMPORTER\.RE|\bIMPORTER\b"),
    ("STAND 64", r"STAND\s?64|STAND64"),
    ("CLAREO", r"CLAREO\s+S\.A\.S|CLAREOLIGHTING\.COM"),
    ("COREDIME", r"COREDIME"),
    ("COMINTER MAYOTTE", r"COMINTER\s+MAYOTTE|COMINTER\.YT"),
    ("COMINTER", r"COMINTER"),
    ("ELECTRIC PLUS", r"ELECTRIC\s+PLUS"),
    ("GMR", r"\bGMR\b"),
    # "109 Holding" : gabarit plus récent du même fournisseur (109 Est/
    # Sud/Nord/Ouest) qui ne contient JAMAIS le mot "DISTRIBUTION" dans
    # le corps du devis — seulement dans le pied de page légal
    # ("Siège social : 109 Holding..."). Cas réel, chantier Kanopée CDC.
    ("109 DISTRIBUTION", r"109\s+DISTRIBUTION|109\s+HOLDING"),
    ("EDOI", r"\bEDOI\b"),
    ("STAND 64", r"STAND\s*64"),
    ("IMPORELEC", r"IMPORELEC"),
    ("IMPORTER", r"\bIMPORTER\b"),
    ("COMPTOIR DU CABLING", r"COMPTOIR\s+DU\s+CABLING"),
    ("TELENCO", r"TELENCO"),
    ("SAGEES", r"SAGEES"),
    ("DEM", r"\bDEM\b"),
    # Le PDF DEM ne contient parfois aucun texte identifiant (logo en image).
    # Sa signature : les prix exprimés au cent -> "8058.85/C"
    ("DEM", r"\d+\.\d{2}/C\s"),
]


def detecter_fournisseur(texte: str) -> str:

    texte = texte.upper()

    for nom, motif in _FOURNISSEURS:

        if re.search(motif, texte):
            return nom

    return "INCONNU"
