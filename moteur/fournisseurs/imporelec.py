"""
Parser Imporelec (imporElec, Pacé - métropole).

Bloc article :
    N° (entier)
    Référence
    [Réf Europe]           (facultatif : absent si ligne suivante = désignation)
    Désignation (1-2 lignes)
    [Oui/Non]              (dispo, facultatif)
    Qté (entier)
    P.U.H.T (x €)
    Montant HT (y €)
Puis éventuellement "Composé de :" / "=> ..." (article composé, ignoré ici).

Code écrit AVANT que la règle d'or (jamais de gabarit sans PDF réel) ne
soit formalisée. Confronté cette session à un vrai PDF (195 lignes,
tests/fixtures/imporelec.pdf) : la structure ci-dessus est globalement
CONFIRMÉE (somme des montants = 44 675,71€, quasi exactement le "Sous
Total" affiché avant remise globale du devis, 44 675,70€).

BUG CONNU, NON CORRIGÉ (voir CLAUDE.md) : ce PDF affiche périodiquement un
sous-total cumulé au milieu du tableau (ex. "22" puis "406,94 €" après le
22e article). Le parser ne le reconnaît pas comme tel : ce sous-total est
alors pris à tort pour la Référence de l'article SUIVANT, et le numéro de
compteur + les vraies références partent dans la Désignation. 6 lignes sur
195 sont touchées ainsi (visibles dans tests/test_parsers_imporelec.py) ;
les montants (donc le comparatif prix) restent corrects, seule la
Référence de ces 6 lignes est inexploitable pour le rapprochement.
"""

import re

from moteur.modele import Article
from moteur.outils import to_float, lignes_propres, chercher_devis

_EURO = re.compile(r"^[\d\s.,]+€$")
_REF = re.compile(r"^[A-Z0-9][A-Z0-9\-]{2,}(\s\*+)?$")


def parse_imporelec(texte: str) -> list[Article]:

    articles = []
    devis = chercher_devis(texte, r"\bDEVIS N°\b")  # placeholder
    m = re.search(r"REFERENCE\s+PAGE.*?\n(\d+)\n", texte, re.S)
    m = re.search(r"\n(\d{5,7})\n\d{2}/\d{2}/\d{4}\n", texte)
    devis = m.group(1) if m else ""

    lignes = lignes_propres(texte)

    # Localiser le début du tableau
    debut = 0
    for k, l in enumerate(lignes):
        if l == "Montant HT":
            debut = k + 1
            break

    i = debut
    n = len(lignes)

    while i < n:
        l = lignes[i]

        if l.startswith("Total HT") or l.startswith("***") or l.startswith("Vous pouvez"):
            break

        # Ignorer les lignes de composition
        if l.startswith("Composé") or l.startswith("=>"):
            i += 1
            continue

        # Un bloc commence par le numéro de ligne (entier seul)
        if not l.isdigit():
            i += 1
            continue

        # num = l ; ref = i+1
        ref = lignes[i + 1] if i + 1 < n else ""

        # Rassembler jusqu'aux deux montants € consécutifs
        j = i + 2
        buffer = []
        while j < n and not (_EURO.match(lignes[j]) and j + 1 < n and _EURO.match(lignes[j + 1])):
            buffer.append(lignes[j])
            j += 1
            if j - i > 12:
                break

        if j + 1 >= n or not _EURO.match(lignes[j]):
            i += 1
            continue

        pu = to_float(lignes[j].replace("€", ""))
        montant = to_float(lignes[j + 1].replace("€", ""))

        # buffer = [ref_europe?] + désignation + [dispo] + qté
        qte = to_float(buffer[-1]) if buffer and buffer[-1].isdigit() else 1.0
        corps = buffer[:-1] if buffer and buffer[-1].isdigit() else buffer

        # dispo éventuelle en fin de corps
        if corps and corps[-1] in ("Oui", "Non"):
            corps = corps[:-1]

        ref_europe = ""
        if corps and _REF.match(corps[0]) and len(corps) > 1:
            ref_europe = corps[0]
            corps = corps[1:]

        designation = " ".join(corps)

        articles.append(
            Article(
                fournisseur="IMPORELEC",
                devis=devis,
                reference_fournisseur=ref.replace("*", "").strip(),
                reference_distributeur=ref_europe,
                designation=designation,
                quantite=qte,
                unite="UN",
                prix_brut=pu,
                prix_net=pu,
                montant=montant,
            )
        )

        i = j + 2

    return articles


FOURNISSEURS = ["IMPORELEC"]
parse = parse_imporelec
