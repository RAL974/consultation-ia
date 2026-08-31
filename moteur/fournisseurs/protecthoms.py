"""
Parser BL PROTECTHOMS (nouveau fournisseur, session Rapprochement AI) —
équipements de protection individuelle / matériel amiante, PAS du matériel
électrique comme les autres fournisseurs de ce projet, mais bien présent
dans la liste Fournisseurs du Suivi (confirmé) : en périmètre. Aucun devis
connu pour ce fournisseur — module dédié BL uniquement. Un seul vrai BL vu
à ce jour (`BL M3.15.399.pdf`, commande M3.15.399).

Structure simple et claire : tableau "Reference produit | Designation |
Quantites | Reste à livrer" — chaque ligne visuelle est déjà UN article
complet (référence, désignation, quantité, sur 3 ou 4 cellules selon
qu'une colonne "Reste à livrer" est renseignée ou non). Pas de prix du
tout sur ce document (comme Coredime) — pas d'autocontrôle Total HT
possible.

Référence produit toujours au format 1 chiffre + 2 lettres + 6 chiffres
(ex. "2VU043003", "6SP040000", "1MA010701", vérifié sur les 8 lignes du
seul document vu) — utilisée comme ancre de ligne plutôt qu'une position
de cellule fixe ou un en-tête/pied de tableau, plus robuste ici.

"Quantites" est la quantité LIVRÉE sur CE bon (comme partout ailleurs dans
ce projet) ; "Reste à livrer" (4e cellule, seulement quand renseignée) est
purement informatif — jamais soustrait ni ajouté, cohérent avec le fait
que ce sont deux colonnes séparées et non un statut à exclure (à la
différence du "Reste à livrer" Coredime/DEM, qui EXCLUT la ligne
entière — ici la ligne EST livrée, à hauteur de "Quantites")."""

import re

from moteur.ocr import regrouper_lignes
from moteur.rapprochement.modele_bl import BonLivraison, LigneBL

# --- GABARIT BL (PROTECTHOMS) ------------------------------------------------
MOTIF_REFERENCE_BL_PROTECTHOMS = re.compile(r"^\d[A-Z]{2}\d{6}$")
MOTIF_NOMBRE_BL_PROTECTHOMS = re.compile(r"^\d+$")
MOTIF_COMMANDE_BL_PROTECTHOMS = re.compile(r"Numero\s*de\s*commande\s*:?\s*(\S+)", re.IGNORECASE)
MOTIF_NUMERO_BL_PROTECTHOMS = re.compile(r"N.\s*:?\s*(BL\d+)", re.IGNORECASE)
MOTIF_DATE_BL_PROTECTHOMS = re.compile(r"Du\s*:?\s*(\d{2}/\d{2}/\d{4})", re.IGNORECASE)
# --- fin GABARIT BL -----------------------------------------------------------


def parse_bl_protecthoms(mots_par_page: list[list[dict]]) -> BonLivraison:

    lignes_plates = [
        " ".join(m["texte"] for m in ligne)
        for mots in mots_par_page
        for ligne in regrouper_lignes(mots)
    ]
    texte = "\n".join(lignes_plates)

    numero_commande = ""
    m = MOTIF_COMMANDE_BL_PROTECTHOMS.search(texte)
    if m:
        numero_commande = m.group(1).upper()

    numero_bl = ""
    m = MOTIF_NUMERO_BL_PROTECTHOMS.search(texte)
    if m:
        numero_bl = m.group(1).upper()

    date_bl = ""
    m = MOTIF_DATE_BL_PROTECTHOMS.search(texte)
    if m:
        date_bl = m.group(1)

    articles = []
    for mots in mots_par_page:
        for ligne in regrouper_lignes(mots):
            cellules = [m["texte"].strip() for m in ligne]
            if not cellules or not MOTIF_REFERENCE_BL_PROTECTHOMS.match(cellules[0]):
                continue

            i_qte = next(
                (j for j, c in enumerate(cellules) if j >= 1 and MOTIF_NOMBRE_BL_PROTECTHOMS.match(c)),
                None,
            )
            if i_qte is None:
                continue

            designation = " ".join(cellules[1:i_qte]).strip()
            quantite = float(cellules[i_qte])

            articles.append(LigneBL(
                reference_fournisseur=cellules[0], designation=designation,
                quantite_livree=quantite, prix_net=None, montant=None,
            ))

    return BonLivraison(
        fournisseur="PROTECTHOMS", fichier="", numero_bl=numero_bl, date_bl=date_bl,
        numero_commande=numero_commande, lignes=articles,
    )


FOURNISSEURS = ["PROTECTHOMS"]
parse_bl = parse_bl_protecthoms
