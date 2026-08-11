from moteur.modele import Article
from moteur.outils import to_float as _to_float, lignes_propres, chercher_devis
from moteur.fournisseurs._gabarit import scan_ancre, disponibilite_apres, diviser_qte_unite

# --- GABARIT (Ravate Elec / Ravate Pro) ---------------------------------
# Structure ancrée sur "Réf. FNR :" : Ravate facture toujours par la Réf.
# FNR, jamais par le code article interne (règle métier — voir CLAUDE.md).
MARQUEUR = ("RÉF. FNR :", "REF. FNR :")
MOTIF_DEVIS = r"\d{8}DV\d{4}"
OFFSETS = {
    "reference_fournisseur": 1,
    "reference_distributeur": 2,
    "designation": 3,
    "qte_unite": 4,
    "prix_brut": 5,
    "prix_net": 6,
    "montant": 7,
}
OFFSET_DISPONIBILITE = 8  # zone libre scannée jusqu'au marqueur suivant
# --- fin GABARIT ---------------------------------------------------------


def parse_ravate_core(texte: str, fournisseur: str) -> list[Article]:
    """Cœur commun aux devis Ravate (Elec et Pro)."""

    articles = []

    devis = chercher_devis(texte, MOTIF_DEVIS)

    lignes = lignes_propres(texte)

    for bloc in scan_ancre(lignes, MARQUEUR, OFFSETS):

        try:
            if any(bloc[champ] is None for champ in OFFSETS):
                raise IndexError("bloc incomplet")

            quantite, unite = diviser_qte_unite(bloc["qte_unite"])

            articles.append(
                Article(
                    fournisseur=fournisseur,
                    devis=devis,
                    reference_fournisseur=bloc["reference_fournisseur"],
                    reference_distributeur=bloc["reference_distributeur"],
                    designation=bloc["designation"],
                    quantite=quantite,
                    unite=unite,
                    prix_brut=_to_float(bloc["prix_brut"]),
                    prix_net=_to_float(bloc["prix_net"]),
                    montant=_to_float(bloc["montant"]),
                    disponibilite=disponibilite_apres(
                        lignes, bloc["_i"] + OFFSET_DISPONIBILITE, MARQUEUR
                    ),
                )
            )
        except Exception as e:
            print(f"Erreur lecture article {fournisseur} : {e}")

    return articles


def parse_ravate(texte: str) -> list[Article]:
    return parse_ravate_core(texte, "RAVATE")


FOURNISSEURS = ["RAVATE"]
parse = parse_ravate
