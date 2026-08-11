"""
Comparateur multi-fournisseurs.

Le rapprochement des offres repose sur une clé d'équivalence :
- si un référentiel (moteur/referentiel.py) est fourni ET connaît la
  référence avec certitude (alias exact, déduit du référentiel achats réel
  ou confirmé par l'acheteur), sa clé normalisée fait autorité ;
- sinon, si une base d'articles est fournie, la clé provient de
  base.groupe() (fabricant + cœur numérique, code interne, ou famille
  manuelle) — inchangé, c'est le repli historique ;
- sinon, on retombe sur le cœur numérique simple (cle_matching).

Le référentiel ne fusionne JAMAIS automatiquement une simple proposition
("propose", pas encore confirmée) : tant qu'elle n'est pas confirmée, la
référence retombe sur base.groupe() comme avant.

Avec un besoin, le comparatif suit l'ordre du besoin (après expansion des
éventuels composés — un besoin qui se décompose en plusieurs références,
ex. coffret + porte). Les lignes de devis non rapprochées sont mises de
côté (onglet "Hors besoin").
"""

from moteur.besoin import cle_matching
from moteur.base import coeur_numerique
from moteur.modele import LigneBesoin
from moteur.normalisation import prix_normalise, code_interne_auto


def _cle(ref, base, referentiel=None, designation=""):
    if referentiel is not None:
        cle, statut = referentiel.resoudre(ref, designation=designation)
        if statut == "connu":
            return f"REF:{cle}"
    return base.groupe(ref, designation) if base is not None else cle_matching(ref)


def _expanser_composes(besoin, referentiel):
    """Remplace toute ligne de besoin qui se décompose (referentiel/composes.csv)
    par une ligne par composant (quantité = qté besoin x quantité membre)."""

    if not referentiel or not besoin:
        return besoin

    expanse = []

    for ligne_besoin in besoin:

        membres = referentiel.composants(ligne_besoin.reference)

        if not membres:
            expanse.append(ligne_besoin)
            continue

        for membre, quantite_membre in membres:
            expanse.append(LigneBesoin(
                reference=membre,
                quantite=ligne_besoin.quantite * quantite_membre,
                designation=ligne_besoin.designation,
            ))

    return expanse


# Familles « commodités » : les références n'y sont JAMAIS partagées entre
# fournisseurs (chacun a la sienne). Un cœur numérique commun y est donc un
# faux ami potentiel (ex. L401227 Arcotec vs 401227 Legrand) : on ne
# consolide jamais par cœur quand un tel code est en jeu d'un seul côté.
_COMMODITES = ("AMP_LED", "EMBOUT_")


def _code_auto(cle, designation):
    """Code interne « sémantique » d'un groupe : la clé CI: elle-même,
    sinon celui que produirait sa désignation représentative ('' si aucun)."""
    if cle and cle.startswith("CI:"):
        return cle[3:]
    return code_interne_auto(designation or "")


def _codes_compatibles(code_besoin, code_offre):
    """
    Deux codes internes autorisent-ils la consolidation d'offres qui
    partagent déjà le même cœur numérique que la référence du besoin ?

      - identiques                       -> oui ;
      - une commodité d'un seul côté     -> NON (faux amis, cf. _COMMODITES) ;
      - l'un des deux vide               -> oui (aucune contradiction) ;
      - même famille, mêmes segments     -> oui, « ? » sert de joker
        (INTERDIFF_1P+N_25A_?_30MA ~ INTERDIFF_1P+N_25A_AC_30MA) ;
      - sinon                            -> non (type A ≠ AC, DISJ ≠ GOULOTTE...).
    """
    if code_besoin == code_offre:
        return True
    for code in (code_besoin, code_offre):
        if code and code.startswith(_COMMODITES):
            return False
    if not code_besoin or not code_offre:
        return True
    seg_b = code_besoin.split("_")
    seg_o = code_offre.split("_")
    if seg_b[0] != seg_o[0] or len(seg_b) != len(seg_o):
        return False
    return all(a == b or a == "?" or b == "?" for a, b in zip(seg_b, seg_o))


def comparer(articles, besoin=None, base=None, referentiel=None):

    besoin = _expanser_composes(besoin, referentiel)

    index = {}
    fournisseurs = set()

    for article in articles:

        cle = _cle(article.reference_fournisseur, base, referentiel, article.designation)

        if not cle:
            continue

        fournisseurs.add(article.fournisseur)

        entree = index.setdefault(
            cle,
            {
                "reference": article.reference_fournisseur,
                "designation": article.designation,
                "unite": article.unite,
                "offres": {},
                "coeurs": set(),
            },
        )

        coeur = coeur_numerique(article.reference_fournisseur)
        if coeur:
            entree["coeurs"].add(coeur)

        if len(article.designation) > len(entree["designation"]):
            entree["designation"] = article.designation

        est_cable = base.est_cable(article.reference_fournisseur) if base else False

        prix_norm, unite_base, fiable = prix_normalise(
            article.prix_net, article.unite, article.designation, est_cable
        )

        offre = {
            "prix": article.prix_net,
            "prix_norm": prix_norm,
            "base": unite_base,
            "fiable": fiable,
            "devis": article.devis,
            "disponibilite": article.disponibilite,
            "reference": article.reference_fournisseur,
            "quantite": article.quantite,
            "unite": article.unite,
            "montant": article.montant,
        }

        existante = entree["offres"].get(article.fournisseur)

        # Même fournisseur, même groupe deux fois -> on garde la moins chère
        if existante is None or offre["prix_norm"] < existante["prix_norm"]:
            entree["offres"][article.fournisseur] = offre

    resultat = {
        "fournisseurs": sorted(fournisseurs),
        "lignes": [],
        "hors_besoin": [],
        "totaux_fournisseurs": {},
    }

    if besoin:

        utilisees = set()

        for ligne_besoin in besoin:

            cle = _cle(ligne_besoin.reference, base, referentiel)

            entree = index.get(cle)

            # Consolidation par cœur : la RÉFÉRENCE du besoin fait autorité.
            # Un groupe resté isolé (le plus souvent par un code interne
            # déduit de la désignation d'UN seul fournisseur, ou parce que
            # la réf n'est pas dans la BDD) qui porte le même cœur numérique
            # que la référence demandée est rattaché à la ligne du besoin —
            # sauf contradiction (voir _codes_compatibles).
            coeur_b = coeur_numerique(ligne_besoin.reference)
            if coeur_b:
                code_b = _code_auto(cle, ligne_besoin.designation)
                for autre_cle, autre in index.items():
                    if autre_cle == cle or autre_cle in utilisees:
                        continue
                    if coeur_b not in autre["coeurs"]:
                        continue
                    code_o = _code_auto(autre_cle, autre["designation"])
                    if not _codes_compatibles(code_b, code_o):
                        continue
                    utilisees.add(autre_cle)
                    if entree is None:
                        entree = autre
                        continue
                    # Fusion : on garde l'offre la moins chère par fournisseur
                    for fournisseur, offre in autre["offres"].items():
                        existante = entree["offres"].get(fournisseur)
                        if existante is None or offre["prix_norm"] < existante["prix_norm"]:
                            entree["offres"][fournisseur] = offre
                    if len(autre["designation"]) > len(entree["designation"]):
                        entree["designation"] = autre["designation"]

            if entree:

                utilisees.add(cle)

                resultat["lignes"].append(
                    {
                        "demande": ligne_besoin.designation,
                        "reference": ligne_besoin.reference,
                        "designation": entree["designation"],
                        "unite": entree["unite"],
                        "qte_besoin": ligne_besoin.quantite,
                        "offres": entree["offres"],
                    }
                )

            else:

                resultat["lignes"].append(
                    {
                        "demande": ligne_besoin.designation,
                        "reference": ligne_besoin.reference,
                        "designation": "",
                        "unite": "",
                        "qte_besoin": ligne_besoin.quantite,
                        "offres": {},
                    }
                )

        for cle, entree in index.items():

            if cle in utilisees:
                continue

            for fournisseur, offre in entree["offres"].items():

                resultat["hors_besoin"].append(
                    {
                        "fournisseur": fournisseur,
                        "reference": entree["reference"],
                        "designation": entree["designation"],
                        "unite": offre["unite"],
                        "quantite": offre["quantite"],
                        "prix": offre["prix"],
                        "montant": offre["montant"],
                        "devis": offre["devis"],
                        "disponibilite": offre["disponibilite"],
                    }
                )

        # Lignes sans aucune offre remontées en tête : jamais perdues en
        # silence, l'acheteur les voit en premier sans avoir à chercher
        # dans tout le comparatif. Tri stable : l'ordre du besoin est
        # conservé à l'intérieur de chaque groupe (sans offre / avec offre).
        resultat["lignes"].sort(key=lambda l: 1 if l["offres"] else 0)

        # Total par fournisseur si tout le besoin lui était pris (somme des
        # lignes où il a effectivement une offre ; les lignes où il n'a pas
        # répondu sont ignorées, pas comptées comme un manque à combler).
        totaux = {f: 0.0 for f in resultat["fournisseurs"]}
        for ligne in resultat["lignes"]:
            for fournisseur, offre in ligne["offres"].items():
                qte = ligne["qte_besoin"] or 0
                totaux[fournisseur] += qte * offre["prix_norm"]
        resultat["totaux_fournisseurs"] = totaux

    else:

        cles = sorted(
            index.keys(),
            key=lambda c: (-len(index[c]["offres"]), index[c]["reference"]),
        )

        for cle in cles:

            entree = index[cle]

            resultat["lignes"].append(
                {
                    "demande": "",
                    "reference": entree["reference"],
                    "designation": entree["designation"],
                    "unite": entree["unite"],
                    "qte_besoin": None,
                    "offres": entree["offres"],
                }
            )

    return resultat


def meilleure_offre(offres: dict):
    """Retourne (fournisseur, offre) la moins chère, ou (None, None)."""

    meilleur = None
    gagnant = None

    for fournisseur, offre in offres.items():

        if meilleur is None or offre["prix_norm"] < meilleur["prix_norm"]:
            meilleur = offre
            gagnant = fournisseur

    return gagnant, meilleur
