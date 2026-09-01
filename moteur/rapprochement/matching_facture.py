"""
Rapprochement d'une facture déjà lue (voir lecture_facture.py) avec les
lignes de la commande correspondante dans le Suivi commandes — étape 2 du
flux factures (session F2, voir CLAUDE.md "Rapprochement factures").

Réutilise le moteur de correspondance de référence déjà éprouvé côté BL
(moteur.rapprochement.matching : `_memes_references`/`_repli_reference_proche`/
`_repli_referentiel`, référentiel articles compris — importés tels quels,
jamais dupliqués) plutôt que d'en écrire un second. Fonctionne par duck
typing : `LigneFacture` porte les mêmes champs `reference_fournisseur`/
`designation` que `LigneBL`, donc ces fonctions BL s'appliquent sans
adaptation. Seule la SÉMANTIQUE DE COMPARAISON change (voir
`_comparer_facture`) : une facture ne "cumule" jamais une quantité comme un
BL (`qte_livree_cumulee`) — elle est confrontée à ce qui est DÉJÀ enregistré
comme livré (Qté livrée) et facturé (N° facture, si déjà renseigné).

Colonnes facture (N° facture / Date facture / Qté facturée / PU facturé,
voir CLAUDE.md "Volet 1") : PAS ENCORE créées dans le vrai Suivi commandes
à ce jour (proposition validée dans son principe, jamais appliquée dans
Excel). `lire_lignes_commande_facture()` les lit si présentes, sinon les 4
champs restent `None` sur chaque `LigneSuiviFacture` — le rapprochement en
LECTURE SEULE reste alors utilisable (diagnostic de résorption sur les 3
colonnes déjà existantes : Qté livrée / Tarif BL / Tarif convenu, voir
`colonnes_facture_disponibles()`) ; seule l'ÉCRITURE réelle exige ces
colonnes, et `moteur.rapprochement.ecriture.appliquer()` refuse déjà
proprement si elles sont absentes des en-têtes du fichier réel
(`ColonneNonModifiable`, voir son bandeau) — rien à dupliquer ici non plus.
"""

from dataclasses import dataclass, field
from enum import Enum

from openpyxl import load_workbook

from moteur.outils import to_float
from moteur.panier import MAPPING_FOURNISSEURS
from moteur.rapprochement.matching import (
    _cle,
    _memes_references,
    _repli_reference_proche,
    _repli_referentiel,
)
from moteur.rapprochement.modele_facture import LigneFacture

FEUILLE_COMMANDES = "Commandes"

# Colonnes qui existent déjà dans le Suivi de tout poste (indépendamment des
# colonnes facture, voir plus bas) — sans elles, aucun rapprochement facture
# n'est possible du tout (export du Suivi différent de celui attendu).
_COLONNES_REQUISES_FACTURE = (
    "Référence", "Désignation", "Qté commandée", "N° de commande", "Fournisseur",
    "Qté livrée", "Tarif BL", "Tarif convenu",
)

# Colonnes facture proprement dites (voir CLAUDE.md, Volet 1) — OPTIONNELLES
# à la LECTURE (leur absence ne bloque pas le diagnostic en lecture seule,
# voir bandeau du module), mais dans moteur.rapprochement.ecriture.
# COLONNES_MODIFIABLES pour l'écriture réelle (qui, elle, échoue proprement
# si elles manquent).
COLONNES_FACTURE_OPTIONNELLES = ("N° facture", "Date facture", "Qté facturée", "PU facturé")


@dataclass
class LigneSuiviFacture:

    ligne_excel: int  # numéro de ligne réel dans la feuille (1-based, pour l'écriture)

    reference: str
    designation: str
    qte_commandee: float
    qte_livree: float
    tarif_bl: float | None
    tarif_convenu: float | None
    # Renseigné uniquement par lire_lignes_fournisseur_facture() (déduction
    # de commande par contenu, même principe que côté BL) — vide sinon.
    numero_commande: str = ""

    # None si les colonnes facture n'existent pas encore dans CE Suivi (voir
    # bandeau du module) — jamais une valeur devinée à leur place.
    numero_facture: str | None = None
    date_facture: object = None
    qte_facturee: float | None = None
    pu_facture: float | None = None


class StatutFacture(Enum):
    SUR = "sûr"                  # à écrire après OK global
    A_CONFIRMER = "à confirmer"    # écart, question par question
    DEJA_A_JOUR = "déjà à jour"    # idempotence : rien à écrire
    INCONNU = "inconnu"          # aucune ligne Suivi correspondante sûre


@dataclass
class CorrespondanceFacture:

    ligne_facture: LigneFacture
    ligne_suivi: LigneSuiviFacture | None
    statut: StatutFacture
    raisons: list = field(default_factory=list)


def colonnes_facture_disponibles(entetes: dict) -> bool:
    return all(c in entetes for c in COLONNES_FACTURE_OPTIONNELLES)


def _ouvrir_feuille_commandes_facture(chemin_suivi):
    wb = load_workbook(chemin_suivi, read_only=True, data_only=True)
    ws = wb[FEUILLE_COMMANDES]
    entetes = {
        c.value: i
        for i, c in enumerate(next(ws.iter_rows(min_row=1, max_row=1)))
    }

    manquantes = [c for c in _COLONNES_REQUISES_FACTURE if c not in entetes]
    if manquantes:
        wb.close()
        raise KeyError(
            f"Colonne(s) introuvable(s) dans « {FEUILLE_COMMANDES} » : "
            f"{', '.join(manquantes)} — export du Suivi commandes différent de celui attendu ?"
        )

    return wb, ws, entetes


def _valeur_texte(v):
    v = str(v).strip() if v is not None else ""
    return v or None


def _ligne_suivi_facture_depuis_row(i, row, entetes, avec_commande=False) -> LigneSuiviFacture:

    facture_ok = colonnes_facture_disponibles(entetes)

    return LigneSuiviFacture(
        ligne_excel=i,
        reference=row[entetes["Référence"]],
        designation=row[entetes["Désignation"]],
        qte_commandee=to_float(row[entetes["Qté commandée"]]),
        qte_livree=to_float(row[entetes["Qté livrée"]]),
        tarif_bl=row[entetes["Tarif BL"]],
        tarif_convenu=row[entetes["Tarif convenu"]],
        numero_commande=str(row[entetes["N° de commande"]] or "").strip() if avec_commande else "",
        numero_facture=_valeur_texte(row[entetes["N° facture"]]) if facture_ok else None,
        date_facture=row[entetes["Date facture"]] if facture_ok else None,
        qte_facturee=(
            to_float(row[entetes["Qté facturée"]]) if facture_ok and row[entetes["Qté facturée"]] is not None else None
        ),
        pu_facture=(
            to_float(row[entetes["PU facturé"]]) if facture_ok and row[entetes["PU facturé"]] is not None else None
        ),
    )


def lire_lignes_commande_facture(chemin_suivi, fournisseur: str, numero_commande: str) -> list[LigneSuiviFacture]:
    """Même principe que moteur.rapprochement.matching.lire_lignes_commande
    (conversion MAPPING_FOURNISSEURS comprise), pour les colonnes facture."""

    fournisseur = MAPPING_FOURNISSEURS.get(fournisseur.upper(), fournisseur)

    wb, ws, entetes = _ouvrir_feuille_commandes_facture(chemin_suivi)
    try:
        i_cde = entetes["N° de commande"]
        i_fourn = entetes["Fournisseur"]

        resultat = []
        for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):

            cde = row[i_cde]
            fourn = row[i_fourn]

            if cde is None or str(cde).strip() != numero_commande:
                continue
            if fourn is None or _cle(fourn) != _cle(fournisseur):
                continue

            resultat.append(_ligne_suivi_facture_depuis_row(i, row, entetes))

        return resultat
    finally:
        wb.close()


def lire_lignes_fournisseur_facture(chemin_suivi, fournisseur: str) -> list[LigneSuiviFacture]:
    """Même principe que moteur.rapprochement.matching.lire_lignes_fournisseur
    — utilisé pour la déduction de commande par contenu quand l'en-tête de
    la facture (N°Réf.Client) ne suffit pas (voir pipeline_facture.py)."""

    fournisseur = MAPPING_FOURNISSEURS.get(fournisseur.upper(), fournisseur)

    wb, ws, entetes = _ouvrir_feuille_commandes_facture(chemin_suivi)
    try:
        i_fourn = entetes["Fournisseur"]

        resultat = []
        for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):

            fourn = row[i_fourn]
            if fourn is None or _cle(fourn) != _cle(fournisseur):
                continue

            resultat.append(_ligne_suivi_facture_depuis_row(i, row, entetes, avec_commande=True))

        return resultat
    finally:
        wb.close()


def _comparer_facture(ligne_facture: LigneFacture, ligne_suivi: LigneSuiviFacture, numero_facture: str) -> CorrespondanceFacture:
    """Une facture n'est jamais "cumulée" comme un BL : elle est confrontée
    à ce qui est DÉJÀ enregistré (Qté livrée, Tarif BL/Tarif convenu, et le
    N° de facture déjà présent le cas échéant — idempotence)."""

    raisons = []

    if ligne_suivi.qte_livree <= 0:
        raisons.append(
            "Aucune quantité livrée enregistrée pour cette ligne — facture arrivée avant "
            "son BL, ou BL pas encore rapproché"
        )
        return CorrespondanceFacture(ligne_facture, ligne_suivi, StatutFacture.A_CONFIRMER, raisons)

    if ligne_suivi.numero_facture:
        if ligne_suivi.numero_facture == numero_facture:
            # Idempotence : ce même n° de facture est déjà enregistré sur
            # cette ligne — un doublon de dépôt/traitement, rien à réécrire
            # (même principe que côté BL, voir moteur.rapprochement.matching._comparer).
            return CorrespondanceFacture(ligne_facture, ligne_suivi, StatutFacture.DEJA_A_JOUR)
        raisons.append(
            f"Cette ligne porte déjà un autre n° de facture ({ligne_suivi.numero_facture}) "
            "— doublon de dépôt ou litige de facturation à vérifier avant d'écraser"
        )
        return CorrespondanceFacture(ligne_facture, ligne_suivi, StatutFacture.A_CONFIRMER, raisons)

    if abs(ligne_facture.quantite_facturee - ligne_suivi.qte_livree) > 0.001:
        raisons.append(
            f"Qté facturée ({ligne_facture.quantite_facturee:g}) différente de la Qté livrée "
            f"déjà enregistrée ({ligne_suivi.qte_livree:g}) — facturation partielle/multiple ?"
        )

    # Aucune tolérance sur l'écart de prix — décision explicite de l'acheteur
    # (voir CLAUDE.md, Volet 3 : "il y a énormément d'articles à très faible
    # valeur, pas de tolérance, prix BL = prix facture"). Tarif BL prioritaire
    # sur Tarif convenu, même repli que "Facturé BL" côté formule Excel.
    tarif_bl = to_float(ligne_suivi.tarif_bl) if ligne_suivi.tarif_bl else 0.0
    tarif_reference = tarif_bl or to_float(ligne_suivi.tarif_convenu)
    pu_facture = to_float(ligne_facture.prix_unitaire_ht)

    if not tarif_reference:
        raisons.append("Aucun tarif de référence (ni Tarif BL, ni Tarif convenu) pour vérifier le prix facturé")
    elif abs(pu_facture - tarif_reference) > 0.001:
        source = "Tarif BL" if tarif_bl else "Tarif convenu"
        raisons.append(
            f"PU facturé ({pu_facture:g}€) différent du tarif de référence "
            f"({tarif_reference:g}€, {source}) — aucune tolérance (décision de l'acheteur)"
        )

    statut = StatutFacture.A_CONFIRMER if raisons else StatutFacture.SUR
    return CorrespondanceFacture(ligne_facture, ligne_suivi, statut, raisons)


def apparier_facture(lignes_facture: list[LigneFacture], lignes_suivi: list[LigneSuiviFacture],
                      numero_facture: str, referentiel=None, fournisseur: str = "",
                      devis: str = "") -> list[CorrespondanceFacture]:
    """Même algorithme deux passes que moteur.rapprochement.matching.apparier()
    (voir son bandeau pour le bug réel qu'il corrige — une correspondance
    EXACTE ne doit jamais être volée par un repli approximatif traité avant
    elle dans l'ordre du document) : 1re passe, uniquement les
    correspondances EXACTES (candidat unique) sur toutes les lignes de la
    facture ; 2e passe, ce qui reste (repli référence proche / référentiel /
    ambiguïté), sur les lignes Suivi déjà réduites."""

    disponibles = list(lignes_suivi)
    resultat_par_index = {}

    for i, lf in enumerate(lignes_facture):
        candidats = [
            ls for ls in disponibles
            if _memes_references(ls.reference, lf.reference_fournisseur, referentiel,
                                  ls.designation, lf.designation, fournisseur, devis)
        ]
        if len(candidats) == 1:
            ls = candidats[0]
            disponibles.remove(ls)
            resultat_par_index[i] = _comparer_facture(lf, ls, numero_facture)

    for i, lf in enumerate(lignes_facture):

        if i in resultat_par_index:
            continue

        candidats = [
            ls for ls in disponibles
            if _memes_references(ls.reference, lf.reference_fournisseur, referentiel,
                                  ls.designation, lf.designation, fournisseur, devis)
        ]

        if not candidats:
            ls_repli, raison_repli = _repli_reference_proche(lf, disponibles)
            if ls_repli is None:
                ls_repli, raison_repli = _repli_referentiel(lf, disponibles, referentiel, fournisseur, devis)
            if ls_repli is not None:
                disponibles.remove(ls_repli)
                c = _comparer_facture(lf, ls_repli, numero_facture)
                if c.statut is StatutFacture.DEJA_A_JOUR:
                    # Même garde-fou que côté BL (voir matching.apparier) :
                    # un repli approximatif vers une ligne déjà à jour ne
                    # doit jamais être "recomparé" comme un nouvel écart —
                    # il ressort déjà_a_jour, raison du repli conservée.
                    resultat_par_index[i] = CorrespondanceFacture(
                        c.ligne_facture, c.ligne_suivi, StatutFacture.DEJA_A_JOUR, [raison_repli],
                    )
                else:
                    resultat_par_index[i] = CorrespondanceFacture(
                        c.ligne_facture, c.ligne_suivi, StatutFacture.A_CONFIRMER, [raison_repli] + c.raisons,
                    )
            else:
                resultat_par_index[i] = CorrespondanceFacture(
                    lf, None, StatutFacture.INCONNU,
                    ["Aucune ligne du Suivi ne correspond à cette référence pour cette commande"],
                )
        else:
            exacts = [
                ls for ls in candidats
                if str(ls.reference).strip().upper() == lf.reference_fournisseur.strip().upper()
            ]
            if len(exacts) == 1:
                ls = exacts[0]
                disponibles.remove(ls)
                resultat_par_index[i] = _comparer_facture(lf, ls, numero_facture)
            else:
                resultat_par_index[i] = CorrespondanceFacture(
                    lf, None, StatutFacture.INCONNU,
                    [f"{len(candidats)} lignes du Suivi correspondent à cette référence (ambigu)"],
                )

    return [resultat_par_index[i] for i in range(len(lignes_facture))]
