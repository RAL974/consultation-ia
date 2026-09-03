"""Modèle de données pour une facture fournisseur lue (voir
moteur/rapprochement/lecture_facture.py) — même principe que
moteur/rapprochement/modele_bl.py, pour les factures plutôt que les BL."""

from dataclasses import dataclass, field


@dataclass
class LigneFacture:

    reference_fournisseur: str

    designation: str

    quantite_facturee: float

    prix_unitaire_ht: float | None = None

    montant_ht: float | None = None

    # N° de commande dont dépend cette ligne — vide au moment du parsing
    # (le parser ne consulte jamais le Suivi), renseigné par
    # moteur.rapprochement.matching_facture lors de la résolution (voir son
    # bandeau : en-tête N°Réf.Client en 1er ressort, déduction par contenu
    # en repli, comme pour un BL au n° de commande illisible).
    numero_commande: str = ""

    # N° du BL auquel cette ligne est rattachée sur LA FACTURE (une facture
    # peut regrouper plusieurs BL, voir Facture.numeros_bl) — permet de
    # confronter Qté facturée à Qté livrée bloc de BL par bloc de BL plutôt
    # que sur la commande entière.
    numero_bl: str = ""


@dataclass
class Facture:

    fournisseur: str

    fichier: str

    numero_facture: str

    date_facture: str  # texte brut "JJ/MM/AAAA", à parser par l'appelant (comme BonLivraison.date_bl)

    date_echeance: str = ""

    # Candidat(s) de n° de commande lus sur l'EN-TÊTE de la facture
    # (N°Réf.Client, voir moteur.fournisseurs.dist109) — 0 ou 1 élément la
    # quasi-totalité du temps (un seul champ imprimé), jamais déduit du
    # contenu ici (le parser n'a pas accès au Suivi) : la déduction par
    # bloc de BL, quand l'en-tête ne suffit pas, est le rôle de
    # moteur.rapprochement.matching_facture.
    numeros_commande: list = field(default_factory=list)

    # Candidat(s) BRUTS trouvés sur l'en-tête (N°Réf.Client/Réf.:), QUE le
    # parser ait réussi ou non à les convertir en numeros_commande — sert
    # UNIQUEMENT à distinguer, en aval (moteur.rapprochement.matching_facture.
    # est_bdc_manuel_24x), une commande "introuvable" d'un bon manuel type
    # "BC 241766"/"BCN 241461" (jamais rattachable, pas une vraie anomalie
    # de rapprochement). Vide si le champ n'est pas renseigné sur la pièce.
    numeros_commande_bruts: list = field(default_factory=list)

    numeros_bl: list = field(default_factory=list)

    lignes: list = field(default_factory=list)  # list[LigneFacture]

    total_ht_affiche: float | None = None

    # "FACTURE" (normal) ou "AVOIR" — jamais rapprochée automatiquement
    # (voir CLAUDE.md, Volet 3 : "Avoirs : jamais automatiques, toujours à
    # confirmer") ; repéré par la mention "FACTURE D'AVOIR" en en-tête chez
    # 109 Distribution (indication de l'acheteur, aucun exemple réel vu à
    # ce jour — voir moteur.fournisseurs.dist109).
    type_document: str = "FACTURE"
