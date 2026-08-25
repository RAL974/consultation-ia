"""Modèle de données pour un bon de livraison lu (voir moteur/rapprochement/lecture_bl.py)."""

from dataclasses import dataclass, field


@dataclass
class LigneBL:

    reference_fournisseur: str

    designation: str

    quantite_livree: float

    prix_net: float | None = None

    montant: float | None = None


@dataclass
class BonLivraison:

    fournisseur: str

    fichier: str

    numero_bl: str

    date_bl: str

    numero_commande: str

    lignes: list[LigneBL] = field(default_factory=list)

    total_ht_affiche: float | None = None

    # Indices de page (0-based) occupés par ce BL dans le fichier source —
    # None pour les fournisseurs à un seul BL par fichier (la notion ne
    # s'applique pas). Renseigné pour Cominter Ouest (voir
    # moteur.fournisseurs.cominter, moteur.ocr.pages_par_identifiant) :
    # permet d'archiver/déplacer chaque BL d'un fichier qui en contient
    # plusieurs INDIVIDUELLEMENT (voir moteur.rapprochement.pipeline_bl,
    # archivage par BL), sans bloquer les autres BL du même fichier.
    pages: list[int] | None = None

    # "BL" (livraison normale) ou "RETOUR" (le fournisseur annule une
    # ligne d'un BL précédent — cas réel signalé par l'acheteur : un
    # article listé sur un BL mais dont la case "livré" n'était PAS cochée
    # à réception ; le fournisseur envoie un "Retour n° X du date" qui
    # référence le BL concerné). JAMAIS une livraison en soi — voir
    # moteur.rapprochement.pipeline_bl, "bons de retour" : rien n'est
    # jamais écrit à partir d'un RETOUR, seule la ligne qu'il annule est
    # exclue de l'écriture du BL d'origine (numero_bl_origine).
    type_document: str = "BL"
    numero_bl_origine: str = ""
