from dataclasses import dataclass


@dataclass
class Article:

    fournisseur: str

    devis: str

    reference_fournisseur: str

    reference_distributeur: str

    designation: str

    quantite: float

    unite: str

    prix_brut: float

    prix_net: float

    montant: float

    disponibilite: str = ""


@dataclass
class LigneBesoin:
    """Une ligne du besoin chantier."""

    reference: str

    quantite: float

    designation: str = ""
