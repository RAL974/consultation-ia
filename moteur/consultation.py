"""
Résolution du "dossier de consultation" à traiter.

Une consultation = un dossier autonome sous consultations/<nom>/, contenant
le fichier besoin (facultatif), un sous-dossier devis/ avec les PDF, et un
sous-dossier resultats/ (créé au premier traitement). Remplace l'ancien
schéma besoins/ + devis/<chantier>/ + resultats/ partagé par un dossier
unique par consultation, rangé et rejouable — voir CLAUDE.md.
"""

from pathlib import Path

from moteur.besoin import trouver_fichier_besoin


class ConsultationIntrouvable(Exception):
    """Le dossier de consultation à traiter ne peut pas être déterminé sans
    ambiguïté — jamais de choix deviné au hasard."""


def lister_consultations(dossier_projet) -> list[Path]:
    """Sous-dossiers de consultations/, triés par nom."""
    dossier = Path(dossier_projet) / "consultations"
    if not dossier.is_dir():
        return []
    return sorted((p for p in dossier.iterdir() if p.is_dir()), key=lambda p: p.name)


def resoudre_dossier_consultation(dossier_projet, cible=None) -> Path:
    """
    Détermine le dossier de consultation à traiter.

    - cible = chemin (relatif ou absolu) vers un dossier existant -> utilisé
      tel quel.
    - cible = nom d'un sous-dossier de consultations/ -> ce sous-dossier.
    - cible = None -> le seul dossier sous consultations/ s'il n'y en a
      qu'un ; sinon lève ConsultationIntrouvable avec la liste des dossiers
      disponibles (aucun choix deviné au hasard).
    """
    dossier_projet = Path(dossier_projet)

    if cible:
        chemin_direct = Path(cible)
        if chemin_direct.is_dir():
            return chemin_direct

        chemin_sous_consultations = dossier_projet / "consultations" / str(cible)
        if chemin_sous_consultations.is_dir():
            return chemin_sous_consultations

        raise ConsultationIntrouvable(
            f"Dossier de consultation introuvable : « {cible} » "
            "(cherché tel quel, puis sous consultations/)."
        )

    disponibles = lister_consultations(dossier_projet)

    if len(disponibles) == 1:
        return disponibles[0]

    if not disponibles:
        raise ConsultationIntrouvable(
            "Aucun dossier de consultation trouvé sous consultations/.\n"
            "Crée-en un (ex. consultations/MonChantier/), dépose-y le "
            "fichier besoin et un sous-dossier devis/ avec les PDF, puis "
            "relance."
        )

    noms = ", ".join(f"« {d.name} »" for d in disponibles)
    raise ConsultationIntrouvable(
        f"Plusieurs dossiers de consultation existent : {noms}.\n"
        "Précise lequel traiter (nom du dossier en argument)."
    )


def fichier_besoin_de(dossier_consultation) -> Path | None:
    """Fichier besoin du dossier de consultation (voir moteur.besoin), ou
    None si absent (comparaison de toutes les lignes des devis)."""
    return trouver_fichier_besoin(Path(dossier_consultation))


def dossier_devis_de(dossier_consultation) -> Path:
    return Path(dossier_consultation) / "devis"


def dossier_resultats_de(dossier_consultation) -> Path:
    return Path(dossier_consultation) / "resultats"
