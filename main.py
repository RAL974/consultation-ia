import sys
from pathlib import Path

from moteur.dependances import verifier_et_installer

print("=" * 60)
print("CONSULTATION AI")
print("=" * 60)

if not verifier_et_installer(print):
    print("\nDépendances manquantes non résolues. Installe-les manuellement puis relance.")
    raise SystemExit(1)

from moteur.consultation import (
    ConsultationIntrouvable,
    dossier_devis_de,
    dossier_resultats_de,
    fichier_besoin_de,
    resoudre_dossier_consultation,
)
from moteur.pipeline import generer_comparatif

dossier_projet = Path(__file__).parent

# Dossier de consultation à traiter : nom (ou chemin) donné en argument,
# sinon le seul dossier sous consultations/ s'il n'y en a qu'un. Chaque
# consultation regroupe son besoin, ses devis (devis/) et ses résultats
# (resultats/) — voir CLAUDE.md.
cible = sys.argv[1] if len(sys.argv) > 1 else None

try:
    dossier_consultation = resoudre_dossier_consultation(dossier_projet, cible)
except ConsultationIntrouvable as e:
    print(f"\n/!\\ {e}")
    raise SystemExit(1)

print(f"\nConsultation : {dossier_consultation.name}  ({dossier_consultation})")

fichier_besoin = fichier_besoin_de(dossier_consultation)
dossier_devis = dossier_devis_de(dossier_consultation)
dossier_resultats = dossier_resultats_de(dossier_consultation)

if not dossier_devis.is_dir() or not any(dossier_devis.glob("*.pdf")):
    print(f"\n/!\\ Aucun PDF trouvé dans {dossier_devis}")
    print("    Dépose les devis PDF de cette consultation dans ce sous-dossier, puis relance.")
    raise SystemExit(1)

if not fichier_besoin:
    print("Aucun fichier besoin trouvé : comparaison de toutes les lignes des devis.")

generer_comparatif(dossier_projet, fichier_besoin, dossier_devis, dossier_resultats)

print("\nTraitement terminé.")
