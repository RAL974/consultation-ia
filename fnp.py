import sys
from datetime import datetime
from pathlib import Path

from moteur.dependances import verifier_et_installer

print("=" * 60)
print("ÉTAT FNP (FACTURES NON PARVENUES)")
print("=" * 60)

if not verifier_et_installer(print):
    print("\nDépendances manquantes non résolues. Installe-les manuellement puis relance.")
    raise SystemExit(1)

from moteur.fnp import SuiviIntrouvable, generer_etat_fnp

if len(sys.argv) < 2:
    print("\nUsage : py -3 fnp.py <AAAA-MM> [AAAA-MM-JJ]")
    print("  <AAAA-MM>     mois de clôture, ex. 2026-08")
    print("  [AAAA-MM-JJ]  optionnel : ne compter que les livraisons à partir de cette date")
    raise SystemExit(1)

mois = sys.argv[1]
depuis = None
if len(sys.argv) > 2:
    try:
        depuis = datetime.strptime(sys.argv[2], "%Y-%m-%d").date()
    except ValueError:
        print(f"\n/!\\ Date de filtre invalide : « {sys.argv[2]} » (attendu AAAA-MM-JJ)")
        raise SystemExit(1)

dossier_projet = Path(__file__).parent

try:
    chemin = generer_etat_fnp(dossier_projet, mois, depuis)
except SuiviIntrouvable as e:
    print(f"\n/!\\ {e}")
    raise SystemExit(1)
except ValueError as e:
    print(f"\n/!\\ Mois invalide « {mois}» : {e}")
    raise SystemExit(1)

print(f"\nÉtat généré : {chemin}")
