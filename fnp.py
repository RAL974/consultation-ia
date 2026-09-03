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

sans_exclusion = "--sans-exclusion" in sys.argv
arguments = [a for a in sys.argv[1:] if not a.startswith("--")]

if not arguments:
    print("\nUsage : py -3 fnp.py <AAAA-MM> [AAAA-MM-JJ] [--sans-exclusion]")
    print("  <AAAA-MM>          mois de clôture, ex. 2026-08")
    print("  [AAAA-MM-JJ]       optionnel : ne compter que les livraisons à partir de cette date")
    print("  [--sans-exclusion] optionnel : désactive l'exclusion « facture reçue non")
    print("                     rapprochée » (étape 4a) — repli si a_traiter/Factures/")
    print("                     est trop volumineux/lent à scanner")
    raise SystemExit(1)

mois = arguments[0]
depuis = None
if len(arguments) > 1:
    try:
        depuis = datetime.strptime(arguments[1], "%Y-%m-%d").date()
    except ValueError:
        print(f"\n/!\\ Date de filtre invalide : « {arguments[1]} » (attendu AAAA-MM-JJ)")
        raise SystemExit(1)

dossier_projet = Path(__file__).parent

if sans_exclusion:
    print("\n(--sans-exclusion : étape 4a désactivée, le volet (a) inclut les factures déjà reçues mais pas encore rapprochées)")

try:
    chemin = generer_etat_fnp(dossier_projet, mois, depuis, appliquer_exclusion=not sans_exclusion)
except SuiviIntrouvable as e:
    print(f"\n/!\\ {e}")
    raise SystemExit(1)
except ValueError as e:
    print(f"\n/!\\ Mois invalide « {mois}» : {e}")
    raise SystemExit(1)

print(f"\nÉtat généré : {chemin}")
