from pathlib import Path
from moteur.audit import auditer

print("=" * 60)
print("AUDIT BASE D'EQUIVALENCES")
print("=" * 60)
auditer(Path(__file__).parent)
