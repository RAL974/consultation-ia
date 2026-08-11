"""
Journal d'exécution : un fichier par génération de comparatif/panier, écrit
dans le dossier resultats/ concerné (Point 3 de la session v1.0 — c'est ce
que l'utilisateur envoie en cas de problème).

Toute la sortie du moteur passe déjà par print() (voir moteur/pipeline.py) ;
ce module se contente de dédoubler sys.stdout le temps d'une exécution, vers
sa destination habituelle (terminal, ou file d'attente du GUI) ET vers un
fichier journal_<prefixe>_<horodatage>.log — sans rien changer au code qui
imprime déjà (lecture des PDF, autocontrôle, alias appris/proposés...).
"""

import sys
from contextlib import contextmanager
from datetime import datetime
from io import StringIO
from pathlib import Path


class _Tee:
    """Fichier factice qui recopie chaque écriture vers plusieurs flux."""

    def __init__(self, *flux):
        self._flux = flux

    def write(self, texte):
        for f in self._flux:
            f.write(texte)

    def flush(self):
        for f in self._flux:
            f.flush()


@contextmanager
def capturer_journal(dossier_resultats, prefixe="execution"):
    """
    Dédouble sys.stdout vers un buffer pendant le bloc `with`, puis écrit ce
    buffer dans <dossier_resultats>/journal_<prefixe>_<AAAAMMJJ_HHMMSS>.log
    à la sortie (même en cas d'exception, pour garder une trace d'un
    plantage inattendu). N'interrompt jamais l'exécution : une erreur
    d'écriture du journal est seulement signalée, jamais levée.
    """
    dossier_resultats = Path(dossier_resultats)
    tampon = StringIO()
    stdout_origine = sys.stdout
    sys.stdout = _Tee(stdout_origine, tampon)

    try:
        yield
    finally:
        sys.stdout = stdout_origine
        horodatage = datetime.now().strftime("%Y%m%d_%H%M%S")
        try:
            dossier_resultats.mkdir(parents=True, exist_ok=True)
            fichier = dossier_resultats / f"journal_{prefixe}_{horodatage}.log"
            fichier.write_text(tampon.getvalue(), encoding="utf-8")
        except OSError as e:
            print(f"/!\\ Journal non enregistré ({e}).")
