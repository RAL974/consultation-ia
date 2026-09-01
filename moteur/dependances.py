"""
Vérifie que les dépendances Python du projet sont installées et les
installe automatiquement si besoin.

Pourquoi : main.py / gui.py peuvent être lancés par des interpréteurs
Python différents selon le contexte (terminal VS Code avec un environnement
déjà configuré, double-clic sur un .bat qui utilise le "python" par défaut
du PATH Windows, poste d'une collègue...). Chaque interpréteur a SES
PROPRES paquets installés — installer openpyxl dans l'un ne le rend pas
disponible dans l'autre. Plutôt que de deviner quel interpréteur est
utilisé, on installe directement dans celui qui exécute CE script
(sys.executable), ce qui fonctionne quel que soit le contexte de lancement.
"""

import importlib
import subprocess
import sys

# nom du module Python -> nom du paquet pip (pas toujours identique :
# PyMuPDF s'importe "fitz")
REQUIS = {
    "openpyxl": "openpyxl",
    "fitz": "PyMuPDF",
    "rapidocr_onnxruntime": "rapidocr-onnxruntime",
}

# win32com (pywin32) — utilisé par moteur/fnp_brouillon.py (brouillon
# Outlook de l'état FNP) et le pipeline Hermes (detecter_demandes.py,
# creer_brouillons.py) — VOLONTAIREMENT PAS dans REQUIS ci-dessus.
# Essayé une fois (voir CLAUDE.md, session FNP) : `pip install pywin32` seul
# ne suffit pas à rendre `win32com.client` importable (il manque l'étape de
# post-installation qui enregistre pywintypes/pythoncom), donc
# verifier_et_installer() retombait en échec PERMANENT ("toujours
# introuvable après installation") — et comme gui.py/main.py/fnp.py
# appellent tous verifier_et_installer() AU DÉMARRAGE, ça bloquait le GUI
# ENTIER (comparatif/panier/BL inclus) pour un besoin qui ne concerne que le
# bouton optionnel "Créer le brouillon". Les modules qui en ont besoin
# l'importent en LOCAL, au moment précis de l'appel (jamais en tête de
# fichier) et laissent l'exception remonter avec un message clair si
# pywin32 est absent/mal installé sur le poste — jamais bloquant pour le
# reste de l'outil.


def verifier_et_installer(log=print):
    """Retourne True si toutes les dépendances sont disponibles (après
    installation automatique si besoin), False sinon."""

    manquants = []

    for module, paquet in REQUIS.items():
        try:
            importlib.import_module(module)
        except ImportError:
            manquants.append(paquet)

    if not manquants:
        return True

    log(f"Dépendance(s) manquante(s) : {', '.join(manquants)}")
    log(f"Installation automatique en cours (interpréteur : {sys.executable})...\n")

    def _pip_install(args_supplementaires):
        return subprocess.run(
            [sys.executable, "-m", "pip", "install", *args_supplementaires, *manquants],
            capture_output=True, text=True,
        )

    try:
        # --user ne fonctionne pas dans un environnement virtuel ("User
        # site-packages are not visible in this virtualenv") : on essaie
        # d'abord sans, puis avec --user en repli (poste sans venv où
        # l'installation globale nécessiterait des droits admin).
        resultat = _pip_install([])
        if resultat.returncode != 0:
            resultat_user = _pip_install(["--user"])
            if resultat_user.returncode == 0:
                resultat = resultat_user
        log(resultat.stdout)
        if resultat.returncode != 0:
            log(resultat.stderr)
            raise RuntimeError(f"pip a échoué (code {resultat.returncode})")
    except Exception as e:
        log(f"\n/!\\ Échec de l'installation automatique : {e}")
        log("Installe-les manuellement, dans une console, avec :")
        log(f"    {sys.executable} -m pip install {' '.join(manquants)}")
        return False

    importlib.invalidate_caches()

    encore_manquants = []
    for module in REQUIS:
        try:
            importlib.import_module(module)
        except ImportError:
            encore_manquants.append(module)

    if encore_manquants:
        log(f"\n/!\\ Toujours introuvable après installation : {', '.join(encore_manquants)}")
        return False

    log("Dépendances installées avec succès.\n")
    return True
