"""
Consultation AI — interface graphique.

Permet de générer un comparatif sans ligne de commande ni VS Code : on
choisit un dossier de consultation (consultations/<nom>/, voir CLAUDE.md),
on clique sur "Générer le comparatif", et le fichier est créé dans son
sous-dossier resultats/.

Lancement : double-clic sur lancer_gui.bat (ou "python gui.py").
"""

import sys
import traceback
from pathlib import Path

DOSSIER_PROJET = Path(__file__).parent
FICHIER_JOURNAL_ERREUR = DOSSIER_PROJET / "gui_erreur.log"


def _consigner_erreur_fatale(texte):
    """Écrit le détail d'une erreur bloquante dans un fichier ET le laisse
    affiché en console (au lieu de fermer la fenêtre en un clin d'œil)."""
    try:
        FICHIER_JOURNAL_ERREUR.write_text(texte, encoding="utf-8")
    except Exception:
        pass
    print("=" * 60)
    print("ERREUR AU DEMARRAGE — Consultation AI ne peut pas s'ouvrir")
    print("=" * 60)
    print(texte)
    print(f"\nCe détail a aussi été enregistré ici :\n{FICHIER_JOURNAL_ERREUR}")
    input("\nAppuie sur Entrée pour fermer cette fenêtre...")


try:
    import os
    import queue
    import threading

    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    from moteur.dependances import verifier_et_installer
except Exception:
    _consigner_erreur_fatale(traceback.format_exc())
    sys.exit(1)

if not verifier_et_installer(print):
    _consigner_erreur_fatale(
        "Des dépendances Python sont manquantes et n'ont pas pu être "
        "installées automatiquement (voir le détail ci-dessus).\n"
        "Installe-les manuellement puis relance ce fichier."
    )
    sys.exit(1)

try:
    from moteur.pipeline import generer_comparatif
    from moteur.panier import generer_panier
    from moteur.consultation import (
        dossier_devis_de, dossier_resultats_de, fichier_besoin_de,
    )
    from gui_rapprochement import FenetreRapprochementBL
except Exception:
    _consigner_erreur_fatale(traceback.format_exc())
    sys.exit(1)


class RedirectionQueue:
    """Fichier factice : chaque écriture part dans une file d'attente,
    relue côté interface pour alimenter le journal affiché à l'écran."""

    def __init__(self, file_attente):
        self.file_attente = file_attente

    def write(self, texte):
        if texte:
            self.file_attente.put(texte)

    def flush(self):
        pass


class ConsultationGUI:

    def __init__(self, root):
        self.root = root
        root.title("Consultation AI — Générer un comparatif")
        root.geometry("780x700")
        root.minsize(640, 560)

        self.dossier_consultation = tk.StringVar()
        self.fichier_comparatif = tk.StringVar()
        self.file_attente = queue.Queue()
        self.dernier_comparatif = None
        self.dernier_panier = None

        self._construire_interface()
        self._verifier_bdd()
        self.root.after(100, self._lire_file_attente)

    # ------------------------------------------------------------------
    # Construction de la fenêtre
    # ------------------------------------------------------------------
    def _construire_interface(self):

        marge = {"padx": 10, "pady": 6}

        cadre_haut = tk.Frame(self.root)
        cadre_haut.pack(fill="x", **marge)
        cadre_haut.columnconfigure(0, weight=1)

        tk.Label(
            cadre_haut, text="1. Dossier de consultation (dans consultations/) :",
            font=("Segoe UI", 10, "bold"),
        ).grid(row=0, column=0, columnspan=2, sticky="w")

        tk.Entry(
            cadre_haut, textvariable=self.dossier_consultation, width=70, state="readonly",
        ).grid(row=1, column=0, sticky="we", padx=(0, 8))

        tk.Button(
            cadre_haut, text="Parcourir…", command=self._choisir_consultation,
        ).grid(row=1, column=1)

        self.bouton_generer = tk.Button(
            self.root, text="Générer le comparatif",
            font=("Segoe UI", 11, "bold"), bg="#1a7f37", fg="white",
            height=2, command=self._lancer_generation,
        )
        self.bouton_generer.pack(fill="x", **marge)

        cadre_panier = tk.Frame(self.root)
        cadre_panier.pack(fill="x", **marge)
        cadre_panier.columnconfigure(0, weight=1)

        tk.Label(
            cadre_panier,
            text="2. Comparatif décidé (détecté automatiquement, ou corrigé et enregistré dans Excel) :",
            font=("Segoe UI", 10, "bold"),
        ).grid(row=0, column=0, columnspan=2, sticky="w")

        tk.Entry(
            cadre_panier, textvariable=self.fichier_comparatif, width=70, state="readonly",
        ).grid(row=1, column=0, sticky="we", padx=(0, 8))

        tk.Button(
            cadre_panier, text="Parcourir…", command=self._choisir_comparatif,
        ).grid(row=1, column=1)

        self.bouton_panier = tk.Button(
            self.root, text="Générer le panier",
            font=("Segoe UI", 11, "bold"), bg="#1a5f9e", fg="white",
            height=2, command=self._lancer_generation_panier,
        )
        self.bouton_panier.pack(fill="x", **marge)

        tk.Label(
            self.root,
            text="3. Rapprochement BL (109 Distribution pour l'instant) — dépose les BL dans a_traiter/BL/ :",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", padx=10, pady=(6, 0))

        self.bouton_rapprochement = tk.Button(
            self.root, text="Rapprocher les BL",
            font=("Segoe UI", 11, "bold"), bg="#8250df", fg="white",
            height=2, command=self._ouvrir_rapprochement_bl,
        )
        self.bouton_rapprochement.pack(fill="x", **marge)

        tk.Label(self.root, text="Journal :").pack(anchor="w", padx=10)

        self.zone_log = scrolledtext.ScrolledText(
            self.root, height=16, font=("Consolas", 9), state="disabled",
        )
        self.zone_log.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        cadre_bas = tk.Frame(self.root)
        cadre_bas.pack(fill="x", **marge)

        self.bouton_ouvrir = tk.Button(
            cadre_bas, text="Ouvrir le comparatif", state="disabled",
            command=self._ouvrir_comparatif,
        )
        self.bouton_ouvrir.pack(side="left")

        self.bouton_ouvrir_panier = tk.Button(
            cadre_bas, text="Ouvrir le panier", state="disabled",
            command=self._ouvrir_panier,
        )
        self.bouton_ouvrir_panier.pack(side="left", padx=(8, 0))

        self.label_statut = tk.Label(cadre_bas, text="Prêt.", anchor="w")
        self.label_statut.pack(side="left", padx=12)

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------
    def _verifier_bdd(self):
        chemin_bdd = DOSSIER_PROJET / "base" / "BDD_articles.csv"
        if not chemin_bdd.exists():
            self._log(
                f"/!\\ Base articles absente : {chemin_bdd}\n"
                "    Le rapprochement automatique sera limité.\n\n"
            )

    def _choisir_consultation(self):
        chemin = filedialog.askdirectory(
            title="Choisir le dossier de consultation (dans consultations/)",
        )
        if not chemin:
            return

        self.dossier_consultation.set(chemin)
        dossier = Path(chemin)

        self.zone_log.configure(state="normal")
        self.zone_log.delete("1.0", "end")
        self.zone_log.configure(state="disabled")

        besoin = fichier_besoin_de(dossier)
        if besoin:
            self._log(f"Besoin détecté : {besoin.name}\n")
        else:
            self._log(
                "/!\\ Aucun fichier besoin dans ce dossier : le comparatif "
                "comparera toutes les lignes des devis, sans rapprochement "
                "avec une demande.\n"
            )

        dossier_devis = dossier_devis_de(dossier)
        nb_pdf = len(list(dossier_devis.glob("*.pdf"))) if dossier_devis.is_dir() else 0
        if nb_pdf:
            self._log(f"{nb_pdf} PDF de devis trouvé(s) dans devis/\n")
        else:
            self._log(f"/!\\ Aucun PDF trouvé dans {dossier_devis}\n")

        # Comparatif déjà généré pour cette consultation (le plus récent) :
        # pré-rempli pour pouvoir aller direct à "Générer le panier" si on
        # rouvre une consultation traitée plus tôt.
        dossier_resultats = dossier_resultats_de(dossier)
        comparatifs = (
            sorted(
                dossier_resultats.glob("Comparatif*.xlsx"),
                key=lambda f: f.stat().st_mtime, reverse=True,
            )
            if dossier_resultats.is_dir() else []
        )
        if comparatifs:
            self.fichier_comparatif.set(str(comparatifs[0]))
            self._log(f"Comparatif existant détecté : {comparatifs[0].name}\n")
        else:
            self.fichier_comparatif.set("")

    def _choisir_comparatif(self):
        chemin = filedialog.askopenfilename(
            title="Choisir le Comparatif (décidé, ou à valider tel quel)",
            filetypes=[
                ("Comparatif Excel", "*.xlsx"),
                ("Tous les fichiers", "*.*"),
            ],
        )
        if chemin:
            self.fichier_comparatif.set(chemin)

    def _log(self, texte):
        self.zone_log.configure(state="normal")
        self.zone_log.insert("end", texte)
        self.zone_log.see("end")
        self.zone_log.configure(state="disabled")

    def _lire_file_attente(self):
        try:
            while True:
                item = self.file_attente.get_nowait()
                if isinstance(item, tuple) and item and item[0] == "__FIN__":
                    self._generation_terminee(item[1])
                elif isinstance(item, tuple) and item and item[0] == "__FIN_PANIER__":
                    self._generation_panier_terminee(item[1])
                else:
                    self._log(item)
        except queue.Empty:
            pass
        self.root.after(100, self._lire_file_attente)

    # ------------------------------------------------------------------
    def _lancer_generation(self):

        dossier_consultation = self.dossier_consultation.get().strip()

        if not dossier_consultation:
            messagebox.showwarning(
                "Dossier de consultation manquant",
                "Choisis d'abord le dossier de consultation (dans consultations/).",
            )
            return

        dossier_devis = dossier_devis_de(dossier_consultation)

        if not any(dossier_devis.glob("*.pdf")):
            if not messagebox.askyesno(
                "Aucun PDF trouvé",
                f"Le dossier choisi ne contient aucun PDF :\n{dossier_devis}\n\n"
                "Continuer quand même ?",
            ):
                return

        self.bouton_generer.config(state="disabled", text="Génération en cours…")
        self.bouton_panier.config(state="disabled")
        self.bouton_ouvrir.config(state="disabled")
        self.bouton_ouvrir_panier.config(state="disabled")
        self.label_statut.config(text="Génération en cours…")

        self.zone_log.configure(state="normal")
        self.zone_log.delete("1.0", "end")
        self.zone_log.configure(state="disabled")

        thread = threading.Thread(
            target=self._generer, args=(dossier_consultation,), daemon=True,
        )
        thread.start()

    def _generer(self, dossier_consultation):
        """Exécuté dans un thread à part pour ne pas geler la fenêtre.

        Important : ce thread ne doit JAMAIS appeler de méthode Tk
        directement (root.after, config sur un widget...) — Tkinter n'est
        pas thread-safe. Le résultat est déposé dans la même file d'attente
        que le journal ; c'est le minuteur du thread principal
        (_lire_file_attente) qui le récupère et déclenche la suite."""

        ancien_stdout = sys.stdout
        sys.stdout = RedirectionQueue(self.file_attente)

        comparatif = None

        try:
            fichier_besoin = fichier_besoin_de(dossier_consultation)
            dossier_devis = dossier_devis_de(dossier_consultation)
            dossier_resultats = dossier_resultats_de(dossier_consultation)
            comparatif = generer_comparatif(
                DOSSIER_PROJET, fichier_besoin, dossier_devis, dossier_resultats,
            )
        except Exception:
            print("\n/!\\ ERREUR :\n")
            print(traceback.format_exc())
        finally:
            sys.stdout = ancien_stdout

        self.file_attente.put(("__FIN__", comparatif))

    def _generation_terminee(self, comparatif):

        self.bouton_generer.config(state="normal", text="Générer le comparatif")
        self.bouton_panier.config(state="normal")

        if comparatif:
            self.dernier_comparatif = Path(comparatif)
            self.fichier_comparatif.set(str(self.dernier_comparatif))
            self.bouton_ouvrir.config(state="normal")
            self.label_statut.config(text=f"Terminé : {self.dernier_comparatif.name}")
            messagebox.showinfo(
                "Comparatif généré",
                f"Le comparatif a été créé :\n\n{self.dernier_comparatif}\n\n"
                "Ouvre-le, choisis le fournisseur retenu par ligne (déjà "
                "pré-rempli au mieux-disant, modifiable), enregistre, puis "
                "clique sur « Générer le panier ».",
            )
        else:
            self.label_statut.config(text="Échec — voir le journal ci-dessus.")
            messagebox.showerror(
                "Échec de la génération",
                "La génération a échoué. Le détail est dans le journal ci-dessus.",
            )

    def _ouvrir_comparatif(self):
        if not self.dernier_comparatif:
            return
        try:
            if hasattr(os, "startfile"):
                os.startfile(self.dernier_comparatif)  # Windows
            else:
                messagebox.showinfo(
                    "Fichier prêt", f"Ouvre-le manuellement :\n{self.dernier_comparatif}",
                )
        except Exception as e:
            messagebox.showerror("Impossible d'ouvrir le fichier", str(e))

    # ------------------------------------------------------------------
    # Panier (étape 4)
    # ------------------------------------------------------------------
    def _lancer_generation_panier(self):

        dossier_consultation = self.dossier_consultation.get().strip()
        fichier_comparatif = self.fichier_comparatif.get().strip()

        if not dossier_consultation:
            messagebox.showwarning(
                "Dossier de consultation manquant",
                "Choisis d'abord le dossier de consultation (en haut) : il "
                "sert à retrouver les offres réelles correspondant à chaque "
                "décision du Comparatif.",
            )
            return

        if not fichier_comparatif:
            messagebox.showwarning(
                "Comparatif manquant",
                "Choisis le Comparatif décidé (généré ci-dessus, éventuellement "
                "corrigé et enregistré dans Excel).",
            )
            return

        self.bouton_generer.config(state="disabled")
        self.bouton_panier.config(state="disabled", text="Génération en cours…")
        self.bouton_ouvrir.config(state="disabled")
        self.bouton_ouvrir_panier.config(state="disabled")
        self.label_statut.config(text="Génération du panier en cours…")

        self.zone_log.configure(state="normal")
        self.zone_log.delete("1.0", "end")
        self.zone_log.configure(state="disabled")

        thread = threading.Thread(
            target=self._generer_panier,
            args=(dossier_consultation, fichier_comparatif),
            daemon=True,
        )
        thread.start()

    def _generer_panier(self, dossier_consultation, fichier_comparatif):
        """Exécuté dans un thread à part (voir _generer : même prudence
        Tkinter — jamais d'appel direct à un widget depuis ce thread)."""

        ancien_stdout = sys.stdout
        sys.stdout = RedirectionQueue(self.file_attente)

        panier = None

        try:
            fichier_besoin = fichier_besoin_de(dossier_consultation)
            dossier_devis = dossier_devis_de(dossier_consultation)
            dossier_resultats = dossier_resultats_de(dossier_consultation)
            panier = generer_panier(
                DOSSIER_PROJET, fichier_besoin, dossier_devis, fichier_comparatif,
                dossier_resultats,
            )
        except Exception:
            print("\n/!\\ ERREUR :\n")
            print(traceback.format_exc())
        finally:
            sys.stdout = ancien_stdout

        self.file_attente.put(("__FIN_PANIER__", panier))

    def _generation_panier_terminee(self, panier):

        self.bouton_generer.config(state="normal")
        self.bouton_panier.config(state="normal", text="Générer le panier")
        if self.dernier_comparatif:
            self.bouton_ouvrir.config(state="normal")

        if panier:
            self.dernier_panier = Path(panier)
            self.bouton_ouvrir_panier.config(state="normal")
            self.label_statut.config(text=f"Terminé : {self.dernier_panier.name}")
            messagebox.showinfo(
                "Panier généré",
                f"Le panier a été créé :\n\n{self.dernier_panier}\n\n"
                "Vérifie l'onglet « Non commandées » (rien perdu en silence) "
                "avant de coller le bloc de lignes dans le Suivi commandes.",
            )
        else:
            self.label_statut.config(text="Échec ou rien à commander — voir le journal.")
            messagebox.showwarning(
                "Aucun panier généré",
                "Voir le journal ci-dessus : soit une erreur, soit aucune "
                "ligne n'a de fournisseur retenu / d'offre à commander.",
            )

    def _ouvrir_panier(self):
        if not self.dernier_panier:
            return
        try:
            if hasattr(os, "startfile"):
                os.startfile(self.dernier_panier)  # Windows
            else:
                messagebox.showinfo(
                    "Fichier prêt", f"Ouvre-le manuellement :\n{self.dernier_panier}",
                )
        except Exception as e:
            messagebox.showerror("Impossible d'ouvrir le fichier", str(e))

    # ------------------------------------------------------------------
    # Rapprochement BL (étape 3, Rapprochement AI)
    # ------------------------------------------------------------------
    def _ouvrir_rapprochement_bl(self):
        try:
            FenetreRapprochementBL(self.root, DOSSIER_PROJET)
        except Exception as e:
            messagebox.showerror("Impossible d'ouvrir la fenêtre de rapprochement", str(e))


def main():
    try:
        root = tk.Tk()
        ConsultationGUI(root)
        root.mainloop()
    except Exception:
        _consigner_erreur_fatale(traceback.format_exc())
        sys.exit(1)


if __name__ == "__main__":
    main()
