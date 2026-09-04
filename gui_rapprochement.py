"""
Fenêtre de rapprochement BL (Rapprochement AI, session R2) — ouverte depuis
un bouton de gui.py. Toujours en 2 temps : lecture seule d'abord (OCR +
matching contre le Suivi commandes, rien n'est modifié), puis écriture
seulement après relecture et confirmation explicite de l'acheteur (voir
moteur/rapprochement/pipeline_bl.py, mode simulation par défaut).

Seul 109 Distribution est couvert pour l'instant (voir CLAUDE.md) — les BL
des autres fournisseurs ressortent listés en "fichiers non traités", pas
perdus en silence.
"""

import queue
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import messagebox

from moteur.rapprochement.ecriture import ClasseurVerrouille, ColonneNonModifiable
from moteur.rapprochement.pieces import FeuillePiecesAbsente
from moteur.rapprochement.pipeline_bl import (
    DOSSIER_A_TRAITER_BL,
    appliquer_et_archiver,
    rapprocher_dossier,
)


class CadreDefilant(tk.Frame):
    """Frame scrollable verticalement (Tkinter n'en a pas nativement)."""

    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)

        canevas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        barre = tk.Scrollbar(self, orient="vertical", command=canevas.yview)
        self.interieur = tk.Frame(canevas)

        self.interieur.bind(
            "<Configure>", lambda e: canevas.configure(scrollregion=canevas.bbox("all")),
        )
        canevas.create_window((0, 0), window=self.interieur, anchor="nw")
        canevas.configure(yscrollcommand=barre.set)

        canevas.pack(side="left", fill="both", expand=True)
        barre.pack(side="right", fill="y")

        canevas.bind_all("<MouseWheel>", lambda e: canevas.yview_scroll(int(-e.delta / 120), "units"))


class FenetreRapprochementBL(tk.Toplevel):

    def __init__(self, parent, dossier_projet):
        super().__init__(parent)

        self.dossier_projet = Path(dossier_projet)
        self.dossier_a_traiter = self.dossier_projet / DOSSIER_A_TRAITER_BL

        self.title("Rapprochement BL — 109 Distribution")
        self.geometry("840x640")
        self.minsize(600, 400)

        self.rapport = None
        self.cases = []  # [(tk.BooleanVar, BonLivraison, Correspondance)]
        self.file_attente = queue.Queue()

        self.label_statut = tk.Label(
            self, text="Lecture des BL en cours (OCR, peut prendre 1-2 min)…",
            font=("Segoe UI", 10, "bold"),
        )
        self.label_statut.pack(anchor="w", padx=10, pady=8)

        self.cadre_defilant = CadreDefilant(self)
        self.cadre_defilant.pack(fill="both", expand=True, padx=10)

        cadre_bas = tk.Frame(self)
        cadre_bas.pack(fill="x", padx=10, pady=8)

        self.bouton_ecrire = tk.Button(
            cadre_bas, text="Écrire les lignes cochées dans le Suivi commandes",
            font=("Segoe UI", 10, "bold"), bg="#1a7f37", fg="white", height=2,
            command=self._confirmer_ecriture, state="disabled",
        )
        self.bouton_ecrire.pack(fill="x")

        self.after(100, self._lire_file_attente)
        threading.Thread(target=self._lire_et_apparier, daemon=True).start()

    # ------------------------------------------------------------------
    def _lire_et_apparier(self):
        try:
            rapport = rapprocher_dossier(self.dossier_a_traiter, self.dossier_projet)
        except Exception:
            self.file_attente.put(("erreur_lecture", traceback.format_exc()))
            return
        self.file_attente.put(("rapport", rapport))

    def _lire_file_attente(self):
        try:
            while True:
                type_, valeur = self.file_attente.get_nowait()
                if type_ == "rapport":
                    self._afficher_rapport(valeur)
                elif type_ == "erreur_lecture":
                    self.label_statut.config(text="Erreur pendant la lecture des BL.")
                    messagebox.showerror("Erreur", valeur, parent=self)
                elif type_ == "resultat_ecriture":
                    self._afficher_resultat_ecriture(valeur)
                elif type_ == "erreur_ecriture":
                    self.bouton_ecrire.config(
                        state="normal", text="Écrire les lignes cochées dans le Suivi commandes",
                    )
                    messagebox.showerror("Erreur lors de l'écriture", valeur, parent=self)
        except queue.Empty:
            pass
        self.after(100, self._lire_file_attente)

    # ------------------------------------------------------------------
    def _afficher_rapport(self, rapport):
        self.rapport = rapport

        if rapport.fichier_suivi is None:
            self.label_statut.config(text="Suivi commandes introuvable à la racine du projet.")
            messagebox.showwarning(
                "Suivi commandes introuvable",
                "Dépose l'export « *Suivi commandes*.xlsx » à la racine du projet, puis relance.",
                parent=self,
            )
            return

        n_sur, n_confirmer = len(rapport.surs), len(rapport.a_confirmer)
        n_deja, n_inconnu = len(rapport.deja_a_jour), len(rapport.inconnus)
        n_anomalies = len(rapport.anomalies_lecture)
        n_anomalies_bl = len(rapport.anomalies_bl)

        self.label_statut.config(
            text=(
                f"{n_sur} ligne(s) sûre(s) · {n_confirmer} à confirmer · "
                f"{n_deja} déjà à jour · {n_inconnu} inconnue(s) · {n_anomalies} fichier(s) non traité(s) · "
                f"{n_anomalies_bl} BL non rapproché(s)"
            )
        )

        conteneur = self.cadre_defilant.interieur

        if n_sur:
            tk.Label(
                conteneur, text="Sûres — cochées par défaut",
                font=("Segoe UI", 10, "bold"), fg="#1a7f37",
            ).pack(anchor="w", pady=(6, 2))
            for bl, c in rapport.surs:
                self._ajouter_case(conteneur, bl, c, coche_defaut=True)

        if n_confirmer:
            tk.Label(
                conteneur, text="À confirmer — décochées par défaut, vérifie avant de cocher",
                font=("Segoe UI", 10, "bold"), fg="#b35900",
            ).pack(anchor="w", pady=(10, 2))
            for bl, c in rapport.a_confirmer:
                self._ajouter_case(conteneur, bl, c, coche_defaut=False)

        if n_deja:
            tk.Label(
                conteneur, text=f"Déjà à jour ({n_deja}) — rien à écrire, doublon évité",
                font=("Segoe UI", 10, "bold"), fg="#57606a",
            ).pack(anchor="w", pady=(10, 2))
            for bl, c in rapport.deja_a_jour:
                tk.Label(
                    conteneur, text=self._texte_ligne(bl, c), anchor="w", justify="left", wraplength=780,
                ).pack(anchor="w", padx=20)

        if n_inconnu:
            tk.Label(
                conteneur, text=f"Inconnues ({n_inconnu}) — à traiter à la main",
                font=("Segoe UI", 10, "bold"), fg="#cf222e",
            ).pack(anchor="w", pady=(10, 2))
            for bl, c in rapport.inconnus:
                texte = f"{bl.fichier} — réf. {c.ligne_bl.reference_fournisseur} : {', '.join(c.raisons)}"
                tk.Label(conteneur, text=texte, anchor="w", justify="left", wraplength=780).pack(anchor="w", padx=20)

        if rapport.anomalies_bl:
            tk.Label(
                conteneur, text=f"BL non rapprochés ({n_anomalies_bl}) — n° de commande introuvable ou absent du Suivi",
                font=("Segoe UI", 10, "bold"), fg="#57606a",
            ).pack(anchor="w", pady=(10, 2))
            for bl, raison in rapport.anomalies_bl:
                tk.Label(
                    conteneur, text=f"{bl.fichier} — commande {bl.numero_commande or '?'} : {raison}",
                    anchor="w", justify="left", wraplength=780,
                ).pack(anchor="w", padx=20)

        if rapport.anomalies_lecture:
            tk.Label(
                conteneur, text=f"Fichiers non traités ({n_anomalies})",
                font=("Segoe UI", 10, "bold"), fg="#57606a",
            ).pack(anchor="w", pady=(10, 2))
            for nom, raison in rapport.anomalies_lecture:
                tk.Label(
                    conteneur, text=f"{nom} — {raison}", anchor="w", justify="left", wraplength=780,
                ).pack(anchor="w", padx=20)

        if not (n_sur or n_confirmer or n_deja or n_inconnu or n_anomalies or n_anomalies_bl):
            tk.Label(conteneur, text=f"Aucun BL dans {self.dossier_a_traiter}.", anchor="w").pack(anchor="w", pady=10)

        self.bouton_ecrire.config(state="normal" if (n_sur or n_confirmer) else "disabled")

    def _texte_ligne(self, bl, c):
        qte_avant = c.ligne_suivi.qte_livree if c.ligne_suivi else 0
        texte = (
            f"{bl.fichier} — commande {bl.numero_commande}, réf. {c.ligne_bl.reference_fournisseur} : "
            f"Qté livrée {qte_avant:g} → {c.qte_livree_cumulee:g}, tarif {c.ligne_bl.prix_net:g}€"
        )
        if c.raisons:
            texte += "  ⚠ " + " ; ".join(c.raisons)
        return texte

    def _ajouter_case(self, conteneur, bl, c, coche_defaut):
        var = tk.BooleanVar(value=coche_defaut)
        tk.Checkbutton(
            conteneur, variable=var, text=self._texte_ligne(bl, c),
            anchor="w", justify="left", wraplength=760,
        ).pack(anchor="w", padx=10, fill="x")
        self.cases.append((var, bl, c))

    # ------------------------------------------------------------------
    def _confirmer_ecriture(self):

        selection = [(bl, c) for var, bl, c in self.cases if var.get()]

        if not selection:
            messagebox.showinfo("Rien de coché", "Coche au moins une ligne avant d'écrire.", parent=self)
            return

        if not messagebox.askyesno(
            "Confirmer l'écriture",
            f"Tu vas écrire {len(selection)} ligne(s) dans :\n\n{self.rapport.fichier_suivi}\n\n"
            "Une sauvegarde horodatée sera faite automatiquement avant toute écriture "
            "(voir backups/). Continuer ?",
            parent=self,
        ):
            return

        self.bouton_ecrire.config(state="disabled", text="Écriture en cours…")

        threading.Thread(target=self._ecrire, args=(selection,), daemon=True).start()

    def _ecrire(self, selection):
        try:
            resume = appliquer_et_archiver(
                self.dossier_projet, self.dossier_a_traiter, self.rapport, selection,
            )
        except (ClasseurVerrouille, ColonneNonModifiable, FeuillePiecesAbsente) as e:
            self.file_attente.put(("erreur_ecriture", str(e)))
            return
        except Exception:
            self.file_attente.put(("erreur_ecriture", traceback.format_exc()))
            return
        self.file_attente.put(("resultat_ecriture", resume))

    def _afficher_resultat_ecriture(self, resume):
        self.bouton_ecrire.config(state="disabled", text="Écrit — ferme et relance pour un nouveau passage")

        texte_echec = ""
        if resume["archivage_echoue"]:
            noms = ", ".join(f for f, _ in resume["archivage_echoue"])
            texte_echec = (
                f"\n⚠ {len(resume['archivage_echoue'])} BL bien écrit(s) mais PAS déplacé(s) "
                f"vers Traités/ (fichier verrouillé) — à ranger à la main : {noms}\n"
            )

        messagebox.showinfo(
            "Rapprochement terminé",
            f"{resume['lignes_ecrites']} ligne(s) écrite(s) dans le Suivi commandes.\n"
            f"Sauvegarde : {resume['sauvegarde']}\n\n"
            f"{len(resume['bl_archives'])} BL archivé(s) dans a_traiter/BL/Traités/.\n"
            f"{len(resume['bl_a_verifier'])} BL déplacé(s) vers a_traiter/BL/À vérifier/ "
            f"(décision humaine nécessaire).\n{texte_echec}\n"
            f"Rapport : {resume['chemin_rapport']}",
            parent=self,
        )
