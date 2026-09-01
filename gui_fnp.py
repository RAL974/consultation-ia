# -*- coding: utf-8 -*-
"""
Fenêtre « État FNP du mois » (clôture comptable, demande de la DAF) — ouverte
depuis un bouton de gui.py. Entièrement LECTURE SEULE : ne modifie jamais le
Suivi commandes ni Commandes spéciales, écrit uniquement
rapports/FNP_<AAAA-MM>.xlsx — pas de confirmation nécessaire avant de
générer (contrairement aux fenêtres de rapprochement BL/facture, qui elles
écrivent dans le classeur vivant).

Le brouillon Outlook (voir moteur/fnp_brouillon.py) reste une action
SÉPARÉE et explicite, avec ses propres champs destinataire(s)/copie saisis
à la main — jamais une adresse devinée ou mémorisée en dur."""

import queue
import threading
import traceback
from datetime import date
from pathlib import Path

import tkinter as tk
from tkinter import messagebox

from moteur.fnp import (
    SuiviIntrouvable, calculer_rapport_fnp, ecrire_classeur_fnp, mois_en_lettres, mois_precedent_complet,
)


class FenetreFNP(tk.Toplevel):

    def __init__(self, parent, dossier_projet):
        super().__init__(parent)

        self.dossier_projet = Path(dossier_projet)
        self.rapport = None
        self.chemin_genere = None
        self.file_attente = queue.Queue()

        self.title("État FNP du mois (Factures Non Parvenues)")
        self.geometry("720x560")
        self.minsize(560, 420)

        marge = {"padx": 10, "pady": 6}

        cadre_haut = tk.Frame(self)
        cadre_haut.pack(fill="x", **marge)

        tk.Label(cadre_haut, text="Mois de clôture (AAAA-MM) :").grid(row=0, column=0, sticky="w")
        self.var_mois = tk.StringVar(value=mois_precedent_complet())
        tk.Entry(cadre_haut, textvariable=self.var_mois, width=12).grid(row=0, column=1, sticky="w", padx=(6, 20))

        tk.Label(cadre_haut, text="Depuis le (AAAA-MM-JJ, optionnel) :").grid(row=0, column=2, sticky="w")
        self.var_depuis = tk.StringVar()
        tk.Entry(cadre_haut, textvariable=self.var_depuis, width=12).grid(row=0, column=3, sticky="w", padx=(6, 0))

        self.bouton_generer = tk.Button(
            self, text="Générer l'état FNP",
            font=("Segoe UI", 11, "bold"), bg="#1a7f37", fg="white", height=2,
            command=self._lancer_generation,
        )
        self.bouton_generer.pack(fill="x", **marge)

        self.label_resume = tk.Label(self, text="", anchor="w", justify="left", wraplength=680)
        self.label_resume.pack(fill="x", **marge)

        cadre_actions = tk.Frame(self)
        cadre_actions.pack(fill="x", **marge)

        self.bouton_ouvrir = tk.Button(
            cadre_actions, text="Ouvrir l'état généré", state="disabled", command=self._ouvrir,
        )
        self.bouton_ouvrir.pack(side="left")

        tk.Label(self, text="Créer un brouillon Outlook pour la DAF (jamais envoyé automatiquement) :",
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(14, 0))

        cadre_brouillon = tk.Frame(self)
        cadre_brouillon.pack(fill="x", **marge)
        cadre_brouillon.columnconfigure(1, weight=1)

        tk.Label(cadre_brouillon, text="Destinataire(s) (; séparés) :").grid(row=0, column=0, sticky="w")
        self.var_destinataires = tk.StringVar()
        tk.Entry(cadre_brouillon, textvariable=self.var_destinataires).grid(row=0, column=1, sticky="we", padx=(6, 0))

        tk.Label(cadre_brouillon, text="Copie (direction, optionnel) :").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.var_copie = tk.StringVar()
        tk.Entry(cadre_brouillon, textvariable=self.var_copie).grid(row=1, column=1, sticky="we", padx=(6, 0), pady=(4, 0))

        self.bouton_brouillon = tk.Button(
            self, text="Créer le brouillon", state="disabled",
            font=("Segoe UI", 10, "bold"), bg="#1a5f9e", fg="white", height=2,
            command=self._creer_brouillon,
        )
        self.bouton_brouillon.pack(fill="x", **marge)

        self.after(100, self._lire_file_attente)

    # ------------------------------------------------------------------
    def _lancer_generation(self):

        mois = self.var_mois.get().strip()
        depuis_texte = self.var_depuis.get().strip()

        depuis = None
        if depuis_texte:
            try:
                annee, m, j = (int(x) for x in depuis_texte.split("-"))
                depuis = date(annee, m, j)
            except ValueError:
                messagebox.showwarning(
                    "Date invalide", "Le filtre \"Depuis le\" doit être au format AAAA-MM-JJ.", parent=self,
                )
                return

        self.bouton_generer.config(state="disabled", text="Génération en cours…")
        self.bouton_ouvrir.config(state="disabled")
        self.bouton_brouillon.config(state="disabled")
        self.label_resume.config(text="Lecture du Suivi commandes et de Commandes spéciales en cours…")

        threading.Thread(target=self._generer, args=(mois, depuis), daemon=True).start()

    def _generer(self, mois, depuis):
        try:
            rapport = calculer_rapport_fnp(self.dossier_projet, mois, depuis)
            chemin = ecrire_classeur_fnp(self.dossier_projet, rapport)
        except SuiviIntrouvable as e:
            self.file_attente.put(("erreur", str(e)))
            return
        except ValueError as e:
            self.file_attente.put(("erreur", f"Mois invalide « {mois} » : {e}"))
            return
        except Exception:
            self.file_attente.put(("erreur", traceback.format_exc()))
            return
        self.file_attente.put(("ok", rapport, chemin))

    def _lire_file_attente(self):
        try:
            while True:
                item = self.file_attente.get_nowait()
                if item[0] == "ok":
                    self._afficher_resultat(item[1], item[2])
                elif item[0] == "erreur":
                    self.bouton_generer.config(state="normal", text="Générer l'état FNP")
                    self.label_resume.config(text="")
                    messagebox.showerror("Erreur", item[1], parent=self)
        except queue.Empty:
            pass
        self.after(100, self._lire_file_attente)

    def _afficher_resultat(self, rapport, chemin):
        self.rapport = rapport
        self.chemin_genere = chemin

        self.bouton_generer.config(state="normal", text="Générer l'état FNP")
        self.bouton_ouvrir.config(state="normal")
        self.bouton_brouillon.config(state="normal")

        total_bl = sum(l.montant_ht for l in rapport.lignes_bl)
        total_marchandise = sum(d.montant_marchandise for d in rapport.dossiers_transitaires)
        total_estime = sum(d.cout_estime for d in rapport.dossiers_transitaires if d.cout_estime is not None)

        texte = (
            f"Généré : {chemin}\n\n"
            f"Volet (a) BL non facturés : {total_bl:,.2f} € sur {len(rapport.lignes_bl)} ligne(s) "
            f"({len(rapport.lignes_sans_prix)} ligne(s) livrée(s) sans prix connu, non valorisées).\n"
            f"Volet (b) Transitaires : {len(rapport.dossiers_transitaires)} dossier(s), "
            f"{total_marchandise:,.2f} € de marchandise, {total_estime:,.2f} € de coût transitaire ESTIMÉ."
        )
        if rapport.transitaire_repli_utilise:
            texte += f"\n⚠ {rapport.transitaire_avertissement}"

        self.label_resume.config(text=texte)

    # ------------------------------------------------------------------
    def _ouvrir(self):
        if not self.chemin_genere:
            return
        import os
        try:
            os.startfile(self.chemin_genere)
        except Exception as e:
            messagebox.showerror("Impossible d'ouvrir le fichier", str(e), parent=self)

    def _creer_brouillon(self):

        destinataires = [a.strip() for a in self.var_destinataires.get().split(";") if a.strip()]
        copie = [a.strip() for a in self.var_copie.get().split(";") if a.strip()]

        if not destinataires:
            messagebox.showwarning(
                "Destinataire manquant", "Renseigne au moins un destinataire avant de créer le brouillon.",
                parent=self,
            )
            return

        if not messagebox.askyesno(
            "Créer le brouillon",
            f"Un brouillon Outlook va être créé (PAS envoyé) avec « {self.chemin_genere.name} » en pièce jointe, "
            f"à destination de :\n{', '.join(destinataires)}\n"
            + (f"Copie : {', '.join(copie)}\n" if copie else "")
            + "\nTu devras l'envoyer toi-même depuis Outlook après relecture. Continuer ?",
            parent=self,
        ):
            return

        try:
            from moteur.fnp_brouillon import creer_brouillon_fnp
            creer_brouillon_fnp(
                self.chemin_genere, destinataires, mois_en_lettres(self.rapport.mois), copie=copie,
            )
        except Exception as e:
            messagebox.showerror(
                "Impossible de créer le brouillon",
                f"{e}\n\nOutlook (et le module pywin32) doivent être installés et Outlook ouvert au moins une fois.",
                parent=self,
            )
            return

        messagebox.showinfo(
            "Brouillon créé",
            "Le brouillon a été créé dans Outlook (dossier Brouillons) — relis-le puis envoie-le toi-même.",
            parent=self,
        )
