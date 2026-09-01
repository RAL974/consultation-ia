"""
Fenêtre de rapprochement factures (Rapprochement AI, session F2) — ouverte
depuis un bouton de gui.py. Même principe que gui_rapprochement.py (BL) :
toujours en 2 temps, lecture seule d'abord (texte PDF natif + matching
contre le Suivi commandes, rien n'est modifié), puis écriture seulement
après relecture et confirmation explicite de l'acheteur (voir
moteur/rapprochement/pipeline_facture.py, mode simulation par défaut).

Seul 109 Distribution est couvert pour l'instant (voir CLAUDE.md) — les
factures des autres fournisseurs ressortent listées en "fichiers non
traités", pas perdues en silence.

Les colonnes facture (N° facture / Date facture / Qté facturée / PU
facturé) ne sont pas encore créées dans le vrai Suivi commandes à ce jour
(voir CLAUDE.md, Volet 1) : le rapprochement en LECTURE SEULE reste
utilisable sans elles (diagnostic), mais toute tentative d'ÉCRITURE échoue
proprement (ColonneNonModifiable, message clair affiché) tant qu'elles
n'existent pas — pas une erreur à contourner ici."""

import queue
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import messagebox

from moteur.rapprochement.ecriture import ClasseurVerrouille, ColonneNonModifiable
from moteur.rapprochement.pipeline_facture import (
    DOSSIER_A_TRAITER_FACTURES,
    appliquer_et_archiver_factures,
    rapprocher_dossier_factures,
)

from gui_rapprochement import CadreDefilant


class FenetreRapprochementFacture(tk.Toplevel):

    def __init__(self, parent, dossier_projet):
        super().__init__(parent)

        self.dossier_projet = Path(dossier_projet)
        self.dossier_a_traiter = self.dossier_projet / DOSSIER_A_TRAITER_FACTURES

        self.title("Rapprochement factures — 109 Distribution")
        self.geometry("880x660")
        self.minsize(600, 400)

        self.rapport = None
        self.cases = []  # [(tk.BooleanVar, Facture, CorrespondanceFacture)]
        self.file_attente = queue.Queue()

        self.label_statut = tk.Label(
            self, text="Lecture des factures en cours…",
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
            rapport = rapprocher_dossier_factures(self.dossier_a_traiter, self.dossier_projet)
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
                    self.label_statut.config(text="Erreur pendant la lecture des factures.")
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
            self.label_statut.config(text="Suivi commandes introuvable.")
            messagebox.showwarning(
                "Suivi commandes introuvable",
                "Le classeur « 1.3.0.1. Suivi commandes... » est introuvable dans "
                "« 1.3.0.1. Commandes courantes/ », à côté du dossier projet.",
                parent=self,
            )
            return

        n_sur, n_confirmer = len(rapport.surs), len(rapport.a_confirmer)
        n_deja, n_inconnu = len(rapport.deja_a_jour), len(rapport.inconnus)
        n_anomalies = len(rapport.anomalies_lecture)
        n_anomalies_facture = len(rapport.anomalies_facture)

        self.label_statut.config(
            text=(
                f"{n_sur} ligne(s) sûre(s) · {n_confirmer} à confirmer · "
                f"{n_deja} déjà à jour · {n_inconnu} inconnue(s) · {n_anomalies} fichier(s) non traité(s) · "
                f"{n_anomalies_facture} bloc(s) non rapproché(s)"
            )
        )

        conteneur = self.cadre_defilant.interieur

        if n_sur:
            tk.Label(
                conteneur, text="Sûres — cochées par défaut",
                font=("Segoe UI", 10, "bold"), fg="#1a7f37",
            ).pack(anchor="w", pady=(6, 2))
            for facture, c in rapport.surs:
                self._ajouter_case(conteneur, facture, c, coche_defaut=True)

        if n_confirmer:
            tk.Label(
                conteneur, text="À confirmer — décochées par défaut, vérifie avant de cocher",
                font=("Segoe UI", 10, "bold"), fg="#b35900",
            ).pack(anchor="w", pady=(10, 2))
            for facture, c in rapport.a_confirmer:
                self._ajouter_case(conteneur, facture, c, coche_defaut=False)

        if n_deja:
            tk.Label(
                conteneur, text=f"Déjà à jour ({n_deja}) — rien à écrire, doublon évité",
                font=("Segoe UI", 10, "bold"), fg="#57606a",
            ).pack(anchor="w", pady=(10, 2))
            for facture, c in rapport.deja_a_jour:
                tk.Label(
                    conteneur, text=self._texte_ligne(facture, c), anchor="w", justify="left", wraplength=820,
                ).pack(anchor="w", padx=20)

        if n_inconnu:
            tk.Label(
                conteneur, text=f"Inconnues ({n_inconnu}) — à traiter à la main",
                font=("Segoe UI", 10, "bold"), fg="#cf222e",
            ).pack(anchor="w", pady=(10, 2))
            for facture, c in rapport.inconnus:
                texte = f"{facture.fichier} — réf. {c.ligne_facture.reference_fournisseur} : {', '.join(c.raisons)}"
                tk.Label(conteneur, text=texte, anchor="w", justify="left", wraplength=820).pack(anchor="w", padx=20)

        if rapport.anomalies_facture:
            tk.Label(
                conteneur, text=f"Blocs non rapprochés ({n_anomalies_facture}) — avoir, commande introuvable...",
                font=("Segoe UI", 10, "bold"), fg="#57606a",
            ).pack(anchor="w", pady=(10, 2))
            for facture, raison in rapport.anomalies_facture:
                tk.Label(
                    conteneur, text=f"{facture.fichier} — facture {facture.numero_facture or '?'} : {raison}",
                    anchor="w", justify="left", wraplength=820,
                ).pack(anchor="w", padx=20)

        if rapport.anomalies_lecture:
            tk.Label(
                conteneur, text=f"Fichiers non traités ({n_anomalies})",
                font=("Segoe UI", 10, "bold"), fg="#57606a",
            ).pack(anchor="w", pady=(10, 2))
            for nom, raison in rapport.anomalies_lecture:
                tk.Label(
                    conteneur, text=f"{nom} — {raison}", anchor="w", justify="left", wraplength=820,
                ).pack(anchor="w", padx=20)

        if not (n_sur or n_confirmer or n_deja or n_inconnu or n_anomalies or n_anomalies_facture):
            tk.Label(conteneur, text=f"Aucune facture dans {self.dossier_a_traiter}.", anchor="w").pack(anchor="w", pady=10)

        self.bouton_ecrire.config(state="normal" if (n_sur or n_confirmer) else "disabled")

    def _texte_ligne(self, facture, c):
        qte_avant = c.ligne_suivi.qte_livree if c.ligne_suivi else 0
        texte = (
            f"{facture.fichier} — facture {facture.numero_facture}, commande "
            f"{c.ligne_facture.numero_commande or '?'}, réf. {c.ligne_facture.reference_fournisseur} : "
            f"Qté livrée {qte_avant:g}, Qté facturée {c.ligne_facture.quantite_facturee:g}, "
            f"PU facturé {c.ligne_facture.prix_unitaire_ht:g}€"
        )
        if c.raisons:
            texte += "  ⚠ " + " ; ".join(c.raisons)
        return texte

    def _ajouter_case(self, conteneur, facture, c, coche_defaut):
        var = tk.BooleanVar(value=coche_defaut)
        tk.Checkbutton(
            conteneur, variable=var, text=self._texte_ligne(facture, c),
            anchor="w", justify="left", wraplength=800,
        ).pack(anchor="w", padx=10, fill="x")
        self.cases.append((var, facture, c))

    # ------------------------------------------------------------------
    def _confirmer_ecriture(self):

        selection = [(facture, c) for var, facture, c in self.cases if var.get()]

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
            resume = appliquer_et_archiver_factures(
                self.dossier_projet, self.dossier_a_traiter, self.rapport, selection,
            )
        except (ClasseurVerrouille, ColonneNonModifiable) as e:
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
                f"\n⚠ {len(resume['archivage_echoue'])} facture(s) bien écrite(s) mais PAS déplacée(s)/copiée(s) "
                f"(fichier verrouillé) — à ranger à la main : {noms}\n"
            )

        texte_resorption = ""
        if resume["resorption"]:
            r = resume["resorption"]
            texte_resorption = (
                f"\nRésorption 109 DISTRIBUTION : {r['a_facturer']} ligne(s) livrée(s) encore sans "
                f"facture sur {r['livrees']} au total.\n"
            )

        messagebox.showinfo(
            "Rapprochement terminé",
            f"{resume['lignes_ecrites']} ligne(s) écrite(s) dans le Suivi commandes.\n"
            f"Sauvegarde : {resume['sauvegarde']}\n\n"
            f"{len(resume['factures_archivees'])} facture(s) archivée(s) dans a_traiter/BL/Traités/.\n"
            f"{len(resume['factures_a_verifier'])} facture(s) déplacée(s) vers a_traiter/Factures/À vérifier/ "
            f"(décision humaine nécessaire).\n{texte_echec}{texte_resorption}\n"
            f"Rapport : {resume['chemin_rapport']}",
            parent=self,
        )
