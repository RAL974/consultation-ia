# -*- coding: utf-8 -*-
"""
Création des BROUILLONS de demandes de devis dans Outlook — JAMAIS d'envoi.
Garde-fou : mail.Save() uniquement (brouillon). Pas de Send(), pas de Display().

Usage :
    python creer_brouillons.py --dry-run              # affiche ce qui serait créé, ne crée rien
    python creer_brouillons.py --create               # crée les brouillons + écrit suivi_consultations.csv

Entrée : brouillons_a_creer.json (dans le dossier du script) :
[
  {
    "to": ["allan.pougary@ravate.com", ...],
    "subject": "[Demande de devis] GYSM GARANCE - ...",
    "body": "...",
    "send_account": "achats@espace-soleil.re",
    "reference": "GYSM GARANCE - Demande de matériel (Mathieu BOISSET, 27/08)"
  }, ...
]
"""
import json
import os
import sys
import datetime

import win32com.client

BASE = os.path.dirname(os.path.abspath(__file__))
IN = os.path.join(BASE, "brouillons_a_creer.json")
SUIVI = os.path.join(BASE, "suivi_consultations.csv")

def find_account(ns, smtp):
    """Retourne l'objet Account dont SmtpAddress == smtp (insensible à la casse)."""
    for acc in ns.Accounts:
        try:
            if (acc.SmtpAddress or "").lower() == smtp.lower():
                return acc
        except Exception:
            pass
    return None

def create_drafts(create: bool):
    if not os.path.exists(IN):
        print("Pas de fichier brouillons_a_creer.json — rien à faire.")
        return
    items = json.load(open(IN, encoding="utf-8"))
    if not items:
        print("brouillons_a_creer.json vide — rien à faire.")
        return

    if not create:
        print("=== DRY-RUN : aucun brouillon créé ===")
        for it in items:
            print("-" * 70)
            print("À:", "; ".join(it.get("to", [])))
            print("Objet:", it.get("subject", ""))
            print("Compte:", it.get("send_account", ""))
            print("Réf.:", it.get("reference", ""))
            print("--- Corps ---")
            print(it.get("body", ""))
        return

    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")

    # Comptage avant (pour vérification réelle de l'effet)
    drafts_folder = ns.GetDefaultFolder(16)  # olFolderDrafts
    before = drafts_folder.Items.Count

    rows = []
    created = []
    for it in items:
        to = [t for t in it.get("to", []) if t]
        subject = it.get("subject", "")
        body = it.get("body", "")
        send_account = it.get("send_account", "achats@espace-soleil.re")
        reference = it.get("reference", "")

        mail = outlook.CreateItem(0)  # olMailItem
        for addr in to:
            mail.Recipients.Add(addr)
        mail.Subject = subject
        mail.Body = body
        acc = find_account(ns, send_account)
        if acc is not None:
            try:
                mail.SendUsingAccount = acc
            except Exception:
                pass
        try:
            mail.Save()  # brouillon — JAMAIS Send()
        except Exception as e:
            print("ÉCHEC brouillon:", subject, "->", repr(e))
            continue

        created.append(subject)
        rows.append({
            "date_creation": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
            "demande_source": reference,
            "articles": it.get("articles", ""),
            "destinataires": "; ".join(to),
            "objet_brouillon": subject,
            "statut": "brouillon créé",
        })

    # Vérification réelle : recompter les brouillons
    after = drafts_folder.Items.Count

    header = ["date_creation", "demande_source", "articles", "destinataires", "objet_brouillon", "statut"]
    write_header = not os.path.exists(SUIVI) or os.path.getsize(SUIVI) == 0
    with open(SUIVI, "a", encoding="utf-8") as f:
        if write_header:
            f.write(";".join(header) + "\n")
        for r in rows:
            f.write(";".join(str(r.get(h, "")).replace(";", ",").replace("\n", " ") for h in header) + "\n")

    print(f"Brouillons créés : {len(created)} (dossier Brouillons : {before} -> {after})")
    for s in created:
        print("  +", s)
    print("Suivi écrit dans:", SUIVI)

if __name__ == "__main__":
    create = "--create" in sys.argv
    create_drafts(create)
