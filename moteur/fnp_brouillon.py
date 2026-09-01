# -*- coding: utf-8 -*-
"""
Brouillon Outlook pour l'état FNP mensuel — même garde-fou que
creer_brouillons.py (pipeline Hermes, racine du projet) : mail.Save()
UNIQUEMENT, jamais Send() ni Display(). L'acheteur relit et envoie
elle-même depuis Outlook.

Import de win32com fait à l'INTÉRIEUR de creer_brouillon_fnp() (pas en tête
de module) : ce module est importé par gui_fnp.py au démarrage du GUI, avant
que moteur.dependances.verifier_et_installer() ait forcément eu l'occasion
d'installer pywin32 — un import en tête de fichier ferait planter tout le
GUI si pywin32 manque encore, alors que ça n'est nécessaire qu'au moment
précis de créer le brouillon.
"""

from pathlib import Path


def creer_brouillon_fnp(chemin_rapport, destinataires: list[str], mois_en_lettres: str,
                         copie: list[str] | None = None, compte_envoi: str = "achats@espace-soleil.re") -> None:
    """Crée un brouillon Outlook avec `chemin_rapport` en pièce jointe,
    adressé à `destinataires` (+ `copie` en Cc). Ne fait QUE Save() — voir
    bandeau du module. Lève une exception si Outlook/pywin32 est indisponible,
    si `chemin_rapport` n'existe pas, ou si la création échoue ; à l'appelant
    d'afficher un message clair."""

    import win32com.client

    chemin_rapport = Path(chemin_rapport)
    if not chemin_rapport.exists():
        raise FileNotFoundError(f"Rapport introuvable : {chemin_rapport}")

    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")

    mail = outlook.CreateItem(0)  # olMailItem
    for addr in destinataires:
        mail.Recipients.Add(addr)
    for addr in (copie or []):
        rec = mail.Recipients.Add(addr)
        rec.Type = 2  # olCC
    mail.Recipients.ResolveAll()

    mail.Subject = f"État des Factures Non Parvenues — {mois_en_lettres}"
    mail.Body = (
        f"Bonjour,\n\n"
        f"Vous trouverez ci-joint l'état des Factures Non Parvenues (FNP) pour {mois_en_lettres} :\n"
        f"- Volet (a) : bons de livraison reçus mais pas encore facturés (voir onglet \"BL non facturés\").\n"
        f"- Volet (b) : dossiers transitaires arrivés mais dont la facture de transport n'est pas encore reçue "
        f"(voir onglet \"Transitaires\").\n\n"
        f"Le détail des montants, le niveau de confiance et les points de vigilance (notamment la part de "
        f"livraisons antérieures à la mise en place du suivi factures) sont repris dans l'onglet \"Synthèse\".\n\n"
        f"Cordialement,\n"
    )

    try:
        acc = next((a for a in ns.Accounts if (a.SmtpAddress or "").lower() == compte_envoi.lower()), None)
        if acc is not None:
            mail.SendUsingAccount = acc
    except Exception:
        pass

    mail.Attachments.Add(str(chemin_rapport))

    mail.Save()  # brouillon — JAMAIS Send()
