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


def _euro_fr(montant: float) -> str:
    """1234.5 -> "1 234,50 €" (séparateur milliers espace, virgule décimale
    — convention française, pour un corps de mail en texte brut)."""
    return f"{montant:,.2f}".replace(",", " ").replace(".", ",") + " €"


def _corps_mail_fnp(rapport, mois_en_lettres: str) -> str:
    """Corps du mail avec les totaux RÉELS du rapport (session S0, étape 6) :
    total volet (a), total ESTIMÉ volet (b), total déclaré hors outil (4b),
    réserves de périmètre en clair (4c) — jamais juste un texte générique,
    la DAF doit pouvoir lire les chiffres sans ouvrir la pièce jointe."""

    total_bl = sum(l.montant_ht for l in rapport.lignes_bl)
    total_estime_transit = sum(d.cout_estime for d in rapport.dossiers_transitaires if d.cout_estime is not None)
    total_ajustements = sum(a.montant_ht for a in rapport.ajustements)

    lignes = [
        "Bonjour,",
        "",
        f"Vous trouverez ci-joint l'état des Factures Non Parvenues (FNP) pour {mois_en_lettres} :",
        "",
        f"- Volet (a) — BL reçus, pas encore facturés : {_euro_fr(total_bl)} HT ({len(rapport.lignes_bl)} ligne(s))",
        f"- Volet (b) — Transitaires, coût de transport ESTIMÉ : {_euro_fr(total_estime_transit)} HT "
        f"({len(rapport.dossiers_transitaires)} dossier(s))",
    ]

    if rapport.ajustements:
        lignes.append(
            f"- Déclaré par l'acheteur, hors outil (bons manuels, régularisations...) : "
            f"{_euro_fr(total_ajustements)} HT ({len(rapport.ajustements)} ligne(s))"
        )

    if rapport.factures_recues_non_rapprochees:
        lignes.append(
            f"- {len(rapport.factures_recues_non_rapprochees)} facture(s) déjà reçue(s) mais pas encore "
            "rapprochée(s) dans l'outil : sorties du volet (a), voir onglet \"Factures reçues\""
        )

    lignes += ["", "Réserves de périmètre :"]
    if rapport.reserves:
        lignes += [
            f"- {rapport.reserves.n_bdc_manuel_24x} facture(s) portant sur un bon manuel (carnet papier) "
            "— matériel livré, commande absente du Suivi, hors périmètre de ce calcul",
            f"- {rapport.reserves.n_transitaires_sans_estimation} dossier(s) transitaire sans estimation de "
            "coût disponible",
            f"- Le volet (b) ne couvre que les {rapport.reserves.n_dossiers_speciales_total} dossier(s) "
            "réellement saisis dans Commandes spéciales (classeur peu alimenté)",
        ]

    lignes += [
        "",
        "Le détail par fournisseur/chantier/ancienneté, ainsi que le niveau de confiance de chaque volet, "
        "sont repris dans l'onglet \"Synthèse\" du classeur joint.",
        "",
        "Cordialement,",
    ]

    return "\n".join(lignes)


def creer_brouillon_fnp(chemin_rapport, rapport, destinataires: list[str],
                         copie: list[str] | None = None, compte_envoi: str = "achats@espace-soleil.re") -> None:
    """Crée un brouillon Outlook avec `chemin_rapport` en pièce jointe,
    adressé à `destinataires` (+ `copie` en Cc), corps construit à partir des
    VRAIS totaux de `rapport` (RapportFNP, voir moteur.fnp — _corps_mail_fnp).
    Ne fait QUE Save() — voir bandeau du module. Lève une exception si
    Outlook/pywin32 est indisponible, si `chemin_rapport` n'existe pas, ou si
    la création échoue ; à l'appelant d'afficher un message clair."""

    import win32com.client
    from moteur.fnp import mois_en_lettres as _mois_en_lettres

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

    mois_texte = _mois_en_lettres(rapport.mois)
    mail.Subject = f"État des Factures Non Parvenues — {mois_texte}"
    mail.Body = _corps_mail_fnp(rapport, mois_texte)

    try:
        acc = next((a for a in ns.Accounts if (a.SmtpAddress or "").lower() == compte_envoi.lower()), None)
        if acc is not None:
            mail.SendUsingAccount = acc
    except Exception:
        pass

    mail.Attachments.Add(str(chemin_rapport))

    mail.Save()  # brouillon — JAMAIS Send()
