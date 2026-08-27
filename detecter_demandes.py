# -*- coding: utf-8 -*-
"""
Détection du flux entrant — demandes d'achat internes + devis fournisseurs reçus.
Lecture SEULE des boîtes ral@espace-soleil.re et achats@espace-soleil.re (Outlook COM).
Aucun envoi, aucune écriture mailbox : sortie = fichiers JSON pour analyse par l'agent.

Usage :
    python detecter_demandes.py            # fenêtre depuis la dernière passe (ou 48h au 1er run)
    python detecter_demandes.py --hours 72 # forcer la fenêtre

Sorties (dans le dossier du script) :
    demandes_a_traiter.json   candidats "demande d'achat / besoin" (dont deja_traite flag)
    devis_recus.json          candidats "devis fournisseur reçu" (PDF joint, expéditeur externe)
    etat.json                 état : last_scan + EntryID déjà vus
"""
import json
import os
import re
import sys
import datetime
import unicodedata

import win32com.client

BASE = os.path.dirname(os.path.abspath(__file__))
ETAT = os.path.join(BASE, "etat.json")
CANDIDATS = os.path.join(BASE, "demandes_a_traiter.json")
DEVIS = os.path.join(BASE, "devis_recus.json")

INTERNE_DOMAINE = "espace-soleil.re"

# Liste OFFICIELLE des demandeurs (conducs/RA/maintenance) — fournie par
# William 27/08/2026. Un mail dont l'expéditeur est dans cette liste EST une
# demande d'achat potentielle, quel que soit le sujet (souvent un nom
# d'affaire). Comparaison insensible casse/accents.
NOMS_DEMANDEURS = [
    "arnaud aujoulat", "bernard riviere", "brasseur air espace soleil",
    "cazambo nicolas", "chahissou daniel", "coutarel nicolas",
    "housseine toiliha", "hugues latra", "jonathan mercher",
    "karim barhoumi-andreani", "ludovic govin", "mathieu boisset",
    "maintenance espace soleil", "payet caroline", "said eldayane",
    "stephane schoch", "wilfried adeux",
]

# Mots-clés sujet d'une demande d'achat / besoin interne (testés en minuscules)
KW_DEMANDE = [
    "demande de matériel", "demande de materiel", "besoin", "besoins",
    "commande de matériel", "commande de materiel", "commande matériel",
    "demande d'achat", "demande de devis", "demande de prix",
    "à commander", "a commander", "besoin urgent",
]
# Mots-clés sujet d'un devis reçu (réponse fournisseur)
KW_DEVIS = ["devis", "offre", "proposition", "quotation", "price"]

# Bruit connu (newsletters, sollicitations, notifications) — testé sur expéditeur + sujet
BRUIT = [
    "lenovo", "alibaba", "abrule", "dqzhanghaoyue", "globalsources", "gsl energy",
    "loxam day", "réponse automatique", "reponse automatique", "automatic reply",
    "noreply", "no-reply", "calendly", "cloudflare", "zeendoc", "colissimo",
    "newsletter", "désinscription", "desinscription", "communauté", "webinaire",
    "surmontez", "backlit panel", "sourcing outlook", "trade show",
]

def norm_name(s: str) -> str:
    """Normalise un nom : minuscules, sans accents, espaces réduits."""
    s = unicodedata.normalize("NFD", s or "")
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s).strip().lower()

def is_demandeur(sender_name: str) -> bool:
    """True si l'expéditeur fait partie des demandeurs officiels."""
    n = norm_name(sender_name)
    return n in NOMS_DEMANDEURS

def norm_subject(s: str) -> str:
    """Normalise un sujet pour comparaison (retire RE:/TR:/FW:, casse, espaces)."""
    s = re.sub(r"^(re|tr|fw|fwd|r\s?|ré|réf)\s*:\s*", "", s, flags=re.IGNORECASE)
    s = re.sub(r"^\[externe\]\s*", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

def get_sender(m):
    try:
        return (m.SenderName or "").strip()
    except Exception:
        return ""

def get_sender_email(m):
    try:
        ea = m.SenderEmailAddress or ""
        if ea.lower().startswith("/o="):
            try:
                eu = m.Sender.GetExchangeUser()
                if eu and eu.PrimarySmtpAddress:
                    return eu.PrimarySmtpAddress.lower()
            except Exception:
                pass
        return ea.lower()
    except Exception:
        return ""

def get_recipients_sent(m):
    """Adresses des destinataires d'un mail envoyé (résolution GAL si besoin)."""
    out = []
    try:
        for r in m.Recipients:
            try:
                ea = r.Address or ""
                if ea.lower().startswith("/o="):
                    try:
                        eu = r.AddressEntry.GetExchangeUser()
                        if eu and eu.PrimarySmtpAddress:
                            ea = eu.PrimarySmtpAddress
                    except Exception:
                        pass
                if ea:
                    out.append(ea.lower())
            except Exception:
                pass
    except Exception:
        pass
    return out

def is_bruit(sender_name, sender_email, subject):
    hay = " ".join([sender_name or "", sender_email or "", subject or ""]).lower()
    return any(b in hay for b in BRUIT)

def scan():
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")

    # Fenêtre
    etat = {"last_scan": None, "processed": []}
    if os.path.exists(ETAT):
        try:
            etat = json.load(open(ETAT, encoding="utf-8"))
        except Exception:
            etat = {"last_scan": None, "processed": []}
    try:
        last_scan = datetime.datetime.fromisoformat(etat["last_scan"]) if etat.get("last_scan") else None
    except Exception:
        last_scan = None
    if last_scan is None:
        last_scan = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(hours=48)
    # Sécurité : ne jamais remonter plus de 72h même si l'état est vieux (évite un raz-de-marée)
    min_start = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(hours=72)
    start = max(last_scan, min_start)

    demandes, devis = [], []
    processed = set(etat.get("processed", []))
    sent_subjects = {}  # (store, norm_subject) -> plus récent SentOn (pour anti-doublon)

    # 1) Index des Envoyés (sujets normalisés -> date) pour l'anti-doublon "déjà traité"
    for store in ns.Stores:
        dn = store.DisplayName or ""
        if "ral@" not in dn and "achats@" not in dn:
            continue
        try:
            sent = store.GetDefaultFolder(5)  # olFolderSentMail
            items = sent.Items
            items.Sort("[SentOn]", True)
            n = min(3000, items.Count)
            for i in range(1, n + 1):
                try:
                    m = items.Item(i)
                    t = m.SentOn
                    if t.tzinfo is None:
                        t = t.replace(tzinfo=datetime.timezone.utc)
                    if t < start - datetime.timedelta(days=30):
                        break
                    key = norm_subject(m.Subject or "")
                    if key:
                        prev = sent_subjects.get((dn, key))
                        if prev is None or t > prev:
                            sent_subjects[(dn, key)] = t
                except Exception:
                    pass
        except Exception:
            pass

    # 2) Scan des Boîtes de réception
    for store in ns.Stores:
        dn = store.DisplayName or ""
        if "ral@" not in dn and "achats@" not in dn:
            continue
        try:
            inbox = store.GetDefaultFolder(6)  # olFolderInbox
            items = inbox.Items
            items.Sort("[ReceivedTime]", True)
            n = min(2000, items.Count)
            for i in range(1, n + 1):
                try:
                    m = items.Item(i)
                    t = m.ReceivedTime
                    if t.tzinfo is None:
                        t = t.replace(tzinfo=datetime.timezone.utc)
                    if t < start:
                        continue
                    try:
                        eid = m.EntryID
                    except Exception:
                        eid = None
                    if eid and eid in processed:
                        continue
                    sender = get_sender(m)
                    semail = get_sender_email(m)
                    subject = m.Subject or ""
                    if is_bruit(sender, semail, subject):
                        continue

                    # Déjà traité ? (un Envoyé au même sujet, postérieur à la réception)
                    skey = norm_subject(subject)
                    handled = False
                    if skey:
                        sdate = sent_subjects.get((dn, skey))
                        if sdate and sdate >= t:
                            handled = True

                    atts = []
                    try:
                        for a in m.Attachments:
                            atts.append(a.FileName)
                    except Exception:
                        pass

                    record = {
                        "store": dn, "entry_id": eid, "sender": sender,
                        "sender_email": semail, "subject": subject,
                        "date": t.isoformat(), "body": (m.Body or "")[:8000],
                        "attachments": atts, "deja_traite": handled,
                    }
                    sl = subject.lower()
                    is_internal = (INTERNE_DOMAINE in (semail or "")) or is_demandeur(sender)
                    has_pdf = any(a.lower().endswith(".pdf") for a in atts)
                    has_devis_kw = any(k in sl for k in KW_DEVIS)
                    is_reply = subject.startswith(("RE:", "TR:", "FW:", "FWD:", "R:"))

                    if is_internal:
                        # Demande d'achat / besoin : le SUJET ne suffit pas
                        # (souvent un nom d'affaire : "DICOM - Hangar Etrac",
                        # "URGENT - ISHOP Saint-Denis"...) — l'expéditeur interne
                        # est le signal. RE:/TR: = fil en cours, à examiner
                        # (nouvelle demande dans le corps ?).
                        record["type"] = "fil" if is_reply else "demande"
                        demandes.append(record)
                    elif has_pdf and (has_devis_kw or is_reply or any(k in sl for k in KW_DEMANDE)):
                        # Réponse fournisseur (devis PDF, souvent RE: <notre objet>)
                        devis.append(record)

                    if eid:
                        processed.add(eid)
                except Exception:
                    pass
        except Exception:
            pass

    # Tri chronologique
    demandes.sort(key=lambda r: r["date"])
    devis.sort(key=lambda r: r["date"])

    with open(CANDIDATS, "w", encoding="utf-8") as f:
        json.dump(demandes, f, ensure_ascii=False, indent=1)
    with open(DEVIS, "w", encoding="utf-8") as f:
        json.dump(devis, f, ensure_ascii=False, indent=1)

    etat["last_scan"] = datetime.datetime.now(datetime.timezone.utc).isoformat()
    etat["processed"] = sorted(processed)[-5000:]
    with open(ETAT, "w", encoding="utf-8") as f:
        json.dump(etat, f, ensure_ascii=False, indent=1)

    print("Fenêtre analysée :", start.isoformat())
    print("Demandes candidates :", len(demandes), "->", CANDIDATS)
    print("Devis reçus candidats :", len(devis), "->", DEVIS)
    for r in demandes:
        print("  D |", r["date"][:16], "|", r["store"].split("@")[0], "|", r["sender"][:22], "|", r["subject"][:60], "| traité:", r["deja_traite"])
    for r in devis:
        print("  V |", r["date"][:16], "|", r["store"].split("@")[0], "|", r["sender"][:22], "|", r["subject"][:60])

if __name__ == "__main__":
    if "--hours" in sys.argv:
        idx = sys.argv.index("--hours")
        h = float(sys.argv[idx + 1])
        # Surcharge : on force la fenêtre en écrasant last_scan dans un état temporaire
        tmp = {"last_scan": (datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(hours=h)).isoformat(), "processed": []}
        with open(ETAT, "w", encoding="utf-8") as f:
            json.dump(tmp, f, ensure_ascii=False, indent=1)
    scan()
