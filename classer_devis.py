# -*- coding: utf-8 -*-
"""
Phase 2 — classement des devis reçus dans Consultation AI.
Lit devis_recus.json, regroupe par affaire, crée consultations/<affaire>/devis/,
et télécharge les PDF depuis Outlook (SaveAsFile — lecture seule, aucun envoi).

Usage :
    python classer_devis.py                    # traite toutes les affaires >= MIN_DEVIS
    python classer_devis.py --min 3            # seuil de devis par affaire
    python classer_devis.py --affaires "Câbles R2V;GYSM"   # seulement ces affaires
"""
import json
import os
import re
import sys
import unicodedata
from collections import defaultdict
from pathlib import Path

import win32com.client

BASE = Path(r"X:\1.3. Logistique et approvisionnement\1.3.0. Commandes\Consultation AI")
DEVIS_JSON = BASE / "devis_recus.json"
CONSULTATIONS = BASE / "consultations"
MIN_DEVIS = 3

def norm_name(s: str) -> str:
    s = unicodedata.normalize("NFD", s or "")
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s).strip().lower()

def affaire(subject: str) -> str:
    """Nom d'affaire depuis un sujet : retire RE:/TR:/FW:/[Externe], garde le reste."""
    s = re.sub(r"^(re|tr|fw|fwd|r)\s*:\s*", "", subject or "", flags=re.I)
    s = re.sub(r"^\[externe\]\s*", "", s, flags=re.I)
    s = re.sub(r"\s+", " ", s).strip()
    # Supprime la mention "URGENT" du nom de dossier (gardée dans le sujet d'origine)
    return s

def clean_dirname(name: str, maxlen: int = 60) -> str:
    """Nom de dossier sûr (sans caractères interdits Windows)."""
    name = re.sub(r'[<>:"/\\|?*]', " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name[:maxlen]

def resolve_item(ns, store_name, entry_id):
    """Résout un mail par EntryID (global au store) — GetItemFromID n'est
    exposé que sur le Namespace en dispatch dynamique pywin32."""
    try:
        return ns.GetItemFromID(entry_id)
    except Exception:
        return None

def download_devis(items_by_affaire):
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    report = []
    for aff, records in sorted(items_by_affaire.items()):
        doss = CONSULTATIONS / clean_dirname(aff)
        devis_dir = doss / "devis"
        devis_dir.mkdir(parents=True, exist_ok=True)
        saved = []
        for r in records:
            item = resolve_item(ns, r["store"], r["entry_id"])
            if item is None:
                report.append(f"  !! item introuvable: {r['subject'][:40]}")
                continue
            try:
                for a in item.Attachments:
                    fn = a.FileName or ""
                    if not fn.lower().endswith(".pdf"):
                        continue
                    dest = devis_dir / fn
                    # Évite les doublons de nom entre fournisseurs
                    if dest.exists():
                        stem, suffix = dest.stem, dest.suffix
                        n = 2
                        while dest.exists():
                            dest = devis_dir / f"{stem} ({n}){suffix}"
                            n += 1
                    a.SaveAsFile(str(dest))
                    saved.append(fn)
            except Exception as e:
                report.append(f"  !! échec PJ {r['subject'][:30]}: {e}")
        report.append(f"{aff} -> {len(saved)} PDF dans {devis_dir.relative_to(BASE)}")
    return report

def main():
    data = json.load(open(DEVIS_JSON, encoding="utf-8"))
    min_devis = MIN_DEVIS
    only = None
    if "--min" in sys.argv:
        min_devis = int(sys.argv[sys.argv.index("--min") + 1])
    if "--affaires" in sys.argv:
        only = [a.strip() for a in sys.argv[sys.argv.index("--affaires") + 1].split(";")]

    groups = defaultdict(list)
    for r in data:
        pdfs = [a for a in r["attachments"] if a.lower().endswith(".pdf")]
        if not pdfs:
            continue
        key = norm_name(affaire(r["subject"]))
        groups[key].append(r)

    selected = {}
    for key, records in groups.items():
        if only is not None and key not in [norm_name(a) for a in only]:
            continue
        if only is None and len(records) < min_devis:
            continue
        # Nom d'affaire lisible : le plus long sujet (le plus informatif)
        display = max((affaire(r["subject"]) for r in records), key=len)
        selected[display] = records

    print(f"{len(selected)} affaire(s) retenue(s) :")
    for aff, records in sorted(selected.items(), key=lambda kv: -len(kv[1])):
        print(f"  [{len(records)} devis] {aff}")
        for r in records:
            pdfs = [a for a in r["attachments"] if a.lower().endswith(".pdf")]
            print(f"      {r['date'][:16]} {r['sender'][:22]:22s} {r['sender_email'][:30]:30s} {pdfs}")

    # Sauvegarde du mapping affaire -> records pour l'étape Besoin
    out = {}
    for aff, records in selected.items():
        out[aff] = [
            {
                "date": r["date"], "sender": r["sender"],
                "sender_email": r["sender_email"], "subject": r["subject"],
                "store": r["store"], "entry_id": r["entry_id"],
                "pdfs": [a for a in r["attachments"] if a.lower().endswith(".pdf")],
                "body": r["body"],
            }
            for r in records
        ]
    with open(BASE / "affaires_devis.json", "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)

    print("\nTéléchargement des PDF...")
    for line in download_devis(selected):
        print(line)
    print("\nMapping affaires -> devis sauvegardé dans affaires_devis.json")

if __name__ == "__main__":
    main()
