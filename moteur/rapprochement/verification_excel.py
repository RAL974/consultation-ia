"""
Vérification d'un classeur PAR EXCEL LUI-MÊME (automation COM, pywin32),
toujours sur une COPIE — jamais sur le Suivi vivant (voir CLAUDE.md, P1 :
les contrôles « aucun message de réparation » et « les formules
calculent » sont faits par l'outil quand l'acheteur n'est pas là pour
ouvrir la copie).

Ce que fait `verifier_dans_excel()` :
1. ouvre le fichier dans une instance Excel invisible (DisplayAlerts
   désactivé : une réparation automatique serait SILENCIEUSE, d'où le
   point 2) ;
2. détecte une réparation par le journal qu'Excel écrit alors dans %TEMP%
   (error<n>_<n>.xml, « Removed Records / Repaired Records ») — tout
   journal créé pendant l'ouverture est remonté dans "reparation" ;
3. recalcule tout (CalculateFullRebuild) — équivaut au recalcul à
   l'ouverture que provoque fullCalcOnLoad ;
4. relève feuilles, tableaux structurés (nom, plage, nb de lignes) et les
   cellules demandées ;
5. enregistre éventuellement une copie RECALCULÉE (SaveAs xlsx) — c'est
   la « copie ouverte-fermée » dont les valeurs en cache se relisent
   ensuite avec openpyxl (contrôle des colonnes calculées, étape 5) ;
6. ferme sans enregistrer l'original, quitte Excel.

Aucune dépendance au reste du moteur : réutilisable tel quel pour M1/T1.
"""

import os
import time
from pathlib import Path

XL_OPENXML_WORKBOOK = 51


def _journaux_reparation(dossier_temp: Path) -> dict:
    return {p: p.stat().st_mtime for p in dossier_temp.glob("error*.xml")}


def verifier_dans_excel(chemin, recalculer=True, enregistrer_sous=None, cellules=(), lire_seul=True) -> dict:
    """Voir bandeau. `cellules` : liste de (feuille, adresse) à lire après
    recalcul. Retourne {"ouvert", "reparation": [journaux], "feuilles",
    "tableaux": {nom: {feuille, ref, lignes, colonnes}}, "cellules":
    {(feuille, adresse): valeur}, "erreur", "duree_s", "enregistre"}.
    Lève RuntimeError si pywin32/Excel sont indisponibles."""

    try:
        import pythoncom
        import win32com.client
    except ImportError as e:  # pragma: no cover - dépend du poste
        raise RuntimeError(f"pywin32 indisponible : {e}")

    chemin = Path(chemin).resolve()
    if not chemin.exists():
        raise FileNotFoundError(chemin)
    dossier_temp = Path(os.environ.get("TEMP", os.environ.get("TMP", ".")))
    avant = _journaux_reparation(dossier_temp)
    debut = time.time()

    resultat = {
        "ouvert": False, "reparation": [], "feuilles": [], "tableaux": {}, "cellules": {},
        "erreur": None, "duree_s": None, "enregistre": None,
    }

    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False
        wb = excel.Workbooks.Open(str(chemin), UpdateLinks=0, ReadOnly=bool(lire_seul))
        resultat["ouvert"] = True

        if recalculer:
            excel.CalculateFullRebuild()

        for ws in wb.Worksheets:
            resultat["feuilles"].append(ws.Name)
            for lo in ws.ListObjects:
                resultat["tableaux"][lo.Name] = {
                    "feuille": ws.Name,
                    "ref": str(lo.Range.Address).replace("$", ""),
                    "lignes": lo.ListRows.Count,
                    "colonnes": lo.ListColumns.Count,
                }

        for feuille, adresse in cellules:
            valeur = wb.Worksheets(feuille).Range(adresse).Value
            resultat["cellules"][(feuille, adresse)] = valeur

        if enregistrer_sous:
            cible = Path(enregistrer_sous).resolve()
            if cible.exists():
                cible.unlink()
            wb.SaveAs(str(cible), FileFormat=XL_OPENXML_WORKBOOK)
            resultat["enregistre"] = cible
    except Exception as e:  # remonté dans le résultat, jamais avalé
        resultat["erreur"] = repr(e)
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        wb = None
        excel = None
        pythoncom.CoUninitialize()

    apres = _journaux_reparation(dossier_temp)
    for p, mtime in apres.items():
        if p not in avant or mtime > avant[p]:
            if mtime >= debut - 1:
                try:
                    extrait = p.read_text(encoding="utf-8", errors="replace")[:800]
                except OSError:
                    extrait = "(illisible)"
                resultat["reparation"].append(f"{p} : {extrait}")

    resultat["duree_s"] = round(time.time() - debut, 1)
    return resultat


def resume_verification(resultat: dict) -> str:
    """Compte rendu court, lisible, d'un résultat de verifier_dans_excel()."""

    lignes = [
        f"Ouverture Excel : {'OK' if resultat['ouvert'] else 'ÉCHEC'}"
        + (f" — erreur : {resultat['erreur']}" if resultat["erreur"] else ""),
        f"Réparation détectée : {'OUI — ' + ' | '.join(resultat['reparation']) if resultat['reparation'] else 'aucune'}",
        f"Feuilles ({len(resultat['feuilles'])}) : {', '.join(resultat['feuilles'])}",
        f"Tableaux ({len(resultat['tableaux'])}) : "
        + ", ".join(f"{n} [{t['feuille']}!{t['ref']}, {t['lignes']} lignes]" for n, t in resultat["tableaux"].items()),
    ]
    for (feuille, adresse), valeur in resultat["cellules"].items():
        lignes.append(f"  {feuille}!{adresse} = {valeur!r}")
    if resultat.get("enregistre"):
        lignes.append(f"Copie recalculée : {resultat['enregistre']}")
    lignes.append(f"Durée : {resultat['duree_s']} s")
    return "\n".join(lignes)
