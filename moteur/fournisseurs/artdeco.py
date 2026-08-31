import re

from moteur.modele import Article
from moteur.outils import to_float

# --- GABARIT (ART DECO — brand "LED'S RUN", showroom LED, domaine
# artdeco.re) ----------------------------------------------------------------
# Nouveau fournisseur (session TRAVAUX_PARSERS.md, expéditeur réel
# idriss@artdeco.re — voir CLAUDE.md). Un seul vrai devis vu à ce jour
# (devis 617004507, 2 lignes) : "ELECTRICITE SERVICES REUNION" qui figure
# dans le nom du fichier et dans le corps du PDF n'est PAS le fournisseur
# (c'est le nom légal de l'acheteuse elle-même, comme sur tous les autres
# devis de ce dépôt — voir "22 RUE PIERRE BROSSOLETTE") : le vrai vendeur
# identifié dans le document est "SHOWROOM LED S RUN" / "LED'S RUN",
# email contact@artdeco.re.
#
# Structure du texte extrait : chaque article est un bloc de 8 lignes à
# offsets fixes depuis son Id (identifiant numérique 3-6 chiffres, seul
# sur sa ligne), plus une 9e ligne de continuation de désignation
# OPTIONNELLE (la désignation déborde sur cette ligne, décalée après tout
# le bloc chiffré — comportement constaté identique sur les 2 lignes du
# seul devis disponible) :
#     Id                          (0)
#     Désignation                 (+1)
#     Pu HT                       (+2)
#     Qté                         (+3)
#     Rem.%                       (+4)
#     PU TTC                      (+5)
#     "Mt.HT Code_TVA" (collés)   (+6)
#     Mt. TTC                     (+7)
#     [suite désignation]         (+8, optionnel)
# Zone de tableau bornée par l'indice, PAS par l'en-tête de colonnes (qui
# apparaît, sur ce document, APRÈS les lignes d'articles plutôt qu'avant —
# artefact de mise en page de ce gabarit) : entre la phrase d'aide
# ("...clic sur l'ID produit...") et "Sous Total :".
MOTIF_ZONE_DEBUT = re.compile(r"clic sur l.ID produit", re.IGNORECASE)
MOTIF_ZONE_FIN = re.compile(r"^Sous Total\s*:?\s*$", re.IGNORECASE)
MOTIF_ID = re.compile(r"^\d{3,6}$")
MOTIF_MONTANT_CODE = re.compile(r"^([\d\s]+,\d{2})\s+(\S+)$")
MOTIF_DEVIS = r"DEVIS\s*N[°o]\s*(\d+)"
MOTIF_SOUS_TOTAL_HT = re.compile(r"Sous Total\s*:\s*\n\s*([\d\s]+,\d{2})")
# --- fin GABARIT -------------------------------------------------------------


def _f(v: str) -> float:
    return to_float(v.replace(" ", ""))


def _zone_tableau_artdeco(lignes: list[str]) -> list[str]:

    i_debut = next((i for i, l in enumerate(lignes) if MOTIF_ZONE_DEBUT.search(l)), None)
    if i_debut is None:
        return []

    i_fin = next(
        (i for i in range(i_debut + 1, len(lignes)) if MOTIF_ZONE_FIN.match(lignes[i].strip())),
        None,
    )

    return lignes[i_debut + 1:(i_fin if i_fin is not None else len(lignes))]


def parse_artdeco(texte: str) -> list[Article]:

    articles = []

    devis = ""
    m = re.search(MOTIF_DEVIS, texte, re.IGNORECASE)
    if m:
        devis = m.group(1)

    lignes = [l.strip() for l in texte.splitlines() if l.strip()]
    zone = _zone_tableau_artdeco(lignes)

    i = 0
    while i < len(zone):

        if not MOTIF_ID.match(zone[i]) or i + 7 >= len(zone):
            i += 1
            continue

        ref = zone[i]
        designation = zone[i + 1]
        m_montant = MOTIF_MONTANT_CODE.match(zone[i + 6])

        if m_montant is None:
            i += 1
            continue

        try:
            pu_ht = _f(zone[i + 2])
            quantite = _f(zone[i + 3])
            montant = _f(m_montant.group(1))
        except ValueError:
            i += 1
            continue

        if not quantite:
            i += 1
            continue

        pas = 8
        if i + 8 < len(zone) and not MOTIF_ID.match(zone[i + 8]):
            designation = f"{designation} {zone[i + 8]}".strip()
            pas = 9

        articles.append(
            Article(
                fournisseur="ART DECO",
                devis=devis,
                reference_fournisseur=ref,
                reference_distributeur="",
                designation=designation,
                quantite=quantite,
                unite="UN",
                prix_brut=pu_ht,
                prix_net=round(montant / quantite, 4),
                montant=montant,
            )
        )
        i += pas

    m = MOTIF_SOUS_TOTAL_HT.search(texte)
    if m:
        total_pdf = to_float(m.group(1))
        total_extrait = round(sum(a.montant for a in articles), 2)
        if abs(total_pdf - total_extrait) > 0.02:
            print(
                f"!! ART DECO : Sous Total HT du PDF ({total_pdf:.2f}€) "
                f"!= somme des lignes extraites ({total_extrait:.2f}€) "
                f"— une ligne a peut-être été oubliée ou mal lue."
            )

    return articles


# Déclaration pour le chargement automatique
FOURNISSEURS = ["ART DECO"]
parse = parse_artdeco
