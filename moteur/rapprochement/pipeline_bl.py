"""
Orchestration du rapprochement BL — commune à un futur CLI et au bouton GUI
(même principe que moteur/pipeline.py côté devis). Séparée en deux temps,
volontairement :

1. `rapprocher_dossier()` — LECTURE SEULE : OCR + matching contre le Suivi,
   ne modifie jamais rien. C'est ce rapport que l'utilisateur doit voir et
   trier (cocher/décocher) avant toute écriture (voir CLAUDE.md,
   "Rapprochement AI" — mode simulation par défaut).
2. `appliquer_et_archiver()` — écrit (via moteur/rapprochement/ecriture.py,
   qui gère verrou/sauvegarde/patch chirurgical), PUIS range chaque BL lu :
   - entièrement résolu (toutes ses lignes déjà à jour ou tout juste
     écrites, ou bon de retour — voir _est_resolu) -> a_traiter/BL/Traités/
     <n° de commande>/, avec une copie du bon de commande correspondant
     s'il est trouvé dans l'archive des BC (voir trouver_bon_de_commande,
     demande explicite de l'acheteur : "ainsi nous aurons tout le flux
     commande-BL-facture facilement consultable") ;
   - au moins une ligne "à confirmer" non cochée ou "inconnue" (ex. une
     référence différente entre le BL et le Suivi) -> a_traiter/BL/À
     vérifier/ (reste À PLAT, pas de sous-dossier par commande — ce sont
     des BL qui attendent encore une décision humaine, pas un flux
     consultable), JAMAIS mélangé avec les BL pas encore traités, pour
     repérer facilement ceux qui attendent une décision humaine.
   Puis écrit un rapport dans rapports/.
"""

import re
import shutil
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import fitz

from moteur.referentiel import Referentiel
from moteur.rapprochement.ecriture import Ecriture, appliquer
from moteur.rapprochement.lecture_bl import analyser_dossier
from moteur.rapprochement.pieces import (
    COMMENTAIRE_MONTANT_RECALCULE,
    MODE_AUTO,
    MODE_CONFIRME,
    TYPE_BL,
    FeuillePiecesAbsente,
    dedoublonner_ids,
    ecrire_pieces,
    feuille_pieces_presente,
    nouvelle_piece,
)
from moteur.rapprochement.matching import (
    Correspondance,
    Statut,
    apparier,
    deduire_commande_par_contenu,
    lire_lignes_commande,
    lire_lignes_fournisseur,
)

DOSSIER_A_TRAITER_BL = "a_traiter/BL"
DOSSIER_TRAITES = "Traités"  # sous-dossier de a_traiter/BL/ (créé par l'acheteur) —
# elle y récupère les BL numérisés déjà rapprochés pour agrafer BdC+BL papier
# et archiver dans les classeurs physiques (demande explicite, session R2 suite).
DOSSIER_A_VERIFIER = "À vérifier"  # sous-dossier de a_traiter/BL/ : BL LUS mais avec
# au moins une ligne inconnue ou "à confirmer" non cochée (ex. référence
# différente entre BL et Suivi, substitution fournisseur...) — jamais
# mélangé avec les BL pas encore traités du tout, pour repérer facilement
# ceux qui attendent une décision humaine (demande explicite, session R2 suite).
DOSSIER_RAPPORTS = "rapports"
DOSSIER_BACKUPS = "backups"
DOSSIER_REFERENTIEL = "referentiel"
NOM_A_CONFIRMER_BL = "A_confirmer_BL.xlsx"  # fichier À PART de A_confirmer.xlsx
# (côté devis, moteur/pipeline.py) — même referentiel/articles.db partagé
# (un alias confirmé vaut pour les deux flux), mais chacun régénère SA
# PROPRE file d'attente de propositions à chaque exécution ; les mélanger
# écraserait les propositions de l'autre flux (voir moteur/referentiel.py,
# Referentiel.ecrire_a_confirmer).

# Nom du dossier voisin (frère du dossier projet) où vit le VRAI Suivi
# commandes — voir CLAUDE.md, "Fichier vivant du Suivi commandes" : PAS le
# même fichier que celui trouvé par moteur.panier.trouver_fichier_suivi() à
# la racine du dépôt, qui n'est qu'un export ponctuel potentiellement
# périmé (utilisé par panier.py juste pour caler l'ordre des colonnes).
DOSSIER_COMMANDES_COURANTES = "1.3.0.1. Commandes courantes"

# Archive des bons de commande eux-mêmes (fichier fourni par l'acheteur,
# demande explicite : "dans traités, il faudra créer un dossier pour chaque
# commande [...] dedans on y met ce bon de commande, tous les BL et bons de
# retours associés"). Arborescence MIXTE constatée réellement sous ce
# dossier : un sous-dossier par année ("2026/", lui-même mixte — fichiers en
# vrac ET sous-dossiers par chantier), ET "BdCPDF/" (nouveau dossier créé
# par l'acheteur pour la génération automatique des BdC récents, à plat —
# précision explicite de l'acheteur, session suivante). Toujours le même
# motif de nom "<Chantier> - BC <numéro> - <fournisseur>.<pdf|xlsx>" partout
# — recherche donc récursive (rglob) depuis LA RACINE "Commandes/" (pas
# narrowée à un sous-dossier particulier), pour couvrir ces deux emplacements
# ET tout futur sous-dossier de la même farine sans jamais coder son nom en
# dur.
DOSSIER_COMMANDES_ARCHIVE = "Commandes"


def trouver_dossier_commandes(dossier_projet) -> Path | None:
    """Racine de l'archive des bons de commande — voir
    DOSSIER_COMMANDES_ARCHIVE (couvre à la fois l'archive historique par
    année et le nouveau dossier "BdCPDF/" de génération automatique, tous
    deux ses sous-dossiers directs). None si absent (ex. poste sans accès
    au dossier réseau), jamais une erreur bloquante : la copie du BC est un
    plus, pas une condition de l'archivage du BL."""

    dossier = Path(dossier_projet).parent / DOSSIER_COMMANDES_COURANTES / DOSSIER_COMMANDES_ARCHIVE
    return dossier if dossier.is_dir() else None


def trouver_bon_de_commande(numero_commande: str, dossier_commandes) -> Path | None:
    """Cherche, dans TOUTE l'arborescence de `dossier_commandes` (plate +
    sous-dossiers par chantier), le fichier "... - BC <numero_commande> -
    <fournisseur>.<pdf|xlsx>" correspondant à ce numéro. Ne retourne un
    résultat que s'il y a EXACTEMENT UN candidat DE CONTENU DISTINCT —
    jamais un choix au hasard entre plusieurs BC réellement ambigus (règle
    d'or du projet) ; None aussi si `dossier_commandes` est introuvable
    (voir trouver_dossier_commandes) ou si `numero_commande` est vide.

    BUG RÉEL CORRIGÉ (signalé par l'acheteur, commande M3.23.033) : le même
    BC se retrouve couramment archivé à LA FOIS dans "Commandes/<année>/"
    ET dans "Commandes/BdCPDF/" (le second étant filé dans le premier après
    coup) — deux fichiers de MÊME NOM ET MÊME TAILLE, donc pas une vraie
    ambiguïté (choisir entre deux copies identiques n'est pas "deviner").
    Dédoublonnés par (nom, taille) avant d'exiger l'unicité ; deux BC de
    même numéro mais de contenu VRAIMENT différent (nom ou taille distincts)
    restent, eux, traités comme ambigus — aucun résultat renvoyé."""

    if not numero_commande or dossier_commandes is None:
        return None

    dossier_commandes = Path(dossier_commandes)
    if not dossier_commandes.is_dir():
        return None

    motif = re.compile(rf"BC\s*{re.escape(numero_commande)}\b", re.IGNORECASE)
    candidats = [
        f for f in dossier_commandes.rglob("*")
        if f.is_file() and f.suffix.lower() in (".pdf", ".xlsx") and motif.search(f.stem)
    ]

    candidats_distincts = list({(f.name, f.stat().st_size): f for f in candidats}.values())

    return candidats_distincts[0] if len(candidats_distincts) == 1 else None


def trouver_fichier_suivi_vivant(dossier_projet):
    """Le classeur RÉELLEMENT ouvert/modifié par l'acheteur au quotidien
    (dossier frère du projet, voir DOSSIER_COMMANDES_COURANTES) — jamais la
    copie potentiellement périmée à la racine du dépôt. Exclut les copies
    manuelles ("copie ...xlsx") et le classeur en refonte ("Suivi
    nouveau..."), qui matcheraient sinon le même motif de nom."""

    dossier = Path(dossier_projet).parent / DOSSIER_COMMANDES_COURANTES
    if not dossier.is_dir():
        return None

    candidats = [
        f for f in dossier.glob("*Suivi commandes*.xlsx")
        if not f.name.startswith("~$") and not f.name.lower().startswith("copie")
    ]
    if not candidats:
        return None

    return max(candidats, key=lambda f: f.stat().st_mtime)


@dataclass
class RapportRapprochement:

    surs: list = field(default_factory=list)          # [(BonLivraison, Correspondance)]
    a_confirmer: list = field(default_factory=list)
    deja_a_jour: list = field(default_factory=list)
    inconnus: list = field(default_factory=list)
    anomalies_lecture: list = field(default_factory=list)  # [(nom_fichier, raison)] — échec de lecture d'UNE PARTIE du fichier (fournisseur non reconnu/sans parser sur une page, voir détection par page, lecture_bl.lire_bl) ; peut coexister avec des BL résolus pour le MÊME fichier depuis ses autres pages — voir appliquer_et_archiver
    anomalies_bl: list = field(default_factory=list)       # [(BonLivraison, raison)] — BL lu mais pas rapprochable (n° de commande introuvable, commande absente du Suivi...), voir "archivage par BL individuel"
    fichier_suivi: Path | None = None


def rapprocher_dossier(dossier_a_traiter, dossier_projet) -> RapportRapprochement:
    """Lecture seule : lit tous les BL du dossier, les rapproche du Suivi.
    Ne modifie ni le Suivi ni les fichiers de `dossier_a_traiter`.

    Ouvre son PROPRE référentiel articles (moteur/referentiel.py, même
    base articles.db que côté devis — demande explicite de l'acheteur,
    voir bandeau de moteur.rapprochement.matching : "il faut créer une
    base des équivalences, ce genre de cas va se présenter très souvent",
    cas réel 59210/CFF1BIS) : un alias déjà CONFIRMÉ (par une exécution
    précédente, devis OU BL) compte comme une correspondance exacte pour
    apparier() ; une nouvelle proposition est écrite dans
    referentiel/A_confirmer_BL.xlsx (fichier À PART de celui du devis,
    voir DOSSIER_REFERENTIEL/NOM_A_CONFIRMER_BL) pour confirmation."""

    dossier_projet = Path(dossier_projet)
    fichier_suivi = trouver_fichier_suivi_vivant(dossier_projet)

    bons_de_livraison, anomalies = analyser_dossier(dossier_a_traiter)

    rapport = RapportRapprochement(anomalies_lecture=list(anomalies), fichier_suivi=fichier_suivi)

    if fichier_suivi is None:
        for bl in bons_de_livraison:
            rapport.anomalies_bl.append((bl, "Suivi commandes introuvable à la racine du projet"))
        return rapport

    dossier_referentiel = dossier_projet / DOSSIER_REFERENTIEL
    referentiel = Referentiel(dossier_projet / "moteur")
    referentiel.importer_bdd(dossier_projet / "base" / "BDD_articles.csv")
    referentiel.importer_equivalences_bl(dossier_referentiel / "equivalences_bl.csv")
    referentiel.appliquer_confirmations(dossier_referentiel / NOM_A_CONFIRMER_BL)

    # Un "bon de retour" (109 Distribution, voir moteur.fournisseurs.
    # dist109) N'EST JAMAIS une livraison : il ANNULE une ligne d'un BL
    # précédent (cas réel signalé par l'acheteur — article listé sur un BL
    # mais pas coché "livré" à réception, le fournisseur envoie un retour
    # qui référence ce BL). Avant le traitement normal, on repère toutes
    # les références ainsi annulées (par n° de BL d'origine) pour ne
    # JAMAIS les proposer à l'écriture quand ce BL d'origine est traité
    # plus bas dans la même passe — sans quoi la quantité "non livrée"
    # aurait été comptée comme une vraie livraison (bug réel trouvé et
    # corrigé avant toute écriture, voir CLAUDE.md).
    references_annulees_par_bl = {}
    for bl in bons_de_livraison:
        if bl.type_document == "RETOUR":
            for ligne in bl.lignes:
                references_annulees_par_bl.setdefault(bl.numero_bl_origine, set()).add(
                    ligne.reference_fournisseur.strip().upper()
                )

    for bl in bons_de_livraison:

        if bl.type_document == "RETOUR":
            # Jamais une livraison en soi : rien n'est jamais écrit à
            # partir d'un retour lui-même — seule la ligne qu'il annule
            # (sur le BL d'origine référencé) est exclue plus bas.
            refs = ", ".join(l.reference_fournisseur for l in bl.lignes) or "?"
            rapport.anomalies_bl.append((
                bl,
                f"Bon de retour — annule {refs} du BL {bl.numero_bl_origine or '(numéro introuvable)'} : "
                "rien à écrire depuis ce document, vérifier que le BL d'origine n'a pas compté cette ligne à tort",
            ))
            continue

        # Ces trois anomalies sont rattachées au BL (pas au fichier) : un
        # fichier Cominter avec plusieurs BL peut avoir UN bl sans commande
        # reconnue à côté d'autres parfaitement résolus — jamais bloquer
        # tout le fichier pour ça (voir appliquer_et_archiver, archivage
        # par BL individuel).
        commande_deduite, score_deduction = False, 0

        if not bl.numero_commande:
            # Cas réel signalé par l'acheteur : les gars de l'atelier
            # perforent les BL pour les classer, ce qui abîme parfois la
            # zone où le n° de commande est imprimé — avant de renoncer,
            # tente de le DÉDUIRE en comparant le contenu du BL (référence
            # + quantité) à TOUTES les commandes de ce fournisseur dans le
            # Suivi (voir matching.deduire_commande_par_contenu, jamais un
            # choix au hasard : au moins 2 lignes concordantes, score
            # strictement meilleur que tout autre candidat).
            try:
                lignes_fournisseur = lire_lignes_fournisseur(fichier_suivi, bl.fournisseur)
            except Exception as e:
                rapport.anomalies_bl.append((bl, f"Erreur de lecture du Suivi ({e})"))
                continue

            numero_deduit, score_deduction = deduire_commande_par_contenu(bl.lignes, lignes_fournisseur)

            if numero_deduit is None:
                rapport.anomalies_bl.append((bl, "N° de commande introuvable sur le BL"))
                continue

            bl.numero_commande = numero_deduit
            commande_deduite = True

        try:
            lignes_suivi = lire_lignes_commande(fichier_suivi, bl.fournisseur, bl.numero_commande)
        except Exception as e:
            rapport.anomalies_bl.append((bl, f"Erreur de lecture du Suivi ({e})"))
            continue

        if not lignes_suivi:
            rapport.anomalies_bl.append((
                bl,
                f"Commande {bl.numero_commande} introuvable dans le Suivi pour « {bl.fournisseur} »",
            ))
            continue

        references_annulees = references_annulees_par_bl.get(bl.numero_bl, set())

        for c in apparier(
            bl.lignes, lignes_suivi, date_bl_reelle=_parser_date_bl(bl.date_bl),
            referentiel=referentiel, fournisseur=bl.fournisseur, devis=bl.numero_bl,
        ):

            if c.ligne_bl.reference_fournisseur.strip().upper() in references_annulees:
                # Cette ligne précise a été ANNULÉE par un bon de retour
                # associé à ce même BL (voir plus haut) — jamais écrite,
                # même si _comparer() l'aurait autrement jugée "sûre".
                c = Correspondance(
                    c.ligne_bl, c.ligne_suivi, Statut.A_CONFIRMER,
                    [
                        f"Cette ligne a été ANNULÉE par un bon de retour associé au BL {bl.numero_bl} "
                        "— ne pas écrire, la vraie livraison doit venir d'un autre document"
                    ] + c.raisons,
                    qte_deja_incluse=c.qte_deja_incluse,
                )

            if commande_deduite:
                # Un n° de commande DÉDUIT (pas lu directement) n'est
                # jamais assez sûr pour un rapprochement "sûr" automatique
                # — la ligne bascule "à confirmer" avec la déduction
                # expliquée en clair, même si _comparer() l'aurait
                # autrement jugée "sûre".
                raison_deduction = (
                    f"N° de commande « {bl.numero_commande} » déduit automatiquement du contenu du BL "
                    f"({score_deduction} référence(s)/quantité(s) concordantes — illisible sur le document) "
                    "— à vérifier"
                )
                statut = Statut.A_CONFIRMER if c.statut is Statut.SUR else c.statut
                c = Correspondance(
                    c.ligne_bl, c.ligne_suivi, statut, [raison_deduction] + c.raisons,
                    qte_deja_incluse=c.qte_deja_incluse,
                )

            if c.statut is Statut.SUR:
                rapport.surs.append((bl, c))
            elif c.statut is Statut.A_CONFIRMER:
                rapport.a_confirmer.append((bl, c))
            elif c.statut is Statut.DEJA_A_JOUR:
                rapport.deja_a_jour.append((bl, c))
            else:
                rapport.inconnus.append((bl, c))

    _desamorcer_conflits_meme_ligne_suivi(rapport)

    referentiel.ecrire_a_confirmer(dossier_referentiel, NOM_A_CONFIRMER_BL)
    referentiel.fermer()

    return rapport


def _desamorcer_conflits_meme_ligne_suivi(rapport: RapportRapprochement) -> None:
    """Si DEUX fichiers DIFFÉRENTS du même lot proposent chacun une
    correspondance "sûre" vers la MÊME ligne du Suivi, ce n'est jamais
    anodin : soit le même BL déposé deux fois par erreur (cas réel,
    session R2 suite — un vieux fichier et un nouveau scan du même BL
    Cominter Ouest cohabitaient dans a_traiter/BL/), soit une livraison
    fractionnée mal enchaînée. Dans les deux cas, empiler deux "+X"
    indépendants sur la MÊME base (qte_livree_cumulee calculée séparément
    par apparier(), sans connaissance de l'autre fichier) corromprait le
    cumul si les deux étaient écrits. Seule la 1ère proposition rencontrée
    reste "sûre" ; toute autre visant la même ligne bascule "à confirmer"."""

    lignes_vues = {}
    surs_filtres = []

    for bl, c in rapport.surs:

        cle = c.ligne_suivi.ligne_excel
        premier_fichier = lignes_vues.get(cle)

        if premier_fichier is not None and premier_fichier != bl.fichier:
            rapport.a_confirmer.append((bl, Correspondance(
                c.ligne_bl, c.ligne_suivi, Statut.A_CONFIRMER,
                [
                    f"Un autre BL de ce même lot ({premier_fichier}) cible aussi cette "
                    f"ligne du Suivi — vérifier qu'il ne s'agit pas du même BL déposé "
                    f"deux fois avant de confirmer"
                ],
            )))
            continue

        lignes_vues.setdefault(cle, bl.fichier)
        surs_filtres.append((bl, c))

    rapport.surs = surs_filtres


def regrouper_par_bl(rapport: RapportRapprochement) -> dict:
    """{id(bl): {"bl": BonLivraison, "sur": [...], "a_confirmer": [...],
    "deja_a_jour": [...], "inconnu": [...], "anomalies": [raison, ...]}} —
    CLÉ PAR OBJET BL (id()), pas par nom de fichier : un même fichier peut
    contenir PLUSIEURS BonLivraison (Cominter Ouest, voir
    moteur.ocr.pages_par_identifiant) — les grouper par nom de fichier
    fusionnerait à tort leurs lignes et empêcherait de savoir si UN SEUL
    d'entre eux est résolu (voir CLAUDE.md, "archivage par BL
    individuel")."""

    groupes = {}

    def _groupe(bl):
        return groupes.setdefault(id(bl), {
            "bl": bl, "sur": [], "a_confirmer": [], "deja_a_jour": [], "inconnu": [], "anomalies": [],
        })

    for bl, c in rapport.surs:
        _groupe(bl)["sur"].append(c)
    for bl, c in rapport.a_confirmer:
        _groupe(bl)["a_confirmer"].append(c)
    for bl, c in rapport.deja_a_jour:
        _groupe(bl)["deja_a_jour"].append(c)
    for bl, c in rapport.inconnus:
        _groupe(bl)["inconnu"].append(c)
    for bl, raison in rapport.anomalies_bl:
        _groupe(bl)["anomalies"].append(raison)

    return groupes


def _parser_date_bl(date_bl: str):
    try:
        return datetime.strptime((date_bl or "").strip(), "%d/%m/%Y").date()
    except ValueError:
        return None


def ecritures_pour(correspondances) -> list[Ecriture]:
    """Construit les Ecriture (voir moteur/rapprochement/ecriture.py) pour
    une liste de (BonLivraison, Correspondance) déjà décidées "à écrire"
    (sûres, ou "à confirmer" cochées par l'acheteur)."""

    ecritures = []

    for bl, c in correspondances:

        # Garde-fou (défense en profondeur) : une correspondance "déjà à
        # jour" n'a RIEN à écrire par construction (voir Statut.DEJA_A_JOUR)
        # — bug réel rencontré (recette, repli référence proche) où une
        # correspondance "déjà à jour" avait failli être traitée comme une
        # livraison fraîche à cumuler, doublant une quantité déjà exacte
        # dans le VRAI classeur. Ignorée ici même si un futur appelant se
        # trompait de liste.
        if c.statut is Statut.DEJA_A_JOUR:
            continue

        ligne = c.ligne_suivi.ligne_excel

        ecritures.append(Ecriture(ligne, "Qté livrée", c.qte_livree_cumulee))

        if c.ligne_bl.prix_net:
            ecritures.append(Ecriture(ligne, "Tarif BL", c.ligne_bl.prix_net))

        d = _parser_date_bl(bl.date_bl)
        ecritures.append(Ecriture(ligne, "Date de livraison", d or date.today()))

    return ecritures


def pieces_pour_bl(correspondances, dossier_traites) -> list:
    """Depuis P1 (feuille Pièces) : UNE ligne de type BL par ligne
    rapprochée écrite (voir ecritures_pour — « Qté livrée » reste écrite en
    cumul dans Commandes comme avant, Pièces garde le détail document par
    document). Qté = quantité livrée par CE BL, PU HT = prix net du BL,
    Montant HT = montant imprimé (sinon Qté × PU, Commentaire « montant
    recalculé »), Chantier/Sous-Chantier/Référence Suivi copiés de la
    ligne Commandes, Fichier = chemin d'archive du BL
    (Traités/<commande>/<nom d'archive>, voir _nom_archive_bl — le même
    nom que produira l'archivage qui suit l'écriture). Une correspondance
    « déjà à jour » ne produit rien."""

    pieces = []
    for bl, c in correspondances:
        if c.statut is Statut.DEJA_A_JOUR or c.ligne_suivi is None:
            continue
        lb, ls = c.ligne_bl, c.ligne_suivi
        montant, commentaire = lb.montant, None
        if montant is None and lb.prix_net:
            montant = round(lb.quantite_livree * lb.prix_net, 2)
            commentaire = COMMENTAIRE_MONTANT_RECALCULE
        suffixe = Path(bl.fichier or "").suffix or ".pdf"
        chemin = _dossier_pour_commande(Path(dossier_traites), bl.numero_commande) / f"{_nom_archive_bl(bl)}{suffixe}"
        pieces.append(nouvelle_piece(
            TYPE_BL, bl.fournisseur, bl.numero_bl, _parser_date_bl(bl.date_bl) or date.today(),
            bl.numero_commande, ls.chantier, ls.sous_chantier, ls.reference,
            lb.reference_fournisseur, lb.designation, lb.quantite_livree, lb.prix_net, montant,
            mode=MODE_AUTO if c.statut is Statut.SUR else MODE_CONFIRME, fichier=chemin, commentaire=commentaire,
        ))
    return dedoublonner_ids(pieces)


def _nom_archive_bl(bl) -> str:
    """"<date> - <fournisseur> - <n° BL> - BC <n° commande>" — pour
    s'y retrouver sans avoir à ouvrir chaque fichier. Partagé par
    archiver_bl() (fichier entier, un seul BL par fichier) et
    _extraire_bl_vers() (BL extrait par découpage de pages, fichier avec
    plusieurs BL). Un bon de retour (voir moteur.fournisseurs.dist109) est
    marqué "RETOUR" (au lieu du nom du fournisseur) avec le BL qu'il annule
    entre parenthèses, pour le repérer sans avoir à l'ouvrir.

    BUG RÉEL CORRIGÉ (nouveau fournisseur YESSS, n° de BL "CAM/040759") :
    numero_bl/numero_commande sont passés par _sans_caracteres_interdits()
    (voir _nom_dossier_commande) avant d'entrer dans le nom de fichier —
    sans ça, le "/" cassait l'archivage (Windows l'interprète comme un
    séparateur de dossier)."""

    jour = _parser_date_bl(bl.date_bl) or date.today()
    numero_bl = _sans_caracteres_interdits(bl.numero_bl or "sans-numero", "-")
    numero_commande = _sans_caracteres_interdits(bl.numero_commande or "inconnue", "-")

    if bl.type_document == "RETOUR":
        numero_bl_origine = _sans_caracteres_interdits(bl.numero_bl_origine, "-") if bl.numero_bl_origine else ""
        origine = f" (annule BL {numero_bl_origine})" if numero_bl_origine else ""
        return (
            f"{jour.isoformat()} - RETOUR - "
            f"{numero_bl}{origine} - BC {numero_commande}"
        )

    return (
        f"{jour.isoformat()} - {bl.fournisseur} - "
        f"{numero_bl} - BC {numero_commande}"
    )


MOTIF_CARACTERES_INTERDITS_NOM_FICHIER = re.compile(r'[<>:"/\\|?*]')


def _sans_caracteres_interdits(texte: str, remplacement: str = "_") -> str:
    """Remplace les caractères interdits dans un nom de fichier/dossier
    Windows (`< > : " / \\ | ? *`) — BUG RÉEL CORRIGÉ (nouveau fournisseur
    YESSS, dont le n° de BL contient un "/" imprimé, ex. "CAM/040759") :
    sans ce nettoyage, `_nom_archive_bl()` produisait un nom contenant un
    "/" que Windows interprète comme un séparateur de dossier, faisant
    échouer l'archivage entier avec une erreur "chemin d'accès
    introuvable" (le Suivi avait pourtant déjà été écrit avec succès à ce
    stade). Partagé avec `_nom_dossier_commande()` (déjà défensif pour le
    même risque, jamais rencontré en pratique jusqu'ici) pour ne pas avoir
    deux logiques de nettoyage légèrement différentes."""

    return MOTIF_CARACTERES_INTERDITS_NOM_FICHIER.sub(remplacement, texte)


def _nom_dossier_commande(numero_commande: str) -> str:
    """Nom de dossier sûr sur le système de fichiers — les n° de commande
    réels ("M3.10.175", "142.033"...) ne contiennent jamais de caractère
    interdit sous Windows, mais on ne prend pas ce risque pour un cas
    futur."""

    nom = (numero_commande or "").strip() or "Commande inconnue"
    return _sans_caracteres_interdits(nom)


def _dossier_pour_commande(dossier_traites: Path, numero_commande: str) -> Path:
    return Path(dossier_traites) / _nom_dossier_commande(numero_commande)


def _copier_bon_de_commande_si_absent(dossier_commande_cible: Path, numero_commande: str,
                                       dossier_commandes_bc) -> Path | None:
    """Copie le bon de commande (voir trouver_bon_de_commande) dans le
    dossier de la commande fraîchement créé/complété — une seule fois
    (idempotent : si un fichier "BC - ..." y est déjà, ne recherche/copie
    rien de plus). Ne casse jamais l'archivage du BL lui-même si le BC est
    introuvable ou ambigu (retourne simplement None)."""

    dossier_commande_cible = Path(dossier_commande_cible)

    if any(dossier_commande_cible.glob("BC - *")):
        return None

    bc = trouver_bon_de_commande(numero_commande, dossier_commandes_bc)
    if bc is None:
        return None

    cible = dossier_commande_cible / f"BC - {bc.name}"
    shutil.copy2(bc, cible)
    return cible


def archiver_bl(chemin_source: Path, bl, dossier_traites: Path, dossier_commandes_bc=None) -> Path:
    """Déplace le PDF/image source ENTIER vers a_traiter/BL/Traités/<n° de
    commande>/ (un sous-dossier par commande, demande explicite de
    l'acheteur : "ainsi nous aurons tout le flux commande-BL-facture
    facilement consultable, et nous pourrons très facilement repérer les
    écarts de facturation") — renommé via _nom_archive_bl(). Copie aussi le
    bon de commande dans ce dossier s'il est trouvé et pas déjà présent
    (voir _copier_bon_de_commande_si_absent). Réservé aux fichiers à UN SEUL
    BL (voir _traiter_bl_multiples_du_fichier pour les fichiers Cominter à
    plusieurs BL, qui doivent être découpés plutôt que déplacés en bloc)."""

    chemin_source = Path(chemin_source)

    dossier = _dossier_pour_commande(dossier_traites, bl.numero_commande)
    dossier.mkdir(parents=True, exist_ok=True)

    base = _nom_archive_bl(bl)
    cible = dossier / f"{base}{chemin_source.suffix}"
    compteur = 1
    while cible.exists():
        compteur += 1
        cible = dossier / f"{base} ({compteur}){chemin_source.suffix}"

    shutil.move(str(chemin_source), str(cible))
    _copier_bon_de_commande_si_absent(dossier, bl.numero_commande, dossier_commandes_bc)
    return cible


def _nombre_pages(chemin_source: Path) -> int:
    with fitz.open(chemin_source) as doc:
        return doc.page_count


def _extraire_bl_vers(chemin_source: Path, indices_pages: list[int], dossier_dest: Path, nom_base: str) -> Path:
    """Extrait un sous-ensemble de PAGES (0-based) de `chemin_source` vers
    un nouveau PDF dans `dossier_dest`, nommé `nom_base` — pour archiver ou
    déplacer UN SEUL BL d'un fichier qui en contient plusieurs, sans
    attendre que ses frères du même fichier soient, eux aussi, résolus
    (voir CLAUDE.md, "archivage par BL individuel"). Ne modifie jamais
    `chemin_source` lui-même (voir _reecrire_avec_pages, appelé séparément
    une fois tous les BL du fichier traités)."""

    dossier_dest = Path(dossier_dest)
    dossier_dest.mkdir(parents=True, exist_ok=True)

    cible = dossier_dest / f"{nom_base}.pdf"
    compteur = 1
    while cible.exists():
        compteur += 1
        cible = dossier_dest / f"{nom_base} ({compteur}).pdf"

    with fitz.open(chemin_source) as doc:
        nouveau = fitz.open()
        for i in indices_pages:
            nouveau.insert_pdf(doc, from_page=i, to_page=i)
        nouveau.save(cible)
        nouveau.close()

    return cible


def _reecrire_avec_pages(chemin_source: Path, indices_pages: list[int]) -> None:
    """Réécrit `chemin_source` pour ne garder QUE les pages listées (0-based)
    — appelé une fois que d'autres BL du même fichier ont été extraits vers
    Traités/À vérifier, pour que le fichier restant dans a_traiter/BL/ ne
    contienne plus que les BL pas encore résolus (sinon il serait relu en
    entier, avec ses BL déjà archivés, à chaque exécution suivante —
    inutile, même si sans risque grâce à l'idempotence du matching)."""

    chemin_source = Path(chemin_source)
    chemin_tmp = chemin_source.with_name(chemin_source.name + ".tmp")

    with fitz.open(chemin_source) as doc:
        nouveau = fitz.open()
        for i in indices_pages:
            nouveau.insert_pdf(doc, from_page=i, to_page=i)
        nouveau.save(chemin_tmp)
        nouveau.close()

    chemin_tmp.replace(chemin_source)


def deplacer_vers_a_verifier(chemin_source: Path, dossier_a_verifier: Path) -> Path:
    """Déplace un BL LU mais pas entièrement résolu (ligne inconnue, ou "à
    confirmer" non cochée) vers a_traiter/BL/À vérifier/ — nom de fichier
    conservé tel quel (pas de renommage, contrairement à archiver_bl : ce
    n'est pas terminé, l'acheteur doit encore le retrouver/comparer)."""

    chemin_source = Path(chemin_source)
    dossier = Path(dossier_a_verifier)
    dossier.mkdir(parents=True, exist_ok=True)

    cible = dossier / chemin_source.name
    compteur = 1
    while cible.exists():
        compteur += 1
        cible = dossier / f"{chemin_source.stem} ({compteur}){chemin_source.suffix}"

    shutil.move(str(chemin_source), str(cible))
    return cible


def ecrire_rapport(dossier_rapports: Path, texte: str) -> Path:
    dossier_rapports = Path(dossier_rapports)
    dossier_rapports.mkdir(parents=True, exist_ok=True)
    horodatage = datetime.now().strftime("%Y%m%d_%H%M%S")
    chemin = dossier_rapports / f"rapprochement_{horodatage}.txt"
    chemin.write_text(texte, encoding="utf-8")
    return chemin


def _est_resolu(g: dict, cles_ecrites: set) -> bool:
    """Un groupe BL (voir regrouper_par_bl) est résolu si TOUTES ses lignes
    sont déjà à jour ou viennent d'être écrites, aucune inconnue, et aucune
    anomalie de rapprochement (n° de commande introuvable...).

    Un bon de retour (voir moteur.fournisseurs.dist109, type_document ==
    "RETOUR") est TOUJOURS considéré résolu : il n'a par nature rien à
    écrire (rapprocher_dossier lui attache systématiquement une anomalie
    "rien à écrire depuis ce document", purement informative — voir
    rapport.anomalies_bl) et doit rejoindre Traités/<commande>/ aux côtés du
    BL qu'il annule, pas rester bloqué indéfiniment dans "à vérifier"."""

    if g["bl"].type_document == "RETOUR":
        return True

    if g["inconnu"] or g["anomalies"]:
        return False

    resolues = len(g["deja_a_jour"]) + sum(1 for c in g["sur"] + g["a_confirmer"] if id(c) in cles_ecrites)
    total = len(g["sur"]) + len(g["a_confirmer"]) + len(g["deja_a_jour"])

    return resolues >= total


def _raisons_non_resolu(g: dict, cles_ecrites: set) -> list:
    return [r for c in g["inconnu"] for r in c.raisons] + [
        r for c in g["a_confirmer"] if id(c) not in cles_ecrites for r in c.raisons
    ] + list(g["anomalies"])


def _traiter_bl_unique_du_fichier(chemin_source: Path, fichier: str, g: dict, cles_ecrites: set,
                                   dossier_a_traiter: Path, resume: dict, dossier_commandes_bc=None) -> None:
    """Fichier à UN SEUL BL : comportement historique, le fichier ENTIER
    est déplacé tel quel (archivé ou "à vérifier")."""

    # Un déplacement (archivage OU "à vérifier") qui échoue (fichier
    # verrouillé par un antivirus/le réseau/Explorer...) ne doit JAMAIS
    # faire perdre le résumé de l'écriture déjà faite dans le Suivi ni
    # empêcher le rapport final — cas réel rencontré en recette (session R2
    # suite) : le classeur avait bien été écrit, mais l'exception
    # d'archivage remontait sans jamais atteindre ecrire_rapport() plus bas.
    if _est_resolu(g, cles_ecrites):
        try:
            cible = archiver_bl(chemin_source, g["bl"], dossier_a_traiter / DOSSIER_TRAITES, dossier_commandes_bc)
            resume["bl_archives"].append((fichier, cible))
        except OSError as e:
            resume["archivage_echoue"].append((fichier, str(e)))
    else:
        raisons = _raisons_non_resolu(g, cles_ecrites)
        try:
            cible = deplacer_vers_a_verifier(chemin_source, dossier_a_traiter / DOSSIER_A_VERIFIER)
            resume["bl_a_verifier"].append((fichier, cible, raisons))
        except OSError as e:
            resume["archivage_echoue"].append((fichier, str(e)))


def _traiter_bl_multiples_du_fichier(chemin_source: Path, fichier: str, gs: list, cles_ecrites: set,
                                      dossier_a_traiter: Path, resume: dict, dossier_commandes_bc=None) -> None:
    """Fichier avec PLUSIEURS BL (Cominter Ouest, 109 Distribution — voir
    moteur.ocr.pages_par_identifiant) : chaque BL résolu est extrait
    INDIVIDUELLEMENT (découpage PDF par pages) vers Traités/<commande>/,
    chaque BL non résolu vers "À vérifier/" — jamais tout le fichier bloqué
    par UN seul BL problématique (cas réel signalé par l'acheteur, voir
    CLAUDE.md). Le fichier source est ensuite réduit aux seules pages pas
    encore redistribuées (ou supprimé si tout a pu être extrait), pour ne
    pas être relu en entier à la prochaine exécution."""

    pages_totales = _nombre_pages(chemin_source)
    pages_consommees = set()

    for g in gs:

        bl = g["bl"]
        indices = bl.pages if bl.pages is not None else list(range(pages_totales))
        resolu = _est_resolu(g, cles_ecrites)
        dossier_dest = (
            _dossier_pour_commande(dossier_a_traiter / DOSSIER_TRAITES, bl.numero_commande)
            if resolu else dossier_a_traiter / DOSSIER_A_VERIFIER
        )

        try:
            cible = _extraire_bl_vers(chemin_source, indices, dossier_dest, _nom_archive_bl(bl))
        except OSError as e:
            resume["archivage_echoue"].append((fichier, str(e)))
            continue

        pages_consommees.update(indices)

        if resolu:
            _copier_bon_de_commande_si_absent(dossier_dest, bl.numero_commande, dossier_commandes_bc)
            resume["bl_archives"].append((fichier, cible))
        else:
            resume["bl_a_verifier"].append((fichier, cible, _raisons_non_resolu(g, cles_ecrites)))

    pages_restantes = sorted(set(range(pages_totales)) - pages_consommees)
    try:
        if not pages_restantes:
            chemin_source.unlink()
        elif len(pages_restantes) < pages_totales:
            _reecrire_avec_pages(chemin_source, pages_restantes)
    except OSError as e:
        resume["archivage_echoue"].append((fichier, f"Réduction du fichier source après extraction impossible : {e}"))


def appliquer_et_archiver(dossier_projet, dossier_a_traiter, rapport: RapportRapprochement,
                           correspondances_a_ecrire: list) -> dict:
    """Écrit `correspondances_a_ecrire` dans le Suivi (verrou/sauvegarde/
    patch chirurgical — voir ecriture.appliquer), puis :
    - archive chaque BL dont TOUTES les lignes sont résolues (écrites ou
      déjà à jour) dans a_traiter/BL/Traités/ ;
    - déplace chaque BL avec au moins une ligne "à confirmer" non cochée ou
      "inconnue" vers a_traiter/BL/À vérifier/ (jamais laissé mélangé avec
      les BL pas encore traités — demande explicite de l'acheteur, session
      R2 suite : elle doit pouvoir repérer d'un coup d'œil les BL qui
      attendent une décision de sa part, ex. référence différente entre BL
      et Suivi).
    Retourne un résumé {sauvegarde, lignes_ecrites, bl_archives,
    bl_a_verifier, archivage_echoue, chemin_rapport}."""

    dossier_projet = Path(dossier_projet)
    dossier_commandes_bc = trouver_dossier_commandes(dossier_projet)

    resume = {
        "sauvegarde": None, "lignes_ecrites": 0, "pieces_ecrites": 0, "pieces_ignorees": [],
        "bl_archives": [], "bl_a_verifier": [],
        "archivage_echoue": [], "chemin_rapport": None,
    }

    if correspondances_a_ecrire:
        # Depuis P1 : la feuille Pièces doit exister AVANT toute écriture
        # (jamais une Qté livrée écrite sans sa ligne de document en face).
        if not feuille_pieces_presente(rapport.fichier_suivi):
            raise FeuillePiecesAbsente(
                f"« {Path(rapport.fichier_suivi).name} » n'a pas de feuille Pièces — rien n'a été écrit."
            )
        ecritures = ecritures_pour(correspondances_a_ecrire)
        resume["sauvegarde"] = appliquer(
            rapport.fichier_suivi, ecritures, dossier_projet / DOSSIER_BACKUPS,
        )
        resume["lignes_ecrites"] = len(correspondances_a_ecrire)
        resultat_pieces = ecrire_pieces(
            rapport.fichier_suivi,
            pieces_pour_bl(correspondances_a_ecrire, Path(dossier_a_traiter) / DOSSIER_TRAITES),
            dossier_projet / DOSSIER_BACKUPS,
        )
        resume["pieces_ecrites"] = resultat_pieces["ajoutees"]
        resume["pieces_ignorees"] = resultat_pieces["ignorees"]

    cles_ecrites = {id(c) for _, c in correspondances_a_ecrire}
    groupes = regrouper_par_bl(rapport)

    # BUG RÉEL CORRIGÉ (détection de fournisseur par page, voir
    # lecture_bl.lire_bl) : un fichier qui mélange plusieurs fournisseurs
    # peut désormais avoir À LA FOIS des BonLivraison résolus (certaines
    # pages) ET une anomalie de lecture (une autre page, ex. fournisseur
    # sans parser BL) — `par_fichier` doit donc être construit AVANT de
    # décider quels fichiers sont en échec TOTAL, sinon un fichier
    # partiellement résolu se ferait déplacer ENTIER vers "à vérifier" par
    # la boucle ci-dessous avant que ses pages résolues n'aient eu la
    # chance d'être archivées individuellement (la 2e boucle trouverait
    # alors `chemin_source` déjà déplacé et ne ferait plus rien).
    par_fichier = {}
    for g in groupes.values():
        par_fichier.setdefault(g["bl"].fichier, []).append(g)

    # Fichiers en échec de lecture TOTAL (aucun BL n'a pu être extrait pour
    # eux — fournisseur non reconnu, pas de parser, PDF illisible, voir
    # lecture_bl.analyser_dossier) : rien à découper par BL, le fichier
    # ENTIER part vers "à vérifier", comme avant. Un fichier qui a AUSSI
    # des BonLivraison résolus (voir ci-dessus) n'est PAS en échec total —
    # son anomalie de lecture reste comptée dans le rapport, mais ses pages
    # résolues suivent leur propre chemin normalement.
    fichiers_en_echec_total = {
        fichier for fichier, _ in rapport.anomalies_lecture
        if fichier not in par_fichier
    }

    dossier_a_traiter = Path(dossier_a_traiter)

    for fichier in fichiers_en_echec_total:

        chemin_source = dossier_a_traiter / fichier
        if not chemin_source.exists():
            continue

        raisons = [r for f, r in rapport.anomalies_lecture if f == fichier]
        try:
            cible = deplacer_vers_a_verifier(chemin_source, dossier_a_traiter / DOSSIER_A_VERIFIER)
            resume["bl_a_verifier"].append((fichier, cible, raisons))
        except OSError as e:
            resume["archivage_echoue"].append((fichier, str(e)))

    # BUG RÉEL CORRIGÉ (session R2 suite, recette Cominter Ouest) : un
    # fichier peut contenir PLUSIEURS BL (voir moteur.ocr.
    # pages_par_identifiant). Regrouper par BL (id(bl)) plutôt que par nom
    # de fichier permet de savoir, BL par BL, s'il est résolu — et donc de
    # n'archiver/déplacer QUE ses pages à lui (voir
    # _traiter_bl_multiples_du_fichier), sans attendre que ses frères du
    # même fichier le soient aussi (cas réel signalé par l'acheteur : sur 8
    # BL scannés dans un seul fichier, un seul avait été traité et renommé
    # correctement — CLAUDE.md, "archivage par BL individuel").
    # (par_fichier déjà construit plus haut, avant le calcul de
    # fichiers_en_echec_total.)

    for fichier, gs in par_fichier.items():

        chemin_source = dossier_a_traiter / fichier
        if not chemin_source.exists():
            continue

        # Un seul BL, et ses pages couvrent le fichier ENTIER (bl.pages non
        # renseigné — historique — ou renseigné mais couvrant déjà toutes
        # les pages, ex. un Cominter/109D/Electric Plus à un seul BL) : cas
        # de loin le plus courant, comportement historique inchangé, on
        # déplace le fichier ENTIER tel quel (pas de réécriture PDF
        # inutile). Si les pages de l'unique BL ne couvrent qu'une PARTIE
        # du fichier (détection de fournisseur par page, voir
        # lecture_bl.lire_bl : d'autres pages appartiennent à un
        # fournisseur différent, ou n'ont pas pu être lues), il faut
        # découper — jamais déplacer tout le fichier en emportant des
        # pages qui n'ont rien à voir avec ce BL.
        bl_unique = gs[0]["bl"] if len(gs) == 1 else None
        couvre_fichier_entier = bl_unique is not None and (
            bl_unique.pages is None
            or set(bl_unique.pages) == set(range(_nombre_pages(chemin_source)))
        )

        if couvre_fichier_entier:
            _traiter_bl_unique_du_fichier(
                chemin_source, fichier, gs[0], cles_ecrites, dossier_a_traiter, resume, dossier_commandes_bc,
            )
        else:
            _traiter_bl_multiples_du_fichier(
                chemin_source, fichier, gs, cles_ecrites, dossier_a_traiter, resume, dossier_commandes_bc,
            )

    lignes_rapport = [
        f"Rapprochement BL — {datetime.now().strftime('%d/%m/%Y %H:%M')}",
        f"{resume['lignes_ecrites']} ligne(s) écrite(s) dans le Suivi commandes — {resume['pieces_ecrites']} ligne(s) BL "
        f"dans la feuille Pièces, {len(resume['pieces_ignorees'])} déjà présente(s) (ID pièce).",
        f"{len(resume['bl_archives'])} BL archivé(s) : " + ", ".join(f for f, _ in resume["bl_archives"]),
        f"{len(resume['bl_a_verifier'])} BL déplacé(s) vers {DOSSIER_A_TRAITER_BL}/{DOSSIER_A_VERIFIER}/ "
        "(décision humaine nécessaire) :",
    ] + [
        f"  - {fichier} : {' ; '.join(raisons) if raisons else '(voir détail)'}"
        for fichier, _, raisons in resume["bl_a_verifier"]
    ] + [
        f"{len(resume['archivage_echoue'])} BL écrit(s) mais PAS déplacé(s) (fichier verrouillé, à ranger à la main) : "
        + ", ".join(f for f, _ in resume["archivage_echoue"]),
        f"{len(rapport.deja_a_jour)} ligne(s) déjà à jour (rien écrit, doublon évité).",
        f"{len(rapport.anomalies_lecture)} anomalie(s) de lecture (fichier illisible/fournisseur non reconnu) : "
        + "; ".join(f"{f} ({r})" for f, r in rapport.anomalies_lecture),
        f"{len(rapport.anomalies_bl)} BL non rapproché(s) (n° de commande/commande introuvable) : "
        + "; ".join(f"{bl.fichier} ({r})" for bl, r in rapport.anomalies_bl),
    ]
    resume["chemin_rapport"] = ecrire_rapport(dossier_projet / DOSSIER_RAPPORTS, "\n".join(lignes_rapport))

    return resume
