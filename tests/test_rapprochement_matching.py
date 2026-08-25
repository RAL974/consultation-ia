"""
Rapprochement BL <-> Suivi commandes (moteur/rapprochement/matching.py,
branche Rapprochement AI, session R2 — voir CLAUDE.md).

Deux familles de tests :
- la logique d'appariement/comparaison (apparier, _comparer) sur des objets
  construits à la main — déterministe, ne dépend d'aucun fichier ;
- la lecture du Suivi (lire_lignes_commande) sur un classeur SYNTHÉTIQUE
  structurellement représentatif (mêmes en-têtes que le vrai), même esprit
  que tests/test_rapprochement_ecriture.py.

La preuve « ça marche sur le vrai Suivi » a été faite manuellement cette
session sur les 4 vrais BL 109 Distribution (tests/fixtures/bl_dist109_*.pdf)
contre le vrai classeur : les 7 lignes s'apparient et ressortent TOUTES
"déjà à jour" — ces 4 BL avaient déjà été saisis à la main par l'acheteur
avant cette session (voir recette).
"""

from openpyxl import Workbook

from moteur.referentiel import Referentiel
from moteur.rapprochement.matching import (
    LigneSuivi,
    Statut,
    apparier,
    deduire_commande_par_contenu,
    lire_lignes_commande,
)
from moteur.rapprochement.modele_bl import LigneBL


def _suivi(reference, qte_commandee, qte_livree=0.0, tarif_bl=None, date_livraison=None, ligne_excel=2, designation=""):
    return LigneSuivi(
        ligne_excel=ligne_excel,
        reference=reference,
        designation=designation,
        qte_commandee=qte_commandee,
        qte_livree=qte_livree,
        tarif_bl=tarif_bl,
        date_livraison=date_livraison,
        statut="",
        note="",
    )


def _bl(reference, quantite_livree, prix_net, designation=""):
    return LigneBL(
        reference_fournisseur=reference,
        designation=designation,
        quantite_livree=quantite_livree,
        prix_net=prix_net,
        montant=round(quantite_livree * prix_net, 2) if prix_net is not None else None,
    )


def test_apparier_ligne_deja_a_jour_idempotence():
    """Cas réel : le BL a déjà été saisi à la main (même qté, même tarif,
    date déjà renseignée) -> rien à écrire, jamais un doublon."""

    suivi = [_suivi("086101L", qte_commandee=20, qte_livree=20, tarif_bl=4.6, date_livraison="2026-08-10")]
    bl = [_bl("86101L", quantite_livree=20, prix_net=4.6)]  # préfixe/zéro de tête différents, cœur numérique identique

    [c] = apparier(bl, suivi)

    assert c.statut is Statut.DEJA_A_JOUR
    assert c.ligne_suivi.ligne_excel == 2


def test_apparier_ligne_deja_a_jour_idempotence_sans_prix_sur_le_bl():
    """Cas réel Coredime : ce fournisseur n'affiche JAMAIS de prix sur le
    BL (réglé à la facture) -> le Tarif BL du Suivi reste vide en
    permanence pour ces lignes. L'idempotence ne doit PAS dépendre d'une
    comparaison de tarif dans ce cas (bug réel corrigé cette session :
    aucune ligne Coredime n'aurait jamais pu ressortir "déjà à jour")."""

    suivi = [_suivi("227059360", qte_commandee=300, qte_livree=300, tarif_bl=None, date_livraison="2026-08-11")]
    bl = [_bl("227059360", quantite_livree=300, prix_net=None)]

    [c] = apparier(bl, suivi)

    assert c.statut is Statut.DEJA_A_JOUR


def test_apparier_date_suivi_incoherente_devient_a_confirmer_sans_recumuler():
    """BUG RÉEL DE FOND trouvé par l'acheteur (voir CLAUDE.md) : "déjà à
    jour" ne vérifiait jamais que la date ENREGISTRÉE dans le Suivi est la
    bonne, seulement qu'une date quelconque est présente — des lignes
    écrites avant que l'extraction de date fonctionne (repli sur
    date.today()) restaient bloquées avec la mauvaise date pour toujours.
    Quand `date_bl_reelle` est fournie et diffère de la date déjà
    enregistrée (qté/tarif par ailleurs identiques), la ligne doit
    ressortir "à confirmer" — MAIS sans jamais recumuler la quantité
    (elle est déjà correcte)."""

    from datetime import date, datetime

    suivi = [_suivi("086101L", qte_commandee=20, qte_livree=20, tarif_bl=4.6, date_livraison=datetime(2026, 8, 13))]
    bl = [_bl("86101L", quantite_livree=20, prix_net=4.6)]

    [c] = apparier(bl, suivi, date_bl_reelle=date(2026, 7, 15))

    assert c.statut is Statut.A_CONFIRMER
    assert "Date de livraison" in c.raisons[0]
    assert "2026-07-15" in c.raisons[0]
    assert c.qte_livree_cumulee == 20  # jamais 40 : la quantité ne doit pas être recumulée


def test_apparier_date_suivi_coherente_avec_date_bl_reelle_reste_deja_a_jour():
    """Sanity check : quand la date enregistrée correspond bien à la vraie
    date du BL, rien ne change (toujours "déjà à jour"). `date_livraison`
    est un `datetime` (comme le renvoie openpyxl sur le vrai classeur, pas
    une chaîne) pour que la comparaison de date ait un sens."""

    from datetime import date, datetime

    suivi = [_suivi("086101L", qte_commandee=20, qte_livree=20, tarif_bl=4.6, date_livraison=datetime(2026, 7, 15))]
    bl = [_bl("86101L", quantite_livree=20, prix_net=4.6)]

    [c] = apparier(bl, suivi, date_bl_reelle=date(2026, 7, 15))

    assert c.statut is Statut.DEJA_A_JOUR


def test_apparier_sans_date_bl_reelle_comportement_inchange():
    """Si l'appelant ne fournit pas date_bl_reelle (ex. date du BL
    illisible à l'OCR), le comportement reste celui d'avant ce correctif —
    aucune vérification de cohérence de date, uniquement présence."""

    suivi = [_suivi("086101L", qte_commandee=20, qte_livree=20, tarif_bl=4.6, date_livraison="2026-08-13")]
    bl = [_bl("86101L", quantite_livree=20, prix_net=4.6)]

    [c] = apparier(bl, suivi)

    assert c.statut is Statut.DEJA_A_JOUR


def test_apparier_ligne_sure_premiere_livraison():
    """Commande encore vide dans le Suivi -> ligne "sûre", prête à écrire."""

    suivi = [_suivi("52302", qte_commandee=32, qte_livree=0, tarif_bl=None, date_livraison=None)]
    bl = [_bl("52302", quantite_livree=32, prix_net=0.95)]

    [c] = apparier(bl, suivi)

    assert c.statut is Statut.SUR
    assert c.raisons == []
    assert c.qte_livree_cumulee == 32


def test_apparier_ligne_sur_livraison_a_confirmer():
    """Le cumul (déjà livré + ce BL) dépasse la quantité commandée ->
    jamais silencieux, part au bac "à confirmer"."""

    suivi = [_suivi("70002", qte_commandee=20, qte_livree=15, tarif_bl=1.5, date_livraison="2026-08-01")]
    bl = [_bl("70002", quantite_livree=10, prix_net=1.5)]  # 15 + 10 = 25 > 20 commandés

    [c] = apparier(bl, suivi)

    assert c.statut is Statut.A_CONFIRMER
    assert any("Sur-livraison" in r for r in c.raisons)
    assert c.qte_livree_cumulee == 25


def test_apparier_ligne_tarif_different_a_confirmer():

    suivi = [_suivi("70002", qte_commandee=20, qte_livree=0, tarif_bl=None, date_livraison=None)]
    suivi[0].tarif_bl = 1.5  # tarif déjà connu (ex. saisi manuellement) mais différent du BL
    bl = [_bl("70002", quantite_livree=20, prix_net=1.8)]

    [c] = apparier(bl, suivi)

    assert c.statut is Statut.A_CONFIRMER
    assert any("Tarif différent" in r for r in c.raisons)


def test_apparier_ligne_sans_correspondance_inconnu():

    suivi = [_suivi("70002", qte_commandee=20)]
    bl = [_bl("999999", quantite_livree=5, prix_net=1.0)]  # référence absente de cette commande

    [c] = apparier(bl, suivi)

    assert c.statut is Statut.INCONNU
    assert c.ligne_suivi is None


def test_apparier_ligne_ambigue_mais_correspondance_exacte_disponible():
    """BUG RÉEL CORRIGÉ (cas réel, commande M3.10.175) : deux références du
    Suivi partagent le même cœur numérique ("R9PRC263"/"R9PRA263" -> "9263"
    toutes les deux, la lettre médiane C/A n'étant pas un chiffre) —
    DIFFÉRENTS articles réels (interrupteur différentiel type A vs AC), pas
    une coïncidence de saisie. Une correspondance de TEXTE EXACT parmi les
    candidats ambigus est toujours préférée à un abandon "ambigu", ici sans
    aucune ambiguïté réelle puisque le BL porte le texte exact de l'un des
    deux."""

    suivi = [
        _suivi("R9PRC263", qte_commandee=2, ligne_excel=2),
        _suivi("R9PRA263", qte_commandee=1, ligne_excel=3),
    ]
    bl = [_bl("R9PRC263", quantite_livree=2, prix_net=48.0)]

    [c] = apparier(bl, suivi)

    assert c.statut is Statut.SUR
    assert c.ligne_suivi.ligne_excel == 2


def test_apparier_ligne_ambigue_inconnu_sans_correspondance_exacte():
    """Deux lignes Suivi partagent le même cœur numérique de référence ET
    aucune ne correspond au texte EXACT du BL (l'écart de zéro de tête
    touche les DEUX candidats) -> toujours pas de choix au hasard, bac
    "inconnu"."""

    suivi = [_suivi("052302", qte_commandee=10, ligne_excel=2), _suivi("0052302", qte_commandee=10, ligne_excel=3)]
    bl = [_bl("52302", quantite_livree=5, prix_net=1.0)]

    [c] = apparier(bl, suivi)

    assert c.statut is Statut.INCONNU
    assert "ambigu" in c.raisons[0]


def test_apparier_une_ligne_suivi_nest_utilisee_quune_fois():
    """Deux lignes de BL, même référence normalisée : la 2e ne doit pas
    réapparier la ligne Suivi déjà prise par la 1ère."""

    suivi = [_suivi("70002", qte_commandee=20, ligne_excel=2)]
    bl = [_bl("70002", quantite_livree=10, prix_net=1.5), _bl("70002", quantite_livree=10, prix_net=1.5)]

    c1, c2 = apparier(bl, suivi)

    assert c1.statut is Statut.SUR
    assert c2.statut is Statut.INCONNU


def test_apparier_repli_reference_proche_1_caractere_ecart():
    """Cas réel signalé par l'acheteur (trou de perforateur sur le BL
    papier) : la référence lue par l'OCR ("L405205") diffère d'UN SEUL
    chiffre de la vraie référence Suivi ("405209") -> repli sur cette
    ligne plutôt que "inconnu", mais TOUJOURS "à confirmer" (jamais un
    rapprochement de repli écrit automatiquement)."""

    suivi = [_suivi(
        "405209", qte_commandee=10, ligne_excel=2,
        designation="Borne de connexion interr diff 63A tête groupe 2 mod",
    )]
    bl = [_bl("L405205", quantite_livree=3, prix_net=4.21, designation="BornedeconnexionDx3-ID63A")]

    [c] = apparier(bl, suivi)

    assert c.statut is Statut.A_CONFIRMER
    assert c.ligne_suivi is suivi[0]
    assert "1 caractère d'écart" in c.raisons[0]
    assert "405209" in c.raisons[0]


def test_apparier_repli_reference_proche_deja_a_jour_ne_cumule_jamais():
    """BUG RÉEL CRITIQUE trouvé lors d'une vraie écriture sur le classeur
    vivant : la ligne Suivi visée par le repli (405209, commande M3.23.034)
    était DÉJÀ à 3 livrées/3 commandées, même tarif, date déjà renseignée —
    un cas "déjà à jour" en tout point, sauf que la référence BL ("L405205",
    trou de perforateur) ne matche pas exactement. Forcer "à confirmer"
    aveuglément faisait calculer qte_livree_cumulee = 3 (déjà) + 3 (BL) = 6
    pour seulement 3 commandées — écrit à tort dans le VRAI Suivi (6/3,
    sur-livraison fantôme) avant d'être repéré et corrigé manuellement.
    Quand le repli pointe vers une ligne dont _comparer() détermine qu'elle
    est déjà à jour, AUCUNE écriture ne doit jamais être proposée."""

    suivi = [_suivi(
        "405209", qte_commandee=3, qte_livree=3, tarif_bl=4.21, date_livraison="2026-08-10",
        ligne_excel=5827, designation="Borne de connexion interr diff 63A tête groupe 2 mod",
    )]
    bl = [_bl("L405205", quantite_livree=3, prix_net=4.21, designation="BornedeconnexionDx3-ID63A")]

    [c] = apparier(bl, suivi)

    assert c.statut is Statut.DEJA_A_JOUR
    assert c.ligne_suivi is suivi[0]
    assert "1 caractère d'écart" in c.raisons[0]
    # Aucune Ecriture ne doit jamais être construite pour une ligne "déjà à
    # jour" (voir pipeline_bl.ecritures_pour, qui ne reçoit que
    # sur+a_confirmer) — vérifié indirectement ici : qte_livree_cumulee
    # resterait fausse (6.0) si jamais réutilisée par erreur, donc on
    # s'assure surtout que le statut empêche cet usage en amont.
    assert c.qte_livree_cumulee == 6.0  # jamais utilisé : DEJA_A_JOUR n'est jamais passé à ecritures_pour


def test_apparier_repli_reference_proche_ignore_si_deux_candidats():
    """Deux lignes Suivi sont chacune à 1 caractère du BL -> aucun repli
    fiable, jamais un choix au hasard entre les deux, bac "inconnu"."""

    suivi = [
        _suivi("405209", qte_commandee=10, ligne_excel=2),
        _suivi("405208", qte_commandee=10, ligne_excel=3),
    ]
    bl = [_bl("L405200", quantite_livree=3, prix_net=4.21)]

    [c] = apparier(bl, suivi)

    assert c.statut is Statut.INCONNU


def test_apparier_repli_reference_proche_ignore_si_ecart_de_plusieurs_caracteres():
    """Deux chiffres différents (pas un trou de perforateur isolé) -> pas
    de repli, bac "inconnu" comme avant."""

    suivi = [_suivi("405209", qte_commandee=10, ligne_excel=2)]
    bl = [_bl("L405118", quantite_livree=3, prix_net=4.21)]

    [c] = apparier(bl, suivi)

    assert c.statut is Statut.INCONNU


def test_apparier_repli_reference_proche_priorite_a_lexact():
    """Une correspondance exacte doit toujours primer sur le repli — même
    si une autre ligne Suivi est proche, elle ne doit jamais lui être
    préférée."""

    suivi = [
        _suivi("405209", qte_commandee=10, ligne_excel=2),
        _suivi("405208", qte_commandee=10, ligne_excel=3),
    ]
    bl = [_bl("L405209", quantite_livree=3, prix_net=4.21)]

    [c] = apparier(bl, suivi)

    assert c.statut is Statut.SUR
    assert c.ligne_suivi.ligne_excel == 2


def test_apparier_repli_chiffre_de_tete_manquant():
    """Cas réel signalé par l'acheteur : les gars de l'atelier perforent
    les BL pour les classer, ce qui efface parfois COMPLÈTEMENT un chiffre
    de tête de la référence imprimée (pas juste l'abîmer/le substituer,
    déjà couvert par le repli 1 caractère) — ex. BL "9894" pour la vraie
    référence Suivi "ALB69894" (le "6" a disparu). Toujours "à confirmer",
    jamais "sûr"."""

    suivi = [_suivi(
        "ALB69894", qte_commandee=1, ligne_excel=2,
        designation="Trépan SDS Ø67",
    )]
    bl = [_bl("9894", quantite_livree=1, prix_net=125.0, designation="SCIECLOCHEBETOND67")]

    [c] = apparier(bl, suivi)

    assert c.statut is Statut.A_CONFIRMER
    assert c.ligne_suivi is suivi[0]
    assert "chiffre(s) de tête manquant(s)" in c.raisons[0]
    assert "ALB69894" in c.raisons[0]


def test_apparier_repli_chiffre_de_tete_manquant_ignore_si_deux_candidats():
    """Deux lignes Suivi ont chacune un chiffre de tête numérique d'écart
    -> aucun repli fiable, bac "inconnu"."""

    suivi = [
        _suivi("19894", qte_commandee=1, ligne_excel=2),
        _suivi("29894", qte_commandee=1, ligne_excel=3),
    ]
    bl = [_bl("9894", quantite_livree=1, prix_net=125.0)]

    [c] = apparier(bl, suivi)

    assert c.statut is Statut.INCONNU


def test_apparier_repli_chiffre_de_tete_manquant_ignore_si_ecart_trop_grand():
    """Plus de 2 chiffres de tête manquants -> pas assez fiable, bac
    "inconnu" comme avant (pas de règle inventée au-delà de ce qui a été
    observé réellement)."""

    suivi = [_suivi("123469894", qte_commandee=1, ligne_excel=2)]
    bl = [_bl("9894", quantite_livree=1, prix_net=125.0)]

    [c] = apparier(bl, suivi)

    assert c.statut is Statut.INCONNU


def test_apparier_repli_reference_courte_ecart_alphanumerique_final():
    """Cas réel signalé par l'acheteur (commande 142.033) : BL "H07VK16BL"
    vs Suivi "H07VK16B" (un "L" en trop en fin de référence) — le cœur
    numérique des deux ("716") est trop court (< 4 chiffres) pour les
    replis existants, qui ne comparent que la partie numérique. Repli sur
    le texte alphanumérique brut (garde les lettres) avec le même critère
    1 caractère d'écart."""

    suivi = [_suivi("H07VK16B", qte_commandee=5, ligne_excel=2)]
    bl = [_bl("H07VK16BL", quantite_livree=5, prix_net=2.75)]

    [c] = apparier(bl, suivi)

    assert c.statut is Statut.A_CONFIRMER
    assert c.ligne_suivi is suivi[0]
    assert "référence trop courte" in c.raisons[0]


def test_apparier_repli_confusion_ocr_1_i():
    """Cas réel signalé par l'acheteur (commande M2.16.010, Ravate Elec) :
    BL "XVR1IISTI" vs Suivi "XVR111STI" ("WittyOne 11kW") — DEUX des trois
    "1" confondus avec la lettre "I" par l'OCR (distance de 2 caractères,
    au-delà du seuil de _distance_courte). Repli dédié par égalité après
    normalisation 1/I, pas juste une distance à 1 caractère près."""

    suivi = [_suivi("XVR111STI", qte_commandee=1, ligne_excel=2, designation="BORNE RECHARGE WITTYONE 11KW")]
    bl = [_bl("XVR1IISTI", quantite_livree=1, prix_net=580.0, designation="BORNE RECHARGE WITTYONE IIKW")]

    [c] = apparier(bl, suivi)

    assert c.statut is Statut.A_CONFIRMER
    assert c.ligne_suivi is suivi[0]
    assert "confusion OCR" in c.raisons[0]


def test_apparier_repli_confusion_ocr_1_i_ignore_si_deja_identique():
    """Pas de repli inutile quand les références sont DÉJÀ identiques (cas
    normal, traité ailleurs comme une correspondance exacte)."""

    suivi = [_suivi("XVR111STI", qte_commandee=1, ligne_excel=2)]
    bl = [_bl("XVR111STI", quantite_livree=1, prix_net=580.0)]

    [c] = apparier(bl, suivi)

    assert c.statut is Statut.SUR
    assert c.raisons == []


def test_deduire_commande_par_contenu_signature_claire():
    """Cas réel signalé par l'acheteur : 3 BL Cominter sans n° de commande
    lisible (même trou de perforateur, sur la zone du n° cette fois),
    déduits sans ambiguïté par leur contenu (référence + quantité
    commandée) — validé sur les 3 vrais cas trouvés en recette."""

    lignes_bl = [
        _bl("ECLLED739814", quantite_livree=22.0, prix_net=14.0),
        _bl("ECLLED7770B", quantite_livree=40.0, prix_net=21.34),
        _bl("LITAZU30120-102", quantite_livree=8.0, prix_net=23.4),
        _bl("ECLLED739795", quantite_livree=8.0, prix_net=27.6),
    ]

    lignes_suivi_fournisseur = [
        # La bonne commande : les 4 références ET quantités concordent exactement.
        LigneSuivi(10, "ECLLED739814", "", 22, 0, None, None, "", "", numero_commande="M3.14.361"),
        LigneSuivi(11, "ECLLED7770B", "", 40, 0, None, None, "", "", numero_commande="M3.14.361"),
        LigneSuivi(12, "LITAZU30120-102", "", 8, 0, None, None, "", "", numero_commande="M3.14.361"),
        LigneSuivi(13, "ECLLED739795", "", 8, 0, None, None, "", "", numero_commande="M3.14.361"),
        # Une commande concurrente qui ne partage que 2 références, avec des
        # quantités différentes -> ne doit jamais l'emporter.
        LigneSuivi(20, "ECLLED739814", "", 30, 30, None, None, "", "", numero_commande="M3.14.347"),
        LigneSuivi(21, "ECLLED7770B", "", 2, 2, None, None, "", "", numero_commande="M3.14.347"),
    ]

    numero, score = deduire_commande_par_contenu(lignes_bl, lignes_suivi_fournisseur)

    assert numero == "M3.14.361"
    assert score == 4


def test_deduire_commande_par_contenu_aucune_deduction_si_une_seule_ligne():
    """Une seule référence en commun -> jamais assez fiable seule (trop de
    références génériques réutilisées dans des dizaines de commandes),
    aucune déduction."""

    lignes_bl = [_bl("ECLLED739814", quantite_livree=22.0, prix_net=14.0)]
    lignes_suivi_fournisseur = [
        LigneSuivi(10, "ECLLED739814", "", 22, 0, None, None, "", "", numero_commande="M3.14.361"),
    ]

    numero, score = deduire_commande_par_contenu(lignes_bl, lignes_suivi_fournisseur)

    assert numero is None
    assert score == 0


def test_deduire_commande_par_contenu_aucune_deduction_si_egalite():
    """Deux commandes ont chacune 2 lignes qui concordent -> égalité,
    jamais un choix au hasard."""

    lignes_bl = [
        _bl("REF1", quantite_livree=5.0, prix_net=1.0),
        _bl("REF2", quantite_livree=3.0, prix_net=1.0),
    ]
    lignes_suivi_fournisseur = [
        LigneSuivi(10, "REF1", "", 5, 0, None, None, "", "", numero_commande="COMMANDE_A"),
        LigneSuivi(11, "REF2", "", 3, 0, None, None, "", "", numero_commande="COMMANDE_A"),
        LigneSuivi(20, "REF1", "", 5, 0, None, None, "", "", numero_commande="COMMANDE_B"),
        LigneSuivi(21, "REF2", "", 3, 0, None, None, "", "", numero_commande="COMMANDE_B"),
    ]

    numero, score = deduire_commande_par_contenu(lignes_bl, lignes_suivi_fournisseur)

    assert numero is None
    assert score == 0


def _classeur_suivi_synthetique(chemin):

    wb = Workbook()
    ws = wb.active
    ws.title = "Commandes"
    ws.append([
        "Référence", "Désignation", "Qté commandée", "N° de commande",
        "Fournisseur", "Date de livraison", "Qté livrée", "Tarif BL",
        "Statut commande", "Note",
    ])
    ws.append(["81000298", "Kit Boulon PVC M8", 2, "123.096", "109 Distribution", None, 0, None, "🔵 En attente", None])
    ws.append(["AUTRE", "Autre article", 5, "999.999", "AUTRE FOURNISSEUR", None, 0, None, "🔵 En attente", None])
    ws.append(["LEG062525", "BAES EVAC IP43", 6, "M3.10.172", "GMR", None, 0, None, "🔵 En attente", None])
    wb.save(chemin)


def test_lire_lignes_commande_filtre_fournisseur_et_commande(tmp_path):

    chemin = tmp_path / "suivi_synthetique.xlsx"
    _classeur_suivi_synthetique(chemin)

    lignes = lire_lignes_commande(chemin, "109 DISTRIBUTION", "123.096")

    assert len(lignes) == 1
    assert lignes[0].reference == "81000298"
    assert lignes[0].ligne_excel == 2
    assert lignes[0].qte_commandee == 2


def test_lire_lignes_commande_convertit_nom_fournisseur_electric_plus_gmr(tmp_path):
    """Cas réel (session R2 suite, fournisseur Electric Plus/GMR) : le nom
    détecteur ("ELECTRIC PLUS", enseigne grand public) n'est pas celui
    écrit dans le Suivi ("GMR", branche grands comptes pros) — sans la
    conversion via moteur.panier.MAPPING_FOURNISSEURS, AUCUNE ligne
    Electric Plus ne pouvait jamais être retrouvée dans le Suivi."""

    chemin = tmp_path / "suivi_synthetique.xlsx"
    _classeur_suivi_synthetique(chemin)

    lignes = lire_lignes_commande(chemin, "ELECTRIC PLUS", "M3.10.172")

    assert len(lignes) == 1
    assert lignes[0].reference == "LEG062525"


def test_lire_lignes_commande_colonne_manquante_leve(tmp_path):

    wb = Workbook()
    ws = wb.active
    ws.title = "Commandes"
    ws.append(["Référence", "Fournisseur"])  # en-têtes incomplets
    chemin = tmp_path / "suivi_incomplet.xlsx"
    wb.save(chemin)

    import pytest
    with pytest.raises(KeyError):
        lire_lignes_commande(chemin, "109 DISTRIBUTION", "123.096")


def test_apparier_exact_nest_jamais_vole_par_un_repli_dune_autre_ligne():
    """BUG RÉEL CORRIGÉ (cas réel, commande 130.036) : "L600001" (aucune
    correspondance exacte, mais à 1 caractère de "600002" côté Suivi) et
    "L600002" (correspondance EXACTE avec "600002") sur le MÊME BL. En une
    seule passe, "L600001" traité EN PREMIER (ordre du document) volait la
    ligne Suivi "600002" par repli, laissant "L600002" — pourtant exact —
    sans rien à apparier (ressorti "inconnu" à tort). apparier() résout
    désormais toutes les correspondances exactes AVANT tout repli, quel que
    soit l'ordre des lignes du BL : "L600002" doit matcher exactement,
    "L600001" doit rester "inconnu" (aucune ligne Suivi disponible qui lui
    corresponde, exactement ou par repli, une fois "600002" déjà prise)."""

    suivi = [_suivi("600002", qte_commandee=20, designation="Inter double va et vient Dooxie")]
    bl = [
        _bl("L600001", quantite_livree=20, prix_net=2.1, designation="Interrupteur Va et Vient"),
        _bl("L600002", quantite_livree=20, prix_net=4.14, designation="Interrupteur Double Va et Vient"),
    ]

    resultat = apparier(bl, suivi)

    assert resultat[0].statut == Statut.INCONNU
    assert resultat[0].ligne_bl.reference_fournisseur == "L600001"

    assert resultat[1].statut == Statut.SUR
    assert resultat[1].ligne_suivi.reference == "600002"


# ------------------------------------------------------------------
# Référentiel articles (demande explicite de l'acheteur : "il faut créer
# une base des équivalences, ce genre de cas va se présenter très souvent"
# — cas réel 59210/CFF1BIS, substitution fournisseur d'un article
# générique, AUCUN rapport textuel ni numérique entre les deux références,
# donc aucun repli ci-dessus ne peut le deviner). Voir moteur/referentiel.py.
# ------------------------------------------------------------------
def test_apparier_alias_referentiel_confirme_traite_comme_exact(tmp_path):
    """Une équivalence déjà CONFIRMÉE (origine 'manuel', comme
    referentiel/equivalences_bl.csv, ou 'confirme' via A_confirmer_BL.xlsx)
    est reconnue automatiquement dès la 1re passe, sans jamais redemander
    confirmation — c'est tout l'intérêt demandé : une fois confirmée une
    fois, elle vaut pour tous les futurs BL."""

    referentiel = Referentiel(tmp_path / "moteur")
    referentiel.cx.execute(
        "INSERT INTO alias VALUES (?,?,?)", ("59210", "CFF1BIS", "manuel"),
    )
    referentiel.cx.execute(
        "INSERT INTO alias VALUES (?,?,?)", ("CFF1BIS", "CFF1BIS", "manuel"),
    )
    referentiel.cx.commit()

    suivi = [_suivi("CFF1BIS", qte_commandee=1, designation="100 colliers à embase RAMTUB 16/32")]
    bl = [_bl("59210", quantite_livree=1, prix_net=19.0, designation="100 COLLIERS A EMBASE RAMTUB 16/32")]

    [c] = apparier(bl, suivi, referentiel=referentiel)

    assert c.statut == Statut.SUR
    assert c.ligne_suivi.reference == "CFF1BIS"

    referentiel.fermer()


def test_apparier_sans_referentiel_reste_inconnu():
    """Comportement par défaut (referentiel=None) inchangé : sans
    référentiel, deux références sans aucun rapport restent 'inconnu' —
    aucun risque de régression pour tout appelant existant."""

    suivi = [_suivi("CFF1BIS", qte_commandee=1, designation="100 colliers à embase RAMTUB 16/32")]
    bl = [_bl("59210", quantite_livree=1, prix_net=19.0, designation="100 COLLIERS A EMBASE RAMTUB 16/32")]

    [c] = apparier(bl, suivi)

    assert c.statut == Statut.INCONNU


def test_apparier_alias_referentiel_propose_reste_a_confirmer(tmp_path):
    """Une correspondance simplement PROPOSÉE par le référentiel (candidat
    structurel plausible, ex. préfixe marque connu — pas encore confirmée)
    ne fusionne JAMAIS automatiquement : reste "à confirmer", exactement
    comme les autres replis approximatifs.

    Référence SANS chiffre (des deux côtés) pour isoler le seul chemin
    référentiel : avec des chiffres, coeur_numerique() aurait déjà fait
    matcher "LEGNILOE"/"NILOE" via le repli existant (préfixe/zéro de tête),
    ce qui n'est PAS ce qu'on veut tester ici."""

    referentiel = Referentiel(tmp_path / "moteur")
    referentiel.cx.execute(
        "INSERT INTO articles VALUES (?,?,?,?,?,?,?,?)",
        ("NILOE", "NILOE", "COMINTER", "Legrand", "Interrupteur simple Legrand Niloe", "", 4.20, "import"),
    )
    referentiel.prefixes = {"Legrand": "LEG"}
    referentiel.cx.commit()

    suivi = [_suivi("NILOE", qte_commandee=5, designation="Interrupteur simple Legrand Niloe")]
    bl = [_bl("LEGNILOE", quantite_livree=5, prix_net=4.20, designation="Interrupteur simple Legrand Niloe")]

    [c] = apparier(bl, suivi, referentiel=referentiel)

    assert c.statut == Statut.A_CONFIRMER
    assert "Référentiel articles" in c.raisons[0]

    referentiel.fermer()
