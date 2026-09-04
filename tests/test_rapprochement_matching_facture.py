"""
moteur.rapprochement.matching_facture — sur des LigneFacture/LigneSuiviFacture
synthétiques (pas besoin d'un vrai classeur Excel, comme
tests/test_rapprochement_matching.py côté BL). Un test de bout en bout sur
un vrai PDF (facture_109_362840_multi_bl_meme_ref.pdf) verrouille en plus
l'agrégation multi-BL (session S0, correction 1a) sur données réelles."""

from pathlib import Path

import pytest

from conftest import ROOT
from moteur.lecture_pdf import lire_pdf
from moteur.fournisseurs.coredime import parse_facture_coredime
from moteur.fournisseurs.dist109 import parse_facture_109
from moteur.rapprochement.matching_facture import (
    CauseFacture,
    LigneSuiviFacture,
    StatutFacture,
    agreger_lignes_meme_reference,
    apparier_facture,
    charger_frais_fournisseurs,
    colonnes_facture_disponibles,
    est_bdc_manuel_24x,
    lire_lignes_commande_facture,
)
from moteur.rapprochement.modele_facture import LigneFacture
from moteur.rapprochement.pipeline_bl import trouver_fichier_suivi_vivant

FIXTURES = Path(__file__).parent / "fixtures"


def _ligne_suivi(**kwargs):
    defaut = dict(
        ligne_excel=2, reference="ABC123", designation="Article test",
        qte_commandee=10.0, qte_livree=10.0, tarif_bl=5.0, tarif_convenu=None,
        numero_commande="", numero_facture=None, date_facture=None,
        qte_facturee=None, pu_facture=None,
    )
    defaut.update(kwargs)
    return LigneSuiviFacture(**defaut)


def _ligne_facture(**kwargs):
    defaut = dict(reference_fournisseur="ABC123", designation="Article test", quantite_facturee=10.0, prix_unitaire_ht=5.0)
    defaut.update(kwargs)
    return LigneFacture(**defaut)


def test_colonnes_facture_disponibles():
    assert colonnes_facture_disponibles(
        {"N° facture": 0, "Date facture": 1, "Qté facturée": 2, "PU facturé": 3, "Montant facturé HT": 4}
    )
    assert not colonnes_facture_disponibles(
        {"N° facture": 0, "Date facture": 1, "Qté facturée": 2, "PU facturé": 3}  # "Montant facturé HT" manquante
    )
    assert not colonnes_facture_disponibles({"N° facture": 0})
    assert not colonnes_facture_disponibles({})


def test_apparier_sur_quand_tout_correspond():
    lf = _ligne_facture()
    ls = _ligne_suivi()

    [c] = apparier_facture([lf], [ls], numero_facture="F1")

    assert c.statut is StatutFacture.SUR
    assert c.raisons == []


def test_apparier_sur_quel_que_soit_lecart_de_prix():
    """Décision explicite de l'acheteur (session F4, suite) : "il faut
    écrire tout ce qui apparaît sur les factures rapprochables à des
    commandes, quel que soit le prix" — revient entièrement sur le "aucune
    tolérance" du cadrage initial (une tolérance de 0,01€ avait été
    envisagée quelques secondes plus tôt dans la même session, puis
    dépassée par cette décision plus large). Un écart important (5,0€ vs
    12,0€, pas juste du bruit d'arrondi) n'empêche plus l'écriture — le PU
    facturé est simplement écrit tel quel, l'écart reste visible dans le
    Suivi via les colonnes Tarif BL/Tarif convenu à côté."""

    lf = _ligne_facture(prix_unitaire_ht=12.0)
    ls = _ligne_suivi(tarif_bl=5.0)

    [c] = apparier_facture([lf], [ls], numero_facture="F1")

    assert c.statut is StatutFacture.SUR
    assert c.raisons == []


def test_apparier_sur_sans_aucun_tarif_de_reference():
    """Même principe sans aucun tarif de référence du tout (ni Tarif BL, ni
    Tarif convenu) — rien à comparer, mais ce n'est plus un motif de
    blocage non plus."""

    lf = _ligne_facture()
    ls = _ligne_suivi(tarif_bl=None, tarif_convenu=None)

    [c] = apparier_facture([lf], [ls], numero_facture="F1")

    assert c.statut is StatutFacture.SUR
    assert c.raisons == []


def test_apparier_a_confirmer_qte_facturee_differente_qte_livree():
    lf = _ligne_facture(quantite_facturee=8.0)
    ls = _ligne_suivi(qte_livree=10.0)

    [c] = apparier_facture([lf], [ls], numero_facture="F1")

    assert c.statut is StatutFacture.A_CONFIRMER
    assert any("Qté facturée" in r for r in c.raisons)


def test_apparier_sur_facture_arrivee_avant_bl_quel_que_soit_le_prix():
    """Décision explicite de l'acheteur (session F4, Coredime) : "ce ne sont
    pas des factures non parvenues puisqu'on les a reçues ! [...] il y a les
    BL manquants là-dedans, ils sont signés" — une facture reçue avant que
    son BL soit rapproché dans le Suivi (Qté livrée encore à 0) N'EST PLUS un
    motif de blocage : le contrôle de quantité est simplement ignoré (rien à
    comparer). Combiné à "quel que soit le prix" (voir plus haut) : même un
    écart de prix important n'empêche pas l'écriture non plus."""

    lf = _ligne_facture(prix_unitaire_ht=12.0)
    ls = _ligne_suivi(qte_livree=0.0)

    [c] = apparier_facture([lf], [ls], numero_facture="F1")

    assert c.statut is StatutFacture.SUR
    assert c.raisons == []


def test_apparier_sur_facture_arrivee_avant_bl_sans_tarif_de_reference():
    """Même situation (Qté livrée à 0) et SANS aucun tarif de référence (ni
    Tarif BL, ni Tarif convenu) — les deux relaxations se combinent, SUR."""

    lf = _ligne_facture()
    ls = _ligne_suivi(qte_livree=0.0, tarif_bl=None, tarif_convenu=None)

    [c] = apparier_facture([lf], [ls], numero_facture="F1")

    assert c.statut is StatutFacture.SUR
    assert c.raisons == []


def test_apparier_deja_a_jour_meme_numero_facture_deja_enregistre():
    """Idempotence : ce même n° de facture est déjà écrit sur la ligne —
    rien à réécrire (double dépôt/traitement du même document)."""

    lf = _ligne_facture()
    ls = _ligne_suivi(numero_facture="F1", qte_facturee=10.0, pu_facture=5.0)

    [c] = apparier_facture([lf], [ls], numero_facture="F1")

    assert c.statut is StatutFacture.DEJA_A_JOUR


def test_apparier_a_confirmer_autre_facture_deja_complete_sur_la_ligne():
    """P1 : une AUTRE facture a déjà facturé toute la quantité livrée —
    l'ajout porterait le cumul facturé au-delà du livré : jamais écrit
    automatiquement, cause QTE_SUPERIEURE (garde-fou double facturation),
    la facture déjà présente citée en clair."""

    lf = _ligne_facture()
    ls = _ligne_suivi(numero_facture="F0", qte_facturee=10.0, pu_facture=5.0)

    [c] = apparier_facture([lf], [ls], numero_facture="F1")

    assert c.statut is StatutFacture.A_CONFIRMER
    assert c.cause is CauseFacture.QTE_SUPERIEURE
    assert any("F0" in r and "cumul" in r for r in c.raisons)


def test_apparier_sur_facturation_en_plusieurs_fois_qui_complete_le_livre():
    """P1 : une ligne livrée 10 déjà facturée 4 par F0 ; F1 facture les 6
    restants -> cumul exact, sûr (feuille Pièces : une ligne par document,
    Commandes[Qté facturée] = somme)."""

    lf = _ligne_facture(quantite_facturee=6.0)
    ls = _ligne_suivi(numero_facture="F0", qte_facturee=4.0, pu_facture=5.0)

    [c] = apparier_facture([lf], [ls], numero_facture="F1")

    assert c.statut is StatutFacture.SUR
    assert ls.numeros_factures == ["F0"]


def test_apparier_inconnu_sans_correspondance():
    lf = _ligne_facture(reference_fournisseur="XYZ999")
    ls = _ligne_suivi(reference="ABC123")

    [c] = apparier_facture([lf], [ls], numero_facture="F1")

    assert c.statut is StatutFacture.INCONNU
    assert c.ligne_suivi is None


def test_apparier_repli_reference_proche_reste_a_confirmer():
    """Repli référence proche (réutilisé depuis matching.py, voir bandeau
    du module) — jamais "sûr" automatiquement, même si le reste concorde."""

    lf = _ligne_facture(reference_fournisseur="ABC124")  # 1 caractère d'écart
    ls = _ligne_suivi(reference="ABC123")

    [c] = apparier_facture([lf], [ls], numero_facture="F1")

    assert c.statut is StatutFacture.A_CONFIRMER
    assert any("proche" in r.lower() for r in c.raisons)


def test_apparier_exact_nest_jamais_vole_par_un_repli():
    """Même garde-fou deux-passes que matching.apparier() (voir son
    bandeau) : une ligne de facture qui matche EXACTEMENT une ligne Suivi
    ne doit jamais se la faire voler par un repli approximatif d'une AUTRE
    ligne de facture traitée avant elle."""

    lf_repli = _ligne_facture(reference_fournisseur="ABC124")  # proche de ABC123 seulement
    lf_exact = _ligne_facture(reference_fournisseur="XYZ999")
    ls_proche = _ligne_suivi(ligne_excel=2, reference="ABC123")
    ls_exacte = _ligne_suivi(ligne_excel=3, reference="XYZ999")

    resultats = apparier_facture([lf_repli, lf_exact], [ls_proche, ls_exacte], numero_facture="F1")

    c_exact = resultats[1]
    assert c_exact.ligne_suivi.ligne_excel == 3
    assert c_exact.statut is StatutFacture.SUR


def test_apparier_sur_meme_si_tarif_bl_et_tarif_convenu_absents():
    """Le prix (et son absence de référence) n'intervient plus du tout dans
    _comparer_facture (voir test_apparier_sur_sans_aucun_tarif_de_reference
    plus haut) — Tarif BL/Tarif convenu restent lus sur LigneSuiviFacture
    pour d'autres usages (ex. l'exception Tarif BL depuis facture, voir
    pipeline_facture.py), mais plus comme condition de blocage ici."""

    lf = _ligne_facture(prix_unitaire_ht=7.0)
    ls = _ligne_suivi(tarif_bl=None, tarif_convenu=None)

    [c] = apparier_facture([lf], [ls], numero_facture="F1")

    assert c.statut is StatutFacture.SUR


# --- causes codées (session S0, correction 1e) ------------------------------


def test_apparier_qte_partielle_et_superieure_causes_distinctes():
    """QTE_PARTIELLE (facturé < déjà livré) et QTE_SUPERIEURE (facturé >
    déjà livré) sont deux causes DISTINCTES — même message "raisons",
    cause différente pour le compte rendu chiffré."""

    lf_partielle = _ligne_facture(quantite_facturee=8.0)
    ls = _ligne_suivi(qte_livree=10.0)
    [c] = apparier_facture([lf_partielle], [ls], numero_facture="F1")
    assert c.statut is StatutFacture.A_CONFIRMER
    assert c.cause is CauseFacture.QTE_PARTIELLE

    lf_superieure = _ligne_facture(quantite_facturee=12.0)
    ls2 = _ligne_suivi(qte_livree=10.0)
    [c2] = apparier_facture([lf_superieure], [ls2], numero_facture="F1")
    assert c2.statut is StatutFacture.A_CONFIRMER
    assert c2.cause is CauseFacture.QTE_SUPERIEURE


def test_apparier_deja_a_jour_parmi_plusieurs_numeros_de_facture():
    """P1 : la ligne porte déjà F0 et F1 (« F0; F1 » tel que la formule
    N° facture de Commandes l'affiche) — F1 redéposée ressort déjà à jour."""

    lf = _ligne_facture()
    ls = _ligne_suivi(numero_facture="F0; F1", qte_facturee=10.0, pu_facture=5.0)

    [c] = apparier_facture([lf], [ls], numero_facture="F1")

    assert c.statut is StatutFacture.DEJA_A_JOUR
    assert ls.numeros_factures == ["F0", "F1"]


def test_apparier_inconnu_a_la_cause_ref_inconnue():
    lf = _ligne_facture(reference_fournisseur="XYZ999")
    ls = _ligne_suivi(reference="ABC123")

    [c] = apparier_facture([lf], [ls], numero_facture="F1")

    assert c.statut is StatutFacture.INCONNU
    assert c.cause is CauseFacture.REF_INCONNUE


def test_apparier_repli_reference_proche_a_la_cause_cle_partielle():
    lf = _ligne_facture(reference_fournisseur="ABC124")
    ls = _ligne_suivi(reference="ABC123")

    [c] = apparier_facture([lf], [ls], numero_facture="F1")

    assert c.statut is StatutFacture.A_CONFIRMER
    assert c.cause is CauseFacture.CLE_PARTIELLE


# --- agrégation multi-BL même référence (session S0, correction 1a) --------


def test_agreger_lignes_meme_reference_somme_les_quantites_meme_pu():
    """Cas de base de l'agrégation : 2 lignes, même référence, même PU
    (ex. réel P03200 réparti sur 2 BL de Facture_362840.pdf) -> UNE seule
    ligne, quantités/montants sommés, n° de BL concaténés (détail
    conservé)."""

    lf1 = _ligne_facture(reference_fournisseur="P03200", quantite_facturee=100.0,
                          prix_unitaire_ht=0.26, montant_ht=26.0)
    lf1.numero_bl = "731835"
    lf2 = _ligne_facture(reference_fournisseur="P03200", quantite_facturee=100.0,
                          prix_unitaire_ht=0.26, montant_ht=26.0)
    lf2.numero_bl = "731846"

    lignes, refs_prix_differents = agreger_lignes_meme_reference([lf1, lf2])

    assert refs_prix_differents == set()
    [l] = lignes
    assert l.quantite_facturee == 200.0
    assert l.montant_ht == 52.0
    assert l.numero_bl == "731835 + 731846"


def test_agreger_lignes_meme_reference_ne_touche_pas_une_reference_unique():
    lf = _ligne_facture(reference_fournisseur="SEUL")
    lignes, refs_prix_differents = agreger_lignes_meme_reference([lf])
    assert lignes == [lf]
    assert refs_prix_differents == set()


def test_agreger_lignes_meme_reference_pu_different_pas_dagregation():
    """PU différent selon le BL -> AUCUNE agrégation pour ce groupe (jamais
    un prix deviné) — les 2 lignes restent séparées, leur clé est ajoutée à
    refs_prix_differents pour qu'apparier_facture() les fasse ressortir
    "à confirmer"."""

    lf1 = _ligne_facture(reference_fournisseur="REFX", prix_unitaire_ht=1.0)
    lf2 = _ligne_facture(reference_fournisseur="REFX", prix_unitaire_ht=1.5)

    lignes, refs_prix_differents = agreger_lignes_meme_reference([lf1, lf2])

    assert lignes == [lf1, lf2]
    assert len(refs_prix_differents) == 1


def test_apparier_agrege_avant_comparaison_a_la_qte_livree():
    """Bout en bout sur apparier_facture() : sans agrégation, chaque bloc
    (50 facturés) comparerait à la Qté livrée TOTALE (100) et ressortirait
    "à confirmer" à tort. Avec l'agrégation (50+50=100), SUR."""

    lf1 = _ligne_facture(reference_fournisseur="REFX", quantite_facturee=50.0, prix_unitaire_ht=2.0)
    lf1.numero_bl = "BL1"
    lf2 = _ligne_facture(reference_fournisseur="REFX", quantite_facturee=50.0, prix_unitaire_ht=2.0)
    lf2.numero_bl = "BL2"
    ls = _ligne_suivi(reference="REFX", qte_livree=100.0)

    [c] = apparier_facture([lf1, lf2], [ls], numero_facture="F1")

    assert c.statut is StatutFacture.SUR
    assert c.ligne_facture.quantite_facturee == 100.0


def test_apparier_prix_differents_meme_reference_a_confirmer_cause_dediee():
    """Qté livrée = 50 (pas 100) pour isoler le SEUL problème testé ici (le
    PU) — sans quoi la 1re ligne (qté facturée 50 vs 100 déjà livré,
    puisque non agrégée) déclencherait AUSSI un écart de quantité réel, et
    _comparer_facture garde alors la cause la plus spécifique déjà posée
    (QTE_PARTIELLE) plutôt que PRIX_DIFF_MEME_REF — comportement voulu
    (voir apparier_facture, "c.cause or CauseFacture.PRIX_DIFF_MEME_REF"),
    mais pas ce que CE test vérifie."""

    lf1 = _ligne_facture(reference_fournisseur="REFX", quantite_facturee=50.0, prix_unitaire_ht=2.0)
    lf2 = _ligne_facture(reference_fournisseur="REFX", quantite_facturee=50.0, prix_unitaire_ht=3.0)
    ls = _ligne_suivi(reference="REFX", qte_livree=50.0)

    resultats = apparier_facture([lf1, lf2], [ls], numero_facture="F1")

    # Une seule ligne Suivi disponible pour 2 lignes facture non agrégées :
    # la 1re la prend (SUR, forcé A_CONFIRMER par le garde-fou prix), la 2e
    # n'a plus de candidat -> INCONNU (jamais un 2e rattachement inventé,
    # voir apparier_facture — le garde-fou n'y touche alors pas, pour ne
    # jamais forcer A_CONFIRMER avec ligne_suivi=None).
    assert len(resultats) == 2
    c_confirmee = next(c for c in resultats if c.ligne_suivi is not None)
    assert c_confirmee.statut is StatutFacture.A_CONFIRMER
    assert c_confirmee.cause is CauseFacture.PRIX_DIFF_MEME_REF
    assert any("PU" in r or "prix" in r.lower() for r in c_confirmee.raisons)


def test_facture_362840_multi_bl_meme_reference_11_lignes_sures():
    """Test de bout en bout sur données RÉELLES (voir CLAUDE.md, session
    S0, correction 1a) : Facture_362840.pdf (109 Distribution, commande
    123.089) a 11 LIGNES BRUTES réparties sur 3 "Bon de livraison" —
    P03200 et F2U15RVVOO sont chacun répartis sur 2 BL différents.
    L'agrégation les ramène à 9 CORRESPONDANCES (une par référence
    DISTINCTE, jamais 2 écritures visant la même ligne Excel) : "362840 ->
    11/11 sûres" (cadrage de session) se vérifie donc au niveau des 11
    lignes brutes, entièrement couvertes par ces 9 correspondances, TOUTES
    sûres. Qté déjà livrée dans le Suivi = qté commandée réelle (lue sur le
    "DETAIL DE LA COMMANDE" de ce même PDF) car la commande est entièrement
    soldée."""

    texte = lire_pdf(FIXTURES / "facture_109_362840_multi_bl_meme_ref.pdf")
    f = parse_facture_109(texte)
    assert len(f.lignes) == 11  # verrou de non-régression sur le parsing

    quantites_suivi = {
        "P03200": 200.0, "20080043": 600.0, "F2U15RVVOO": 300.0,
        "10041540": 100.0, "10041940": 700.0, "10042724": 100.0,
        "10043324": 150.0, "10043924": 150.0, "PVCORANGE": 5.0,
    }
    lignes_suivi = [
        _ligne_suivi(ligne_excel=i + 2, reference=ref, designation="", qte_livree=qte, tarif_bl=None)
        for i, (ref, qte) in enumerate(quantites_suivi.items())
    ]

    correspondances = apparier_facture(f.lignes, lignes_suivi, numero_facture=f.numero_facture)

    # 9 références distinctes (P03200 et F2U15RVVOO agrégées depuis 2
    # lignes brutes chacune : 11 - 2 = 9) — jamais 11, une agrégation
    # réussie réduit le nombre de lignes à écrire, elle ne le préserve pas.
    assert len(correspondances) == 9
    assert all(c.statut is StatutFacture.SUR for c in correspondances)

    par_ref = {c.ligne_facture.reference_fournisseur: c.ligne_facture for c in correspondances}
    assert par_ref["P03200"].quantite_facturee == 200.0
    assert par_ref["P03200"].numero_bl == "731835 + 731846"
    assert par_ref["F2U15RVVOO"].quantite_facturee == 300.0
    assert par_ref["F2U15RVVOO"].numero_bl == "731835 + 731934"


# --- repli "premier token" : référence Suivi à suffixe libre (1b) ----------


def test_apparier_repli_premier_token_suffixe_libre():
    """Cas réel (session S0, Facture_6108234.pdf, commande M3.14.342) :
    Suivi "SIXGPCP35 PVC" vs facture "SIXGPCP35" — le premier terme du
    Suivi correspond exactement. Toujours "à confirmer"."""

    lf = _ligne_facture(reference_fournisseur="SIXGPCP35")
    ls = _ligne_suivi(reference="SIXGPCP35 PVC")

    [c] = apparier_facture([lf], [ls], numero_facture="F1")

    assert c.statut is StatutFacture.A_CONFIRMER
    assert c.cause is CauseFacture.CLE_PARTIELLE
    assert any("suffixe" in r.lower() for r in c.raisons)


def test_apparier_repli_premier_token_ignore_si_reference_suivi_sans_espace():
    """Une référence Suivi SANS espace n'est pas concernée par ce repli
    (déjà couverte par la comparaison exacte si elle correspond, sinon
    vraiment une référence différente) — jamais une comparaison "premier
    token" sur une référence qui n'a pas de suffixe du tout."""

    lf = _ligne_facture(reference_fournisseur="SIXGPCP35")
    ls = _ligne_suivi(reference="AUTREREF")

    [c] = apparier_facture([lf], [ls], numero_facture="F1")

    assert c.statut is StatutFacture.INCONNU


def test_apparier_repli_premier_token_ignore_si_plusieurs_candidats():
    """2 lignes Suivi partagent le même premier token -> ambigu, jamais un
    choix au hasard (même garde-fou que _repli_reference_proche)."""

    lf = _ligne_facture(reference_fournisseur="SIXGPCP35")
    ls1 = _ligne_suivi(ligne_excel=2, reference="SIXGPCP35 PVC")
    ls2 = _ligne_suivi(ligne_excel=3, reference="SIXGPCP35 AUTRE")

    [c] = apparier_facture([lf], [ls1, ls2], numero_facture="F1")

    assert c.statut is StatutFacture.INCONNU


# --- frais connus (session S0, correction 1c) -------------------------------


def test_apparier_frais_connu_jamais_bloquant_et_ne_consomme_pas_de_ligne_suivi():

    lf_frais = _ligne_facture(reference_fournisseur="ECO-23", quantite_facturee=10.0,
                               prix_unitaire_ht=0.08, montant_ht=0.8)
    lf_normale = _ligne_facture(reference_fournisseur="ABC123")
    ls = _ligne_suivi(reference="ABC123")

    frais_connus = {"COREDIME": {"ECO-23": "Éco-participation"}}

    resultats = apparier_facture(
        [lf_frais, lf_normale], [ls], numero_facture="F1",
        fournisseur="COREDIME", frais_connus=frais_connus,
    )

    assert len(resultats) == 2
    c_frais = next(c for c in resultats if c.ligne_facture is lf_frais)
    assert c_frais.statut is StatutFacture.FRAIS
    assert c_frais.cause is CauseFacture.FRAIS
    assert c_frais.ligne_suivi is None

    c_normale = next(c for c in resultats if c.ligne_facture is lf_normale)
    assert c_normale.statut is StatutFacture.SUR  # la ligne Suivi n'a pas été "mangée" par le frais


def test_apparier_frais_connu_uniquement_pour_le_bon_fournisseur():
    """Le whitelist est PAR FOURNISSEUR — une référence "ECO-23" connue
    chez COREDIME ne doit rien changer pour un autre fournisseur (jamais un
    frais générique appliqué à l'aveugle)."""

    lf = _ligne_facture(reference_fournisseur="ECO-23")
    ls = _ligne_suivi(reference="ECO-23")

    frais_connus = {"COREDIME": {"ECO-23": "Éco-participation"}}

    [c] = apparier_facture(
        [lf], [ls], numero_facture="F1", fournisseur="109 DISTRIBUTION", frais_connus=frais_connus,
    )

    assert c.statut is not StatutFacture.FRAIS


def test_charger_frais_fournisseurs_fichier_absent_dict_vide(tmp_path):
    assert charger_frais_fournisseurs(tmp_path / "inexistant.csv") == {}


def test_charger_frais_fournisseurs_lit_le_csv(tmp_path):

    chemin = tmp_path / "frais_fournisseurs.csv"
    chemin.write_text(
        "Fournisseur;Reference;Libelle\n"
        "COREDIME;ECO-23;Éco-participation\n"
        "COREDIME;9993;Livraison avion\n",
        encoding="utf-8",
    )

    frais = charger_frais_fournisseurs(chemin)

    assert frais["COREDIME"]["ECO-23"] == "Éco-participation"
    assert frais["COREDIME"]["9993"] == "Livraison avion"


# --- bon manuel "BC/BCN 24XXXX" (session S0, correction 1e) ----------------


def test_est_bdc_manuel_24x_reconnait_les_deux_formats_reels():
    assert est_bdc_manuel_24x(["BC 241766"])       # cas réel Facture_362777.pdf (109)
    assert est_bdc_manuel_24x(["BCN 241461"])      # cas réel 6100226.pdf (Coredime, "Réf.: BCN 241461")


def test_est_bdc_manuel_24x_faux_sur_un_format_suivi_normal():
    assert not est_bdc_manuel_24x(["123.089"])
    assert not est_bdc_manuel_24x(["M3.14.342"])
    assert not est_bdc_manuel_24x([])
    assert not est_bdc_manuel_24x([""])


# --- résiduel unique : "substitution probable" (étape 1) -------------------


def test_apparier_residuel_unique_substitution_probable():
    """1 ligne facturée exacte (ABC123) + 1 ligne facturée sans AUCUN
    rapport textuel/numérique avec la seule ligne Suivi restante (référence
    totalement différente) -> "substitution probable", à confirmer, jamais
    sûre. Même quantité des deux côtés (10.0), condition nécessaire."""

    lf_exacte = _ligne_facture(reference_fournisseur="ABC123", quantite_facturee=5.0, prix_unitaire_ht=1.0)
    lf_residuelle = _ligne_facture(reference_fournisseur="ZZZ999", quantite_facturee=10.0, prix_unitaire_ht=0.37)
    ls_exacte = _ligne_suivi(ligne_excel=2, reference="ABC123", qte_livree=5.0, tarif_bl=1.0)
    ls_residuelle = _ligne_suivi(
        ligne_excel=3, reference="5120", designation="ICT 20 BLEU TURBO G-ROUL 100M",
        qte_livree=10.0, tarif_bl=0.37,
    )

    corr = apparier_facture([lf_exacte, lf_residuelle], [ls_exacte, ls_residuelle], numero_facture="F1")

    par_ref = {c.ligne_facture.reference_fournisseur: c for c in corr}
    assert par_ref["ABC123"].statut is StatutFacture.SUR

    c = par_ref["ZZZ999"]
    assert c.statut is StatutFacture.A_CONFIRMER
    assert c.cause is CauseFacture.SUBSTITUTION_PROBABLE
    assert c.ligne_suivi.reference == "5120"
    assert any("substitution probable" in r.lower() for r in c.raisons)


def _scenario_residuel(**ecart):
    """1 paire exacte (ABC123, jamais concernée) + 1 paire résiduelle
    (ZZZ999 facture / 5120 Suivi, sans aucun rapport textuel) — mêmes
    qté/PU par défaut (concordants), `ecart` permute un ou plusieurs
    champs du côté facture ou Suivi pour tester un garde-fou précis. Le
    couple exact est nécessaire pour que `len(lignes_a_apparier) > 1`
    (garde-fou "vraie élimination", voir _residuel_unique) — un scénario à
    1 seule ligne de chaque côté ne peut structurellement plus déclencher
    ce repli."""

    lf_exacte = _ligne_facture(reference_fournisseur="ABC123", quantite_facturee=5.0, prix_unitaire_ht=1.0)
    ls_exacte = _ligne_suivi(ligne_excel=2, reference="ABC123", qte_livree=5.0, tarif_bl=1.0)

    defaut_lf = dict(reference_fournisseur="ZZZ999", quantite_facturee=10.0, prix_unitaire_ht=0.37)
    defaut_ls = dict(ligne_excel=3, reference="5120", qte_livree=10.0, tarif_bl=0.37, numero_facture=None)
    defaut_lf.update(ecart.get("lf", {}))
    defaut_ls.update(ecart.get("ls", {}))

    lf_residuelle = _ligne_facture(**defaut_lf)
    ls_residuelle = _ligne_suivi(**defaut_ls)

    corr = apparier_facture([lf_exacte, lf_residuelle], [ls_exacte, ls_residuelle], numero_facture="F1")
    return {c.ligne_facture.reference_fournisseur: c for c in corr}["ZZZ999"]


def test_apparier_residuel_unique_ignore_si_ligne_suivi_deja_facturee():
    """La seule ligne Suivi restante porte déjà un n° de facture (d'une
    AUTRE facture) : jamais proposée comme substitution — "1 ligne Suivi
    SANS facture" est une condition stricte."""

    c = _scenario_residuel(ls={"numero_facture": "AUTRE-FACTURE"})

    assert c.statut is StatutFacture.INCONNU
    assert c.cause is CauseFacture.REF_INCONNUE


def test_apparier_residuel_unique_ignore_si_quantite_differente():
    c = _scenario_residuel(ls={"qte_livree": 25.0})

    assert c.statut is StatutFacture.INCONNU


def test_apparier_residuel_unique_ignore_si_prix_connus_trop_differents():
    """Même quantité, mais PU connus des deux côtés et trop éloignés
    (> 0,02€) : le résiduel unique n'efface pas ce garde-fou."""

    c = _scenario_residuel(lf={"prix_unitaire_ht": 5.0}, ls={"tarif_bl": 1.0})

    assert c.statut is StatutFacture.INCONNU


def test_apparier_residuel_unique_accepte_ecart_prix_sous_le_seuil():
    c = _scenario_residuel(lf={"prix_unitaire_ht": 1.01}, ls={"tarif_bl": 1.0})

    assert c.statut is StatutFacture.A_CONFIRMER
    assert c.cause is CauseFacture.SUBSTITUTION_PROBABLE


def test_apparier_residuel_unique_ignore_si_une_seule_ligne_de_chaque_cote():
    """Garde-fou "vraie élimination" : une commande à 1 SEULE ligne facture
    dès le départ (aucun autre pair à avoir résolu pour se convaincre que
    facture et commande se correspondent) ne déclenche jamais le résiduel
    unique, même si qté/PU concordent par coïncidence — reste "inconnu"
    (cas réel : les tests génériques de ce module utilisent tous la même
    qté/PU par défaut des deux côtés, sans rapport avec ce garde-fou)."""

    lf = _ligne_facture(reference_fournisseur="ZZZ999", quantite_facturee=10.0)
    ls = _ligne_suivi(reference="5120", qte_livree=10.0)

    [c] = apparier_facture([lf], [ls], numero_facture="F1")

    assert c.statut is StatutFacture.INCONNU


def test_apparier_residuel_unique_ignore_si_plusieurs_inconnues():
    """2 lignes facture inconnues pour 1 seule ligne Suivi restante :
    ambigu, jamais un choix au hasard entre les deux."""

    lf1 = _ligne_facture(reference_fournisseur="ZZZ999", quantite_facturee=10.0)
    lf2 = _ligne_facture(reference_fournisseur="YYY888", quantite_facturee=10.0)
    ls = _ligne_suivi(reference="5120", qte_livree=10.0)

    corr = apparier_facture([lf1, lf2], [ls], numero_facture="F1")

    assert all(c.statut is StatutFacture.INCONNU for c in corr)


def test_apparier_residuel_unique_ignore_si_plusieurs_disponibles():
    """1 seule ligne facture inconnue (mais une 2e, exacte, prouve qu'il y
    a bien eu élimination), 2 lignes Suivi encore disponibles : ambigu,
    jamais un choix au hasard."""

    # Reconstruit le scénario à la main (le helper _scenario_residuel n'a
    # qu'1 ligne résiduelle Suivi) pour ajouter une 2e ligne disponible.
    lf_exacte = _ligne_facture(reference_fournisseur="ABC123", quantite_facturee=5.0, prix_unitaire_ht=1.0)
    lf_residuelle = _ligne_facture(reference_fournisseur="ZZZ999", quantite_facturee=10.0, prix_unitaire_ht=0.37)
    ls_exacte = _ligne_suivi(ligne_excel=2, reference="ABC123", qte_livree=5.0, tarif_bl=1.0)
    ls1 = _ligne_suivi(ligne_excel=3, reference="5120", qte_livree=10.0, tarif_bl=0.37)
    ls2 = _ligne_suivi(ligne_excel=4, reference="6130", qte_livree=10.0, tarif_bl=0.37)

    corr = apparier_facture([lf_exacte, lf_residuelle], [ls_exacte, ls1, ls2], numero_facture="F1")
    c = {c.ligne_facture.reference_fournisseur: c for c in corr}["ZZZ999"]

    assert c.statut is StatutFacture.INCONNU


@pytest.mark.skipif(
    trouver_fichier_suivi_vivant(ROOT) is None,
    reason="Classeur Suivi commandes VIVANT introuvable depuis ce poste",
)
def test_apparier_residuel_unique_sur_la_vraie_piece_6108234():
    """Bout en bout sur données réelles (facture_coredime_6108234_suffixe.pdf,
    commande M3.14.342) : après le repli premier-token (SIXGPCP35 ->
    SIXGPCP35 PVC) et les correspondances exactes/cœur-numérique déjà
    écrites lors d'une session précédente, il ne reste plus qu'une seule
    ligne facture inconnue ("LEG06620") et qu'une seule ligne Suivi non
    facturée ("5120", même désignation "ICT 20 BLEU TURBO G-ROUL 100M",
    même quantité 100, même tarif 0,37€) — voir CLAUDE.md, étape 1."""

    texte = lire_pdf(FIXTURES / "facture_coredime_6108234_suffixe.pdf")
    f = parse_facture_coredime(texte)

    suivi = trouver_fichier_suivi_vivant(ROOT)
    lignes_suivi = lire_lignes_commande_facture(suivi, "COREDIME", "M3.14.342")

    corr = apparier_facture(f.lignes, lignes_suivi, numero_facture=f.numero_facture)

    par_ref = {c.ligne_facture.reference_fournisseur: c for c in corr}
    c = par_ref["LEG06620"]

    assert c.statut is StatutFacture.A_CONFIRMER
    assert c.cause is CauseFacture.SUBSTITUTION_PROBABLE
    # La référence "5120" est purement numérique : Excel/openpyxl la relit
    # comme un int, pas une chaîne (déjà le cas ailleurs dans ce projet).
    assert str(c.ligne_suivi.reference) == "5120"
    assert c.ligne_suivi.qte_livree == 100.0
