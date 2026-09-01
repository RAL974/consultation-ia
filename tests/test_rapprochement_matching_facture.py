"""
moteur.rapprochement.matching_facture — sur des LigneFacture/LigneSuiviFacture
synthétiques (pas besoin d'un vrai classeur Excel, comme
tests/test_rapprochement_matching.py côté BL)."""

from moteur.rapprochement.matching_facture import (
    LigneSuiviFacture,
    StatutFacture,
    apparier_facture,
    colonnes_facture_disponibles,
)
from moteur.rapprochement.modele_facture import LigneFacture


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


def test_apparier_a_confirmer_autre_numero_facture_deja_present():
    """Une AUTRE facture est déjà enregistrée sur cette ligne — jamais
    écrasée silencieusement, doublon/litige à trancher à la main."""

    lf = _ligne_facture()
    ls = _ligne_suivi(numero_facture="F0", qte_facturee=10.0, pu_facture=5.0)

    [c] = apparier_facture([lf], [ls], numero_facture="F1")

    assert c.statut is StatutFacture.A_CONFIRMER
    assert any("F0" in r for r in c.raisons)


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
