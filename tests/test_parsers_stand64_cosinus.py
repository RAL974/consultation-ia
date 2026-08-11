"""
STAND 64 — 2e PDF réel (chantier Cosinus), avec le placeholder ZARTICLENP.

Ce devis (contrairement à tests/fixtures/stand64.pdf) n'a AUCUNE référence
propre par ligne pour la plupart des articles : le champ référence imprimé
est un placeholder générique ("ZARTICLENP" = article non prédéfini par le
commercial, absent du catalogue fournisseur ; "Alternative:ZARTICLENP"
pour les propositions alternatives). Constaté en le rejouant tel quel :
34 lignes -> seulement 5 clés distinctes au comparateur, la plupart des
offres s'écrasant silencieusement entre elles.

La vraie référence fabricant, quand elle existe, est alors donnée
uniquement entre parenthèses en fin de désignation (ex.
"...(KUBIA-ART00031180)"). moteur/fournisseurs/stand64.py l'extrait
désormais dans ce cas (_reference_reelle) — ce test fige la sortie
complète avec l'extraction active : 27 références distinctes sur 34
lignes (les doublons restants sont légitimes, le même luminaire réutilisé
à plusieurs endroits du chantier)."""

from moteur.fournisseurs.stand64 import parse_stand64
from moteur.modele import Article

from conftest import FIXTURES


def _texte(nom):
    import fitz
    doc = fitz.open(FIXTURES / nom)
    return "\n".join(p.get_text() for p in doc)


def test_parse_stand64_cosinus_extrait_la_reference_integree():
    articles = parse_stand64(_texte("stand64_cosinus.pdf"))

    assert articles == [
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='WESTI-73044', reference_distributeur='', designation="BRASSEUR D'AIR : COMET III 5 PALES Ø 132 BLANC+ERABLE / BLANC + chainette", quantite=23.0, unite='UN', prix_brut=91.0, prix_net=91.0, montant=2093.0, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031180', reference_distributeur='', designation='TYPE A1 - SUSPENSION CIRCULAIRE : SD-WOOD RING SUSPENSION Ø700MM 70W 8000LM 4000°K IP20 BOIS (KUBIA-ART00031180)', quantite=3.0, unite='UN', prix_brut=2145.33, prix_net=2145.33, montant=6435.99, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031182', reference_distributeur='', designation='TYPE A2 - SUSPENSION CIRCULAIRE DECORATIVE : SD-WOOD ROUND SUSPENSION Ø850MM 109W 10350LM 4000°K IP20 BOIS MASSIF CHENE (KUBIA-ART00031182)', quantite=1.0, unite='UN', prix_brut=3714.72, prix_net=3714.72, montant=3714.72, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00029865', reference_distributeur='', designation='TYPE A3 - DOWNLIGHT ENCASTRÉ DALLE ACOUSTIQUE : DL-BACKDOOR DOWNLIGHT 8W 1000LM 4000°K UGR<17 IP54 ON-OFF (KUBIA-ART00029865)', quantite=26.0, unite='UN', prix_brut=39.19, prix_net=39.19, montant=1018.94, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00029833', reference_distributeur='', designation='TYPE A4 - DOWNLIGHT ENCASTRÉ FAUX PLAFOND : DL-BACKDOOR DOWNLIGHT 14W 1750LM 4000°K UGR<17 IP54 ON-OFF BLANC (KUBIA-ART00029833)', quantite=9.0, unite='UN', prix_brut=50.95, prix_net=50.95, montant=458.55, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031184', reference_distributeur='', designation='SD-WOOD ROUND SUSPENSION Ø650MM 76W 7750LM 4000°K IP20 BOIS MASSIF CHENE (KUBIA-ART00031184)', quantite=2.0, unite='UN', prix_brut=2702.54, prix_net=2702.54, montant=5405.08, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00030917', reference_distributeur='', designation='TYPE A6 - SUSPENSION CYLINDRIQUE : SD-XLINE SUSPENSION TUBULAIRE Ø46MM H250MM 10W 750LM 3000°K IP20 IK03 CLI ALU NOIR (KUBIA-ART00030917)', quantite=5.0, unite='UN', prix_brut=166.56, prix_net=166.56, montant=832.8, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00030923', reference_distributeur='', designation='TYPE A7 - SUSPENSION LINÉAIRE : LS-PLUGIN N SUSPENSION 600MM 13W 1090LM 4000°K  IP40 VASQUE OPALE - RAL SUR MESURE (KUBIA-ART00030923)', quantite=20.0, unite='UN', prix_brut=253.9, prix_net=253.9, montant=5078.0, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031208', reference_distributeur='', designation='LS-PLUGIN XJ/ML SUSPENSION 600MM 20W 4000°K  IP40 VASQUE OPALE - RAL BLANC/NOIR/GRIS (KUBIA-ART00031208)', quantite=20.0, unite='UN', prix_brut=133.25, prix_net=133.25, montant=2665.0, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='ACB-A2033070B', reference_distributeur='', designation='TYPE A8 - APPLIQUE MURALE : KOWA APPLIQUE LED Ø136X100mm 2X5.5W 1060LM 2700-3000°K IP65 IK06 ON/OFF FINITION BLANC UP&DOWN', quantite=4.0, unite='UN', prix_brut=82.33, prix_net=82.33, montant=329.32, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031186', reference_distributeur='', designation='TYPE A9 - SUSPENSION LINÉAIRE SD-WOOD LINEAR 50 SUSPENSION 27W 2750LM 4000°K DIRECT BOIS MASSIF CHENE (KUBIA-ART00031186)', quantite=1.0, unite='UN', prix_brut=1559.16, prix_net=1559.16, montant=1559.16, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00020341', reference_distributeur='', designation='SD-WOOD BALL LED SUSPENSION Ø180MM 12W 1150LM 4000°K CLII (KUBIA-ART00020341)', quantite=3.0, unite='UN', prix_brut=841.77, prix_net=841.77, montant=2525.31, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='YUANH-YL18-36W-1220-CCT-SENSOR', reference_distributeur='', designation='TYPE C1 - RÉGLETTE ÉTANCHE : PLAFONNIER ETANCHE LED 120cm 24/30/(36)/42W 140lm/W 3000/4000/5700°K IK08 IP66 + DETECTEUR ON/OFF*', quantite=2.0, unite='UN', prix_brut=44.0, prix_net=44.0, montant=88.0, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031199', reference_distributeur='', designation='TYPE S1 - SUSPENSION DÉCORATIVE : SD-WIDGET 3140 SUSPENSION CINTRÉ S RAYON 2000MM 140W 11200LM 4000°K SATINÉ NOIR  NOIR (KUBIA-ART00031199)', quantite=4.0, unite='UN', prix_brut=2057.58, prix_net=2057.58, montant=8230.32, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00027277', reference_distributeur='', designation='SD-WIDGET XJ SUSPENSION COURBE 48W 4800LM 4000°K DIFFUSEUR OPALE PROFILE NOIR (KUBIA-ART00027277)', quantite=8.0, unite='UN', prix_brut=715.68, prix_net=715.68, montant=5725.44, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='ACB-A343227OC', reference_distributeur='', designation='TYPE R1 - RÉGLETTE MIROIR : ALDO 830 APP 32W 2330LM 3000°K-4000°K  IP44 OPAL ET CHROME (ACB-A343227OC)', quantite=2.0, unite='UN', prix_brut=130.82, prix_net=130.82, montant=261.64, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031190', reference_distributeur='', designation='TYPE L1 - SUSPENSION : CHIFFRAGE SELON IMPLANTATION PLAN  A DETERMINER SELON IMPLANTATION REELLE LIGNE DE 3200MM - CIRCULATION ENTRE LE REFECTOIRE  : LS-PLUGIN N DEBUT SUSPENSION 1692MM 30W 4000°K IP40 PROFILE NOIR (KUBIA-ART00031190)', quantite=1.0, unite='UN', prix_brut=410.66, prix_net=410.66, montant=410.66, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031191', reference_distributeur='', designation='LS-PLUGIN N FIN SUSPENSION 1412MM 25W 4000°K IP40 PROFILE NOIR (KUBIA-ART00031191)', quantite=1.0, unite='UN', prix_brut=270.93, prix_net=270.93, montant=270.93, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031192', reference_distributeur='', designation="LIGNE DE 9000MM - EN FOND D'OPEN SPACE : LS-PLUGIN N DEBUT SUSPENSION 40W 4000°K IP40 PROFILE NOIR (KUBIA-ART00031192)", quantite=1.0, unite='UN', prix_brut=468.6, prix_net=468.6, montant=468.6, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031197', reference_distributeur='', designation='LS-PLUGIN N MILIEU  SUSPENSION 40W 4000°K IP40 PROFILE NOIR (KUBIA-ART00031197)', quantite=2.0, unite='UN', prix_brut=429.4, prix_net=429.4, montant=858.8, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031198', reference_distributeur='', designation='LS-PLUGIN N FIN SUSPENSION 40W 4000°K IP40 PROFILE NOIR (KUBIA-ART00031198)', quantite=1.0, unite='UN', prix_brut=390.21, prix_net=390.21, montant=390.21, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031192', reference_distributeur='', designation="LIGNE DE 9000MM -  DERIERE LES BUREAUX A COTE DE L'ESPACE DETENTE LS-PLUGIN N DEBUT SUSPENSION 40W 4000°K IP40 PROFILE NOIR (KUBIA-ART00031192)", quantite=1.0, unite='UN', prix_brut=468.6, prix_net=468.6, montant=468.6, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031197', reference_distributeur='', designation='LS-PLUGIN N MILIEU  SUSPENSION 40W 4000°K IP40 PROFILE NOIR (KUBIA-ART00031197)', quantite=2.0, unite='UN', prix_brut=429.4, prix_net=429.4, montant=858.8, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031198', reference_distributeur='', designation='LS-PLUGIN N FIN SUSPENSION 40W 4000°K IP40 PROFILE NOIR (KUBIA-ART00031198)', quantite=1.0, unite='UN', prix_brut=390.21, prix_net=390.21, montant=390.21, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031192', reference_distributeur='', designation='CIRCULATION PRINCIPALE : LS-PLUGIN N DEBUT SUSPENSION 40W 4000°K IP40 PROFILE NOIR (KUBIA-ART00031192)', quantite=1.0, unite='UN', prix_brut=468.6, prix_net=468.6, montant=468.6, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031193', reference_distributeur='', designation='LS-PLUGIN N MILIEU  SUSPENSION 13W 4000°K IP40 PROFILE NOIR (KUBIA-ART00031193)', quantite=1.0, unite='UN', prix_brut=269.23, prix_net=269.23, montant=269.23, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031197', reference_distributeur='', designation='LS-PLUGIN N MILIEU  SUSPENSION 21W 4000°K IP40 PROFILE NOIR (KUBIA-ART00031197)', quantite=1.0, unite='UN', prix_brut=277.75, prix_net=277.75, montant=277.75, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031195', reference_distributeur='', designation='LS-PLUGIN N MILIEU  SUSPENSION 30W 4000°K IP40 PROFILE NOIR (KUBIA-ART00031195)', quantite=1.0, unite='UN', prix_brut=374.88, prix_net=374.88, montant=374.88, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031196', reference_distributeur='', designation='LS-PLUGIN N MILIEU  SUSPENSION 35W 4000°K IP40 PROFILE NOIR (KUBIA-ART00031196)', quantite=2.0, unite='UN', prix_brut=410.66, prix_net=410.66, montant=821.32, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031197', reference_distributeur='', designation='LS-PLUGIN N MILIEU  SUSPENSION 40W 4000°K IP40 PROFILE NOIR (KUBIA-ART00031197)', quantite=3.0, unite='UN', prix_brut=429.4, prix_net=429.4, montant=1288.2, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031204', reference_distributeur='', designation='LS-PLUGIN N FIN  SUSPENSION 17W 4000°K IP40 PROFILE NOIR (KUBIA-ART00031204)', quantite=1.0, unite='UN', prix_brut=245.37, prix_net=245.37, montant=245.37, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031198', reference_distributeur='', designation='LS-PLUGIN N FIN  SUSPENSION 40W 4000°K IP40 PROFILE NOIR (KUBIA-ART00031198)', quantite=3.0, unite='UN', prix_brut=390.21, prix_net=390.21, montant=1170.63, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031200', reference_distributeur='', designation='LS-PLUGIN N JONCTION L 597X597 21W 4000°K IP40 PROFILE NOIR (KUBIA-ART00031200)', quantite=1.0, unite='UN', prix_brut=253.9, prix_net=253.9, montant=253.9, disponibilite=''),
        Article(fournisseur='STAND 64', devis='73650', reference_fournisseur='KUBIA-ART00031202', reference_distributeur='', designation='LS-PLUGIN N JONCTION T 1154X597 35W 4000°K IP40 PROFILE NOIR (KUBIA-ART00031202)', quantite=2.0, unite='UN', prix_brut=478.82, prix_net=478.82, montant=957.64, disponibilite=''),
    ]

    # Plus aucune référence "placeholder" brute dans la sortie : soit un
    # vrai code intégré a été extrait, soit la référence imprimée était
    # déjà réelle (WESTI-73044, ACB-A2033070B, YUANH-...).
    assert not any("ZARTICLE" in a.reference_fournisseur.upper() for a in articles)

    # 34 lignes, 27 clés distinctes (les doublons restants sont légitimes :
    # le même luminaire réutilisé à plusieurs endroits du chantier — pas
    # une perte de données).
    assert len(articles) == 34
    assert len({a.reference_fournisseur for a in articles}) == 27
