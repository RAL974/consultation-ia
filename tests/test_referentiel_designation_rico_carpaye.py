"""
Régression réelle (consultation Rico Carpaye, août 2026) : besoin de
9 lignes SANS référence face à 4 devis (25 lignes) -> 0 rapprochement,
0 proposition dans A_confirmer.xlsx, tout en "Hors besoin".

Trois causes dans _tokens :
  1. accents non pliés -> "dérivations" tokenisé "RIVATIONS",
     "câblage" -> "BLAGE" ;
  2. singulier/pluriel non normalisé -> EMBOUTS != EMBOUT ;
  3. nombres/unités non normalisés -> "1.5mm" != "1.5 mm²",
     virgule décimale française "1,5mm²" coupée en "5MM".

Et une limite structurelle : "inter sectionneur 2x40A Schneider" vs
"ACTI9 ISW 2P 40A" ont un Jaccard de 0 — seuls les attributs techniques
(calibre + pôles) peuvent les rapprocher, comme l'étiquette "Type" le
fait déjà pour les bordereaux architecte.
"""

import pytest

from moteur.referentiel import (
    Referentiel,
    _analyser,
    _attributs_contradictoires,
    _tokens,
)


@pytest.fixture
def referentiel(tmp_path):
    r = Referentiel(tmp_path / "moteur")
    yield r
    r.fermer()


# ----------------------------------------------------------------------
# Normalisation des tokens
# ----------------------------------------------------------------------
def test_tokens_plie_les_accents():
    # Avant correction : {"RIVATIONS", "TANCHE", ...} — mots mutilés.
    assert {"BOITE", "DERIVATION", "ETANCHE", "LISSE"} <= _tokens(
        "Boite dérivations étanche lisse"
    )


def test_tokens_singulier_et_nombre_unite():
    t = _tokens("Sachet  embouts\xa0câblage 1.5mm simple")
    assert {"SACHET", "EMBOUT", "CABLAGE", "1.5", "MM", "SIMPLE"} <= t
    # La virgule décimale française et le ² ne cassent plus le token.
    assert {"EMBOUT", "CABLAGE", "1.5"} <= _tokens(
        "Embout de câblage isolé noir 1,5mm² X8"
    )


# ----------------------------------------------------------------------
# Attributs techniques
# ----------------------------------------------------------------------
def test_attributs_calibre_et_poles():
    _, a = _analyser("inter sectionneur 2x40A Schneider")
    assert a["POLES"] == "2" and a["CAL"] == "40"
    _, b = _analyser("ACTI9 ISW 2P 40A 415VAC A9S65240")
    assert b["POLES"] == "2" and b["CAL"] == "40"
    assert not _attributs_contradictoires(a, b)


def test_attributs_section_et_double():
    _, a = _analyser("Sachet embouts câblage 1.5mm double")
    assert a["SEC"] == "1.5" and a["NBR"] == "DOUBLE"
    _, b = _analyser("Embouts de cablage preisoles 2 x 1.5mm² noir")
    assert b["SEC"] == "1.5" and b["NBR"] == "DOUBLE"
    # "EMBOUT DOUBLE 2, 5MM2" (espace après la virgule, cas réel 109).
    _, c = _analyser("EMBOUT DOUBLE 2, 5MM2 (SACHET DE 100)")
    assert c["SEC"] == "2.5" and c["NBR"] == "DOUBLE"
    # Section sans unité en contexte embout (cas réel RAVATE).
    _, d = _analyser("EMBOUT  1,5  NOIR - 50")
    assert d["SEC"] == "1.5"


def test_contradictions_excluent():
    _, simple15 = _analyser("Sachet embouts câblage 1.5mm simple")
    _, double15 = _analyser("Embouts de cablage preisoles 2 x 1.5mm² noir")
    _, simple6 = _analyser("EMBOUT DE CABLAGE 6 mm² VERT (sachet de 100)")
    assert _attributs_contradictoires(simple15, double15)   # simple vs double
    assert _attributs_contradictoires(simple15, simple6)    # 1.5 vs 6 mm²


# ----------------------------------------------------------------------
# Propositions sur le cas réel complet
# ----------------------------------------------------------------------
BESOIN_DEVIS = [
    # (désignation besoin, candidats {fournisseur: désignation devis},
    #  fournisseurs qui DOIVENT être proposés)
    ("Boite dérivations étanche lisse 150x110x70",
     {"COMINTER": "BOITE DERIVATION 155 PAR 110",
      "109 DISTRIBUTION": "BOITE DERIVATION INDUS. IP55 155X110X80 960°"},
     {"COMINTER", "109 DISTRIBUTION"}),
    ("Sachet  embouts\xa0câblage 1.5mm simple",
     {"ELECTRIC PLUS": "EMBOUT MOYEN 1,5MM2 NOIR NF",
      "COMINTER": "Embout de câblage isolé noir 1,5mm² X8",
      "109 DISTRIBUTION": "EMBOUT DE CABLAGE 1.5 mm² NOIR (sachet de 100)",
      "RAVATE": "EMBOUT  1,5  NOIR - 50"},
     {"ELECTRIC PLUS", "COMINTER", "109 DISTRIBUTION", "RAVATE"}),
    ("Sachet  embouts\xa0câblage 6mm double",
     {"COMINTER": "T Embouts de cablage preisoles 2 x 6mm² vert",
      "109 DISTRIBUTION": "EMBOUT DOUBLE PREISOLE 6 MM2 VERT"},
     {"COMINTER", "109 DISTRIBUTION"}),
    ("inter sectionneur 2x40A Schneider",
     {"ELECTRIC PLUS": "ACTI9, ISW",
      "COMINTER": "ACTI9 ISW 2P 40A 415VAC A9S65240",
      "109 DISTRIBUTION": "ACTI9 ISW 2P 40A 415VCA",
      "RAVATE": "ACTI9 ISW 2P 40A 415VAC"},
     # "ACTI9, ISW" seul ne porte ni calibre ni pôles ni texte commun :
     # non proposable, à confirmer via la BDD ou une référence.
     {"COMINTER", "109 DISTRIBUTION", "RAVATE"}),
    ("tube silicone",
     {"COMINTER": "Cartouche silicone acrylique Blanc 290ML",
      "109 DISTRIBUTION": "SILICONE SIKACRYL 300mL"},
     {"COMINTER", "109 DISTRIBUTION"}),
]


def test_cas_reel_rico_carpaye_propose(referentiel):
    for i, (besoin, devis, attendus) in enumerate(BESOIN_DEVIS):
        candidats = [
            {"fournisseur": f, "reference": f"REF{i}{j}", "designation": d,
             "quantite": 1, "devis": "D1"}
            for j, (f, d) in enumerate(devis.items())
        ]
        retenus = referentiel.proposer_correspondances_designation(
            besoin, 1, candidats,
        )
        assert {c["fournisseur"] for c in retenus} == attendus, besoin


def test_cas_reel_mauvaise_section_jamais_proposee(referentiel):
    """Un besoin 6 mm² ne doit JAMAIS se voir proposer un embout 1.5 mm²,
    même si le texte autour est identique."""
    retenus = referentiel.proposer_correspondances_designation(
        "Sachet  embouts\xa0câblage 6mm simple", 1,
        [{"fournisseur": "109 DISTRIBUTION", "reference": "7170151",
          "designation": "EMBOUT DE CABLAGE 1.5 mm² NOIR (sachet de 100)",
          "quantite": 1, "devis": "320787"}],
    )
    assert retenus == []
