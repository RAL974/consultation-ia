"""
Normalisation des articles.

1. CODE INTERNE — code canonique unifiant un même article, quel que soit le
   fournisseur ET le fabricant.
   Exemple câbles : R2V 3G1.5 / U1000 R02V CU 3G1,5 / R2V 3G1 5  ->  R2V3G1.5

   Les câbles sont normalisés automatiquement (désignation systématique).
   Les autres familles (disjoncteurs multi-marques, etc.) reçoivent leur
   Code interne à la main dans la BDD : c'est un choix d'acheteur.

2. UNITÉ DE VENTE — ramène chaque prix à une base comparable :
   - câbles      -> prix au mètre
   - autres      -> prix à l'unité
   Le facteur de conversion (barre = 2,10 m, boîte de 50, kit de 10, sachet
   de 100...) est déduit de la désignation quand il y figure.
"""

import re


# ----------------------------------------------------------------------
# 1. CODE INTERNE — câbles
# ----------------------------------------------------------------------
_FAMILLES_CABLE = [
    ("H07RNF", r"H[O0]?7\s?RN[\s\-]?F"),
    ("H05VVF", r"H[O0]?5\s?V\s?V[\s\-]?F"),
    ("H07VK",  r"H[O0]?7\s?V[\s\-]?K"),
    ("H07VR",  r"H[O0]?7\s?V[\s\-]?R"),
    ("H07VU",  r"H[O0]?7\s?V[\s\-]?U"),
    ("H05VK",  r"H[O0]?5\s?V[\s\-]?K"),
    ("H1XDV",  r"H1[\s\-]?XD[\s\-]?V"),
    ("AR2V",   r"A\s?R\s?[O0]?2\s?V"),
    ("R2V",    r"(?:U[\s\-]?1000\s?)?R\s?[O0]?2\s?V"),
    ("RVFV",   r"R\s?V[\s\-]?F\s?V"),
    ("VVF",    r"\bVVF\b"),
    ("LIYCY",  r"LI\s?Y\s?C\s?Y"),
    ("CR1C1",  r"CR1[\s\-]?C1"),
    ("SYT",    r"SYT\d?"),
    ("CUNU",   r"CU(?:IVRE)?\s*(?:NU|CABLETTE|\(S\)\s*RECUIT)|CABLETTE\s+CU"),
]


def _section(txt: str):
    """
    Config conducteurs : 3G1.5, 5G2.5, 4X95, 1X50, ou mono -> 1X16.
    Gère les décimales OCR écrites avec un espace ("1 5" -> 1.5) et évite
    d'attraper les chiffres du type de câble (H07, HO5...).
    """
    t = txt.upper()

    # 1) Multi-conducteurs : 3G1.5, 3G1 5, 4X95, 5G2,5
    m = re.search(r"(\d+)\s?([GX])\s?(\d+)(?:[.,]\s?(\d)|\s(\d)(?!\d))?", t)
    if m:
        n, sep, ent = m.group(1), m.group(2), m.group(3)
        dec = m.group(4) or m.group(5)
        sec = f"{ent}.{dec}" if dec else ent
        return f"{n}{sep}{sec}"

    # 2) Mono explicite : 1X50, 1X2.50, 1X 6.00
    m = re.search(r"\b1\s?X\s?(\d+(?:[.,]\d+)?)", t)
    if m:
        return "1X" + m.group(1).replace(",", ".")

    # 3) Mono : la section suit DIRECTEMENT le type de câble (VK/VR/VU/RN-F)
    #    Décimale possible avec virgule, point ou espace OCR ("1,5", "1 5").
    m = re.search(r"(?:V[\s\-]?[KRU]|RN[\s\-]?F)\s+(\d+)(?:[.,]\s?(\d+)|\s(\d)(?!\d))?", t)
    if m:
        ent, d1, d2 = m.group(1), m.group(2), m.group(3)
        dec = d1 or d2
        return f"1X{ent}.{dec}" if dec else f"1X{ent}"

    # 4) Cuivre nu / cablette : section après NU/RECUIT/CUIVRE
    m = re.search(r"(?:NU|RECUIT|CUIVRE|CABLETTE)\D*?(\d+(?:[.,]\d+)?)", t)
    if m:
        return "1X" + m.group(1).replace(",", ".")

    return None


def _clean_section(sec):
    if sec is None:
        return None

    def fix(x):
        if "." in x:
            x = x.rstrip("0").rstrip(".")
        return x

    return re.sub(r"\d+(?:\.\d+)?", lambda mm: fix(mm.group(0)), sec)


def code_interne_cable(designation: str):
    """Retourne le code interne d'un câble, ou None si non reconnu."""

    d = (designation or "").upper()

    if any(x in d for x in ("CRAYON", "COSSE", "EMBOUT", "GAINE", "PASSE",
                            "ENROULEUR", "RALLONGE", "PROLONGATEUR", "MULTIPRISE",
                            "IP44", "IP54", "TAMBOUR", "COFFRET", "ARMOIRE",
                            "REELPRO", "ENROULE")):
        return None

    for code, motif in _FAMILLES_CABLE:
        if re.search(motif, d):
            sec = _clean_section(_section(d))
            return f"{code}{sec}" if sec else None

    return None




# ----------------------------------------------------------------------
# 1b. CODE INTERNE — commodités identifiées par la désignation
#     (ampoules, embouts de câblage...). Références sans lien entre
#     fournisseurs : seule la désignation permet de les unifier.
# ----------------------------------------------------------------------
_CULOTS = ["E27", "E14", "B22", "GU10", "GU5.3", "G9", "G4", "GX53"]


def _temperature_couleur(d: str):
    d = d.upper()
    m = re.search(r"(\d{4})\s*°?\s*K\b", d)
    if m:
        return m.group(1)
    m = re.search(r"\b([2-6])\s*K\b", d)     # 4K -> 4000K
    if m:
        return m.group(1) + "000"
    return None


def _culot(d: str):
    # Neutralise la notation pack "E27x10" avant recherche
    d = re.sub(r"(E\d+)X\d+", r"\1 ", d.upper())
    for c in _CULOTS:
        if re.search(r"\b" + re.escape(c) + r"\b", d):
            return c
    return None


def code_interne_ampoule(designation: str):
    du = (designation or "").upper()

    a_mot_cle = re.search(r"AMPOULE|LAMPE|\bLED\b|BULB|TUBE", du)
    cu = _culot(designation or "")

    # Cas cryptique : culot + puissance + lumens sans mot-clé (ex. Sylvania)
    if not a_mot_cle:
        if not (cu and re.search(r"\d+[.,]?\d*\s*W", du) and re.search(r"\d+\s*LM", du)):
            return None

    if not cu:
        return None

    t = _temperature_couleur(designation or "")
    cu_code = cu.replace(".", "")
    return f"AMP_LED_{cu_code}_{t}K" if t else f"AMP_LED_{cu_code}"


# Sections de conducteur normalisées (mm²) : borne le repli "premier
# nombre de la désignation" pour ne pas prendre un conditionnement
# ("- 50", "X8") pour une section.
_SECTIONS_EMBOUT = {"0.5", "0.75", "1", "1.5", "2.5", "4", "6", "10", "16",
                    "25", "35", "50", "70", "95", "120", "150"}

# Couleurs du code NF des embouts : leur présence identifie un embout de
# câblage même sans le mot CABLAGE (cas réel RAVATE "EMBOUT  6  VERT - 50").
_COULEURS_EMBOUT = ("NOIR", "GRIS", "VERT", "ROUGE", "BLEU", "JAUNE",
                    "BLANC", "ORANGE", "IVOIRE", "BEIGE", "TURQUOISE", "VIOLET")


def code_interne_embout(designation: str):
    # Pliage d'accents indispensable : "CÂBLAGE" ne contient pas "CABL"
    # (cas réel COMINTER "Embout de câblage isolé noir 1,5mm²" -> aucun
    # code, donc jamais regroupé avec les embouts des autres fournisseurs).
    import unicodedata
    du = "".join(
        c for c in unicodedata.normalize("NFKD", (designation or "").upper())
        if not unicodedata.combining(c)
    )
    du = re.sub(r"(\d)\s*,\s*(\d)", r"\1.\2", du)   # "2, 5MM2" -> "2.5MM2"

    if "EMBOUT" not in du:
        return None

    # Embouts d'obturation (goulotte, rail, tube...) : pas des ferrules.
    if any(x in du for x in ("GOULOTTE", "RAIL", "TUBE", "MOULURE", "PLINTHE")):
        return None

    # embout de câblage / repérage / ferrule, identifié par la section
    if (any(x in du for x in ("CABL", "REP", "DOUBLE", "PREISOL", "ISOL"))
            or any(c in du for c in _COULEURS_EMBOUT)
            or re.search(r"\bC\.?\s?\d", du)):
        m = re.search(r"(\d+(?:\.\d+)?)\s*(?:MM²|MM2|MM|²)", du)
        if not m:
            # Section sans unité : premier nombre appartenant aux sections
            # normalisées ("EMBOUT  1.5  NOIR - 50" -> 1.5, pas 50... sauf
            # désignation réduite à "EMBOUT ... 50" où 50 mm² est retenu).
            for cand in re.findall(r"\b\d+(?:\.\d+)?\b", du):
                cand = cand.rstrip("0").rstrip(".") if "." in cand else cand
                if cand in _SECTIONS_EMBOUT:
                    class _M:  # même interface que re.Match pour la suite
                        def group(self, _):
                            return cand
                    m = _M()
                    break
        if m:
            sec = m.group(1).replace(",", ".")
            if "." in sec:
                sec = sec.rstrip("0").rstrip(".")
            if sec not in _SECTIONS_EMBOUT:
                return None
            # Simple et double (préisolé 2 x <section>) sont des articles
            # DIFFÉRENTS : sans ce suffixe, "EMBOUT DOUBLE 2.5MM2" et
            # "EMBOUT DE CABLAGE 2.5" tombaient dans le même groupe et le
            # comparatif affichait des prix de simples face à un besoin de
            # doubles (cas réel, consultation Rico Carpaye).
            double = "DOUBLE" in du or re.search(r"\b2\s*X\s*" + re.escape(sec), du)
            return f"EMBOUT_{sec}_DOUBLE" if double else f"EMBOUT_{sec}"

    return None




# ----------------------------------------------------------------------
# 1c. CODE INTERNE — disjoncteurs, gaines ICTA, goulottes
# ----------------------------------------------------------------------
def _pouvoir_coupure(txt: str):
    """Pouvoir de coupure en kA. kA explicite prioritaire ; sinon A>=1000."""
    t = txt.upper()
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*KA\b", t)
    if m:
        v = float(m.group(1).replace(",", "."))
        return ("%g" % v) + "kA"
    m = re.search(r"\b(\d{4,})\s*A\b", t)      # >=1000 A (pas le calibre)
    if m:
        return ("%g" % (int(m.group(1)) / 1000)) + "kA"
    return ""


def _poles(txt: str):
    t = txt.upper().replace("1P+NG", "1P+N").replace("P+NG", "P+N")
    for p in ["4P+N", "3P+N", "1P+N", "4PG", "4P", "3P", "2P", "1P"]:
        if re.search(r"\b" + re.escape(p) + r"\b", t):
            return p.replace("4PG", "4P")
    return ""


def _calibre(txt: str):
    t = txt.upper()
    m = re.search(r"\b[BCD]\s?(\d+)\s*A\b", t)   # après la courbe
    if m:
        return m.group(1) + "A"
    m = re.search(r"\b(\d{1,3})\s*A\b", t)         # calibre <1000 A
    if m:
        return m.group(1) + "A"
    return ""


def _courbe(txt: str):
    t = txt.upper()
    m = re.search(r"\b([BCD])\s?\d+\s*A\b", t)
    if m:
        return m.group(1)
    m = re.search(r"COURBE\s*([BCD])\b", t)
    if m:
        return m.group(1)
    return ""


def _sensibilite(txt: str):
    m = re.search(r"(\d+)\s*MA\b", txt.upper())
    return m.group(1) + "MA" if m else ""


def _type_diff(txt: str):
    t = txt.upper()
    m = re.search(r"TYPE\s*(AC|A|B|F)\b", t)
    if m:
        return m.group(1)
    m = re.search(r"\b(AC|A|B|F)\s*\d*\s*MA", t)
    if m:
        return m.group(1)
    # Type placé APRÈS la sensibilité : « ... 30MA AC », « 30 MA TYPE A »
    m = re.search(r"\d+\s*MA\s*(?:TYPE\s*)?(AC|A|B|F)\b", t)
    return m.group(1) if m else ""


def code_interne_disjoncteur(designation: str):
    """
    Disjoncteur modulaire       -> DISJ_<poles>_<courbe><calibre>_<pdc>
    Inter/disj différentiel     -> INTERDIFF_<poles>_<calibre>_<type>_<sensi>
    (le pouvoir de coupure et le type A/AC sont volontairement dans la clé :
     ce sont des caractéristiques non interchangeables.)
    """
    du = (designation or "").upper()

    # Différentiel : présence d'une sensibilité en mA
    est_diff = bool(re.search(r"\d+\s*MA\b", du)) and (
        "ID" in du or "DIFF" in du or "INTER" in du or "30MA" in du or "300MA" in du
    )

    p = _poles(du)
    cal = _calibre(du)

    if est_diff:
        s = _sensibilite(du)
        ty = _type_diff(du)
        if p and cal and s:
            return f"INTERDIFF_{p}_{cal}_{ty or '?'}_{s}"
        return None

    c = _courbe(du)
    if p and cal and c:
        pdc = _pouvoir_coupure(du)
        return f"DISJ_{p}_{c}{cal}" + (f"_{pdc}" if pdc else "")
    return None


_DIAM_ICTA = ("16", "20", "25", "32", "40", "50", "63")


def code_interne_icta(designation: str):
    du = (designation or "").upper()
    if "ICTA" not in du:
        return None
    m = re.search(r"(?:DIAM|Ø|ATF|D)\s*(\d{2})\b", du)
    if not m:
        m = re.search(r"\b(16|20|25|32|40|50|63)\b", du)
    return f"ICTA_{m.group(1)}" if m else None


_ACC_GOULOTTE = (
    "EMBOUT", "EMBT", "CADRE", "PRISE", "COUDE", "BOUCHON", "CHEVILLE", "BOITE",
    "BOÎTE", "ANGLE", "SORTIE", "DERIVATION", "DÉRIVATION", "PORTE",
    "ACCESSOIRE", "SUPPORT", "SEPARATEUR", "SÉPARATEUR", "JONCTION",
    "OBTURATEUR", "JOINT", "CLOISON", "FOND", "COUVERCLE SEUL",
)


def code_interne_goulotte(designation: str):
    du = (designation or "").upper()
    if any(a in du for a in _ACC_GOULOTTE):
        return None
    fam = None
    if re.search(r"\bGOULOTTE\b", du):
        fam = "GOULOTTE"
    elif re.search(r"\bMOULURE\b", du):
        fam = "MOULURE"
    elif re.search(r"\bPLINTHE\b", du):
        fam = "PLINTHE"
    elif re.search(r"\bGTL\b", du):
        fam = "GTL"
    if not fam:
        return None
    m = re.search(r"(\d+)\s*[X×*]\s*(\d+(?:[.,]\d+)?)", du)
    if m:
        b = m.group(2).replace(",", ".")
        if "." in b:
            b = b.rstrip("0").rstrip(".")
        return f"{fam}_{m.group(1)}X{b}"
    return None



def code_interne_auto(designation: str, categorie: str = ""):
    """
    Code interne automatique, par ordre de priorité :
    câbles, ampoules, embouts, ICTA, goulottes, disjoncteurs.
    Retourne "" si non reconnu.
    """
    for fn in (
        code_interne_cable,
        code_interne_ampoule,
        code_interne_embout,
        code_interne_icta,
        code_interne_goulotte,
        code_interne_disjoncteur,
    ):
        code = fn(designation)
        if code:
            return code
    return ""


# ----------------------------------------------------------------------
# 2. UNITÉ DE VENTE — facteur de conversion vers l'unité de base
# ----------------------------------------------------------------------

# Unités reconnues -> famille d'unité de base
_UNITE_BASE = {
    "UN": "U", "U": "U", "PCE": "U", "PIECE": "U",
    "MT": "M", "M": "M", "ML": "M", "MTR": "M",
    "BARRE": "BARRE", "BAR": "BARRE",
    "BTE": "BOITE", "BOITE": "BOITE", "BOÎTE": "BOITE", "B": "BOITE",
    "SCH": "SACHET", "SACHET": "SACHET",
    "KIT": "KIT", "ROUL": "M", "TOURET": "M",
}


def famille_unite(unite: str) -> str:
    return _UNITE_BASE.get((unite or "").strip().upper(), (unite or "").strip().upper())


def conditionnement_designation(designation: str):
    """
    Cherche un conditionnement chiffré dans la désignation.
    Retourne (nombre_unites_par_colis, explicite) ou (None, False).
      - explicite=True  : mot-clé clair ("sachet de 100", "boîte de 50")
      - explicite=False : indice plus faible ("X12", "- 50") -> à vérifier
    """
    d = (designation or "").upper()

    m = re.search(r"(?:BO[IÎ]TE|SACHET|KIT|LOT|PACK|BLISTER)\s*(?:DE\s*)?(\d+)", d)
    if m:
        return int(m.group(1)), True

    m = re.search(r"\bX\s?(\d+)\b", d)
    if m:
        return int(m.group(1)), False

    m = re.search(r"[-–]\s*(\d+)\s*$", d.strip())
    if m:
        return int(m.group(1)), False

    return None, False


def longueur_barre(designation: str):
    """Longueur d'une barre/moulure en mètres si présente (ex. 2.10m)."""
    m = re.search(r"(\d(?:[.,]\d+)?)\s*M\b", (designation or "").upper())
    if m:
        return float(m.group(1).replace(",", "."))
    return None


def prix_normalise(prix_net, unite, designation, est_cable=False):
    """
    Ramène un prix à l'unité de base comparable.

    Retourne (prix_norm, unite_base, fiable)
      - fiable=False quand le facteur de conversion est incertain
        (ex. boîte sans quantité connue) : le prix est laissé tel quel
        et signalé.
    """
    fam = famille_unite(unite)

    # Câbles : base = mètre. MT direct ; barre/touret = longueur si connue.
    if est_cable:
        if fam == "M":
            return prix_net, "€/m", True
        if fam == "BARRE":
            lg = longueur_barre(designation)
            if lg:
                return round(prix_net / lg, 4), "€/m", True
            return prix_net, "€/barre", False
        return prix_net, f"€/{fam.lower()}", False

    # Composants : base = unité.
    n, explicite = conditionnement_designation(designation)

    if fam == "U":
        # Pack vendu comme "unité" : on ne divise QUE sur un conditionnement
        # EXPLICITE ("boîte/sachet/kit de N"). Un simple "x12" est trop
        # souvent une dimension (32x12, 20x32...) pour diviser sans risque.
        if n and n > 1 and explicite:
            return round(prix_net / n, 4), "€/u", True
        return prix_net, "€/u", True

    if fam == "BARRE":
        lg = longueur_barre(designation)
        if lg:
            return round(prix_net / lg, 4), "€/m", True
        return prix_net, "€/barre", False

    if fam in ("BOITE", "SACHET", "KIT"):
        if n:
            return round(prix_net / n, 4), "€/u", explicite
        return prix_net, f"€/{fam.lower()}", False

    return prix_net, f"€/{fam.lower()}" if fam else "€/u", True
