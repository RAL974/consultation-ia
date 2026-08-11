"""
Base d'équivalences articles (SQLite).

Rôle : reconnaître qu'une même fourniture porte des références différentes
selon les fournisseurs, et regrouper les offres correspondantes sur une
seule ligne du comparatif.

Deux niveaux de rapprochement :

1. AUTOMATIQUE — par (Fabricant + cœur numérique).
   411651 (FNR Ravate) = L411651 (Cominter) = LEG411651 (Coredime, DEM...)
   se regroupent car ils partagent le fabricant "Legrand" et le cœur 411651.
   Le fabricant est indispensable : il évite de confondre deux références
   numériquement proches de marques différentes.

2. MANUEL — par "famille".
   Pour les équivalences que seul l'acheteur peut décider (ex. une douille
   DCL Legrand, Capri, Eurohm et BLM sont interchangeables mais portent des
   références sans aucun lien). Tu les déclares dans base/familles.csv :

        Famille;Référence
        DOUILLE_DCL_E27;710100
        DOUILLE_DCL_E27;CAP715000
        DOUILLE_DCL_E27;LEG060135

   Toutes les références d'une même famille sont alors regroupées.

La base s'enrichit automatiquement : les références rencontrées dans les
devis mais absentes de la BDD y sont ajoutées (source = 'devis').

------------------------------------------------------------------
Fichiers (dossier base/) :
    BDD_articles.csv   ta base articles (séparateur ;, en-têtes en 1re ligne)
    familles.csv       tes équivalences manuelles (facultatif)
    consultation.db    base SQLite générée automatiquement
------------------------------------------------------------------
"""

import csv
import re
import sqlite3
from pathlib import Path

from moteur.normalisation import code_interne_auto


# ----------------------------------------------------------------------
# Normalisation
# ----------------------------------------------------------------------
def normaliser_ref(ref: str) -> str:
    return (
        str(ref)
        .strip()
        .upper()
        .replace(" ", "")
        .replace("-", "")
        .replace(".", "")
    )


def coeur_numerique(ref: str):
    ch = "".join(re.findall(r"\d+", str(ref).upper())).lstrip("0")
    return ch if len(ch) >= 4 else ""


def normaliser_fabricant(fab: str) -> str:
    return (
        str(fab)
        .strip()
        .upper()
        .replace("É", "E")
        .replace("È", "E")
        .replace("'", "")
    )


def _to_float(v):
    v = str(v).strip().replace(" ", "").replace("\u202f", "").replace("€", "")
    v = v.replace(",", ".")
    try:
        return float(v)
    except ValueError:
        return None


class BaseArticles:
    """Accès à la base d'équivalences."""

    def __init__(self, dossier_base):
        self.dossier = Path(dossier_base)
        self.dossier.mkdir(exist_ok=True)
        self.chemin_db = self.dossier / "consultation.db"
        self.cx = sqlite3.connect(self.chemin_db)
        self._creer_tables()

    # Colonnes attendues (sert à détecter une base d'une version antérieure)
    _COLS_ARTICLES = {
        "ref_norm", "ref_origine", "coeur", "code_interne", "fabricant",
        "designation", "categorie", "tarif", "fournisseur", "source",
    }
    _COLS_FAMILLES = {"famille", "ref_norm", "coeur"}

    def _migrer_si_obsolete(self):
        """
        Une base créée par une version antérieure peut avoir un schéma
        différent. La base n'étant qu'un cache reconstruit depuis le CSV,
        on supprime simplement les tables au schéma obsolète : elles seront
        recréées à jour, puis réalimentées par l'import.
        """
        for table, attendues in (
            ("articles", self._COLS_ARTICLES),
            ("familles", self._COLS_FAMILLES),
        ):
            cols = {
                r[1]
                for r in self.cx.execute(f"PRAGMA table_info({table})").fetchall()
            }
            if cols and cols != attendues:
                self.cx.execute(f"DROP TABLE IF EXISTS {table}")

        self.cx.commit()

    def _creer_tables(self):
        self._migrer_si_obsolete()
        self.cx.executescript(
            """
            CREATE TABLE IF NOT EXISTS articles (
                ref_norm    TEXT,
                ref_origine TEXT,
                coeur       TEXT,
                code_interne TEXT,
                fabricant   TEXT,
                designation TEXT,
                categorie   TEXT,
                tarif       REAL,
                fournisseur TEXT,
                source      TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_ref   ON articles(ref_norm);
            CREATE INDEX IF NOT EXISTS idx_coeur ON articles(coeur);
            CREATE INDEX IF NOT EXISTS idx_ci    ON articles(code_interne);

            CREATE TABLE IF NOT EXISTS familles (
                famille  TEXT,
                ref_norm TEXT,
                coeur    TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_fam_ref   ON familles(ref_norm);
            CREATE INDEX IF NOT EXISTS idx_fam_coeur ON familles(coeur);
            """
        )
        self.cx.commit()

    # ------------------------------------------------------------------
    # Import
    # ------------------------------------------------------------------
    def importer_bdd(self, csv_path):
        """(Ré)importe la BDD articles. Remplace les lignes source='bdd'."""

        csv_path = Path(csv_path)

        if not csv_path.exists():
            print(f"BDD absente : {csv_path.name} (rapprochement automatique limité)")
            return 0

        self.cx.execute("DELETE FROM articles WHERE source='bdd'")

        lignes = csv.DictReader(
            open(csv_path, encoding="utf-8-sig"), delimiter=";"
        )

        n = 0
        vus = set()

        for r in lignes:

            ref = (r.get("Référence") or "").strip()

            if not ref:
                continue

            ref_norm = normaliser_ref(ref)

            # Dédoublonnage : une seule entrée bdd par référence normalisée
            if ref_norm in vus:
                continue
            vus.add(ref_norm)

            self.cx.execute(
                "INSERT INTO articles VALUES (?,?,?,?,?,?,?,?,?,?)",
                (
                    ref_norm,
                    ref,
                    coeur_numerique(ref),
                    (r.get("Code interne") or "").strip().upper(),
                    normaliser_fabricant(r.get("Fabricant", "")),
                    (r.get("Désignation") or "").strip(),
                    (r.get("Catégorie") or "").strip(),
                    _to_float(r.get("Tarif approximatif", "")),
                    (r.get("Fournisseur") or "").strip(),
                    "bdd",
                ),
            )
            n += 1

        self.cx.commit()
        print(f"BDD articles importée : {n} références")
        return n

    def importer_familles(self, csv_path):
        """(Ré)importe les familles manuelles depuis un CSV Famille;Référence."""

        csv_path = Path(csv_path)

        self.cx.execute("DELETE FROM familles")

        if not csv_path.exists():
            self.cx.commit()
            return 0

        n = 0

        for r in csv.DictReader(open(csv_path, encoding="utf-8-sig"), delimiter=";"):

            fam = (r.get("Famille") or "").strip()
            ref = (r.get("Référence") or r.get("Reference") or "").strip()

            if not fam or not ref:
                continue

            ref_norm = normaliser_ref(ref)

            self.cx.execute(
                "INSERT INTO familles VALUES (?,?,?)",
                (fam.upper(), ref_norm, coeur_numerique(ref)),
            )
            n += 1

        self.cx.commit()

        if n:
            print(f"Familles manuelles importées : {n} références")

        return n

    def enrichir_depuis_devis(self, articles):
        """Ajoute les références des devis absentes de la base."""

        ajouts = 0

        for a in articles:

            ref_norm = normaliser_ref(a.reference_fournisseur)

            deja = self.cx.execute(
                "SELECT 1 FROM articles WHERE ref_norm=? LIMIT 1", (ref_norm,)
            ).fetchone()

            if deja:
                continue

            self.cx.execute(
                "INSERT INTO articles VALUES (?,?,?,?,?,?,?,?,?,?)",
                (
                    ref_norm,
                    a.reference_fournisseur,
                    coeur_numerique(a.reference_fournisseur),
                    "",
                    "",
                    a.designation,
                    "",
                    a.prix_net,
                    a.fournisseur,
                    "devis",
                ),
            )
            ajouts += 1

        self.cx.commit()

        if ajouts:
            print(f"Base enrichie : {ajouts} nouvelle(s) référence(s) issue(s) des devis")

        return ajouts

    # ------------------------------------------------------------------
    # Résolution d'équivalence
    # ------------------------------------------------------------------
    def _fabricant(self, ref_norm, coeur):
        """Fabricant connu pour une référence (via ref exacte puis cœur)."""

        row = self.cx.execute(
            "SELECT fabricant FROM articles WHERE ref_norm=? AND fabricant<>'' LIMIT 1",
            (ref_norm,),
        ).fetchone()

        if row:
            return row[0]

        if coeur:
            row = self.cx.execute(
                "SELECT fabricant FROM articles WHERE coeur=? AND fabricant<>'' LIMIT 1",
                (coeur,),
            ).fetchone()

            if row:
                return row[0]

        return ""

    def _code_interne(self, ref_norm, coeur):
        """Code interne connu pour une référence (via ref exacte puis cœur)."""

        row = self.cx.execute(
            "SELECT code_interne FROM articles WHERE ref_norm=? AND code_interne<>'' LIMIT 1",
            (ref_norm,),
        ).fetchone()

        if row:
            return row[0]

        if coeur:
            row = self.cx.execute(
                "SELECT code_interne FROM articles WHERE coeur=? AND code_interne<>'' LIMIT 1",
                (coeur,),
            ).fetchone()

            if row:
                return row[0]

        return ""

    def _designation(self, ref_norm, coeur):
        row = self.cx.execute(
            "SELECT designation FROM articles WHERE ref_norm=? AND designation<>'' LIMIT 1",
            (ref_norm,),
        ).fetchone()
        if row:
            return row[0]
        if coeur:
            row = self.cx.execute(
                "SELECT designation FROM articles WHERE coeur=? AND designation<>'' LIMIT 1",
                (coeur,),
            ).fetchone()
            if row:
                return row[0]
        return ""

    def est_cable(self, ref) -> bool:
        """Vrai si la référence est un câble (catégorie ou code interne)."""

        ref_norm = normaliser_ref(ref)
        coeur = coeur_numerique(ref)

        row = self.cx.execute(
            "SELECT categorie, code_interne FROM articles WHERE ref_norm=? LIMIT 1",
            (ref_norm,),
        ).fetchone()

        if not row and coeur:
            row = self.cx.execute(
                "SELECT categorie, code_interne FROM articles WHERE coeur=? LIMIT 1",
                (coeur,),
            ).fetchone()

        if not row:
            return False

        cat = (row[0] or "").strip().lower()
        ci = (row[1] or "").upper()

        prefixes = ("R2V", "AR2V", "H07", "H05", "H1XDV", "RVFV", "VVF",
                    "LIYCY", "CR1", "SYT", "CUNU")

        return cat == "câble" or ci.startswith(prefixes)

    def _famille(self, ref_norm, coeur):
        """Famille manuelle éventuelle (via ref exacte puis cœur)."""

        row = self.cx.execute(
            "SELECT famille FROM familles WHERE ref_norm=? LIMIT 1", (ref_norm,)
        ).fetchone()

        if row:
            return row[0]

        if coeur:
            row = self.cx.execute(
                "SELECT famille FROM familles WHERE coeur=? LIMIT 1", (coeur,)
            ).fetchone()

            if row:
                return row[0]

        return ""

    def groupe(self, ref: str, designation: str = "") -> str:
        """
        Clé de regroupement d'une référence.

        Priorité (de la plus fiable à la plus faible) :
          1. Famille manuelle                 -> "FAM:<nom>"
          2. Code interne renseigné dans BDD   -> "CI:<code>"
          3. Code interne AUTO (désignation)   -> "CI:<code>"
             (câbles, ampoules, embouts... : rapprochement sémantique,
              insensible aux faux amis de référence type L401227)
          4. Fabricant + cœur numérique        -> "<FABRICANT>:<coeur>"
          5. Cœur numérique seul               -> "<coeur>"
          6. Référence normalisée              -> "<ref_norm>"
        """
        ref_norm = normaliser_ref(ref)
        coeur = coeur_numerique(ref)

        fam = self._famille(ref_norm, coeur)
        if fam:
            return f"FAM:{fam}"

        ci = self._code_interne(ref_norm, coeur)
        if ci:
            return f"CI:{ci}"

        # Référence fabricant réelle (cœur numérique + fabricant connu) : elle
        # fait AUTORITÉ. Deux fournisseurs qui citent la même réf Legrand sont
        # le même produit, quelle que soit la façon dont chacun l'écrit. On ne
        # laisse donc PAS le code par désignation écraser ce rapprochement.
        fab = self._fabricant(ref_norm, coeur)
        if fab and coeur:
            return f"{fab}:{coeur}"

        # Rapprochement sémantique par désignation : réservé aux commodités
        # SANS référence fabricant partagée (ampoules, embouts... où chaque
        # fournisseur a sa propre référence).
        des = designation or self._designation(ref_norm, coeur)
        ci_auto = code_interne_auto(des) if des else ""
        if ci_auto:
            return f"CI:{ci_auto}"

        if coeur:
            return coeur

        return ref_norm

    def info(self, ref: str) -> dict:
        """Fabricant / catégorie / tarif de référence pour une référence."""

        ref_norm = normaliser_ref(ref)
        coeur = coeur_numerique(ref)

        row = self.cx.execute(
            """SELECT fabricant, categorie, tarif, designation, code_interne
               FROM articles WHERE ref_norm=? LIMIT 1""",
            (ref_norm,),
        ).fetchone()

        if not row and coeur:
            row = self.cx.execute(
                """SELECT fabricant, categorie, tarif, designation, code_interne
                   FROM articles WHERE coeur=? LIMIT 1""",
                (coeur,),
            ).fetchone()

        if not row:
            return {"fabricant": "", "categorie": "", "tarif": None,
                    "designation": "", "code_interne": ""}

        return {
            "fabricant": row[0],
            "categorie": row[1],
            "tarif": row[2],
            "designation": row[3],
            "code_interne": row[4],
        }

    def fermer(self):
        self.cx.close()
