
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path
from datetime import date
from io import BytesIO
import sqlite3
import calendar
import re

st.set_page_config(
    page_title="Gestione Finanze - Arletti",
    page_icon="💰",
    layout="centered",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent
DB_FILE = BASE_DIR / "finanze_famiglia.db"
DEFAULT_EXCEL_FILE = BASE_DIR / "Spese casa -2.xlsx"

MONTHS = [
    "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
    "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"
]

# ============================================================
# LE UNICHE 9 CATEGORIE DI SPESA
# ============================================================
EXPENSE_CATEGORIES = [
    "Costo alimentare mensile",
    "Tempo libero e viaggi (ristoranti aperitivi)",
    "Utenze",
    "Scuola e sport",
    "Trasporti e auto",
    "Prelievi contanti (Alla etc...)",
    "Casa e assicurazioni",
    "Shopping",
    "Farmacia e cura della persona",
]

INCOME_CATEGORIES = [
    "Stipendio",
    "Bonus",
    "Rimborso",
    "Affitto",
    "Altre entrate",
]

PEOPLE = ["Famiglia", "Mattia", "Virginia"]

st.markdown("""
<style>
.block-container {
    max-width: 1100px;
    padding-top: 1rem;
    padding-bottom: 3rem;
}
[data-testid="stMetricValue"] {
    font-size: 1.35rem;
}
</style>
""", unsafe_allow_html=True)


# ============================================================
# UTILITY
# ============================================================
def euro(x):
    try:
        x = float(x)
    except Exception:
        x = 0.0
    return f"{x:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")


def to_number(x):
    if x is None:
        return None

    try:
        if pd.isna(x):
            return None
    except Exception:
        pass

    if isinstance(x, (int, float)):
        return float(x)

    s = str(x).strip().replace("€", "").replace(" ", "")

    if not s:
        return None

    if "," in s:
        s = s.replace(".", "").replace(",", ".")

    try:
        return float(s)
    except Exception:
        return None


def month_key(year, month):
    return f"{year:04d}-{month:02d}"


def month_label(year, month):
    return f"{MONTHS[month - 1]} {year}"


def first_day(year, month):
    return date(year, month, 1)


def qdf(sql, params=()):
    c = sqlite3.connect(DB_FILE)

    try:
        return pd.read_sql_query(sql, c, params=params)

    finally:
        c.close()


def execute(sql, params=()):
    c = sqlite3.connect(DB_FILE)

    try:
        cur = c.cursor()
        cur.execute(sql, params)
        c.commit()
        return cur.lastrowid

    finally:
        c.close()


def get_setting(key, default=0.0):
    df = qdf(
        "SELECT valore FROM impostazioni WHERE chiave=?",
        (key,)
    )

    if df.empty:
        return default

    try:
        return float(df.iloc[0]["valore"])

    except Exception:
        return default


def set_setting(key, value):

    exists = qdf(
        "SELECT chiave FROM impostazioni WHERE chiave=?",
        (key,)
    )

    if exists.empty:

        execute(
            """
            INSERT INTO impostazioni
            (chiave,valore)
            VALUES (?,?)
            """,
            (key, str(value))
        )

    else:

        execute(
            """
            UPDATE impostazioni
            SET valore=?
            WHERE chiave=?
            """,
            (str(value), key)
        )


def section_title(title, subtitle=None):

    st.title(title)

    if subtitle:
        st.caption(subtitle)


# ============================================================
# DATABASE
# ============================================================
def init_db():

    c = sqlite3.connect(DB_FILE)
    cur = c.cursor()

    cur.executescript("""
    CREATE TABLE IF NOT EXISTS movimenti (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tipo TEXT NOT NULL,
        data TEXT NOT NULL,
        categoria TEXT NOT NULL,
        descrizione TEXT,
        persona TEXT,
        importo REAL NOT NULL,
        pagato INTEGER DEFAULT 1,
        fonte TEXT DEFAULT 'manuale',
        note TEXT
    );

    CREATE TABLE IF NOT EXISTS spese_mensili (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        mese TEXT NOT NULL,
        categoria TEXT NOT NULL,
        importo REAL NOT NULL DEFAULT 0,
        fonte TEXT DEFAULT 'manuale',
        note TEXT
    );

    CREATE TABLE IF NOT EXISTS modalita_mese (
        mese TEXT PRIMARY KEY,
        modalita TEXT NOT NULL DEFAULT 'singola'
    );

    CREATE TABLE IF NOT EXISTS ricorrenti (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        descrizione TEXT NOT NULL,
        categoria TEXT NOT NULL,
        importo REAL NOT NULL,
        giorno INTEGER DEFAULT 1,
        persona TEXT,
        attiva INTEGER DEFAULT 1
    );

    CREATE TABLE IF NOT EXISTS spese_future (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        data TEXT NOT NULL,
        categoria TEXT NOT NULL,
        descrizione TEXT,
        persona TEXT,
        importo REAL NOT NULL,
        pagata INTEGER DEFAULT 0,
        note TEXT
    );

    CREATE TABLE IF NOT EXISTS budget_mensile (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        mese TEXT NOT NULL,
        categoria TEXT NOT NULL,
        importo REAL NOT NULL DEFAULT 0
    );

    CREATE TABLE IF NOT EXISTS obiettivi (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL,
        obiettivo REAL NOT NULL,
        accumulato REAL DEFAULT 0,
        scadenza TEXT,
        attivo INTEGER DEFAULT 1
    );

    CREATE TABLE IF NOT EXISTS impostazioni (
        chiave TEXT PRIMARY KEY,
        valore TEXT
    );

    CREATE TABLE IF NOT EXISTS persone (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT UNIQUE NOT NULL,
        attiva INTEGER DEFAULT 1
    );
    """)

    cur.execute(
        """
        INSERT OR IGNORE INTO impostazioni
        (chiave,valore)
        VALUES ('risparmio_mensile_target','0')
        """
    )

    cur.execute(
        """
        INSERT OR IGNORE INTO impostazioni
        (chiave,valore)
        VALUES ('fondo_sicurezza','500')
        """
    )

    for p in PEOPLE:

        cur.execute(
            """
            INSERT OR IGNORE INTO persone(nome)
            VALUES (?)
            """,
            (p,)
        )

    c.commit()
    c.close()


init_db()


# ============================================================
# EXCEL
# ============================================================
def normalize_excel_category(value):

    if value is None:
        return None

    s = re.sub(
        r"\s+",
        " ",
        str(value).strip()
    )

    mapping = {
        "Costo alimentare mensile":
            "Costo alimentare mensile",

        "Tempo libero e viaggi":
            "Tempo libero e viaggi (ristoranti aperitivi)",

        "Tempo libero e viaggi (ristoranti aperitivi)":
            "Tempo libero e viaggi (ristoranti aperitivi)",

        "Utenze":
            "Utenze",

        "Scuola e sport":
            "Scuola e sport",

        "Trasporti e auto":
            "Trasporti e auto",

        "Prelievi contanti":
            "Prelievi contanti (Alla etc...)",

        "Prelievi contanti(Alla etc...)":
            "Prelievi contanti (Alla etc...)",

        "Prelievi contanti (Alla etc...)":
            "Prelievi contanti (Alla etc...)",

        "Casa e assicurazioni":
            "Casa e assicurazioni",

        "Shopping":
            "Shopping",

        "Farmacia e cura della persona":
            "Farmacia e cura della persona",
    }

    return mapping.get(s)


def normalize_category(value):
    """Restituisce sempre una delle 9 categorie ufficiali."""
    if value is None:
        return None

    s = re.sub(r"\s+", " ", str(value).strip())
    n = s.lower()

    aliases = {
        "costo alimentare mensile": "Costo alimentare mensile",
        "tempo libero e viaggi": "Tempo libero e viaggi (ristoranti aperitivi)",
        "tempo libero e viaggi (ristoranti aperitivi)": "Tempo libero e viaggi (ristoranti aperitivi)",
        "utenze": "Utenze",
        "scuola e sport": "Scuola e sport",
        "trasporti e auto": "Trasporti e auto",
        "prelievi contanti": "Prelievi contanti (Alla etc...)",
        "prelievi contanti(alla etc...)": "Prelievi contanti (Alla etc...)",
        "prelievi contanti (alla etc...)": "Prelievi contanti (Alla etc...)",
        "casa e assicurazioni": "Casa e assicurazioni",
        "shopping": "Shopping",
        "farmacia e cura della persona": "Farmacia e cura della persona",
    }
    return aliases.get(n)


def repair_categories():
    """Ripara eventuali vecchie varianti/duplicati presenti nel DB."""
    c = sqlite3.connect(DB_FILE)
    cur = c.cursor()

    for table in ("spese_mensili", "movimenti", "budget_mensile", "ricorrenti", "spese_future"):
        try:
            rows = cur.execute(f"SELECT id, categoria FROM {table}").fetchall()
        except sqlite3.OperationalError:
            continue
        for row_id, old_cat in rows:
            new_cat = normalize_category(old_cat)
            if new_cat and new_cat != old_cat:
                cur.execute(f"UPDATE {table} SET categoria=? WHERE id=?", (new_cat, row_id))

    # Elimina duplicati mensili lasciando la riga con l'id più basso e
    # somma prima gli importi per non perdere dati.
    try:
        dup_groups = cur.execute("""
            SELECT mese, categoria, SUM(importo)
            FROM spese_mensili
            GROUP BY mese, categoria
            HAVING COUNT(*) > 1
        """).fetchall()
        for mese, categoria, totale in dup_groups:
            keep = cur.execute("""
                SELECT id FROM spese_mensili
                WHERE mese=? AND categoria=? ORDER BY id LIMIT 1
            """, (mese, categoria)).fetchone()[0]
            cur.execute("DELETE FROM spese_mensili WHERE mese=? AND categoria=? AND id<>?", (mese, categoria, keep))
            cur.execute("UPDATE spese_mensili SET importo=? WHERE id=?", (float(totale), keep))
    except sqlite3.OperationalError:
        pass

    try:
        cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS ux_spese_mensili_mese_categoria ON spese_mensili(mese,categoria)")
    except sqlite3.OperationalError:
        pass

    try:
        cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS ux_budget_mensile_mese_categoria ON budget_mensile(mese,categoria)")
    except sqlite3.OperationalError:
        pass

    c.commit()
    c.close()


repair_categories()


REGISTRATION_MODES = {
    'mensile': '📅 Mensile per categoria',
    'singola': '🧾 Singole spese',
    'mista': '🔀 Mista',
}

def set_month_mode(mese, mode):
    if mode not in REGISTRATION_MODES:
        return
    execute(
        "INSERT INTO modalita_mese(mese, modalita) VALUES (?, ?) ON CONFLICT(mese) DO UPDATE SET modalita=excluded.modalita",
        (mese, mode),
    )

def get_month_mode(year, month):
    key = month_key(year, month)
    saved = qdf("SELECT modalita FROM modalita_mese WHERE mese=?", (key,))
    if not saved.empty and str(saved.iloc[0]['modalita']) in REGISTRATION_MODES:
        return str(saved.iloc[0]['modalita'])
    has_monthly = not qdf("SELECT id FROM spese_mensili WHERE mese=? LIMIT 1", (key,)).empty
    has_single = not qdf("SELECT id FROM movimenti WHERE tipo='uscita' AND substr(data,1,7)=? LIMIT 1", (key,)).empty
    if has_monthly and has_single:
        return 'mista'
    if has_monthly:
        return 'mensile'
    return 'singola'

def mode_description(mode):
    if mode == 'mensile':
        return 'Il totale usa solo i valori mensili per categoria.'
    if mode == 'singola':
        return 'Il totale usa solo le singole spese pagate.'
    return 'Il totale somma entrambe le fonti. Usala solo se rappresentano spese diverse.'


def find_excel_file():

    if "excel_bytes" in st.session_state:
        return st.session_state["excel_bytes"]

    if DEFAULT_EXCEL_FILE.exists():
        return DEFAULT_EXCEL_FILE

    return None


def parse_excel_history(source):
    """
    Legge il foglio 'Costi famiglia'.

    Cerca intestazioni del tipo:
    Spese SETTEMBRE 2026

    e prende l'importo dalla colonna B.

    Vengono importate esclusivamente le 9 categorie
    definite in EXPENSE_CATEGORIES.
    """

    if source is None:

        return pd.DataFrame(
            columns=[
                "mese",
                "categoria",
                "importo",
                "fonte"
            ]
        )

    try:

        import openpyxl

        if isinstance(source, (str, Path)):

            wb = openpyxl.load_workbook(
                source,
                data_only=True
            )

        else:

            wb = openpyxl.load_workbook(
                BytesIO(source),
                data_only=True
            )

        if "Costi famiglia" not in wb.sheetnames:
            return pd.DataFrame()

        ws = wb["Costi famiglia"]

        month_map = {
            m.upper(): i
            for i, m in enumerate(
                MONTHS,
                start=1
            )
        }

        current_year = None
        current_month = None

        rows = []

        for r in range(
            1,
            ws.max_row + 1
        ):

            a = ws.cell(
                r,
                1
            ).value

            if a is None:
                continue

            text = re.sub(
                r"\s+",
                " ",
                str(a).strip()
            )

            upper = text.upper()

            match = re.search(
                r"SPESE\s+([A-ZÀ-Ù]+)\s+(\d{4})",
                upper
            )

            if match:

                month_name = match.group(1)

                if month_name in month_map:

                    current_month = month_map[
                        month_name
                    ]

                    current_year = int(
                        match.group(2)
                    )

                continue

            if (
                current_year is None
                or current_month is None
            ):
                continue

            category = normalize_excel_category(
                text
            )

            if category is None:
                continue

            amount = to_number(
                ws.cell(
                    r,
                    2
                ).value
            )

            if amount is None:
                continue

            rows.append(
                {
                    "mese": month_key(
                        current_year,
                        current_month
                    ),

                    "categoria": category,

                    "importo": amount,

                    "fonte": "Excel"
                }
            )

        if not rows:

            return pd.DataFrame(
                columns=[
                    "mese",
                    "categoria",
                    "importo",
                    "fonte"
                ]
            )

        df = pd.DataFrame(
            rows
        )

        df = (
            df
            .groupby(
                [
                    "mese",
                    "categoria"
                ],
                as_index=False
            )["importo"]
            .sum()
        )

        df["fonte"] = "Excel"

        return df

    except Exception as e:

        st.error(
            f"Errore nella lettura dell'Excel: {e}"
        )

        return pd.DataFrame()


def import_excel_to_db(
    source,
    overwrite_excel_rows=True
):

    df = parse_excel_history(
        source
    )

    if df.empty:
        return 0

    count = 0

    for _, row in df.iterrows():

        mese = row["mese"]
        categoria = row["categoria"]
        importo = float(
            row["importo"]
        )

        existing = qdf(
            """
            SELECT id, fonte

            FROM spese_mensili

            WHERE mese=?
            AND categoria=?

            ORDER BY id DESC

            LIMIT 1
            """,
            (
                mese,
                categoria
            )
        )

        if existing.empty:

            execute(
                """
                INSERT INTO spese_mensili
                (
                    mese,
                    categoria,
                    importo,
                    fonte
                )

                VALUES
                (?, ?, ?, 'Excel')
                """,
                (
                    mese,
                    categoria,
                    importo
                )
            )

            count += 1

        elif (
            overwrite_excel_rows
            and str(existing.iloc[0]["fonte"]) == "Excel"
        ):

            execute(
                """
                UPDATE spese_mensili

                SET importo=?,
                    fonte='Excel'

                WHERE id=?
                """,
                (
                    importo,
                    int(existing.iloc[0]["id"])
                )
            )

            count += 1

    for mese in df['mese'].drop_duplicates().tolist():
        has_single = not qdf("SELECT id FROM movimenti WHERE tipo='uscita' AND substr(data,1,7)=? LIMIT 1", (mese,)).empty
        set_month_mode(mese, 'mista' if has_single else 'mensile')

    return count


# Import automatico dell'Excel se presente.
excel_source = find_excel_file()

if excel_source is not None:
    import_excel_to_db(
        excel_source
    )


# ============================================================
# DATI FINANZIARI
# ============================================================
def monthly_category_totals(year, month):
    key = month_key(year, month)
    mode = get_month_mode(year, month)
    result = {category: 0.0 for category in EXPENSE_CATEGORIES}

    if mode in ('mensile', 'mista'):
        monthly = qdf("SELECT categoria, SUM(importo) AS importo FROM spese_mensili WHERE mese=? GROUP BY categoria", (key,))
        for _, row in monthly.iterrows():
            category = normalize_category(row['categoria'])
            if category in result:
                result[category] += float(row['importo'] or 0)

    if mode in ('singola', 'mista'):
        individual = qdf("SELECT categoria, SUM(importo) AS importo FROM movimenti WHERE tipo='uscita' AND pagato=1 AND substr(data,1,7)=? GROUP BY categoria", (key,))
        for _, row in individual.iterrows():
            category = normalize_category(row['categoria'])
            if category in result:
                result[category] += float(row['importo'] or 0)

    return (pd.DataFrame([{'categoria': c, 'importo': v} for c, v in result.items()])
            .sort_values('importo', ascending=False).reset_index(drop=True))


def monthly_total(
    year,
    month
):

    df = monthly_category_totals(
        year,
        month
    )

    if df.empty:
        return 0.0

    return float(
        df["importo"].sum()
    )


def monthly_income(
    year,
    month
):

    key = month_key(
        year,
        month
    )

    df = qdf(
        """
        SELECT
            COALESCE(
                SUM(importo),
                0
            ) AS totale

        FROM movimenti

        WHERE tipo='entrata'
        AND pagato=1
        AND substr(data,1,7)=?
        """,
        (key,)
    )

    return float(
        df.iloc[0]["totale"]
    )


def recurring_remaining(
    year,
    month
):

    today = date.today()

    df = qdf(
        """
        SELECT *
        FROM ricorrenti
        WHERE attiva=1
        """
    )

    if df.empty:
        return 0.0

    total = 0.0

    for _, row in df.iterrows():

        try:
            day = min(
                max(
                    int(row["giorno"]),
                    1
                ),
                28
            )

        except Exception:
            day = 1

        if day < today.day:
            continue

        category = row["categoria"]
        description = row["descrizione"]
        amount = float(row["importo"])

        already = qdf(
            """
            SELECT COUNT(*) AS n

            FROM movimenti

            WHERE tipo='uscita'

            AND substr(data,1,7)=?

            AND categoria=?

            AND ABS(importo-?) < 0.01

            AND descrizione=?
            """,
            (
                month_key(
                    year,
                    month
                ),
                category,
                amount,
                description
            )
        )

        if int(
            already.iloc[0]["n"]
        ) == 0:

            total += amount

    return total


def future_remaining(
    year,
    month
):

    today = date.today()

    start = max(
        today,
        first_day(
            year,
            month
        )
    )

    last_day = calendar.monthrange(
        year,
        month
    )[1]

    end = date(
        year,
        month,
        last_day
    )

    if start > end:
        return 0.0

    df = qdf(
        """
        SELECT
            COALESCE(
                SUM(importo),
                0
            ) AS totale

        FROM spese_future

        WHERE pagata=0

        AND data>=?

        AND data<=?
        """,
        (
            start.isoformat(),
            end.isoformat()
        )
    )

    return float(
        df.iloc[0]["totale"]
    )


# ============================================================
# ANALISI RISPARMIO
# ============================================================
def historical_category_analysis(
    months_back=6
):
    today = date.today()
    values_by_category = {category: [] for category in EXPENSE_CATEGORIES}

    for offset in range(1, months_back + 1):
        month_number = today.month - offset
        year = today.year
        while month_number <= 0:
            month_number += 12
            year -= 1

        df = monthly_category_totals(year, month_number)
        values = dict(zip(df["categoria"], df["importo"])) if not df.empty else {}

        for category in EXPENSE_CATEGORIES:
            values_by_category[category].append(float(values.get(category, 0.0)))

    rows = []
    for category in EXPENSE_CATEGORIES:
        rows.append({
            "categoria": category,
            "media": sum(values_by_category[category]) / len(values_by_category[category])
        })

    return pd.DataFrame(rows)


def saving_plan(
    months_back=6,
    reduction_pct=0.10
):

    today = date.today()

    historical = historical_category_analysis(
        months_back
    )

    current = monthly_category_totals(
        today.year,
        today.month
    )

    current_dict = {}

    if not current.empty:

        current_dict = dict(
            zip(
                current["categoria"],
                current["importo"]
            )
        )

    if historical.empty:
        return pd.DataFrame()

    rows = []

    for _, row in historical.iterrows():

        category = row["categoria"]

        average = float(
            row["media"]
        )

        current_value = float(
            current_dict.get(
                category,
                0
            )
        )

        if average <= 0:

            target = 0.0
            saving = 0.0

        else:

            target = (
                average
                * (
                    1 - reduction_pct
                )
            )

            if (
                current_value
                > average * 1.20
            ):

                target = (
                    average
                    * 0.95
                )

            target = max(
                target,
                0.0
            )

            saving = max(
                average - target,
                0.0
            )

        if (
            current_value
            > average * 1.15
        ):

            status = (
                "🔴 Molto sopra la media"
            )

        elif (
            current_value
            > average * 1.05
        ):

            status = (
                "🟠 Sopra la media"
            )

        elif (
            current_value > 0
            and current_value <= average
        ):

            status = (
                "🟢 Sotto/in linea"
            )

        else:

            status = (
                "⚪ Da osservare"
            )

        rows.append(
            {
                "Categoria":
                    category,

                "Media storica":
                    average,

                "Mese corrente":
                    current_value,

                "Obiettivo":
                    target,

                "Risparmio possibile":
                    saving,

                "Stato":
                    status
            }
        )

    result = pd.DataFrame(rows)

    # Difesa finale contro qualsiasi duplicato storico: una riga per categoria.
    result = (
        result.groupby("Categoria", as_index=False)
        .agg({
            "Media storica": "max",
            "Mese corrente": "max",
            "Obiettivo": "max",
            "Risparmio possibile": "max",
            "Stato": "first",
        })
    )

    return result.sort_values(
        "Risparmio possibile",
        ascending=False
    )


# ============================================================
# SIDEBAR
# ============================================================
st.sidebar.title(
    "💰 Gestione Finanze"
)

menu = st.sidebar.radio(
    "SEZIONI",
    [
        "🏠 Home",
        "💰 Entrate",
        "💳 Spese",
        "🔁 Spese ricorrenti",
        "📌 Spese future",
        "📅 Budget",
        "💡 Piano risparmio",
        "🎯 Obiettivi",
        "💸 Quanto posso spendere oggi?",
        "📊 Analisi",
        "📥 Import / Export",
        "⚙️ Impostazioni",
    ]
)

st.sidebar.markdown("---")

if excel_source is not None:

    st.sidebar.success(
        "🟢 Excel collegato"
    )

    st.sidebar.caption(
        "Costi famiglia"
    )

else:

    st.sidebar.warning(
        "🟡 Excel non collegato"
    )

    st.sidebar.caption(
        "Puoi caricarlo in Import / Export"
    )


# ============================================================
# HOME
# ============================================================
if menu == "🏠 Home":

    section_title(
        "🏠 Dashboard Finanze",
        "Controllo delle spese familiari, budget e risparmio"
    )

    selected = st.date_input(
        "📅 Mese",
        date.today().replace(day=1),
        key="home_month"
    )

    year = selected.year
    month = selected.month

    income = monthly_income(
        year,
        month
    )

    expenses = monthly_total(
        year,
        month
    )

    savings = income - expenses

    percentage = (
        savings / income * 100
        if income
        else 0
    )

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "💰 Entrate",
        euro(income)
    )

    c2.metric(
        "💳 Spese",
        euro(expenses)
    )

    c3.metric(
        "💚 Risparmio",
        euro(savings)
    )

    c4.metric(
        "📈 Risparmio %",
        f"{percentage:.1f}%"
    )

    st.markdown("---")

    st.subheader(
        f"📊 Spese {month_label(year, month)}"
    )

    categories = monthly_category_totals(
        year,
        month
    )

    if categories.empty:

        st.info(
            "Nessuna spesa registrata per questo mese."
        )

    else:

        st.dataframe(
            categories.rename(
                columns={
                    "categoria":
                        "Categoria",

                    "importo":
                        "Importo"
                }
            ).style.format(
                {
                    "Importo":
                        euro
                }
            ),
            use_container_width=True,
            hide_index=True
        )

        fig, ax = plt.subplots(
            figsize=(8, 5)
        )

        ax.pie(
            categories["importo"],
            labels=categories["categoria"],
            autopct="%1.1f%%",
            startangle=90
        )

        ax.axis("equal")

        st.pyplot(
            fig,
            clear_figure=True
        )

    st.markdown("---")

    st.subheader(
        "💡 Dove puoi risparmiare?"
    )

    plan = saving_plan(
        6,
        0.10
    )

    if plan.empty:

        st.info(
            "Servono alcuni mesi di storico "
            "per creare suggerimenti."
        )

    else:

        potential = float(
            plan["Risparmio possibile"].sum()
        )

        c1, c2 = st.columns(2)

        c1.metric(
            "Risparmio potenziale mensile",
            euro(potential)
        )

        c2.metric(
            "Risparmio potenziale annuale",
            euro(potential * 12)
        )

        for _, row in plan.head(3).iterrows():

            if row["Risparmio possibile"] > 0:

                st.write(
                    f"**{row['Categoria']}** — "
                    f"potenziale: "
                    f"**{euro(row['Risparmio possibile'])}/mese**"
                )


# ============================================================
# ENTRATE
# ============================================================
elif menu == "💰 Entrate":

    section_title(
        "💰 Entrate",
        "Registra le entrate familiari"
    )

    with st.form(
        "form_entrata"
    ):

        transaction_date = st.date_input(
            "Data",
            date.today()
        )

        category = st.selectbox(
            "Categoria",
            INCOME_CATEGORIES
        )

        person = st.selectbox(
            "Persona",
            PEOPLE
        )

        description = st.text_input(
            "Descrizione"
        )

        amount = st.number_input(
            "Importo (€)",
            min_value=0.0,
            step=50.0
        )

        paid = st.checkbox(
            "Entrata già ricevuta",
            value=True
        )

        note = st.text_input(
            "Note"
        )

        save = st.form_submit_button(
            "💾 Salva entrata"
        )

        if save:

            if amount <= 0:

                st.error(
                    "Inserisci un importo maggiore di zero."
                )

            elif not description.strip():

                st.error(
                    "Inserisci una descrizione."
                )

            else:

                execute(
                    """
                    INSERT INTO movimenti
                    (
                        tipo,
                        data,
                        categoria,
                        descrizione,
                        persona,
                        importo,
                        pagato,
                        fonte,
                        note
                    )

                    VALUES
                    (
                        'entrata',
                        ?, ?, ?, ?, ?, ?,
                        'manuale',
                        ?
                    )
                    """,
                    (
                        transaction_date.isoformat(),
                        category,
                        description,
                        person,
                        amount,
                        int(paid),
                        note
                    )
                )

                st.success(
                    "Entrata salvata."
                )

                st.rerun()

    st.markdown("---")

    st.subheader(
        "📋 Entrate"
    )

    df = qdf(
        """
        SELECT
            id AS ID,
            data AS Data,
            categoria AS Categoria,
            descrizione AS Descrizione,
            persona AS Persona,
            importo AS Importo,
            pagato AS Pagata

        FROM movimenti

        WHERE tipo='entrata'

        ORDER BY data DESC, id DESC
        """
    )

    if df.empty:

        st.info(
            "Nessuna entrata."
        )

    else:

        df["Pagata"] = df[
            "Pagata"
        ].map(
            lambda x:
            "Sì" if x else "No"
        )

        st.dataframe(
            df.style.format(
                {
                    "Importo":
                        euro
                }
            ),
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# SPESE
# ============================================================
elif menu == "💳 Spese":

    section_title(
        "💳 Spese",
        "Puoi registrare una spesa singola oppure il totale mensile per categoria"
    )

    tab_single, tab_monthly = st.tabs(
        [
            "🧾 Singola spesa",
            "📅 Spese mensili"
        ]
    )

    # --------------------------------------------------------
    # SINGOLA SPESA
    # --------------------------------------------------------
    with tab_single:

        with st.form(
            "form_spesa_singola"
        ):

            transaction_date = st.date_input(
                "Data",
                date.today()
            )

            category = st.selectbox(
                "Categoria",
                EXPENSE_CATEGORIES,
                key="single_expense_category"
            )

            person = st.selectbox(
                "Persona",
                PEOPLE,
                key="single_expense_person"
            )

            description = st.text_input(
                "Descrizione",
                placeholder="Es. Supermercato"
            )

            amount = st.number_input(
                "Importo (€)",
                min_value=0.0,
                step=5.0
            )

            paid = st.checkbox(
                "Spesa già pagata",
                value=True
            )

            note = st.text_input(
                "Note"
            )

            save = st.form_submit_button(
                "💾 Salva spesa"
            )

            if save:

                if amount <= 0:

                    st.error(
                        "Inserisci un importo maggiore di zero."
                    )

                elif not description.strip():

                    st.error(
                        "Inserisci una descrizione."
                    )

                else:

                    execute(
                        """
                        INSERT INTO movimenti
                        (
                            tipo,
                            data,
                            categoria,
                            descrizione,
                            persona,
                            importo,
                            pagato,
                            fonte,
                            note
                        )

                        VALUES
                        (
                            'uscita',
                            ?, ?, ?, ?, ?, ?,
                            'manuale',
                            ?
                        )
                        """,
                        (
                            transaction_date.isoformat(),
                            category,
                            description,
                            person,
                            amount,
                            int(paid),
                            note
                        )
                    )

                    month_key_value = month_key(transaction_date.year, transaction_date.month)
                    if get_month_mode(transaction_date.year, transaction_date.month) == 'mensile':
                        set_month_mode(month_key_value, 'mista')

                    st.success(
                        "Spesa salvata."
                    )

                    st.rerun()

        st.markdown("---")

        st.subheader(
            "📋 Spese singole"
        )

        df = qdf(
            """
            SELECT
                id AS ID,
                data AS Data,
                categoria AS Categoria,
                descrizione AS Descrizione,
                persona AS Persona,
                importo AS Importo,
                pagato AS Pagata

            FROM movimenti

            WHERE tipo='uscita'

            ORDER BY data DESC, id DESC
            """
        )

        if df.empty:

            st.info(
                "Nessuna spesa singola."
            )

        else:

            df["Pagata"] = df[
                "Pagata"
            ].map(
                lambda x:
                "Sì" if x else "No"
            )

            st.dataframe(
                df.style.format(
                    {
                        "Importo":
                            euro
                    }
                ),
                use_container_width=True,
                hide_index=True
            )

    # --------------------------------------------------------
    # SPESE MENSILI
    # --------------------------------------------------------
    with tab_monthly:

        selected = st.date_input(
            "📅 Mese",
            date.today().replace(day=1),
            key="monthly_expense_month"
        )

        key = month_key(
            selected.year,
            selected.month
        )

        current_mode = get_month_mode(selected.year, selected.month)
        mode_keys = list(REGISTRATION_MODES.keys())
        mode_labels = list(REGISTRATION_MODES.values())
        selected_mode_label = st.selectbox(
            "Modalità di conteggio del mese",
            mode_labels,
            index=mode_keys.index(current_mode),
            key=f"mode_{key}",
        )
        selected_mode = mode_keys[mode_labels.index(selected_mode_label)]
        if selected_mode != current_mode:
            set_month_mode(key, selected_mode)
            current_mode = selected_mode
        st.info(mode_description(current_mode))
        if current_mode == 'mista':
            st.warning("⚠️ Modalità Mista: i valori mensili e le singole spese vengono sommati. Usala solo per spese diverse.")

        placeholders = ",".join(
            "?" * len(EXPENSE_CATEGORIES)
        )

        existing = qdf(
            f"""
            SELECT
                id,
                categoria,
                importo,
                fonte

            FROM spese_mensili

            WHERE mese=?

            AND categoria IN ({placeholders})

            ORDER BY id DESC
            """,
            tuple(
                [key]
                + EXPENSE_CATEGORIES
            )
        )

        existing_values = {}
        existing_source = {}

        for _, row in existing.iterrows():

            if row["categoria"] not in existing_values:

                existing_values[
                    row["categoria"]
                ] = float(
                    row["importo"]
                )

                existing_source[
                    row["categoria"]
                ] = row["fonte"]

        if any(
            value == "Excel"
            for value in existing_source.values()
        ):

            st.info(
                "📘 I valori Excel sono caricati come "
                "totali mensili. Premendo 'Salva mese' "
                "diventano valori manuali."
            )

        with st.form(
            "form_spese_mensili"
        ):

            values = {}

            columns = st.columns(2)

            for i, category in enumerate(
                EXPENSE_CATEGORIES
            ):

                with columns[i % 2]:

                    values[category] = st.number_input(
                        category,
                        min_value=0.0,
                        value=float(
                            existing_values.get(
                                category,
                                0.0
                            )
                        ),
                        step=10.0,
                        key=f"monthly_{key}_{i}"
                    )

            save = st.form_submit_button(
                "💾 Salva mese"
            )

            if save:

                for category, amount in values.items():

                    existing_row = qdf(
                        """
                        SELECT id

                        FROM spese_mensili

                        WHERE mese=?
                        AND categoria=?

                        ORDER BY id DESC

                        LIMIT 1
                        """,
                        (
                            key,
                            category
                        )
                    )

                    if existing_row.empty:

                        execute(
                            """
                            INSERT INTO spese_mensili
                            (
                                mese,
                                categoria,
                                importo,
                                fonte
                            )

                            VALUES
                            (
                                ?,
                                ?,
                                ?,
                                'manuale'
                            )
                            """,
                            (
                                key,
                                category,
                                amount
                            )
                        )

                    else:

                        row_id = int(
                            existing_row.iloc[0]["id"]
                        )

                        execute(
                            """
                            UPDATE spese_mensili

                            SET importo=?,
                                fonte='manuale'

                            WHERE id=?
                            """,
                            (
                                amount,
                                row_id
                            )
                        )

                has_single = not qdf("SELECT id FROM movimenti WHERE tipo='uscita' AND substr(data,1,7)=? LIMIT 1", (key,)).empty
                set_month_mode(key, 'mista' if has_single else 'mensile')

                st.success(
                    f"Spese di "
                    f"{month_label(selected.year, selected.month)} "
                    f"salvate."
                )

                st.rerun()

        st.markdown("---")

        view = qdf(
            f"""
            SELECT
                categoria AS Categoria,
                importo AS Importo,
                fonte AS Fonte

            FROM spese_mensili

            WHERE mese=?

            AND categoria IN ({placeholders})

            ORDER BY importo DESC
            """,
            tuple(
                [key]
                + EXPENSE_CATEGORIES
            )
        )

        if not view.empty:

            st.metric(
                "💳 Totale mensile",
                euro(
                    view["Importo"].sum()
                )
            )

            st.dataframe(
                view.style.format(
                    {
                        "Importo":
                            euro
                    }
                ),
                use_container_width=True,
                hide_index=True
            )

        else:

            st.info(
                "Nessuna spesa mensile."
            )


# ============================================================
# SPESE RICORRENTI
# ============================================================
elif menu == "🔁 Spese ricorrenti":

    section_title(
        "🔁 Spese ricorrenti",
        "Spese che si ripetono ogni mese"
    )

    with st.form(
        "form_ricorrente"
    ):

        description = st.text_input(
            "Descrizione"
        )

        category = st.selectbox(
            "Categoria",
            EXPENSE_CATEGORIES,
            key="recurring_category"
        )

        amount = st.number_input(
            "Importo mensile (€)",
            min_value=0.0,
            step=10.0
        )

        day = st.number_input(
            "Giorno previsto",
            min_value=1,
            max_value=31,
            value=1
        )

        person = st.selectbox(
            "Persona",
            PEOPLE,
            key="recurring_person"
        )

        save = st.form_submit_button(
            "➕ Aggiungi"
        )

        if save:

            if (
                not description.strip()
                or amount <= 0
            ):

                st.error(
                    "Inserisci descrizione e importo."
                )

            else:

                execute(
                    """
                    INSERT INTO ricorrenti
                    (
                        descrizione,
                        categoria,
                        importo,
                        giorno,
                        persona,
                        attiva
                    )

                    VALUES
                    (?, ?, ?, ?, ?, 1)
                    """,
                    (
                        description,
                        category,
                        amount,
                        int(day),
                        person
                    )
                )

                st.success(
                    "Ricorrente aggiunta."
                )

                st.rerun()

    st.markdown("---")

    df = qdf(
        """
        SELECT
            id AS ID,
            descrizione AS Descrizione,
            categoria AS Categoria,
            importo AS Importo,
            giorno AS Giorno,
            persona AS Persona,
            attiva AS Attiva

        FROM ricorrenti

        ORDER BY giorno, id
        """
    )

    if df.empty:

        st.info(
            "Nessuna spesa ricorrente."
        )

    else:

        st.metric(
            "Totale ricorrenti attive",
            euro(
                df.loc[
                    df["Attiva"] == 1,
                    "Importo"
                ].sum()
            )
        )

        df["Attiva"] = df[
            "Attiva"
        ].map(
            lambda x:
            "Sì" if x else "No"
        )

        st.dataframe(
            df.style.format(
                {
                    "Importo":
                        euro
                }
            ),
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# SPESE FUTURE
# ============================================================
elif menu == "📌 Spese future":

    section_title(
        "📌 Spese future",
        "Le spese previste vengono considerate nel calcolo di quanto puoi spendere."
    )

    with st.form("future_expense_form", clear_on_submit=True):
        future_date = st.date_input("Data prevista", date.today())
        future_category = st.selectbox("Categoria", EXPENSE_CATEGORIES)
        future_person = st.selectbox("Persona", PEOPLE)
        future_description = st.text_input("Descrizione", placeholder="Es. assicurazione auto")
        future_amount = st.number_input("Importo (€)", min_value=0.0, step=10.0)
        save_future = st.form_submit_button("💾 Salva spesa futura")

    if save_future:
        if future_amount <= 0:
            st.error("Inserisci un importo maggiore di zero.")
        else:
            execute(
                """
                INSERT INTO spese_future
                (data,categoria,descrizione,persona,importo,pagata,note)
                VALUES (?,?,?,?,?,0,'')
                """,
                (future_date.isoformat(), future_category, future_description, future_person, float(future_amount))
            )
            st.success("Spesa futura salvata.")
            st.rerun()

    st.markdown("---")
    future_df = qdf(
        """
        SELECT id AS ID, data AS Data, categoria AS Categoria,
               descrizione AS Descrizione, persona AS Persona,
               importo AS Importo, pagata AS Pagata
        FROM spese_future
        ORDER BY data ASC, id ASC
        """
    )

    if future_df.empty:
        st.info("Nessuna spesa futura registrata.")
    else:
        future_df["Pagata"] = future_df["Pagata"].map(lambda x: "Sì" if x else "No")
        st.dataframe(
            future_df.style.format({"Importo": euro}),
            use_container_width=True,
            hide_index=True
        )

# ============================================================
# BUDGET
# ============================================================
elif menu == "📅 Budget":

    section_title(
        "📅 Budget",
        "Imposta un limite mensile per le 9 categorie"
    )

    selected = st.date_input(
        "📅 Mese",
        date.today().replace(day=1),
        key="budget_month"
    )

    key = month_key(
        selected.year,
        selected.month
    )

    existing = qdf(
        """
        SELECT
            categoria,
            importo

        FROM budget_mensile

        WHERE mese=?
        """,
        (key,)
    )

    budget_values = {}

    if not existing.empty:

        budget_values = dict(
            zip(
                existing["categoria"],
                existing["importo"]
            )
        )

    with st.form(
        "budget_form"
    ):

        values = {}

        columns = st.columns(2)

        for i, category in enumerate(
            EXPENSE_CATEGORIES
        ):

            with columns[i % 2]:

                values[category] = st.number_input(
                    category,
                    min_value=0.0,
                    value=float(
                        budget_values.get(
                            category,
                            0.0
                        )
                    ),
                    step=25.0,
                    key=f"budget_{key}_{i}"
                )

        save = st.form_submit_button(
            "💾 Salva budget"
        )

        if save:

            for category, amount in values.items():

                existing_row = qdf(
                    """
                    SELECT id

                    FROM budget_mensile

                    WHERE mese=?
                    AND categoria=?

                    ORDER BY id DESC

                    LIMIT 1
                    """,
                    (
                        key,
                        category
                    )
                )

                if existing_row.empty:

                    execute(
                        """
                        INSERT INTO budget_mensile
                        (
                            mese,
                            categoria,
                            importo
                        )

                        VALUES (?, ?, ?)
                        """,
                        (
                            key,
                            category,
                            amount
                        )
                    )

                else:

                    execute(
                        """
                        UPDATE budget_mensile

                        SET importo=?

                        WHERE id=?
                        """,
                        (
                            amount,
                            int(
                                existing_row.iloc[0]["id"]
                            )
                        )
                    )

            st.success(
                "Budget salvato."
            )

            st.rerun()

    st.markdown("---")

    st.subheader(
        "📊 Budget vs spesa"
    )

    spent = monthly_category_totals(
        selected.year,
        selected.month
    )

    spent_dict = {}

    if not spent.empty:

        spent_dict = dict(
            zip(
                spent["categoria"],
                spent["importo"]
            )
        )

    budget_df = qdf(
        """
        SELECT
            categoria,
            importo

        FROM budget_mensile

        WHERE mese=?
        """,
        (key,)
    )

    if budget_df.empty:

        st.info(
            "Nessun budget impostato."
        )

    else:

        rows = []

        for _, row in budget_df.iterrows():

            category = row["categoria"]

            budget_value = float(
                row["importo"]
            )

            spent_value = float(
                spent_dict.get(
                    category,
                    0
                )
            )

            residual = (
                budget_value
                - spent_value
            )

            percentage = (
                spent_value
                / budget_value
                * 100
                if budget_value > 0
                else 0
            )

            rows.append(
                {
                    "Categoria":
                        category,

                    "Budget":
                        budget_value,

                    "Speso":
                        spent_value,

                    "Residuo":
                        residual,

                    "% utilizzata":
                        percentage
                }
            )

        result = pd.DataFrame(
            rows
        )

        st.dataframe(
            result.style.format(
                {
                    "Budget":
                        euro,

                    "Speso":
                        euro,

                    "Residuo":
                        euro,

                    "% utilizzata":
                        lambda x:
                        f"{x:.1f}%"
                }
            ),
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# PIANO RISPARMIO
# ============================================================
elif menu == "💡 Piano risparmio":

    section_title(
        "💡 Come posso risparmiare?",
        "L'app confronta le tue spese con la tua media storica e crea un piano pratico"
    )

    months_back = st.slider(
        "Mesi storici da considerare",
        min_value=3,
        max_value=12,
        value=6
    )

    reduction = st.slider(
        "Riduzione obiettivo sulle categorie variabili",
        min_value=0,
        max_value=30,
        value=10,
        step=5,
        format="%d%%"
    ) / 100

    plan = saving_plan(
        months_back,
        reduction
    )

    if plan.empty:

        st.warning(
            "Non ci sono ancora abbastanza dati storici. "
            "Importa il tuo Excel e/o inserisci alcuni mesi."
        )

    else:

        potential = float(
            plan["Risparmio possibile"].sum()
        )

        st.markdown(
            "### 💰 Potenziale di risparmio"
        )

        c1, c2, c3 = st.columns(3)

        c1.metric(
            "Risparmio mensile",
            euro(potential)
        )

        c2.metric(
            "Risparmio annuale",
            euro(potential * 12)
        )

        c3.metric(
            "Categorie analizzate",
            len(plan)
        )

        st.info(
            "Il valore è una stima matematica basata sulla tua "
            "storia. È un suggerimento, non un obbligo."
        )

        st.markdown("---")

        st.subheader(
            "🎯 Il tuo piano"
        )

        st.dataframe(
            plan.style.format(
                {
                    "Media storica":
                        euro,

                    "Mese corrente":
                        euro,

                    "Obiettivo":
                        euro,

                    "Risparmio possibile":
                        euro,
                }
            ),
            use_container_width=True,
            hide_index=True
        )

        st.markdown("---")

        st.subheader(
            "🔎 Dove intervenire"
        )

        actionable = plan[
            plan["Risparmio possibile"] > 0
        ].head(5)

        if actionable.empty:

            st.success(
                "Le categorie risultano già abbastanza "
                "vicine alla media storica."
            )

        else:

            for _, row in actionable.iterrows():

                category = row["Categoria"]

                average = float(
                    row["Media storica"]
                )

                current = float(
                    row["Mese corrente"]
                )

                target = float(
                    row["Obiettivo"]
                )

                saving = float(
                    row["Risparmio possibile"]
                )

                if current > average * 1.15:

                    message = (
                        f"Sei sopra la tua media di circa "
                        f"{euro(current - average)}. "
                        f"Un primo obiettivo potrebbe essere "
                        f"portare la spesa verso "
                        f"{euro(target)}."
                    )

                else:

                    message = (
                        f"Puoi provare a ridurre gradualmente "
                        f"la media da {euro(average)} "
                        f"a circa {euro(target)}."
                    )

                st.markdown(
                    f"**{category}**  \n"
                    f"{message}  \n"
                    f"💰 Risparmio potenziale: "
                    f"**{euro(saving)}/mese** "
                    f"(**{euro(saving * 12)}/anno**)"
                )

                st.markdown("---")

        st.subheader(
            "📋 Obiettivo mensile personalizzato"
        )

        income = monthly_income(
            date.today().year,
            date.today().month
        )

        current_expenses = monthly_total(
            date.today().year,
            date.today().month
        )

        target_saving = st.number_input(
            "Quanto vuoi riuscire a risparmiare ogni mese?",
            min_value=0.0,
            value=float(
                get_setting(
                    "risparmio_mensile_target",
                    0
                )
            ),
            step=50.0
        )

        if income > 0:

            max_expenses = (
                income
                - target_saving
            )

            st.metric(
                "Spesa massima mensile "
                "per raggiungere l'obiettivo",
                euro(
                    max(
                        0,
                        max_expenses
                    )
                )
            )

            difference = (
                max_expenses
                - current_expenses
            )

            if difference >= 0:

                st.success(
                    f"Al ritmo attuale sei dentro il limite "
                    f"di {euro(difference)}."
                )

            else:

                st.warning(
                    f"Per raggiungere l'obiettivo dovresti "
                    f"ridurre le spese di circa "
                    f"{euro(abs(difference))}."
                )

        if st.button(
            "💾 Salva questo obiettivo di risparmio"
        ):

            set_setting(
                "risparmio_mensile_target",
                target_saving
            )

            st.success(
                "Obiettivo salvato."
            )


# ============================================================
# OBIETTIVI
# ============================================================
elif menu == "🎯 Obiettivi":

    section_title(
        "🎯 Obiettivi di risparmio"
    )

    with st.form(
        "goal_form"
    ):

        name = st.text_input(
            "Nome",
            placeholder="Es. Vacanza"
        )

        target = st.number_input(
            "Obiettivo (€)",
            min_value=0.0,
            step=500.0
        )

        saved = st.number_input(
            "Già accumulato (€)",
            min_value=0.0,
            step=100.0
        )

        deadline = st.date_input(
            "Scadenza",
            date.today()
        )

        save = st.form_submit_button(
            "🎯 Crea obiettivo"
        )

        if save:

            if (
                not name.strip()
                or target <= 0
            ):

                st.error(
                    "Inserisci nome e importo."
                )

            else:

                execute(
                    """
                    INSERT INTO obiettivi
                    (
                        nome,
                        obiettivo,
                        accumulato,
                        scadenza,
                        attivo
                    )

                    VALUES
                    (?, ?, ?, ?, 1)
                    """,
                    (
                        name,
                        target,
                        saved,
                        deadline.isoformat()
                    )
                )

                st.success(
                    "Obiettivo creato."
                )

                st.rerun()

    st.markdown("---")

    df = qdf(
        """
        SELECT *
        FROM obiettivi
        WHERE attivo=1
        ORDER BY scadenza
        """
    )

    if df.empty:

        st.info(
            "Nessun obiettivo."
        )

    else:

        for _, row in df.iterrows():

            target = float(
                row["obiettivo"]
            )

            saved = float(
                row["accumulato"]
            )

            progress = min(
                max(
                    saved / target
                    if target
                    else 0,
                    0
                ),
                1
            )

            st.subheader(
                f"🎯 {row['nome']}"
            )

            st.progress(
                progress
            )

            st.write(
                f"{euro(saved)} / "
                f"{euro(target)} "
                f"— {progress * 100:.1f}%"
            )

            st.caption(
                f"Scadenza: {row['scadenza']}"
            )


# ============================================================
# QUANTO POSSO SPENDERE OGGI
# ============================================================
elif menu == "💸 Quanto posso spendere oggi?":

    section_title(
        "💸 Quanto posso spendere oggi?",
        "Stima della disponibilità fino alla fine del mese"
    )

    today = date.today()

    year = today.year
    month = today.month

    income = monthly_income(
        year,
        month
    )

    spent = monthly_total(
        year,
        month
    )

    recurring = recurring_remaining(
        year,
        month
    )

    future = future_remaining(
        year,
        month
    )

    saving_target = get_setting(
        "risparmio_mensile_target",
        0
    )

    safety = get_setting(
        "fondo_sicurezza",
        500
    )

    available = (
        income
        - spent
        - recurring
        - future
        - saving_target
        - safety
    )

    last_day = calendar.monthrange(
        year,
        month
    )[1]

    days_remaining = max(
        last_day - today.day + 1,
        1
    )

    daily = (
        max(
            0,
            available
        )
        / days_remaining
    )

    if available >= 0:

        st.success(
            f"💚 Disponibilità residua: "
            f"**{euro(available)}**"
        )

    else:

        st.error(
            f"🔴 Sei oltre il limite programmato "
            f"di **{euro(abs(available))}**"
        )

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Disponibilità",
        euro(available)
    )

    c2.metric(
        "Giorni rimanenti",
        days_remaining
    )

    c3.metric(
        "Budget giornaliero",
        euro(daily)
    )

    st.markdown("---")

    st.subheader(
        "🔎 Come è calcolato"
    )

    calculation = pd.DataFrame(
        {
            "Voce": [
                "Entrate del mese",
                "Spese già registrate",
                "Ricorrenti ancora previste",
                "Spese future",
                "Risparmio programmato",
                "Fondo sicurezza",
                "Disponibilità"
            ],

            "Importo": [
                income,
                -spent,
                -recurring,
                -future,
                -saving_target,
                -safety,
                available
            ]
        }
    )

    st.dataframe(
        calculation.style.format(
            {
                "Importo":
                    euro
            }
        ),
        use_container_width=True,
        hide_index=True
    )

    st.markdown("---")

    st.subheader(
        "🧮 Simulazione"
    )

    simulation = st.number_input(
        "Se oggi spendessi...",
        min_value=0.0,
        step=10.0
    )

    after = (
        available
        - simulation
    )

    if after >= 0:

        st.success(
            f"Dopo la spesa avresti ancora "
            f"**{euro(after)}**."
        )

    else:

        st.warning(
            f"Dopo la spesa saresti sotto di "
            f"**{euro(abs(after))}**."
        )


# ============================================================
# ANALISI
# ============================================================
elif menu == "📊 Analisi":

    section_title(
        "📊 Analisi",
        "Andamento delle spese e confronto tra categorie"
    )

    selected = st.date_input(
        "Mese",
        date.today().replace(day=1),
        key="analysis_month"
    )

    year = selected.year
    month = selected.month

    income = monthly_income(
        year,
        month
    )

    expenses = monthly_total(
        year,
        month
    )

    categories = monthly_category_totals(
        year,
        month
    )

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Entrate",
        euro(income)
    )

    c2.metric(
        "Spese",
        euro(expenses)
    )

    c3.metric(
        "Risparmio",
        euro(income - expenses)
    )

    if not categories.empty:

        st.markdown("---")

        st.subheader(
            "🥧 Distribuzione"
        )

        fig, ax = plt.subplots(
            figsize=(8, 5)
        )

        ax.pie(
            categories["importo"],
            labels=categories["categoria"],
            autopct="%1.1f%%",
            startangle=90
        )

        ax.axis("equal")

        st.pyplot(
            fig,
            clear_figure=True
        )

        st.dataframe(
            categories.rename(
                columns={
                    "categoria":
                        "Categoria",

                    "importo":
                        "Importo"
                }
            ).style.format(
                {
                    "Importo":
                        euro
                }
            ),
            use_container_width=True,
            hide_index=True
        )

    st.markdown("---")

    st.subheader(
        "📈 Ultimi 12 mesi"
    )

    rows = []

    today = date.today()

    for offset in range(
        11,
        -1,
        -1
    ):

        month_number = (
            today.month - offset
        )

        year = today.year

        while month_number <= 0:

            month_number += 12
            year -= 1

        rows.append(
            {
                "Mese":
                    month_label(
                        year,
                        month_number
                    ),

                "Spese":
                    monthly_total(
                        year,
                        month_number
                    ),

                "Entrate":
                    monthly_income(
                        year,
                        month_number
                    )
            }
        )

    history = pd.DataFrame(
        rows
    )

    fig, ax = plt.subplots(
        figsize=(10, 4)
    )

    ax.plot(
        history["Mese"],
        history["Spese"],
        marker="o",
        label="Spese"
    )

    ax.plot(
        history["Mese"],
        history["Entrate"],
        marker="o",
        label="Entrate"
    )

    ax.tick_params(
        axis="x",
        rotation=45
    )

    ax.set_ylabel("€")
    ax.grid(alpha=0.2)
    ax.legend()

    st.pyplot(
        fig,
        clear_figure=True
    )

    st.dataframe(
        history.style.format(
            {
                "Spese":
                    euro,

                "Entrate":
                    euro
            }
        ),
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# IMPORT / EXPORT
# ============================================================
elif menu == "📥 Import / Export":

    section_title(
        "📥 Import / Export",
        "Gestione dello storico Excel e backup"
    )

    st.subheader(
        "📘 Importa il tuo Excel"
    )

    uploaded = st.file_uploader(
        "Carica 'Spese casa -2.xlsx'",
        type=["xlsx"]
    )

    if uploaded is not None:

        st.session_state[
            "excel_bytes"
        ] = uploaded.getvalue()

        preview = parse_excel_history(
            st.session_state[
                "excel_bytes"
            ]
        )

        if preview.empty:

            st.error(
                "Non ho trovato dati validi "
                "nel foglio 'Costi famiglia'."
            )

        else:

            st.success(
                f"Trovate {len(preview)} "
                "registrazioni mensili nelle 9 categorie."
            )

            st.dataframe(
                preview.style.format(
                    {
                        "importo":
                            euro
                    }
                ),
                use_container_width=True,
                hide_index=True
            )

            if st.button(
                "📥 Importa nello storico"
            ):

                n = import_excel_to_db(
                    st.session_state[
                        "excel_bytes"
                    ]
                )

                st.success(
                    f"Importazione completata: "
                    f"{n} righe elaborate."
                )

                st.rerun()

    elif DEFAULT_EXCEL_FILE.exists():

        st.success(
            "🟢 Ho trovato automaticamente "
            "'Spese casa -2.xlsx'."
        )

        preview = parse_excel_history(
            DEFAULT_EXCEL_FILE
        )

        if not preview.empty:

            st.metric(
                "Registrazioni Excel",
                len(preview)
            )

            st.dataframe(
                preview.style.format(
                    {
                        "importo":
                            euro
                    }
                ),
                use_container_width=True,
                hide_index=True
            )

    else:

        st.warning(
            "Nessun Excel disponibile. "
            "Caricalo qui sopra."
        )

    st.markdown("---")

    st.subheader(
        "⬇️ Backup completo"
    )

    tables = {
        "Movimenti":
            qdf(
                "SELECT * FROM movimenti"
            ),

        "Spese mensili":
            qdf(
                "SELECT * FROM spese_mensili"
            ),

        "Modalita mesi":
            qdf(
                "SELECT * FROM modalita_mese"
            ),

        "Ricorrenti":
            qdf(
                "SELECT * FROM ricorrenti"
            ),

        "Spese future":
            qdf(
                "SELECT * FROM spese_future"
            ),

        "Budget":
            qdf(
                "SELECT * FROM budget_mensile"
            ),

        "Obiettivi":
            qdf(
                "SELECT * FROM obiettivi"
            ),

        "Impostazioni":
            qdf(
                "SELECT * FROM impostazioni"
            ),
    }

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        for sheet, df in tables.items():

            df.to_excel(
                writer,
                sheet_name=sheet[:31],
                index=False
            )

    st.download_button(
        "⬇️ Scarica backup",
        data=output.getvalue(),
        file_name="backup_finanze_famiglia.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


# ============================================================
# IMPOSTAZIONI
# ============================================================
elif menu == "⚙️ Impostazioni":

    section_title(
        "⚙️ Impostazioni",
        "Parametri utilizzati nei calcoli"
    )

    st.subheader(
        "🎯 Risparmio"
    )

    saving = st.number_input(
        "Obiettivo di risparmio mensile (€)",
        min_value=0.0,
        value=float(
            get_setting(
                "risparmio_mensile_target",
                0
            )
        ),
        step=50.0
    )

    safety = st.number_input(
        "Fondo sicurezza da lasciare disponibile (€)",
        min_value=0.0,
        value=float(
            get_setting(
                "fondo_sicurezza",
                500
            )
        ),
        step=50.0
    )

    if st.button(
        "💾 Salva impostazioni"
    ):

        set_setting(
            "risparmio_mensile_target",
            saving
        )

        set_setting(
            "fondo_sicurezza",
            safety
        )

        st.success(
            "Impostazioni salvate."
        )

    st.markdown("---")

    st.subheader(
        "🏷️ Le 9 categorie dell'app"
    )

    st.dataframe(
        pd.DataFrame(
            {
                "N.": range(1, 10),
                "Categoria":
                    EXPENSE_CATEGORIES
            }
        ),
        use_container_width=True,
        hide_index=True
    )

    st.info(
        "Queste sono le sole categorie di spesa utilizzate "
        "dall'app e dall'importazione dello storico Excel."
    )

    st.markdown("---")

    st.subheader(
        "🗑️ Pulizia dati"
    )

    st.warning(
        "Attenzione: queste operazioni eliminano "
        "definitivamente dati dal database."
    )

    action = st.selectbox(
        "Seleziona cosa eliminare",
        [
            "Nessuna operazione",
            "Tutte le spese singole",
            "Tutte le entrate",
            "Tutte le spese mensili",
            "Tutte le modalità dei mesi",
            "Tutti i budget",
            "Tutti gli obiettivi",
            "Tutte le ricorrenti",
            "Tutte le spese future",
        ]
    )

    confirm = st.checkbox(
        "Confermo l'eliminazione"
    )

    if st.button(
        "🗑️ Esegui eliminazione"
    ):

        if action == "Nessuna operazione":

            st.info(
                "Nessuna operazione."
            )

        elif not confirm:

            st.error(
                "Devi confermare."
            )

        else:

            sql_map = {
                "Tutte le spese singole":
                    "DELETE FROM movimenti WHERE tipo='uscita'",

                "Tutte le entrate":
                    "DELETE FROM movimenti WHERE tipo='entrata'",

                "Tutte le spese mensili":
                    "DELETE FROM spese_mensili",

                "Tutte le modalità dei mesi":
                    "DELETE FROM modalita_mese",

                "Tutti i budget":
                    "DELETE FROM budget_mensile",

                "Tutti gli obiettivi":
                    "DELETE FROM obiettivi",

                "Tutte le ricorrenti":
                    "DELETE FROM ricorrenti",

                "Tutte le spese future":
                    "DELETE FROM spese_future",
            }

            execute(
                sql_map[action]
            )

            st.success(
                "Eliminazione completata."
            )

            st.rerun()
