import streamlit as st
import pandas as pd
import sqlite3
import matplotlib.pyplot as plt

from pathlib import Path
from datetime import date, datetime
from io import BytesIO
import calendar
import re


# ============================================================
# CONFIGURAZIONE
# ============================================================

st.set_page_config(
    page_title="Gestione Finanze - Arletti",
    page_icon="💰",
    layout="centered",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent

DB_FILE = BASE_DIR / "finanze_famiglia.db"
EXCEL_FILE = BASE_DIR / "Spese casa -2.xlsx"


# ============================================================
# MESI
# ============================================================

MONTHS = [
    "Gennaio",
    "Febbraio",
    "Marzo",
    "Aprile",
    "Maggio",
    "Giugno",
    "Luglio",
    "Agosto",
    "Settembre",
    "Ottobre",
    "Novembre",
    "Dicembre",
]

MONTH_MAP = {
    month.upper(): i
    for i, month in enumerate(MONTHS, start=1)
}


# ============================================================
# LE 9 CATEGORIE DEFINITIVE
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


# ============================================================
# CATEGORIE ENTRATE
# ============================================================

INCOME_CATEGORIES = [
    "Stipendio",
    "Bonus",
    "Rimborso",
    "Affitto",
    "Altre entrate",
]


# ============================================================
# PERSONE
# ============================================================

PEOPLE = [
    "Famiglia",
    "Mattia",
    "Partner",
]


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
    <style>

    .block-container {
        max-width: 1100px;
        padding-top: 1rem;
        padding-bottom: 3rem;
    }

    [data-testid="stMetricValue"] {
        font-size: 1.4rem;
    }

    .small-text {
        font-size: 0.85rem;
        color: #777;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# FUNZIONI GENERALI
# ============================================================

def euro(value):
    """Formatta un numero in euro."""

    try:
        value = float(value)
    except Exception:
        value = 0.0

    return (
        f"{value:,.2f} €"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


def safe_float(value, default=0.0):

    try:

        if value is None:
            return default

        if pd.isna(value):
            return default

    except Exception:
        pass

    try:
        return float(value)
    except Exception:
        return default


def parse_number(value):

    if value is None:
        return None

    try:

        if pd.isna(value):
            return None

    except Exception:
        pass

    if isinstance(value, (int, float)):

        return float(value)

    text = str(value).strip()

    if not text:
        return None

    text = (
        text
        .replace("€", "")
        .replace(" ", "")
    )

    # formato italiano:
    # 1.234,56
    if "," in text:

        text = (
            text
            .replace(".", "")
            .replace(",", ".")
        )

    try:
        return float(text)
    except Exception:
        return None


def month_key(year, month):

    return f"{year:04d}-{month:02d}"


def month_label(year, month):

    return f"{MONTHS[month - 1]} {year}"


def first_day_of_month():

    today = date.today()

    return today.replace(day=1)


# ============================================================
# DATABASE
# ============================================================

def get_connection():

    conn = sqlite3.connect(
        DB_FILE,
        timeout=30
    )

    conn.execute(
        "PRAGMA busy_timeout = 30000"
    )

    return conn


def query_df(sql, params=()):

    conn = get_connection()

    try:

        return pd.read_sql_query(
            sql,
            conn,
            params=params
        )

    finally:

        conn.close()


def execute(sql, params=()):

    conn = get_connection()

    try:

        cursor = conn.cursor()

        cursor.execute(
            sql,
            params
        )

        conn.commit()

        return cursor.lastrowid

    finally:

        conn.close()


# ============================================================
# INIZIALIZZAZIONE DATABASE
# ============================================================

def create_database():

    conn = get_connection()

    cursor = conn.cursor()

    cursor.executescript(
        """
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

            importo REAL NOT NULL,

            fonte TEXT DEFAULT 'manuale',

            note TEXT
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

            importo REAL NOT NULL
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


        CREATE TABLE IF NOT EXISTS categorie (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            nome TEXT UNIQUE NOT NULL,

            tipo TEXT NOT NULL,

            attiva INTEGER DEFAULT 1
        );
        """
    )

    conn.commit()

    conn.close()


create_database()


# ============================================================
# MIGRAZIONE / CONTROLLO COLONNE
# ============================================================

def ensure_column(table, column, definition):

    conn = get_connection()

    try:

        columns = pd.read_sql_query(
            f"PRAGMA table_info({table})",
            conn
        )

        existing = columns["name"].tolist()

        if column not in existing:

            conn.execute(
                f"""
                ALTER TABLE {table}
                ADD COLUMN {column} {definition}
                """
            )

            conn.commit()

    finally:

        conn.close()


# Se il database era stato creato con una versione precedente
# aggiungiamo eventuali colonne mancanti.

ensure_column(
    "movimenti",
    "note",
    "TEXT"
)

ensure_column(
    "spese_mensili",
    "fonte",
    "TEXT DEFAULT 'manuale'"
)

ensure_column(
    "spese_mensili",
    "note",
    "TEXT"
)

ensure_column(
    "spese_future",
    "note",
    "TEXT"
)


# ============================================================
# IMPOSTAZIONI
# ============================================================

def get_setting(key, default=0):

    df = query_df(
        """
        SELECT valore

        FROM impostazioni

        WHERE chiave=?
        """,
        (key,)
    )

    if df.empty:

        return default

    return safe_float(
        df.iloc[0]["valore"],
        default
    )


def set_setting(key, value):

    # Non usiamo ON CONFLICT.
    # Prima eliminiamo l'eventuale valore.

    execute(
        """
        DELETE FROM impostazioni
        WHERE chiave=?
        """,
        (key,)
    )

    execute(
        """
        INSERT INTO impostazioni
        (
            chiave,
            valore
        )

        VALUES (?, ?)
        """,
        (
            key,
            str(value)
        )
    )


# ============================================================
# DATI INIZIALI
# ============================================================

def initialize_default_data():

    conn = get_connection()

    cursor = conn.cursor()

    # --------------------------------------------------------
    # Impostazioni
    # --------------------------------------------------------

    cursor.execute(
        """
        INSERT OR IGNORE INTO impostazioni
        (
            chiave,
            valore
        )

        VALUES
        (
            'risparmio_mensile_target',
            '0'
        )
        """
    )

    cursor.execute(
        """
        INSERT OR IGNORE INTO impostazioni
        (
            chiave,
            valore
        )

        VALUES
        (
            'fondo_sicurezza',
            '500'
        )
        """
    )

    # --------------------------------------------------------
    # Persone
    # --------------------------------------------------------

    for person in PEOPLE:

        cursor.execute(
            """
            INSERT OR IGNORE INTO persone
            (
                nome
            )

            VALUES (?)
            """,
            (person,)
        )

    # --------------------------------------------------------
    # Categorie spese
    # --------------------------------------------------------

    for category in EXPENSE_CATEGORIES:

        cursor.execute(
            """
            INSERT OR IGNORE INTO categorie
            (
                nome,
                tipo
            )

            VALUES (?, 'uscita')
            """,
            (category,)
        )

    # --------------------------------------------------------
    # Categorie entrate
    # --------------------------------------------------------

    for category in INCOME_CATEGORIES:

        cursor.execute(
            """
            INSERT OR IGNORE INTO categorie
            (
                nome,
                tipo
            )

            VALUES (?, 'entrata')
            """,
            (category,)
        )

    conn.commit()

    conn.close()


initialize_default_data()


# ============================================================
# NORMALIZZAZIONE CATEGORIE EXCEL
# ============================================================

def normalize_excel_category(text):

    if text is None:
        return None

    text = str(text).strip()

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    # --------------------------------------------------------
    # Tutte le possibili varianti che vogliamo riconoscere
    # --------------------------------------------------------

    normalized = {

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

    return normalized.get(
        text
    )


# ============================================================
# LETTURA EXCEL
# ============================================================

def read_excel_history():

    if not EXCEL_FILE.exists():

        return pd.DataFrame(
            columns=[
                "mese",
                "categoria",
                "importo"
            ]
        )

    try:

        import openpyxl

        workbook = openpyxl.load_workbook(
            EXCEL_FILE,
            data_only=True
        )

    except Exception:

        return pd.DataFrame(
            columns=[
                "mese",
                "categoria",
                "importo"
            ]
        )

    if "Costi famiglia" not in workbook.sheetnames:

        return pd.DataFrame(
            columns=[
                "mese",
                "categoria",
                "importo"
            ]
        )

    ws = workbook["Costi famiglia"]

    rows = []

    current_year = None
    current_month = None

    for row_number in range(
        1,
        ws.max_row + 1
    ):

        value_a = ws.cell(
            row_number,
            1
        ).value

        if value_a is None:
            continue

        text = str(
            value_a
        ).strip()

        upper = text.upper()

        # ----------------------------------------------------
        # Cerca:
        #
        # Spese SETTEMBRE 2026
        # ----------------------------------------------------

        match = re.search(
            r"SPESE\s+([A-ZÀ-Ù]+)\s+(\d{4})",
            upper
        )

        if match:

            month_name = match.group(1)

            year = int(
                match.group(2)
            )

            if month_name in MONTH_MAP:

                current_month = MONTH_MAP[
                    month_name
                ]

                current_year = year

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

        # Colonna B = importo
        amount = parse_number(
            ws.cell(
                row_number,
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

                "importo": amount
            }
        )

    if not rows:

        return pd.DataFrame(
            columns=[
                "mese",
                "categoria",
                "importo"
            ]
        )

    df = pd.DataFrame(
        rows
    )

    # --------------------------------------------------------
    # Se per qualche motivo ci sono più righe della stessa
    # categoria nello stesso mese, le sommiamo.
    # --------------------------------------------------------

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

    return df


# ============================================================
# IMPORTAZIONE EXCEL
# ============================================================

def import_excel():

    df = read_excel_history()

    if df.empty:
        return 0

    count = 0

    for _, row in df.iterrows():

        mese = row["mese"]
        categoria = row["categoria"]
        importo = safe_float(
            row["importo"]
        )

        existing = query_df(
            """
            SELECT
                id,
                fonte

            FROM spese_mensili

            WHERE mese=?
            AND categoria=?

            ORDER BY id DESC
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
                (
                    ?,
                    ?,
                    ?,
                    'Excel'
                )
                """,
                (
                    mese,
                    categoria,
                    importo
                )
            )

        else:

            # Se è ancora un dato Excel,
            # lo aggiorniamo.

            fonte = str(
                existing.iloc[0]["fonte"]
            )

            if fonte.lower() == "excel":

                execute(
                    """
                    UPDATE spese_mensili

                    SET
                        importo=?,
                        fonte='Excel'

                    WHERE mese=?
                    AND categoria=?
                    """,
                    (
                        importo,
                        mese,
                        categoria
                    )
                )

        count += 1

    return count


# ============================================================
# IMPORT AUTOMATICO
# ============================================================

if EXCEL_FILE.exists():

    import_excel()


# ============================================================
# SALVATAGGIO SPESA MENSILE
# ============================================================

def save_monthly_category(
    mese,
    categoria,
    importo
):

    """
    Questa funzione NON usa ON CONFLICT.

    Elimina prima eventuali vecchi record della stessa
    categoria/mese e poi inserisce quello nuovo.

    In questo modo funziona anche con vecchi database SQLite
    che non avevano la UNIQUE(mese, categoria).
    """

    conn = get_connection()

    try:

        cursor = conn.cursor()

        # ----------------------------------------------------
        # Elimina eventuali duplicati precedenti
        # ----------------------------------------------------

        cursor.execute(
            """
            DELETE FROM spese_mensili

            WHERE mese=?
            AND categoria=?
            """,
            (
                mese,
                categoria
            )
        )

        # ----------------------------------------------------
        # Inserisce il nuovo valore
        # ----------------------------------------------------

        cursor.execute(
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
                mese,
                categoria,
                float(importo)
            )
        )

        conn.commit()

    finally:

        conn.close()


# ============================================================
# ELIMINA SPESA MENSILE
# ============================================================

def delete_monthly_category(
    mese,
    categoria
):

    execute(
        """
        DELETE FROM spese_mensili

        WHERE mese=?
        AND categoria=?
        """,
        (
            mese,
            categoria
        )
    )


# ============================================================
# SPESE MENSILI PER CATEGORIA
# ============================================================

def get_monthly_category_totals(
    year,
    month
):

    mese = month_key(
        year,
        month
    )

    # --------------------------------------------------------
    # Totali mensili
    # --------------------------------------------------------

    monthly = query_df(
        """
        SELECT
            categoria,
            SUM(importo) AS importo

        FROM spese_mensili

        WHERE mese=?

        GROUP BY categoria
        """,
        (mese,)
    )

    # --------------------------------------------------------
    # Singole spese
    # --------------------------------------------------------

    individual = query_df(
        """
        SELECT
            categoria,
            SUM(importo) AS importo

        FROM movimenti

        WHERE tipo='uscita'

        AND substr(data, 1, 7)=?

        AND pagato=1

        GROUP BY categoria
        """,
        (mese,)
    )

    frames = []

    if not monthly.empty:

        monthly["fonte_calcolo"] = "Mensile"

        frames.append(
            monthly
        )

    if not individual.empty:

        individual["fonte_calcolo"] = "Singole"

        frames.append(
            individual
        )

    if not frames:

        return pd.DataFrame(
            columns=[
                "categoria",
                "importo"
            ]
        )

    combined = pd.concat(
        frames,
        ignore_index=True
    )

    result = (
        combined
        .groupby(
            "categoria",
            as_index=False
        )["importo"]
        .sum()
    )

    return result


def get_monthly_total(
    year,
    month
):

    df = get_monthly_category_totals(
        year,
        month
    )

    if df.empty:
        return 0.0

    return float(
        df["importo"].sum()
    )


# ============================================================
# ENTRATE MENSILI
# ============================================================

def get_monthly_income(
    year,
    month
):

    mese = month_key(
        year,
        month
    )

    df = query_df(
        """
        SELECT
            COALESCE(
                SUM(importo),
                0
            ) AS totale

        FROM movimenti

        WHERE tipo='entrata'

        AND substr(data,1,7)=?

        AND pagato=1
        """,
        (mese,)
    )

    return safe_float(
        df.iloc[0]["totale"]
    )


# ============================================================
# SPESE RICORRENTI
# ============================================================

def get_recurring():

    return query_df(
        """
        SELECT *

        FROM ricorrenti

        ORDER BY giorno, descrizione
        """
    )


def get_recurring_remaining():

    today = date.today()

    df = query_df(
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

        day = int(
            safe_float(
                row["giorno"],
                1
            )
        )

        if day >= today.day:

            total += safe_float(
                row["importo"]
            )

    return total


# ============================================================
# SPESE FUTURE
# ============================================================

def get_future_expenses_current_month():

    today = date.today()

    last_day = calendar.monthrange(
        today.year,
        today.month
    )[1]

    end = date(
        today.year,
        today.month,
        last_day
    )

    df = query_df(
        """
        SELECT *

        FROM spese_future

        WHERE pagata=0

        AND data>=?

        AND data<=?

        ORDER BY data
        """,
        (
            today.isoformat(),
            end.isoformat()
        )
    )

    if df.empty:
        return 0.0

    return float(
        df["importo"].sum()
    )


# ============================================================
# MEDIA STORICA PER CATEGORIA
# ============================================================

def get_historical_category_data():

    """
    Crea uno storico mensile per categoria.

    Include sia:
    - dati Excel / spese mensili
    - singole spese registrate nell'app
    """

    monthly = query_df(
        """
        SELECT
            mese,
            categoria,
            SUM(importo) AS importo

        FROM spese_mensili

        GROUP BY
            mese,
            categoria
        """
    )

    individual = query_df(
        """
        SELECT
            substr(data,1,7) AS mese,
            categoria,
            SUM(importo) AS importo

        FROM movimenti

        WHERE tipo='uscita'

        AND pagato=1

        GROUP BY
            substr(data,1,7),
            categoria
        """
    )

    frames = []

    if not monthly.empty:
        frames.append(
            monthly
        )

    if not individual.empty:
        frames.append(
            individual
        )

    if not frames:

        return pd.DataFrame(
            columns=[
                "mese",
                "categoria",
                "importo"
            ]
        )

    df = pd.concat(
        frames,
        ignore_index=True
    )

    # Se nella stessa categoria/mese esistono entrambe
    # le modalità, vengono sommate.
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

    return df


# ============================================================
# ANALISI RISPARMIO
# ============================================================

def calculate_saving_plan():

    history = get_historical_category_data()

    if history.empty:

        return pd.DataFrame()

    # --------------------------------------------------------
    # Usiamo gli ultimi 6 mesi disponibili.
    # --------------------------------------------------------

    available_months = sorted(
        history["mese"].unique()
    )

    last_months = available_months[-6:]

    history = history[
        history["mese"].isin(
            last_months
        )
    ].copy()

    # --------------------------------------------------------
    # Media mensile per categoria
    # --------------------------------------------------------

    averages = (
        history
        .groupby(
            "categoria"
        )["importo"]
        .mean()
        .reset_index()
    )

    averages.rename(
        columns={
            "importo":
                "media"
        },
        inplace=True
    )

    # --------------------------------------------------------
    # Ultimo mese disponibile
    # --------------------------------------------------------

    latest_month = max(
        last_months
    )

    latest = history[
        history["mese"] == latest_month
    ][
        [
            "categoria",
            "importo"
        ]
    ].copy()

    latest.rename(
        columns={
            "importo":
                "ultimo_mese"
        },
        inplace=True
    )

    result = averages.merge(
        latest,
        on="categoria",
        how="left"
    )

    result["ultimo_mese"] = (
        result["ultimo_mese"]
        .fillna(0)
    )

    # --------------------------------------------------------
    # Differenza dall'ultimo mese
    # --------------------------------------------------------

    result["differenza"] = (
        result["ultimo_mese"]
        - result["media"]
    )

    # --------------------------------------------------------
    # Volatilità
    # --------------------------------------------------------

    volatility = (
        history
        .groupby(
            "categoria"
        )["importo"]
        .std()
        .reset_index()
    )

    volatility.rename(
        columns={
            "importo":
                "variabilita"
        },
        inplace=True
    )

    result = result.merge(
        volatility,
        on="categoria",
        how="left"
    )

    result["variabilita"] = (
        result["variabilita"]
        .fillna(0)
    )

    # --------------------------------------------------------
    # Obiettivo di risparmio.
    #
    # Usiamo una riduzione prudente:
    #
    # 10% per categorie normali
    # 15% per categorie dove l'ultimo mese è molto sopra
    # la media.
    #
    # Evitiamo di suggerire tagli su categorie quasi nulle.
    # --------------------------------------------------------

    targets = []
    savings = []

    for _, row in result.iterrows():

        avg = safe_float(
            row["media"]
        )

        latest = safe_float(
            row["ultimo_mese"]
        )

        difference = safe_float(
            row["differenza"]
        )

        if avg <= 0:

            target = 0
            saving = 0

        elif difference > avg * 0.20:

            # Ultimo mese >20% sopra la media
            reduction = 0.15

            target = avg * (
                1 - reduction
            )

            saving = avg - target

        else:

            reduction = 0.10

            target = avg * (
                1 - reduction
            )

            saving = avg - target

        # arrotondamento
        target = round(
            max(target, 0),
            2
        )

        saving = round(
            max(saving, 0),
            2
        )

        targets.append(
            target
        )

        savings.append(
            saving
        )

    result["obiettivo"] = targets

    result["risparmio_possibile"] = savings

    # --------------------------------------------------------
    # Percentuale rispetto alla media
    # --------------------------------------------------------

    result["variazione_percentuale"] = result.apply(
        lambda row:
        (
            (
                row["ultimo_mese"]
                - row["media"]
            )
            / row["media"]
            * 100
        )
        if row["media"] > 0
        else 0,
        axis=1
    )

    return result.sort_values(
        "risparmio_possibile",
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
        "📅 Budget",
        "💡 Come posso risparmiare?",
        "🎯 Obiettivi",
        "💸 Quanto posso spendere oggi?",
        "📊 Analisi",
        "📥 Import / Export",
        "⚙️ Impostazioni",
    ]
)

st.sidebar.markdown("---")

if EXCEL_FILE.exists():

    st.sidebar.success(
        "🟢 Excel collegato"
    )

    st.sidebar.caption(
        "Spese casa -2.xlsx"
    )

else:

    st.sidebar.warning(
        "🟡 Excel non trovato"
    )

    st.sidebar.caption(
        "Inserisci Spese casa -2.xlsx "
        "nella cartella dell'app"
    )


# ============================================================
# HOME
# ============================================================

if menu == "🏠 Home":

    st.title(
        "🏠 Dashboard Finanze"
    )

    st.caption(
        "Controllo delle spese familiari"
    )

    selected = st.date_input(
        "📅 Mese",
        first_day_of_month()
    )

    year = selected.year
    month = selected.month

    income = get_monthly_income(
        year,
        month
    )

    expenses = get_monthly_total(
        year,
        month
    )

    savings = (
        income
        - expenses
    )

    saving_percentage = (
        savings / income * 100
        if income > 0
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
        f"{saving_percentage:.1f}%"
    )

    st.markdown("---")

    st.subheader(
        f"📊 {month_label(year, month)}"
    )

    categories = get_monthly_category_totals(
        year,
        month
    )

    if categories.empty:

        st.info(
            "Non ci sono ancora spese registrate."
        )

    else:

        categories_display = categories.rename(
            columns={
                "categoria": "Categoria",
                "importo": "Importo"
            }
        )

        st.dataframe(
            categories_display.style.format(
                {
                    "Importo": euro
                }
            ),
            use_container_width=True,
            hide_index=True
        )

        fig, ax = plt.subplots(
            figsize=(8, 6)
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

    saving_target = get_setting(
        "risparmio_mensile_target",
        0
    )

    if saving_target > 0:

        st.subheader(
            "🎯 Obiettivo di risparmio"
        )

        difference = (
            savings
            - saving_target
        )

        if difference >= 0:

            st.success(
                f"Hai raggiunto l'obiettivo: "
                f"sei sopra di {euro(difference)}."
            )

        else:

            st.warning(
                f"Ti mancano "
                f"{euro(abs(difference))} "
                f"per raggiungere l'obiettivo."
            )


# ============================================================
# ENTRATE
# ============================================================

elif menu == "💰 Entrate":

    st.title(
        "💰 Entrate"
    )

    st.caption(
        "Inserisci stipendio, bonus, rimborsi o altre entrate."
    )

    with st.form(
        "new_income"
    ):

        description = st.text_input(
            "Descrizione",
            placeholder="Es. Stipendio ottobre"
        )

        col1, col2 = st.columns(2)

        with col1:

            category = st.selectbox(
                "Categoria",
                INCOME_CATEGORIES
            )

        with col2:

            person = st.selectbox(
                "Persona",
                PEOPLE
            )

        amount = st.number_input(
            "Importo (€)",
            min_value=0.0,
            step=50.0
        )

        transaction_date = st.date_input(
            "Data",
            date.today()
        )

        paid = st.checkbox(
            "Già ricevuta",
            value=True
        )

        note = st.text_input(
            "Note"
        )

        save = st.form_submit_button(
            "💾 Salva entrata"
        )

        if save:

            if not description.strip():

                st.error(
                    "Inserisci una descrizione."
                )

            elif amount <= 0:

                st.error(
                    "Inserisci un importo maggiore di zero."
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
                        ?,
                        ?,
                        ?,
                        ?,
                        ?,
                        ?,
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
        "📋 Entrate registrate"
    )

    df = query_df(
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
            "Nessuna entrata registrata."
        )

    else:

        df["Pagata"] = df[
            "Pagata"
        ].apply(
            lambda x:
            "Sì" if x else "No"
        )

        st.dataframe(
            df.style.format(
                {
                    "Importo": euro
                }
            ),
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# SPESE
# ============================================================

elif menu == "💳 Spese":

    st.title(
        "💳 Spese"
    )

    st.caption(
        "Puoi registrare una spesa singola oppure il totale mensile per categoria."
    )

    tab_single, tab_monthly = st.tabs(
        [
            "🧾 Singola spesa",
            "📅 Spese mensili"
        ]
    )

    # ========================================================
    # SINGOLA
    # ========================================================

    with tab_single:

        st.subheader(
            "🧾 Nuova spesa"
        )

        with st.form(
            "single_expense"
        ):

            expense_date = st.date_input(
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
                "Già pagata",
                value=True
            )

            note = st.text_input(
                "Note",
                key="single_note"
            )

            save = st.form_submit_button(
                "💾 Salva spesa"
            )

            if save:

                if not description.strip():

                    st.error(
                        "Inserisci una descrizione."
                    )

                elif amount <= 0:

                    st.error(
                        "Inserisci un importo."
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
                            ?,
                            ?,
                            ?,
                            ?,
                            ?,
                            ?,
                            'manuale',
                            ?
                        )
                        """,
                        (
                            expense_date.isoformat(),
                            category,
                            description,
                            person,
                            amount,
                            int(paid),
                            note
                        )
                    )

                    st.success(
                        "Spesa salvata."
                    )

                    st.rerun()

        st.markdown("---")

        st.subheader(
            "📋 Spese singole"
        )

        df = query_df(
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
            ].apply(
                lambda x:
                "Sì" if x else "No"
            )

            st.dataframe(
                df.style.format(
                    {
                        "Importo": euro
                    }
                ),
                use_container_width=True,
                hide_index=True
            )

    # ========================================================
    # MENSILE
    # ========================================================

    with tab_monthly:

        st.subheader(
            "📅 Totale mensile per categoria"
        )

        selected = st.date_input(
            "Mese",
            first_day_of_month(),
            key="monthly_month"
        )

        mese = month_key(
            selected.year,
            selected.month
        )

        existing = query_df(
            """
            SELECT
                categoria,
                importo,
                fonte

            FROM spese_mensili

            WHERE mese=?
            """,
            (mese,)
        )

        existing_values = {}

        existing_sources = {}

        for _, row in existing.iterrows():

            existing_values[
                row["categoria"]
            ] = safe_float(
                row["importo"]
            )

            existing_sources[
                row["categoria"]
            ] = row["fonte"]

        if not existing.empty:

            if any(
                str(x).lower() == "excel"
                for x in existing["fonte"]
            ):

                st.info(
                    "📘 Questo mese contiene dati importati dall'Excel. "
                    "Se salvi il mese, i valori modificati diventano manuali."
                )

        with st.form(
            "monthly_expenses"
        ):

            values = {}

            col1, col2 = st.columns(2)

            for index, category in enumerate(
                EXPENSE_CATEGORIES
            ):

                with (
                    col1
                    if index % 2 == 0
                    else col2
                ):

                    values[category] = st.number_input(
                        category,
                        min_value=0.0,
                        value=existing_values.get(
                            category,
                            0.0
                        ),
                        step=10.0,
                        key=(
                            "monthly_value_"
                            + mese
                            + "_"
                            + str(index)
                        )
                    )

            save = st.form_submit_button(
                "💾 Salva mese"
            )

            if save:

                try:

                    for category, amount in values.items():

                        save_monthly_category(
                            mese,
                            category,
                            amount
                        )

                    st.success(
                        f"Spese di "
                        f"{month_label(selected.year, selected.month)} "
                        f"salvate correttamente."
                    )

                    st.rerun()

                except Exception as error:

                    st.error(
                        "Errore durante il salvataggio: "
                        f"{error}"
                    )

        st.markdown("---")

        monthly_data = query_df(
            """
            SELECT
                categoria AS Categoria,
                importo AS Importo,
                fonte AS Fonte

            FROM spese_mensili

            WHERE mese=?

            ORDER BY importo DESC
            """,
            (mese,)
        )

        if monthly_data.empty:

            st.info(
                "Nessun dato mensile."
            )

        else:

            total = monthly_data[
                "Importo"
            ].sum()

            st.metric(
                "💳 Totale mese",
                euro(total)
            )

            st.dataframe(
                monthly_data.style.format(
                    {
                        "Importo": euro
                    }
                ),
                use_container_width=True,
                hide_index=True
            )

            st.warning(
                "⚠️ Se inserisci sia il totale mensile "
                "di una categoria sia singole spese della stessa "
                "categoria nello stesso mese, l'app le considera "
                "entrambe."
            )


# ============================================================
# SPESE RICORRENTI
# ============================================================

elif menu == "🔁 Spese ricorrenti":

    st.title(
        "🔁 Spese ricorrenti"
    )

    st.caption(
        "Spese che si ripetono ogni mese."
    )

    with st.form(
        "recurring_form"
    ):

        description = st.text_input(
            "Descrizione",
            placeholder="Es. Mutuo"
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
            "Giorno del mese",
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

            if not description.strip():

                st.error(
                    "Inserisci una descrizione."
                )

            elif amount <= 0:

                st.error(
                    "Inserisci un importo."
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

                    VALUES (?, ?, ?, ?, ?, 1)
                    """,
                    (
                        description,
                        category,
                        amount,
                        day,
                        person
                    )
                )

                st.success(
                    "Spesa ricorrente aggiunta."
                )

                st.rerun()

    st.markdown("---")

    recurring = get_recurring()

    if recurring.empty:

        st.info(
            "Nessuna spesa ricorrente."
        )

    else:

        active_total = recurring.loc[
            recurring["attiva"] == 1,
            "importo"
        ].sum()

        st.metric(
            "💳 Ricorrenti mensili",
            euro(active_total)
        )

        display = recurring.rename(
            columns={
                "id": "ID",
                "descrizione": "Descrizione",
                "categoria": "Categoria",
                "importo": "Importo",
                "giorno": "Giorno",
                "persona": "Persona",
                "attiva": "Attiva"
            }
        )

        display["Attiva"] = display[
            "Attiva"
        ].apply(
            lambda x:
            "Sì" if x else "No"
        )

        st.dataframe(
            display.style.format(
                {
                    "Importo": euro
                }
            ),
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# BUDGET
# ============================================================

elif menu == "📅 Budget":

    st.title(
        "📅 Budget"
    )

    st.caption(
        "Imposta quanto vuoi spendere al massimo per ogni categoria."
    )

    selected = st.date_input(
        "Mese",
        first_day_of_month(),
        key="budget_month"
    )

    mese = month_key(
        selected.year,
        selected.month
    )

    current_budget = query_df(
        """
        SELECT
            categoria,
            importo

        FROM budget_mensile

        WHERE mese=?
        """,
        (mese,)
    )

    budget_values = {}

    for _, row in current_budget.iterrows():

        budget_values[
            row["categoria"]
        ] = safe_float(
            row["importo"]
        )

    with st.form(
        "budget_form"
    ):

        values = {}

        col1, col2 = st.columns(2)

        for index, category in enumerate(
            EXPENSE_CATEGORIES
        ):

            with (
                col1
                if index % 2 == 0
                else col2
            ):

                values[category] = st.number_input(
                    category,
                    min_value=0.0,
                    value=budget_values.get(
                        category,
                        0.0
                    ),
                    step=25.0,
                    key=(
                        "budget_"
                        + mese
                        + "_"
                        + str(index)
                    )
                )

        save = st.form_submit_button(
            "💾 Salva budget"
        )

        if save:

            for category, amount in values.items():

                # ------------------------------------------------
                # Niente ON CONFLICT:
                # prima cancelliamo il vecchio record.
                # ------------------------------------------------

                execute(
                    """
                    DELETE FROM budget_mensile

                    WHERE mese=?
                    AND categoria=?
                    """,
                    (
                        mese,
                        category
                    )
                )

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
                        mese,
                        category,
                        amount
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

    budget = query_df(
        """
        SELECT
            categoria,
            importo

        FROM budget_mensile

        WHERE mese=?
        """,
        (mese,)
    )

    spent = get_monthly_category_totals(
        selected.year,
        selected.month
    )

    if budget.empty:

        st.info(
            "Non hai ancora impostato il budget."
        )

    else:

        rows = []

        for _, row in budget.iterrows():

            category = row["categoria"]

            budget_value = safe_float(
                row["importo"]
            )

            spent_value = 0.0

            if not spent.empty:

                match = spent[
                    spent["categoria"]
                    == category
                ]

                if not match.empty:

                    spent_value = safe_float(
                        match.iloc[0]["importo"]
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
                    "Budget": euro,
                    "Speso": euro,
                    "Residuo": euro,
                    "% utilizzata":
                        lambda x:
                        f"{x:.1f}%"
                }
            ),
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# COME POSSO RISPARMIARE?
# ============================================================

elif menu == "💡 Come posso risparmiare?":

    st.title(
        "💡 Come posso risparmiare?"
    )

    st.caption(
        "L'app analizza le tue spese storiche e individua "
        "le categorie dove esiste maggiore possibilità di riduzione."
    )

    plan = calculate_saving_plan()

    if plan.empty:

        st.warning(
            "Non ho ancora abbastanza dati per creare un piano "
            "di risparmio. Inserisci almeno qualche mese di spese."
        )

    else:

        # ----------------------------------------------------
        # RISPARMIO TOTALE
        # ----------------------------------------------------

        total_saving = plan[
            "risparmio_possibile"
        ].sum()

        annual_saving = (
            total_saving
            * 12
        )

        c1, c2 = st.columns(2)

        c1.metric(
            "💰 Possibile risparmio mensile",
            euro(total_saving)
        )

        c2.metric(
            "📅 Possibile risparmio annuale",
            euro(annual_saving)
        )

        st.markdown("---")

        # ----------------------------------------------------
        # CATEGORIE PRIORITARIE
        # ----------------------------------------------------

        st.subheader(
            "🎯 Dove puoi intervenire"
        )

        for _, row in plan.iterrows():

            category = row["categoria"]

            average = safe_float(
                row["media"]
            )

            latest = safe_float(
                row["ultimo_mese"]
            )

            target = safe_float(
                row["obiettivo"]
            )

            possible = safe_float(
                row["risparmio_possibile"]
            )

            variation = safe_float(
                row["variazione_percentuale"]
            )

            if possible <= 0:
                continue

            # ------------------------------------------------
            # Stato
            # ------------------------------------------------

            if variation > 20:

                status = "🔴"

                message = (
                    "Questa categoria è stata "
                    "recentemente molto sopra la tua media."
                )

            elif variation > 5:

                status = "🟠"

                message = (
                    "Questa categoria è leggermente "
                    "sopra la tua media."
                )

            else:

                status = "🟡"

                message = (
                    "Qui esiste un piccolo margine "
                    "di ottimizzazione."
                )

            with st.container():

                st.markdown(
                    f"### {status} {category}"
                )

                c1, c2, c3 = st.columns(3)

                c1.metric(
                    "Media",
                    euro(average)
                )

                c2.metric(
                    "Ultimo mese",
                    euro(latest)
                )

                c3.metric(
                    "Obiettivo",
                    euro(target)
                )

                st.write(
                    message
                )

                st.success(
                    f"💰 Risparmio potenziale: "
                    f"**{euro(possible)} al mese** "
                    f"(**{euro(possible * 12)} all'anno**)"
                )

                st.progress(
                    min(
                        max(
                            target / average
                            if average > 0
                            else 0,
                            0
                        ),
                        1
                    )
                )

                st.markdown("---")

        # ----------------------------------------------------
        # PIANO RIASSUNTIVO
        # ----------------------------------------------------

        st.subheader(
            "📋 Il tuo piano di risparmio"
        )

        summary = plan[
            [
                "categoria",
                "media",
                "ultimo_mese",
                "obiettivo",
                "risparmio_possibile"
            ]
        ].copy()

        summary.rename(
            columns={
                "categoria":
                    "Categoria",

                "media":
                    "Media mensile",

                "ultimo_mese":
                    "Ultimo mese",

                "obiettivo":
                    "Nuovo obiettivo",

                "risparmio_possibile":
                    "Risparmio possibile"
            },
            inplace=True
        )

        summary = summary[
            summary["Risparmio possibile"] > 0
        ]

        st.dataframe(
            summary.style.format(
                {
                    "Media mensile": euro,
                    "Ultimo mese": euro,
                    "Nuovo obiettivo": euro,
                    "Risparmio possibile": euro
                }
            ),
            use_container_width=True,
            hide_index=True
        )

        st.info(
            "💡 Gli obiettivi sono stime automatiche basate "
            "sulla tua spesa storica. Non sono limiti obbligatori: "
            "servono per capire dove esiste un possibile margine."
        )


# ============================================================
# OBIETTIVI
# ============================================================

elif menu == "🎯 Obiettivi":

    st.title(
        "🎯 Obiettivi di risparmio"
    )

    with st.form(
        "goal_form"
    ):

        name = st.text_input(
            "Nome obiettivo",
            placeholder="Es. Vacanza"
        )

        target = st.number_input(
            "Obiettivo (€)",
            min_value=0.0,
            step=500.0
        )

        accumulated = st.number_input(
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

            if not name.strip():

                st.error(
                    "Inserisci il nome dell'obiettivo."
                )

            elif target <= 0:

                st.error(
                    "Inserisci un importo."
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

                    VALUES (?, ?, ?, ?, 1)
                    """,
                    (
                        name,
                        target,
                        accumulated,
                        deadline.isoformat()
                    )
                )

                st.success(
                    "Obiettivo creato."
                )

                st.rerun()

    st.markdown("---")

    goals = query_df(
        """
        SELECT *

        FROM obiettivi

        WHERE attivo=1

        ORDER BY scadenza
        """
    )

    if goals.empty:

        st.info(
            "Nessun obiettivo."
        )

    else:

        for _, goal in goals.iterrows():

            target = safe_float(
                goal["obiettivo"]
            )

            accumulated = safe_float(
                goal["accumulato"]
            )

            progress = (
                accumulated / target
                if target > 0
                else 0
            )

            progress = min(
                max(progress, 0),
                1
            )

            st.subheader(
                f"🎯 {goal['nome']}"
            )

            st.progress(
                progress
            )

            st.write(
                f"**{euro(accumulated)}** "
                f"/ {euro(target)} "
                f"— {progress * 100:.1f}%"
            )

            st.caption(
                f"Scadenza: {goal['scadenza']}"
            )


# ============================================================
# QUANTO POSSO SPENDERE OGGI?
# ============================================================

elif menu == "💸 Quanto posso spendere oggi?":

    st.title(
        "💸 Quanto posso spendere oggi?"
    )

    st.caption(
        "Calcolo della disponibilità residua fino alla fine del mese."
    )

    today = date.today()

    year = today.year
    month = today.month

    income = get_monthly_income(
        year,
        month
    )

    spent = get_monthly_total(
        year,
        month
    )

    recurring = get_recurring_remaining()

    future = get_future_expenses_current_month()

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

    days_remaining = (
        last_day
        - today.day
        + 1
    )

    daily = (
        available / days_remaining
        if days_remaining > 0
        else 0
    )

    # --------------------------------------------------------
    # RISULTATO
    # --------------------------------------------------------

    st.markdown("---")

    if available >= 0:

        st.success(
            f"💚 Puoi spendere ancora "
            f"**{euro(available)}** "
            f"nel resto del mese."
        )

    else:

        st.error(
            f"🔴 La disponibilità prevista è negativa di "
            f"**{euro(abs(available))}**."
        )

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "💰 Disponibilità",
        euro(available)
    )

    c2.metric(
        "📅 Giorni rimanenti",
        days_remaining
    )

    c3.metric(
        "💸 Budget giornaliero",
        euro(
            max(
                0,
                daily
            )
        )
    )

    st.markdown("---")

    st.subheader(
        "🔎 Come viene calcolato"
    )

    calculation = pd.DataFrame(
        {
            "Voce": [
                "Entrate ricevute",
                "Spese già sostenute",
                "Spese ricorrenti ancora previste",
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
                "Importo": euro
            }
        ),
        use_container_width=True,
        hide_index=True
    )

    st.markdown("---")

    st.subheader(
        "🧮 Simula una spesa"
    )

    simulation = st.number_input(
        "Se oggi spendessi...",
        min_value=0.0,
        step=10.0
    )

    simulated_available = (
        available
        - simulation
    )

    if simulated_available >= 0:

        st.success(
            f"Dopo questa spesa avresti "
            f"ancora **{euro(simulated_available)}**."
        )

    else:

        st.warning(
            f"Dopo questa spesa andresti sotto "
            f"di **{euro(abs(simulated_available))}**."
        )


# ============================================================
# ANALISI
# ============================================================

elif menu == "📊 Analisi":

    st.title(
        "📊 Analisi"
    )

    selected = st.date_input(
        "Mese",
        first_day_of_month(),
        key="analysis_month"
    )

    year = selected.year
    month = selected.month

    income = get_monthly_income(
        year,
        month
    )

    expenses = get_monthly_total(
        year,
        month
    )

    savings = (
        income
        - expenses
    )

    c1, c2, c3 = st.columns(3)

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

    st.markdown("---")

    categories = get_monthly_category_totals(
        year,
        month
    )

    if categories.empty:

        st.info(
            "Nessuna spesa registrata."
        )

    else:

        st.subheader(
            "📊 Spese per categoria"
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
                    "Importo": euro
                }
            ),
            use_container_width=True,
            hide_index=True
        )

        fig, ax = plt.subplots(
            figsize=(8, 6)
        )

        ax.bar(
            categories["categoria"],
            categories["importo"]
        )

        ax.set_ylabel(
            "Euro"
        )

        ax.tick_params(
            axis="x",
            rotation=60
        )

        ax.grid(
            axis="y",
            alpha=0.2
        )

        st.pyplot(
            fig,
            clear_figure=True
        )

    # --------------------------------------------------------
    # STORICO
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "📈 Storico mensile"
    )

    history = query_df(
        """
        SELECT
            mese,
            SUM(importo) AS totale

        FROM spese_mensili

        GROUP BY mese

        ORDER BY mese
        """
    )

    if not history.empty:

        history = history.tail(
            12
        )

        history["label"] = history[
            "mese"
        ].apply(
            lambda x:
            f"{MONTHS[int(x[5:7]) - 1][:3]} {x[:4]}"
        )

        fig, ax = plt.subplots(
            figsize=(10, 4)
        )

        ax.plot(
            history["label"],
            history["totale"],
            marker="o"
        )

        ax.set_ylabel(
            "€"
        )

        ax.tick_params(
            axis="x",
            rotation=45
        )

        ax.grid(
            alpha=0.2
        )

        st.pyplot(
            fig,
            clear_figure=True
        )


# ============================================================
# IMPORT / EXPORT
# ============================================================

elif menu == "📥 Import / Export":

    st.title(
        "📥 Import / Export"
    )

    st.subheader(
        "📘 Importazione Excel"
    )

    if EXCEL_FILE.exists():

        excel_history = read_excel_history()

        st.success(
            "🟢 Spese casa -2.xlsx trovato."
        )

        if not excel_history.empty:

            st.write(
                f"Rilevate **{len(excel_history)} "
                f"combinazioni mese/categoria**."
            )

            st.dataframe(
                excel_history.rename(
                    columns={
                        "mese":
                            "Mese",

                        "categoria":
                            "Categoria",

                        "importo":
                            "Importo"
                    }
                ).style.format(
                    {
                        "Importo": euro
                    }
                ),
                use_container_width=True,
                hide_index=True
            )

        if st.button(
            "🔄 Reimporta Excel"
        ):

            count = import_excel()

            st.success(
                f"Importazione completata: "
                f"{count} record elaborati."
            )

            st.rerun()

    else:

        st.warning(
            "Il file Spese casa -2.xlsx non è stato trovato."
        )

    st.markdown("---")

    st.subheader(
        "⬇️ Backup completo"
    )

    tables = {

        "Movimenti":
            query_df(
                """
                SELECT *
                FROM movimenti
                ORDER BY data
                """
            ),

        "Spese mensili":
            query_df(
                """
                SELECT *
                FROM spese_mensili
                ORDER BY mese, categoria
                """
            ),

        "Ricorrenti":
            query_df(
                """
                SELECT *
                FROM ricorrenti
                """
            ),

        "Spese future":
            query_df(
                """
                SELECT *
                FROM spese_future
                ORDER BY data
                """
            ),

        "Budget":
            query_df(
                """
                SELECT *
                FROM budget_mensile
                ORDER BY mese, categoria
                """
            ),

        "Obiettivi":
            query_df(
                """
                SELECT *
                FROM obiettivi
                """
            ),

        "Categorie":
            query_df(
                """
                SELECT *
                FROM categorie
                """
            ),

        "Persone":
            query_df(
                """
                SELECT *
                FROM persone
                """
            ),

        "Impostazioni":
            query_df(
                """
                SELECT *
                FROM impostazioni
                """
            ),
    }

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        for name, dataframe in tables.items():

            dataframe.to_excel(
                writer,
                sheet_name=name[:31],
                index=False
            )

    st.download_button(
        "⬇️ Scarica backup Excel",
        data=output.getvalue(),
        file_name="backup_finanze_arletti.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


# ============================================================
# IMPOSTAZIONI
# ============================================================

elif menu == "⚙️ Impostazioni":

    st.title(
        "⚙️ Impostazioni"
    )

    st.subheader(
        "🎯 Risparmio"
    )

    saving_target = st.number_input(
        "Risparmio mensile desiderato (€)",
        min_value=0.0,
        value=get_setting(
            "risparmio_mensile_target",
            0
        ),
        step=50.0
    )

    safety = st.number_input(
        "Fondo sicurezza da mantenere (€)",
        min_value=0.0,
        value=get_setting(
            "fondo_sicurezza",
            500
        ),
        step=50.0
    )

    if st.button(
        "💾 Salva impostazioni"
    ):

        set_setting(
            "risparmio_mensile_target",
            saving_target
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
        "🏷️ Le tue categorie"
    )

    categories_df = pd.DataFrame(
        {
            "N°":
                range(
                    1,
                    len(EXPENSE_CATEGORIES) + 1
                ),

            "Categoria":
                EXPENSE_CATEGORIES
        }
    )

    st.dataframe(
        categories_df,
        use_container_width=True,
        hide_index=True
    )

    st.info(
        "Queste sono le uniche 9 categorie di spesa "
        "utilizzate dall'app."
    )

    st.markdown("---")

    st.subheader(
        "🗄️ Database"

    )

    st.write(
        f"Database locale: `{DB_FILE.name}`"
    )

    st.write(
        f"File Excel: `{EXCEL_FILE.name}`"
    )

    st.markdown("---")

    st.subheader(
        "🗑️ Pulizia dati"
    )

    st.warning(
        "Questa funzione elimina definitivamente i dati selezionati."
    )

    delete_choice = st.selectbox(
        "Seleziona cosa eliminare",
        [
            "Niente",
            "Tutte le spese singole",
            "Tutte le entrate",
            "Tutte le spese mensili",
            "Tutti i budget",
            "Tutti gli obiettivi",
            "Tutte le spese ricorrenti",
            "Tutte le spese future"
        ]
    )

    confirm = st.checkbox(
        "Confermo di voler eliminare i dati"
    )

    if st.button(
        "🗑️ Elimina dati"
    ):

        if delete_choice == "Niente":

            st.info(
                "Nessuna operazione."
            )

        elif not confirm:

            st.error(
                "Devi confermare l'operazione."
            )

        else:

            if delete_choice == "Tutte le spese singole":

                execute(
                    """
                    DELETE FROM movimenti
                    WHERE tipo='uscita'
                    """
                )

            elif delete_choice == "Tutte le entrate":

                execute(
                    """
                    DELETE FROM movimenti
                    WHERE tipo='entrata'
                    """
                )

            elif delete_choice == "Tutte le spese mensili":

                execute(
                    """
                    DELETE FROM spese_mensili
                    """
                )

            elif delete_choice == "Tutti i budget":

                execute(
                    """
                    DELETE FROM budget_mensile
                    """
                )

            elif delete_choice == "Tutti gli obiettivi":

                execute(
                    """
                    DELETE FROM obiettivi
                    """
                )

            elif delete_choice == "Tutte le spese ricorrenti":

                execute(
                    """
                    DELETE FROM ricorrenti
                    """
                )

            elif delete_choice == "Tutte le spese future":

                execute(
                    """
                    DELETE FROM spese_future
                    """
                )

            st.success(
                "Dati eliminati."
            )

            st.rerun()
