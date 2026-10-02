import streamlit as st
import pandas as pd
import sqlite3
import matplotlib.pyplot as plt

from pathlib import Path
from datetime import date
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


# ============================================================
# CATEGORIE SPESE
# ============================================================
#
# ATTENZIONE:
# Queste sono le UNICHE categorie utilizzate dall'app.
#

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
    "Virginia",
]


# ============================================================
# STILE
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
        font-size: 1.45rem;
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
    """
    Formatta un numero come valuta italiana.
    """

    try:
        value = float(value)
    except Exception:
        value = 0

    return (
        f"{value:,.2f} €"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


def num(value):
    """
    Converte valori Excel in numeri.
    """

    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except Exception:
        pass

    if isinstance(value, (int, float)):
        return float(value)

    value = str(value).strip()

    value = value.replace("€", "")
    value = value.replace(" ", "")

    if not value:
        return None

    if "," in value:
        value = value.replace(".", "")
        value = value.replace(",", ".")

    try:
        return float(value)
    except Exception:
        return None


def month_key(year, month):
    return f"{year}-{month:02d}"


def month_label(year, month):
    return f"{MONTHS[month - 1]} {year}"


def section_title(title, subtitle=None):

    st.title(title)

    if subtitle:
        st.caption(subtitle)


# ============================================================
# DATABASE
# ============================================================

def get_connection():

    return sqlite3.connect(DB_FILE)


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
# CREAZIONE DATABASE
# ============================================================

def init_database():

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

            note TEXT,

            UNIQUE(mese, categoria)
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

            importo REAL NOT NULL,

            UNIQUE(mese, categoria)
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

    # --------------------------------------------------------
    # Impostazioni iniziali
    # --------------------------------------------------------

    cursor.execute(
        """
        INSERT OR IGNORE INTO impostazioni
        (chiave, valore)
        VALUES ('risparmio_mensile_target', '0')
        """
    )

    cursor.execute(
        """
        INSERT OR IGNORE INTO impostazioni
        (chiave, valore)
        VALUES ('fondo_sicurezza', '500')
        """
    )

    # --------------------------------------------------------
    # Persone
    # --------------------------------------------------------

    for person in PEOPLE:

        cursor.execute(
            """
            INSERT OR IGNORE INTO persone(nome)
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
            INSERT OR IGNORE INTO categorie(nome, tipo)
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
            INSERT OR IGNORE INTO categorie(nome, tipo)
            VALUES (?, 'entrata')
            """,
            (category,)
        )

    conn.commit()

    conn.close()


init_database()


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

    try:
        return float(
            df.iloc[0]["valore"]
        )

    except Exception:
        return default


def set_setting(key, value):

    execute(
        """
        INSERT INTO impostazioni
        (chiave, valore)

        VALUES (?, ?)

        ON CONFLICT(chiave)
        DO UPDATE SET
            valore=excluded.valore
        """,
        (
            key,
            str(value)
        )
    )


# ============================================================
# NORMALIZZAZIONE CATEGORIE EXCEL
# ============================================================

def normalize_excel_category(category):

    if category is None:
        return None

    original = str(category).strip()

    # Normalizzazione spazi
    normalized = re.sub(
        r"\s+",
        " ",
        original
    ).strip()

    # --------------------------------------------------------
    # Corrispondenze esatte / varianti del file Excel
    # --------------------------------------------------------

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

    return mapping.get(
        normalized,
        None
    )


# ============================================================
# LETTURA EXCEL
# ============================================================

def parse_excel_history():

    if not EXCEL_FILE.exists():

        return pd.DataFrame(
            columns=[
                "anno",
                "mese",
                "mese_key",
                "categoria",
                "importo",
                "fonte"
            ]
        )

    try:

        import openpyxl

        workbook = openpyxl.load_workbook(
            EXCEL_FILE,
            data_only=True
        )

    except Exception as error:

        st.error(
            f"Errore apertura Excel: {error}"
        )

        return pd.DataFrame()

    if "Costi famiglia" not in workbook.sheetnames:

        return pd.DataFrame()

    ws = workbook["Costi famiglia"]

    rows = []

    current_year = None
    current_month = None

    month_map = {
        month.upper(): index
        for index, month in enumerate(
            MONTHS,
            start=1
        )
    }

    # --------------------------------------------------------
    # Scansione foglio
    # --------------------------------------------------------

    for row in range(
        1,
        ws.max_row + 1
    ):

        first_cell = ws.cell(
            row,
            1
        ).value

        if first_cell is None:
            continue

        text = str(
            first_cell
        ).strip()

        upper_text = text.upper()

        # ----------------------------------------------------
        # Ricerca intestazione mese
        #
        # Esempio:
        # Spese SETTEMBRE 2026
        # ----------------------------------------------------

        match = re.search(
            r"SPESE\s+([A-ZÀ-Ù]+)\s+(\d{4})",
            upper_text
        )

        if match:

            month_name = match.group(1)

            year = int(
                match.group(2)
            )

            if month_name in month_map:

                current_month = month_map[
                    month_name
                ]

                current_year = year

            continue

        if (
            current_year is None
            or current_month is None
        ):
            continue

        # ----------------------------------------------------
        # Solo le categorie che abbiamo deciso di mantenere
        # ----------------------------------------------------

        category = normalize_excel_category(
            text
        )

        if category is None:
            continue

        # ----------------------------------------------------
        # Importo nella colonna B
        # ----------------------------------------------------

        amount = num(
            ws.cell(
                row,
                2
            ).value
        )

        if amount is None:
            continue

        rows.append(
            {
                "anno": current_year,

                "mese": current_month,

                "mese_key": month_key(
                    current_year,
                    current_month
                ),

                "categoria": category,

                "importo": amount,

                "fonte": "Excel"
            }
        )

    return pd.DataFrame(
        rows
    )


# ============================================================
# IMPORTAZIONE EXCEL NEL DATABASE
# ============================================================

def import_excel_history():

    df = parse_excel_history()

    if df.empty:
        return 0

    # --------------------------------------------------------
    # Raggruppiamo eventuali righe duplicate
    # --------------------------------------------------------

    df = (
        df
        .groupby(
            [
                "mese_key",
                "categoria"
            ],
            as_index=False
        )["importo"]
        .sum()
    )

    imported = 0

    for _, row in df.iterrows():

        mese = row["mese_key"]
        categoria = row["categoria"]
        importo = float(row["importo"])

        # ----------------------------------------------------
        # Se il record esiste già:
        #
        # - se è Excel → aggiorna
        # - se è manuale → NON sovrascrivere
        #
        # In questo modo eventuali modifiche manuali
        # rimangono.
        # ----------------------------------------------------

        existing = query_df(
            """
            SELECT id, fonte
            FROM spese_mensili
            WHERE mese=?
            AND categoria=?
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
                VALUES (?, ?, ?, 'Excel')
                """,
                (
                    mese,
                    categoria,
                    importo
                )
            )

        else:

            fonte = existing.iloc[0]["fonte"]

            if fonte == "Excel":

                execute(
                    """
                    UPDATE spese_mensili

                    SET importo=?,
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

        imported += 1

    return imported


# ============================================================
# IMPORT AUTOMATICO
# ============================================================

if EXCEL_FILE.exists():

    import_excel_history()


# ============================================================
# CALCOLO SPESE MENSILI
# ============================================================

def monthly_expenses_by_category(
    year,
    month
):

    key = month_key(
        year,
        month
    )

    # --------------------------------------------------------
    # Spese mensili
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
        (key,)
    )

    # --------------------------------------------------------
    # Spese singole
    #
    # ATTENZIONE:
    # Le spese singole vengono aggiunte alle mensili.
    #
    # Quindi la modalità mensile e quella singola possono
    # coesistere.
    # --------------------------------------------------------

    individual = query_df(
        """
        SELECT
            categoria,
            SUM(importo) AS importo

        FROM movimenti

        WHERE tipo='uscita'

        AND substr(data,1,7)=?

        AND pagato=1

        GROUP BY categoria
        """,
        (key,)
    )

    frames = []

    if not monthly.empty:
        frames.append(monthly)

    if not individual.empty:
        frames.append(individual)

    if not frames:

        return pd.DataFrame(
            columns=[
                "categoria",
                "importo"
            ]
        )

    result = pd.concat(
        frames,
        ignore_index=True
    )

    result = (
        result
        .groupby(
            "categoria",
            as_index=False
        )["importo"]
        .sum()
        .sort_values(
            "importo",
            ascending=False
        )
    )

    return result


def monthly_expenses_total(
    year,
    month
):

    df = monthly_expenses_by_category(
        year,
        month
    )

    if df.empty:
        return 0

    return float(
        df["importo"].sum()
    )


def monthly_income_total(
    year,
    month
):

    key = month_key(
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
        (key,)
    )

    return float(
        df.iloc[0]["totale"]
    )


# ============================================================
# RICORRENTI
# ============================================================

def recurring_total():

    df = query_df(
        """
        SELECT
            COALESCE(
                SUM(importo),
                0
            ) AS totale

        FROM ricorrenti

        WHERE attiva=1
        """
    )

    return float(
        df.iloc[0]["totale"]
    )


def recurring_remaining_today():

    today = date.today()

    df = query_df(
        """
        SELECT *

        FROM ricorrenti

        WHERE attiva=1
        """
    )

    if df.empty:
        return 0

    total = 0

    for _, row in df.iterrows():

        try:

            day = int(
                row["giorno"]
            )

        except Exception:

            day = 1

        if day >= today.day:

            total += float(
                row["importo"]
            )

    return total


# ============================================================
# SPESE FUTURE
# ============================================================

def future_expenses_remaining(
    year,
    month
):

    today = date.today()

    start = date(
        year,
        month,
        1
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

    first_date = max(
        today,
        start
    )

    df = query_df(
        """
        SELECT *

        FROM spese_future

        WHERE pagata=0

        AND data>=?

        AND data<=?
        """,
        (
            first_date.isoformat(),
            end.isoformat()
        )
    )

    if df.empty:
        return 0

    return float(
        df["importo"].sum()
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
        "Metti Spese casa -2.xlsx "
        "nella stessa cartella di app.py"
    )


# ============================================================
# HOME
# ============================================================

if menu == "🏠 Home":

    section_title(
        "🏠 Dashboard Finanze",
        "Controllo delle spese familiari e del risparmio"
    )

    selected_month = st.date_input(
        "📅 Mese",
        date.today().replace(
            day=1
        )
    )

    year = selected_month.year
    month = selected_month.month

    income = monthly_income_total(
        year,
        month
    )

    expenses = monthly_expenses_total(
        year,
        month
    )

    savings = income - expenses

    savings_percentage = (
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
        f"{savings_percentage:.1f}%"
    )

    st.markdown("---")

    # --------------------------------------------------------
    # CATEGORIE DEL MESE
    # --------------------------------------------------------

    st.subheader(
        f"📊 Spese di {month_label(year, month)}"
    )

    categories = monthly_expenses_by_category(
        year,
        month
    )

    if categories.empty:

        st.info(
            "Non ci sono ancora spese registrate per questo mese."
        )

    else:

        st.dataframe(
            categories.rename(
                columns={
                    "categoria": "Categoria",
                    "importo": "Importo"
                }
            ).style.format(
                {
                    "Importo": lambda x: euro(x)
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

    # --------------------------------------------------------
    # RISPARMIO
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "🎯 Obiettivo di risparmio"
    )

    saving_target = get_setting(
        "risparmio_mensile_target",
        0
    )

    if saving_target > 0:

        st.metric(
            "Risparmio programmato",
            euro(saving_target)
        )

        difference = savings - saving_target

        if difference >= 0:

            st.success(
                f"Sei sopra l'obiettivo di "
                f"{euro(difference)}."
            )

        else:

            st.warning(
                f"Sei sotto l'obiettivo di "
                f"{euro(abs(difference))}."
            )

    else:

        st.info(
            "Imposta un obiettivo di risparmio "
            "nella sezione Impostazioni."
        )


# ============================================================
# ENTRATE
# ============================================================

elif menu == "💰 Entrate":

    section_title(
        "💰 Entrate",
        "Registra stipendio, bonus, rimborsi e altre entrate"
    )

    with st.form(
        "form_entrata"
    ):

        description = st.text_input(
            "Descrizione",
            placeholder="Es. Stipendio ottobre"
        )

        c1, c2 = st.columns(2)

        with c1:

            category = st.selectbox(
                "Categoria",
                INCOME_CATEGORIES
            )

        with c2:

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
                    "Importo":
                        lambda x: euro(x)
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
        "Inserisci una singola spesa oppure il totale mensile per categoria"
    )

    tab_single, tab_monthly = st.tabs(
        [
            "🧾 Singola spesa",
            "📅 Spese mensili"
        ]
    )


    # ========================================================
    # SINGOLA SPESA
    # ========================================================

    with tab_single:

        st.subheader(
            "🧾 Inserisci una singola spesa"
        )

        st.caption(
            "Esempio: supermercato €45, benzina €60, "
            "ristorante €80."
        )

        with st.form(
            "form_singola_spesa"
        ):

            expense_date = st.date_input(
                "Data",
                date.today()
            )

            category = st.selectbox(
                "Categoria",
                EXPENSE_CATEGORIES,
                key="single_category"
            )

            person = st.selectbox(
                "Persona",
                PEOPLE,
                key="single_person"
            )

            description = st.text_input(
                "Descrizione",
                placeholder="Es. Spesa supermercato"
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
                            'uscita',
                            ?, ?, ?, ?, ?, ?,
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
            "📋 Spese singole registrate"
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
                "Nessuna spesa singola registrata."
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
                        "Importo":
                            lambda x: euro(x)
                    }
                ),
                use_container_width=True,
                hide_index=True
            )


    # ========================================================
    # SPESE MENSILI
    # ========================================================

    with tab_monthly:

        st.subheader(
            "📅 Spese mensili per categoria"
        )

        st.caption(
            "Inserisci direttamente il totale del mese "
            "per ciascuna delle 9 categorie."
        )

        selected = st.date_input(
            "Mese",
            date.today().replace(
                day=1
            ),
            key="monthly_expense_month"
        )

        selected_key = month_key(
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
            (selected_key,)
        )

        existing_dict = {}

        if not existing.empty:

            for _, row in existing.iterrows():

                existing_dict[
                    row["categoria"]
                ] = float(
                    row["importo"]
                )

        if not existing.empty:

            excel_rows = existing[
                existing["fonte"] == "Excel"
            ]

            if not excel_rows.empty:

                st.info(
                    "📘 I valori mostrati provengono dallo storico Excel. "
                    "Se li modifichi e salvi, diventeranno valori manuali "
                    "dell'app."
                )

        with st.form(
            "form_spese_mensili"
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
                        value=float(
                            existing_dict.get(
                                category,
                                0
                            )
                        ),
                        step=10.0,
                        key=(
                            "monthly_"
                            + selected_key
                            + "_"
                            + category
                        )
                    )

            save_month = st.form_submit_button(
                "💾 Salva mese"
            )

            if save_month:

                # ------------------------------------------------
                # Salviamo un SOLO valore per categoria/mese.
                # ------------------------------------------------

                for category, amount in values.items():

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
                            ?, ?, ?, 'manuale'
                        )

                        ON CONFLICT(mese, categoria)

                        DO UPDATE SET
                            importo=excluded.importo,
                            fonte='manuale'
                        """,
                        (
                            selected_key,
                            category,
                            amount
                        )
                    )

                st.success(
                    f"Spese di "
                    f"{month_label(selected.year, selected.month)} "
                    f"salvate."
                )

                st.rerun()

        st.markdown("---")

        monthly_view = query_df(
            """
            SELECT
                categoria AS Categoria,
                importo AS Importo,
                fonte AS Fonte

            FROM spese_mensili

            WHERE mese=?

            ORDER BY importo DESC
            """,
            (selected_key,)
        )

        if monthly_view.empty:

            st.info(
                "Nessuna spesa mensile registrata."
            )

        else:

            total = monthly_view[
                "Importo"
            ].sum()

            st.metric(
                "💳 Totale mese",
                euro(total)
            )

            st.dataframe(
                monthly_view.style.format(
                    {
                        "Importo":
                            lambda x: euro(x)
                    }
                ),
                use_container_width=True,
                hide_index=True
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
            "➕ Aggiungi ricorrente"
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

                    VALUES
                    (?, ?, ?, ?, ?, 1)
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

    df = query_df(
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

        ORDER BY giorno
        """
    )

    if df.empty:

        st.info(
            "Nessuna spesa ricorrente."
        )

    else:

        total = df.loc[
            df["Attiva"] == 1,
            "Importo"
        ].sum()

        st.metric(
            "💳 Totale ricorrenti mensili",
            euro(total)
        )

        df["Attiva"] = df[
            "Attiva"
        ].apply(
            lambda x:
            "Sì" if x else "No"
        )

        st.dataframe(
            df.style.format(
                {
                    "Importo":
                        lambda x: euro(x)
                }
            ),
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# BUDGET
# ============================================================

elif menu == "📅 Budget":

    section_title(
        "📅 Budget mensile",
        "Imposta un limite di spesa per ciascuna delle 9 categorie"
    )

    selected = st.date_input(
        "📅 Mese",
        date.today().replace(
            day=1
        ),
        key="budget_month"
    )

    key = month_key(
        selected.year,
        selected.month
    )

    existing = query_df(
        """
        SELECT
            categoria,
            importo

        FROM budget_mensile

        WHERE mese=?
        """,
        (key,)
    )

    budget_dict = {}

    if not existing.empty:

        for _, row in existing.iterrows():

            budget_dict[
                row["categoria"]
            ] = float(
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
                    value=float(
                        budget_dict.get(
                            category,
                            0
                        )
                    ),
                    step=25.0,
                    key=(
                        "budget_"
                        + key
                        + "_"
                        + category
                    )
                )

        save = st.form_submit_button(
            "💾 Salva budget"
        )

        if save:

            for category, amount in values.items():

                execute(
                    """
                    INSERT INTO budget_mensile
                    (
                        mese,
                        categoria,
                        importo
                    )

                    VALUES (?, ?, ?)

                    ON CONFLICT(mese, categoria)

                    DO UPDATE SET
                        importo=excluded.importo
                    """,
                    (
                        key,
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
        "📊 Situazione reale"
    )

    budget = query_df(
        """
        SELECT
            categoria,
            importo

        FROM budget_mensile

        WHERE mese=?
        """,
        (key,)
    )

    spent = monthly_expenses_by_category(
        selected.year,
        selected.month
    )

    if budget.empty:

        st.info(
            "Inserisci il budget delle categorie."
        )

    else:

        rows = []

        for _, row in budget.iterrows():

            category = row["categoria"]

            budget_value = float(
                row["importo"]
            )

            spent_value = 0

            if not spent.empty:

                match = spent[
                    spent["categoria"]
                    == category
                ]

                if not match.empty:

                    spent_value = float(
                        match.iloc[0]["importo"]
                    )

            remaining = (
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
                        remaining,

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
                        lambda x: euro(x),

                    "Speso":
                        lambda x: euro(x),

                    "Residuo":
                        lambda x: euro(x),

                    "% utilizzata":
                        lambda x:
                        f"{x:.1f}%"
                }
            ),
            use_container_width=True,
            hide_index=True
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
            "Nome obiettivo",
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

            if not name.strip():

                st.error(
                    "Inserisci un nome."
                )

            elif target <= 0:

                st.error(
                    "Inserisci un obiettivo."
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
            "Nessun obiettivo attivo."
        )

    else:

        for _, goal in goals.iterrows():

            target = float(
                goal["obiettivo"]
            )

            saved = float(
                goal["accumulato"]
            )

            progress = (
                saved / target
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
                f"**{euro(saved)}** / "
                f"{euro(target)} "
                f"— {progress * 100:.1f}%"
            )

            st.caption(
                f"Scadenza: {goal['scadenza']}"
            )


# ============================================================
# QUANTO POSSO SPENDERE OGGI?
# ============================================================

elif menu == "💸 Quanto posso spendere oggi?":

    section_title(
        "💸 Quanto posso spendere oggi?",
        "Calcolo della disponibilità residua fino alla fine del mese"
    )

    today = date.today()

    year = today.year
    month = today.month

    key = month_key(
        year,
        month
    )

    # --------------------------------------------------------
    # ENTRATE
    # --------------------------------------------------------

    income = monthly_income_total(
        year,
        month
    )

    # --------------------------------------------------------
    # SPESE GIÀ SOSTENUTE
    # --------------------------------------------------------

    spent = monthly_expenses_total(
        year,
        month
    )

    # --------------------------------------------------------
    # RICORRENTI
    # --------------------------------------------------------

    recurring = recurring_remaining_today()

    # --------------------------------------------------------
    # SPESE FUTURE
    # --------------------------------------------------------

    future = future_expenses_remaining(
        year,
        month
    )

    # --------------------------------------------------------
    # RISPARMIO
    # --------------------------------------------------------

    saving_target = get_setting(
        "risparmio_mensile_target",
        0
    )

    # --------------------------------------------------------
    # FONDO SICUREZZA
    # --------------------------------------------------------

    safety_fund = get_setting(
        "fondo_sicurezza",
        500
    )

    # --------------------------------------------------------
    # CALCOLO
    # --------------------------------------------------------

    available = (
        income
        - spent
        - recurring
        - future
        - saving_target
        - safety_fund
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

    daily_budget = (
        available
        / days_remaining
        if days_remaining > 0
        else 0
    )

    # --------------------------------------------------------
    # RISULTATO
    # --------------------------------------------------------

    st.markdown("---")

    if available >= 0:

        st.success(
            f"💚 Puoi ancora spendere "
            f"**{euro(available)}** "
            f"nel mese."
        )

    else:

        st.error(
            f"🔴 Hai superato la disponibilità "
            f"programmata di "
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
                daily_budget
            )
        )
    )

    st.markdown("---")

    st.subheader(
        "🔎 Dettaglio del calcolo"
    )

    calculation = pd.DataFrame(
        {
            "Voce": [
                "Entrate del mese",
                "Spese già sostenute",
                "Spese ricorrenti ancora previste",
                "Spese future",
                "Risparmio programmato",
                "Fondo sicurezza",
                "Disponibilità residua",
            ],

            "Importo": [
                income,
                -spent,
                -recurring,
                -future,
                -saving_target,
                -safety_fund,
                available,
            ]
        }
    )

    st.dataframe(
        calculation.style.format(
            {
                "Importo":
                    lambda x: euro(x)
            }
        ),
        use_container_width=True,
        hide_index=True
    )

    st.markdown("---")

    st.subheader(
        "🧮 Simula una spesa"
    )

    simulated = st.number_input(
        "Se oggi spendessi...",
        min_value=0.0,
        step=10.0
    )

    after = (
        available
        - simulated
    )

    if after >= 0:

        st.success(
            f"Dopo questa spesa avresti "
            f"ancora **{euro(after)}**."
        )

    else:

        st.warning(
            f"Dopo questa spesa saresti sotto "
            f"di **{euro(abs(after))}**."
        )


# ============================================================
# ANALISI
# ============================================================

elif menu == "📊 Analisi":

    section_title(
        "📊 Analisi",
        "Analisi delle tue 9 categorie di spesa"
    )

    selected = st.date_input(
        "Mese da analizzare",
        date.today().replace(
            day=1
        ),
        key="analysis_month"
    )

    year = selected.year
    month = selected.month

    expenses = monthly_expenses_by_category(
        year,
        month
    )

    income = monthly_income_total(
        year,
        month
    )

    total_expenses = (
        expenses["importo"].sum()
        if not expenses.empty
        else 0
    )

    savings = (
        income
        - total_expenses
    )

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "💰 Entrate",
        euro(income)
    )

    c2.metric(
        "💳 Spese",
        euro(total_expenses)
    )

    c3.metric(
        "💚 Risparmio",
        euro(savings)
    )

    if not expenses.empty:

        st.markdown("---")

        st.subheader(
            "🥧 Distribuzione delle spese"
        )

        fig, ax = plt.subplots(
            figsize=(8, 6)
        )

        ax.pie(
            expenses["importo"],
            labels=expenses["categoria"],
            autopct="%1.1f%%",
            startangle=90
        )

        ax.axis("equal")

        st.pyplot(
            fig,
            clear_figure=True
        )

        st.subheader(
            "📋 Dettaglio"
        )

        st.dataframe(
            expenses.rename(
                columns={
                    "categoria":
                        "Categoria",

                    "importo":
                        "Importo"
                }
            ).style.format(
                {
                    "Importo":
                        lambda x: euro(x)
                }
            ),
            use_container_width=True,
            hide_index=True
        )

        st.markdown("---")

        st.subheader(
            "📊 Confronto categorie"
        )

        fig, ax = plt.subplots(
            figsize=(10, 5)
        )

        ax.bar(
            expenses["categoria"],
            expenses["importo"]
        )

        ax.set_ylabel(
            "€"
        )

        ax.tick_params(
            axis="x",
            rotation=45
        )

        ax.grid(
            axis="y",
            alpha=0.2
        )

        st.pyplot(
            fig,
            clear_figure=True
        )

    else:

        st.info(
            "Nessuna spesa per questo mese."
        )

    # --------------------------------------------------------
    # STORICO MENSILE
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "📈 Andamento storico"
    )

    history = query_df(
        """
        SELECT
            mese,
            SUM(importo) AS importo

        FROM spese_mensili

        GROUP BY mese

        ORDER BY mese
        """
    )

    if not history.empty:

        history = history.tail(
            12
        )

        history["Label"] = history[
            "mese"
        ].apply(
            lambda x:
            f"{MONTHS[int(x[5:7]) - 1][:3]} "
            f"{x[:4]}"
        )

        fig, ax = plt.subplots(
            figsize=(10, 4)
        )

        ax.plot(
            history["Label"],
            history["importo"],
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

    section_title(
        "📥 Import / Export",
        "Backup completo dei dati"
    )

    # --------------------------------------------------------
    # STATO EXCEL
    # --------------------------------------------------------

    if EXCEL_FILE.exists():

        st.success(
            "🟢 Spese casa -2.xlsx trovato."
        )

        historical = parse_excel_history()

        if not historical.empty:

            st.metric(
                "Righe storiche Excel",
                len(historical)
            )

            st.caption(
                "Sono considerate esclusivamente le 9 categorie "
                "definite nell'app."
            )

    else:

        st.warning(
            "Il file Spese casa -2.xlsx non è presente."
        )

    # --------------------------------------------------------
    # RIEPILOGO DATABASE
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "📊 Stato database"
    )

    stats = {

        "Spese singole":
            query_df(
                """
                SELECT COUNT(*) n
                FROM movimenti
                WHERE tipo='uscita'
                """
            ).iloc[0]["n"],

        "Entrate":
            query_df(
                """
                SELECT COUNT(*) n
                FROM movimenti
                WHERE tipo='entrata'
                """
            ).iloc[0]["n"],

        "Registrazioni mensili":
            query_df(
                """
                SELECT COUNT(*) n
                FROM spese_mensili
                """
            ).iloc[0]["n"],

        "Ricorrenti":
            query_df(
                """
                SELECT COUNT(*) n
                FROM ricorrenti
                """
            ).iloc[0]["n"],

        "Budget":
            query_df(
                """
                SELECT COUNT(*) n
                FROM budget_mensile
                """
            ).iloc[0]["n"],

        "Obiettivi":
            query_df(
                """
                SELECT COUNT(*) n
                FROM obiettivi
                """
            ).iloc[0]["n"],
    }

    stats_df = pd.DataFrame(
        list(stats.items()),
        columns=[
            "Voce",
            "Numero"
        ]
    )

    st.dataframe(
        stats_df,
        use_container_width=True,
        hide_index=True
    )

    # --------------------------------------------------------
    # EXPORT EXCEL
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "⬇️ Backup Excel"
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

        for sheet_name, dataframe in tables.items():

            dataframe.to_excel(
                writer,
                sheet_name=sheet_name[:31],
                index=False
            )

        # Aggiunge anche lo storico Excel
        if EXCEL_FILE.exists():

            historical = parse_excel_history()

            if not historical.empty:

                historical.to_excel(
                    writer,
                    sheet_name="Storico Excel",
                    index=False
                )

    st.download_button(
        label="⬇️ Scarica backup completo",
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
        "Parametri utilizzati nei calcoli dell'app"
    )

    # --------------------------------------------------------
    # RISPARMIO
    # --------------------------------------------------------

    st.subheader(
        "🎯 Risparmio"
    )

    saving_target = st.number_input(
        "Risparmio mensile programmato (€)",
        min_value=0.0,
        value=float(
            get_setting(
                "risparmio_mensile_target",
                0
            )
        ),
        step=50.0
    )

    safety_fund = st.number_input(
        "Fondo sicurezza (€)",
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
            saving_target
        )

        set_setting(
            "fondo_sicurezza",
            safety_fund
        )

        st.success(
            "Impostazioni salvate."
        )

    # --------------------------------------------------------
    # CATEGORIE
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "🏷️ Categorie utilizzate"
    )

    categories_df = pd.DataFrame(
        {
            "N°": range(
                1,
                len(
                    EXPENSE_CATEGORIES
                ) + 1
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
        "Queste sono le uniche categorie di spesa utilizzate "
        "dall'app."
    )

    # --------------------------------------------------------
    # GESTIONE DATABASE
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "🗑️ Gestione dati"
    )

    st.warning(
        "Le operazioni seguenti cancellano dati dal database."
    )

    delete_type = st.selectbox(
        "Cosa vuoi eliminare?",
        [
            "Nessuna operazione",
            "Tutte le spese singole",
            "Tutte le entrate",
            "Tutte le spese mensili manuali",
            "Tutti i budget",
            "Tutti gli obiettivi",
            "Tutte le spese ricorrenti",
            "Tutte le spese future",
        ]
    )

    confirm = st.checkbox(
        "Confermo di voler eliminare i dati selezionati"
    )

    if st.button(
        "🗑️ Elimina"
    ):

        if delete_type == "Nessuna operazione":

            st.info(
                "Nessuna operazione selezionata."
            )

        elif not confirm:

            st.error(
                "Conferma prima di procedere."
            )

        else:

            if delete_type == "Tutte le spese singole":

                execute(
                    """
                    DELETE FROM movimenti
                    WHERE tipo='uscita'
                    """
                )

            elif delete_type == "Tutte le entrate":

                execute(
                    """
                    DELETE FROM movimenti
                    WHERE tipo='entrata'
                    """
                )

            elif delete_type == "Tutte le spese mensili manuali":

                execute(
                    """
                    DELETE FROM spese_mensili
                    WHERE fonte='manuale'
                    """
                )

            elif delete_type == "Tutti i budget":

                execute(
                    """
                    DELETE FROM budget_mensile
                    """
                )

            elif delete_type == "Tutti gli obiettivi":

                execute(
                    """
                    DELETE FROM obiettivi
                    """
                )

            elif delete_type == "Tutte le spese ricorrenti":

                execute(
                    """
                    DELETE FROM ricorrenti
                    """
                )

            elif delete_type == "Tutte le spese future":

                execute(
                    """
                    DELETE FROM spese_future
                    """
                )

            st.success(
                "Operazione completata."
            )

            st.rerun()
