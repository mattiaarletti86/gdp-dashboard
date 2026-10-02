import streamlit as st
import pandas as pd
import sqlite3
import matplotlib.pyplot as plt
from pathlib import Path
from datetime import date, datetime, timedelta
from io import BytesIO
import calendar
import re


# ============================================================
# CONFIGURAZIONE
# ============================================================

st.set_page_config(
    page_title="Gestione Finanze - Arletti",
    page_icon="💰",
    layout="centered"
)

BASE_DIR = Path(__file__).resolve().parent
DB_FILE = BASE_DIR / "finanze_famiglia.db"
EXCEL_FILE = BASE_DIR / "Spese casa -2.xlsx"

MONTHS = [
    "Gennaio", "Febbraio", "Marzo", "Aprile",
    "Maggio", "Giugno", "Luglio", "Agosto",
    "Settembre", "Ottobre", "Novembre", "Dicembre"
]

EXPENSE_CATEGORIES = [
    "Alimentari",
    "Casa",
    "Utenze",
    "Auto e trasporti",
    "Figli",
    "Scuola e sport",
    "Shopping",
    "Salute e cura persona",
    "Vacanze e viaggi",
    "Ristoranti",
    "Tempo libero",
    "Arredi",
    "Prelievi contanti",
    "Altro"
]

INCOME_CATEGORIES = [
    "Stipendio",
    "Bonus",
    "Rimborso",
    "Affitto",
    "Altre entrate"
]

PEOPLE = [
    "Io",
    "Partner",
    "Famiglia"
]


# ============================================================
# STILE
# ============================================================

st.markdown("""
<style>

.block-container {
    padding-top: 1rem;
    padding-bottom: 3rem;
    max-width: 1100px;
}

h1 {
    margin-bottom: 0.2rem;
}

[data-testid="stMetricValue"] {
    font-size: 1.45rem;
}

div[data-testid="stForm"] {
    border-radius: 12px;
}

.small-note {
    font-size: 0.85rem;
    color: #777;
}

.big-number {
    font-size: 2rem;
    font-weight: 700;
}

</style>
""", unsafe_allow_html=True)


# ============================================================
# FUNZIONI GENERALI
# ============================================================

def euro(value):
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
    if value is None:
        return None

    if isinstance(value, float) and pd.isna(value):
        return None

    if isinstance(value, (int, float)):
        return float(value)

    s = str(value).strip()
    s = s.replace("€", "").replace(" ", "")

    if not s:
        return None

    if "," in s:
        s = s.replace(".", "").replace(",", ".")

    try:
        return float(s)
    except Exception:
        return None


def month_label(year, month):
    return f"{MONTHS[month - 1]} {year}"


def month_key(year, month):
    return f"{year}-{month:02d}"


def section_title(title, subtitle=None):
    st.title(title)

    if subtitle:
        st.caption(subtitle)


# ============================================================
# DATABASE
# ============================================================

def get_connection():
    return sqlite3.connect(DB_FILE)


def execute(sql, params=()):
    conn = get_connection()

    try:
        cursor = conn.cursor()
        cursor.execute(sql, params)
        conn.commit()
        return cursor.lastrowid

    finally:
        conn.close()


def query_df(sql, params=()):
    conn = get_connection()

    try:
        return pd.read_sql_query(sql, conn, params=params)

    finally:
        conn.close()


def init_database():

    conn = get_connection()
    cursor = conn.cursor()

    cursor.executescript("""
    
    CREATE TABLE IF NOT EXISTS movimenti (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tipo TEXT NOT NULL,
        data TEXT NOT NULL,
        categoria TEXT NOT NULL,
        sottocategoria TEXT,
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
        source_key TEXT UNIQUE
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

    """)

    # Impostazioni iniziali
    cursor.execute("""
        INSERT OR IGNORE INTO impostazioni(chiave, valore)
        VALUES ('risparmio_mensile_target', '0')
    """)

    cursor.execute("""
        INSERT OR IGNORE INTO impostazioni(chiave, valore)
        VALUES ('fondo_sicurezza', '500')
    """)

    # Persone
    for person in PEOPLE:
        cursor.execute(
            "INSERT OR IGNORE INTO persone(nome) VALUES (?)",
            (person,)
        )

    # Categorie spese
    for category in EXPENSE_CATEGORIES:
        cursor.execute(
            "INSERT OR IGNORE INTO categorie(nome, tipo) VALUES (?, 'uscita')",
            (category,)
        )

    # Categorie entrate
    for category in INCOME_CATEGORIES:
        cursor.execute(
            "INSERT OR IGNORE INTO categorie(nome, tipo) VALUES (?, 'entrata')",
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
    execute("""
        INSERT INTO impostazioni(chiave, valore)
        VALUES (?, ?)
        ON CONFLICT(chiave)
        DO UPDATE SET valore=excluded.valore
    """, (key, str(value)))


# ============================================================
# EXCEL - STORICO
# ============================================================

def parse_excel_family():

    if not EXCEL_FILE.exists():
        return pd.DataFrame()

    try:
        import openpyxl

        workbook = openpyxl.load_workbook(
            EXCEL_FILE,
            data_only=True
        )

        if "Costi famiglia" not in workbook.sheetnames:
            return pd.DataFrame()

        ws = workbook["Costi famiglia"]

        rows = []

        current_year = None
        current_month = None

        month_map = {
            month.upper(): index
            for index, month in enumerate(MONTHS, 1)
        }

        for row in range(1, ws.max_row + 1):

            first_cell = ws.cell(row, 1).value

            if first_cell is None:
                continue

            text = str(first_cell).strip()
            upper_text = text.upper()

            # Cerca intestazioni tipo:
            # SPESE SETTEMBRE 2025
            match = re.match(
                r"SPESE\s+([A-ZÀ-Ù]+)\s+(\d{4})",
                upper_text
            )

            if match:

                month_name = match.group(1)
                year = int(match.group(2))

                if month_name in month_map:
                    current_month = month_map[month_name]
                    current_year = year

                continue

            if current_year is None or current_month is None:
                continue

            amount = num(ws.cell(row, 2).value)

            if amount is None:
                continue

            if amount == 0:
                continue

            if upper_text in ["MEDIA", "TOTALE"]:
                continue

            rows.append({
                "anno": current_year,
                "mese": current_month,
                "mese_key": month_key(
                    current_year,
                    current_month
                ),
                "categoria": text,
                "importo": amount,
                "fonte": "Excel"
            })

        return pd.DataFrame(rows)

    except Exception as error:

        st.warning(
            f"Errore nella lettura dell'Excel: {error}"
        )

        return pd.DataFrame()


def import_excel_history():

    df = parse_excel_family()

    if df.empty:
        return 0

    imported = 0

    for _, row in df.iterrows():

        source_key = (
            f"excel|{row['mese_key']}|"
            f"{row['categoria']}"
        )

        execute("""
            INSERT INTO spese_mensili
            (
                mese,
                categoria,
                importo,
                fonte,
                source_key
            )
            VALUES (?, ?, ?, 'Excel', ?)

            ON CONFLICT(source_key)
            DO UPDATE SET
                importo=excluded.importo,
                categoria=excluded.categoria,
                mese=excluded.mese
        """, (
            row["mese_key"],
            row["categoria"],
            float(row["importo"]),
            source_key
        ))

        imported += 1

    return imported


# Import automatico dello storico
if EXCEL_FILE.exists():
    import_excel_history()


# ============================================================
# CALCOLI SPESE
# ============================================================

def monthly_expenses_total(year, month):
    key = month_key(year, month)

    monthly = query_df("""
        SELECT COALESCE(SUM(importo),0) AS totale
        FROM spese_mensili
        WHERE mese=?
    """, (key,))

    individual = query_df("""
        SELECT COALESCE(SUM(importo),0) AS totale
        FROM movimenti
        WHERE tipo='uscita'
        AND substr(data,1,7)=?
        AND pagato=1
    """, (key,))

    monthly_total = float(monthly.iloc[0]["totale"])
    individual_total = float(individual.iloc[0]["totale"])

    return monthly_total + individual_total


def monthly_income_total(year, month):

    key = month_key(year, month)

    result = query_df("""
        SELECT COALESCE(SUM(importo),0) AS totale
        FROM movimenti
        WHERE tipo='entrata'
        AND substr(data,1,7)=?
        AND pagato=1
    """, (key,))

    return float(result.iloc[0]["totale"])


def monthly_expenses_by_category(year, month):

    key = month_key(year, month)

    monthly = query_df("""
        SELECT
            categoria,
            SUM(importo) AS importo
        FROM spese_mensili
        WHERE mese=?
        GROUP BY categoria
    """, (key,))

    individual = query_df("""
        SELECT
            categoria,
            SUM(importo) AS importo
        FROM movimenti
        WHERE tipo='uscita'
        AND substr(data,1,7)=?
        AND pagato=1
        GROUP BY categoria
    """, (key,))

    frames = []

    if not monthly.empty:
        frames.append(monthly)

    if not individual.empty:
        frames.append(individual)

    if not frames:
        return pd.DataFrame(
            columns=["categoria", "importo"]
        )

    result = pd.concat(frames)

    result = (
        result
        .groupby("categoria", as_index=False)["importo"]
        .sum()
        .sort_values("importo", ascending=False)
    )

    return result


# ============================================================
# RICORRENTI
# ============================================================

def recurring_monthly_total():

    df = query_df("""
        SELECT COALESCE(SUM(importo),0) AS totale
        FROM ricorrenti
        WHERE attiva=1
    """)

    return float(df.iloc[0]["totale"])


def recurring_remaining_today():

    today = date.today()

    df = query_df("""
        SELECT *
        FROM ricorrenti
        WHERE attiva=1
    """)

    if df.empty:
        return 0

    total = 0

    for _, row in df.iterrows():

        try:
            day = int(row["giorno"])
        except Exception:
            day = 1

        # Se il pagamento ricorrente è previsto oggi o nei prossimi giorni
        # viene considerato ancora da sostenere.
        if day >= today.day:
            total += float(row["importo"])

    return total


# ============================================================
# SPESE FUTURE
# ============================================================

def future_expenses_remaining(year, month):

    today = date.today()

    start = date(year, month, 1)

    last_day = calendar.monthrange(year, month)[1]
    end = date(year, month, last_day)

    df = query_df("""
        SELECT *
        FROM spese_future
        WHERE pagata=0
        AND data>=?
        AND data<=?
    """, (
        max(today, start).isoformat(),
        end.isoformat()
    ))

    if df.empty:
        return 0

    return float(df["importo"].sum())


# ============================================================
# BUDGET
# ============================================================

def get_budget(month_key_value):

    return query_df("""
        SELECT categoria, importo
        FROM budget_mensile
        WHERE mese=?
        ORDER BY categoria
    """, (month_key_value,))


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.title("💰 Gestione Finanze")

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
        "⚙️ Impostazioni",
        "📥 Import / Export"
    ]
)

st.sidebar.markdown("---")

if EXCEL_FILE.exists():
    st.sidebar.success("🟢 Excel collegato")
    st.sidebar.caption("Spese casa -2.xlsx")
else:
    st.sidebar.warning(
        "🟡 Excel non trovato"
    )
    st.sidebar.caption(
        "Lo storico Excel non sarà disponibile."
    )


# ============================================================
# HOME
# ============================================================

if menu == "🏠 Home":

    section_title(
        "🏠 Dashboard Finanze",
        "Controllo semplice delle entrate, delle spese e del risparmio familiare"
    )

    selected_month = st.date_input(
        "📅 Mese",
        date.today().replace(day=1)
    )

    year = selected_month.year
    month = selected_month.month
    key = month_key(year, month)

    income = monthly_income_total(year, month)
    expenses = monthly_expenses_total(year, month)
    savings = income - expenses

    savings_pct = (
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
        f"{savings_pct:.1f}%"
    )

    st.markdown("---")

    # --------------------------------------------------------
    # SITUAZIONE BUDGET
    # --------------------------------------------------------

    st.subheader("📅 Situazione del mese")

    budget_df = get_budget(key)

    if budget_df.empty:

        st.info(
            "Non hai ancora impostato un budget per questo mese."
        )

    else:

        spent_df = monthly_expenses_by_category(
            year,
            month
        )

        rows = []

        for _, budget_row in budget_df.iterrows():

            category = budget_row["categoria"]
            budget = float(budget_row["importo"])

            spent = 0

            if not spent_df.empty:

                match = spent_df[
                    spent_df["categoria"] == category
                ]

                if not match.empty:
                    spent = float(
                        match.iloc[0]["importo"]
                    )

            rows.append({
                "Categoria": category,
                "Budget": budget,
                "Speso": spent,
                "Residuo": budget - spent
            })

        budget_view = pd.DataFrame(rows)

        st.dataframe(
            budget_view.style.format({
                "Budget": lambda x: euro(x),
                "Speso": lambda x: euro(x),
                "Residuo": lambda x: euro(x)
            }),
            use_container_width=True,
            hide_index=True
        )

    # --------------------------------------------------------
    # ULTIMI MESI
    # --------------------------------------------------------

    st.markdown("---")
    st.subheader("📈 Andamento spese")

    history = query_df("""
        SELECT
            mese,
            SUM(importo) AS importo
        FROM spese_mensili
        GROUP BY mese
        ORDER BY mese
    """)

    individual_history = query_df("""
        SELECT
            substr(data,1,7) AS mese,
            SUM(importo) AS importo
        FROM movimenti
        WHERE tipo='uscita'
        AND pagato=1
        GROUP BY substr(data,1,7)
        ORDER BY mese
    """)

    frames = []

    if not history.empty:
        frames.append(history)

    if not individual_history.empty:
        frames.append(individual_history)

    if frames:

        chart_df = pd.concat(frames)

        chart_df = (
            chart_df
            .groupby("mese", as_index=False)["importo"]
            .sum()
            .sort_values("mese")
            .tail(12)
        )

        chart_df["label"] = chart_df["mese"].apply(
            lambda x: MONTHS[int(x[5:7]) - 1][:3]
            + " "
            + x[:4]
        )

        fig, ax = plt.subplots(
            figsize=(10, 4)
        )

        ax.plot(
            chart_df["label"],
            chart_df["importo"],
            marker="o"
        )

        ax.set_ylabel("€")
        ax.tick_params(
            axis="x",
            rotation=45
        )

        ax.grid(alpha=0.2)

        st.pyplot(
            fig,
            clear_figure=True
        )


# ============================================================
# ENTRATE
# ============================================================

elif menu == "💰 Entrate":

    section_title(
        "💰 Entrate",
        "Inserisci stipendio, bonus, rimborsi e altre entrate"
    )

    with st.form("nuova_entrata"):

        c1, c2 = st.columns(2)

        with c1:
            description = st.text_input(
                "Descrizione",
                placeholder="Es. Stipendio ottobre"
            )

        with c2:
            category = st.selectbox(
                "Categoria",
                INCOME_CATEGORIES
            )

        c3, c4 = st.columns(2)

        with c3:
            person = st.selectbox(
                "Persona",
                PEOPLE
            )

        with c4:
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

                execute("""
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
                        ?, ?, ?, ?, ?,  ?, 'manuale', ?
                    )
                """, (
                    transaction_date.isoformat(),
                    category,
                    description,
                    person,
                    amount,
                    int(paid),
                    note
                ))

                st.success(
                    "Entrata salvata."
                )

                st.rerun()

    st.markdown("---")
    st.subheader("📋 Entrate inserite")

    df = query_df("""
        SELECT
            id AS ID,
            data AS Data,
            categoria AS Categoria,
            descrizione AS Descrizione,
            persona AS Persona,
            importo AS Importo,
            note AS Note
        FROM movimenti
        WHERE tipo='entrata'
        ORDER BY data DESC, id DESC
    """)

    if df.empty:

        st.info(
            "Nessuna entrata inserita."
        )

    else:

        st.dataframe(
            df.style.format({
                "Importo": lambda x: euro(x)
            }),
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

    tabs = st.tabs([
        "🧾 Singola spesa",
        "📅 Spese mensili"
    ])

    # ========================================================
    # SINGOLA SPESA
    # ========================================================

    with tabs[0]:

        st.subheader(
            "🧾 Inserisci una singola spesa"
        )

        st.caption(
            "Esempio: supermercato €42,50, benzina €60, ristorante €75."
        )

        with st.form("singola_spesa"):

            c1, c2 = st.columns(2)

            with c1:

                expense_date = st.date_input(
                    "Data",
                    date.today()
                )

                category = st.selectbox(
                    "Categoria",
                    EXPENSE_CATEGORIES
                )

            with c2:

                person = st.selectbox(
                    "Persona",
                    PEOPLE
                )

                amount = st.number_input(
                    "Importo (€)",
                    min_value=0.0,
                    step=5.0
                )

            description = st.text_input(
                "Descrizione",
                placeholder="Es. Spesa supermercato"
            )

            note = st.text_input(
                "Note"
            )

            paid = st.checkbox(
                "Spesa già pagata",
                value=True
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

                    execute("""
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
                            ?, ?, ?, ?, ?, ?, 'manuale', ?
                        )
                    """, (
                        expense_date.isoformat(),
                        category,
                        description,
                        person,
                        amount,
                        int(paid),
                        note
                    ))

                    st.success(
                        "Spesa salvata."
                    )

                    st.rerun()

        st.markdown("---")

        st.subheader(
            "📋 Spese singole inserite"
        )

        df = query_df("""
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
        """)

        if df.empty:

            st.info(
                "Nessuna spesa singola inserita."
            )

        else:

            df["Pagata"] = df["Pagata"].map(
                lambda x: "Sì" if x else "No"
            )

            st.dataframe(
                df.style.format({
                    "Importo": lambda x: euro(x)
                }),
                use_container_width=True,
                hide_index=True
            )

            st.caption(
                "Per cancellare una spesa usa la sezione Impostazioni → Gestione dati."
            )

    # ========================================================
    # SPESE MENSILI
    # ========================================================

    with tabs[1]:

        st.subheader(
            "📅 Inserisci le spese del mese per categoria"
        )

        st.caption(
            "Questa modalità è ideale quando conosci già il totale mensile "
            "senza voler inserire ogni singola ricevuta."
        )

        selected = st.date_input(
            "Mese",
            date.today().replace(day=1),
            key="monthly_expense_date"
        )

        selected_key = month_key(
            selected.year,
            selected.month
        )

        existing = query_df("""
            SELECT
                categoria,
                importo,
                fonte
            FROM spese_mensili
            WHERE mese=?
        """, (selected_key,))

        existing_dict = {}

        if not existing.empty:

            for _, row in existing.iterrows():

                category = row["categoria"]

                existing_dict[category] = (
                    existing_dict.get(category, 0)
                    + float(row["importo"])
                )

        st.info(
            "⚠️ Importante: se inserisci il totale mensile di una categoria "
            "e anche singole spese della stessa categoria nello stesso mese, "
            "l'app considererà entrambi gli importi."
        )

        with st.form("spese_mensili"):

            values = {}

            columns = st.columns(2)

            for index, category in enumerate(
                EXPENSE_CATEGORIES
            ):

                with columns[index % 2]:

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
                        key=f"monthly_{selected_key}_{category}"
                    )

            save_month = st.form_submit_button(
                "💾 Salva mese"
            )

            if save_month:

                # Elimina i valori manuali del mese.
                # Quelli Excel vengono preservati.
                execute("""
                    DELETE FROM spese_mensili
                    WHERE mese=?
                    AND fonte='manuale'
                """, (selected_key,))

                for category, amount in values.items():

                    if amount > 0:

                        execute("""
                            INSERT INTO spese_mensili
                            (
                                mese,
                                categoria,
                                importo,
                                fonte,
                                source_key
                            )
                            VALUES (?, ?, ?, 'manuale', ?)
                        """, (
                            selected_key,
                            category,
                            amount,
                            f"manuale|{selected_key}|{category}"
                        ))

                st.success(
                    f"Spese di {month_label(selected.year, selected.month)} salvate."
                )

                st.rerun()

        st.markdown("---")

        monthly_view = query_df("""
            SELECT
                categoria AS Categoria,
                SUM(importo) AS Importo,
                GROUP_CONCAT(DISTINCT fonte) AS Fonte
            FROM spese_mensili
            WHERE mese=?
            GROUP BY categoria
            ORDER BY Importo DESC
        """, (selected_key,))

        if not monthly_view.empty:

            total = monthly_view["Importo"].sum()

            st.metric(
                "Totale spese mensili",
                euro(total)
            )

            st.dataframe(
                monthly_view.style.format({
                    "Importo": lambda x: euro(x)
                }),
                use_container_width=True,
                hide_index=True
            )

        else:

            st.info(
                "Nessuna spesa mensile registrata per questo mese."
            )


# ============================================================
# RICORRENTI
# ============================================================

elif menu == "🔁 Spese ricorrenti":

    section_title(
        "🔁 Spese ricorrenti",
        "Costi che si ripetono ogni mese"
    )

    with st.form("nuova_ricorrente"):

        description = st.text_input(
            "Descrizione",
            placeholder="Es. Mutuo"
        )

        c1, c2 = st.columns(2)

        with c1:

            category = st.selectbox(
                "Categoria",
                EXPENSE_CATEGORIES
            )

            amount = st.number_input(
                "Importo mensile (€)",
                min_value=0.0,
                step=10.0
            )

        with c2:

            day = st.number_input(
                "Giorno previsto",
                min_value=1,
                max_value=31,
                value=1
            )

            person = st.selectbox(
                "Persona",
                PEOPLE
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

                execute("""
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
                """, (
                    description,
                    category,
                    amount,
                    day,
                    person
                ))

                st.success(
                    "Spesa ricorrente aggiunta."
                )

                st.rerun()

    st.markdown("---")

    df = query_df("""
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
    """)

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

        df["Attiva"] = df["Attiva"].map(
            lambda x: "Sì" if x else "No"
        )

        st.dataframe(
            df.style.format({
                "Importo": lambda x: euro(x)
            }),
            use_container_width=True,
            hide_index=True
        )

        st.caption(
            "Le spese ricorrenti vengono considerate nel calcolo "
            "di 'Quanto posso spendere oggi?'."
        )


# ============================================================
# BUDGET
# ============================================================

elif menu == "📅 Budget":

    section_title(
        "📅 Budget mensile",
        "Imposta quanto vuoi spendere al massimo per ogni categoria"
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

    existing_budget = query_df("""
        SELECT categoria, importo
        FROM budget_mensile
        WHERE mese=?
    """, (key,))

    budget_dict = {}

    if not existing_budget.empty:

        for _, row in existing_budget.iterrows():

            budget_dict[row["categoria"]] = float(
                row["importo"]
            )

    with st.form("budget_form"):

        budget_values = {}

        cols = st.columns(2)

        for index, category in enumerate(
            EXPENSE_CATEGORIES
        ):

            with cols[index % 2]:

                budget_values[category] = st.number_input(
                    category,
                    min_value=0.0,
                    value=float(
                        budget_dict.get(
                            category,
                            0
                        )
                    ),
                    step=25.0,
                    key=f"budget_{key}_{category}"
                )

        save_budget = st.form_submit_button(
            "💾 Salva budget"
        )

        if save_budget:

            execute("""
                DELETE FROM budget_mensile
                WHERE mese=?
            """, (key,))

            for category, amount in budget_values.items():

                if amount > 0:

                    execute("""
                        INSERT INTO budget_mensile
                        (
                            mese,
                            categoria,
                            importo
                        )
                        VALUES (?, ?, ?)
                    """, (
                        key,
                        category,
                        amount
                    ))

            st.success(
                "Budget salvato."
            )

            st.rerun()

    st.markdown("---")

    st.subheader(
        f"📊 Situazione {month_label(selected.year, selected.month)}"
    )

    spent = monthly_expenses_by_category(
        selected.year,
        selected.month
    )

    budget = get_budget(key)

    if budget.empty:

        st.info(
            "Inserisci almeno un budget per vedere il confronto."
        )

    else:

        rows = []

        for _, row in budget.iterrows():

            category = row["categoria"]
            budget_value = float(row["importo"])

            spent_value = 0

            if not spent.empty:

                match = spent[
                    spent["categoria"] == category
                ]

                if not match.empty:

                    spent_value = float(
                        match.iloc[0]["importo"]
                    )

            remaining = budget_value - spent_value

            percentage = (
                spent_value / budget_value * 100
                if budget_value > 0
                else 0
            )

            rows.append({
                "Categoria": category,
                "Budget": budget_value,
                "Speso": spent_value,
                "Residuo": remaining,
                "% utilizzata": percentage
            })

        result = pd.DataFrame(rows)

        st.dataframe(
            result.style.format({
                "Budget": lambda x: euro(x),
                "Speso": lambda x: euro(x),
                "Residuo": lambda x: euro(x),
                "% utilizzata": lambda x: f"{x:.1f}%"
            }),
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# OBIETTIVI
# ============================================================

elif menu == "🎯 Obiettivi":

    section_title(
        "🎯 Obiettivi di risparmio",
        "Tieni sotto controllo i soldi che vuoi accantonare"
    )

    with st.form("nuovo_obiettivo"):

        name = st.text_input(
            "Nome obiettivo",
            placeholder="Es. Vacanza"
        )

        c1, c2 = st.columns(2)

        with c1:

            target = st.number_input(
                "Obiettivo (€)",
                min_value=0.0,
                step=500.0
            )

        with c2:

            saved = st.number_input(
                "Già accumulato (€)",
                min_value=0.0,
                step=100.0
            )

        deadline = st.date_input(
            "Scadenza",
            date.today()
        )

        create = st.form_submit_button(
            "🎯 Crea obiettivo"
        )

        if create:

            if not name.strip():
                st.error(
                    "Inserisci un nome."
                )

            elif target <= 0:
                st.error(
                    "Inserisci un obiettivo maggiore di zero."
                )

            else:

                execute("""
                    INSERT INTO obiettivi
                    (
                        nome,
                        obiettivo,
                        accumulato,
                        scadenza,
                        attivo
                    )
                    VALUES (?, ?, ?, ?, 1)
                """, (
                    name,
                    target,
                    saved,
                    deadline.isoformat()
                ))

                st.success(
                    "Obiettivo creato."
                )

                st.rerun()

    st.markdown("---")

    goals = query_df("""
        SELECT *
        FROM obiettivi
        WHERE attivo=1
        ORDER BY scadenza
    """)

    if goals.empty:

        st.info(
            "Nessun obiettivo attivo."
        )

    else:

        for _, goal in goals.iterrows():

            target = float(goal["obiettivo"])
            saved = float(goal["accumulato"])

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

            st.progress(progress)

            st.write(
                f"**{euro(saved)}** / {euro(target)} "
                f"({progress * 100:.1f}%)"
            )

            st.caption(
                f"Scadenza: {goal['scadenza']}"
            )


# ============================================================
# QUANTO POSSO SPENDERE OGGI
# ============================================================

elif menu == "💸 Quanto posso spendere oggi?":

    section_title(
        "💸 Quanto posso spendere oggi?",
        "Calcolo della disponibilità reale considerando spese già sostenute e impegni previsti"
    )

    today = date.today()

    st.info(
        f"📅 Calcolo riferito a oggi: {today.strftime('%d/%m/%Y')}"
    )

    # --------------------------------------------------------
    # MESE
    # --------------------------------------------------------

    year = today.year
    month = today.month
    key = month_key(year, month)

    # --------------------------------------------------------
    # ENTRATE
    # --------------------------------------------------------

    income = monthly_income_total(
        year,
        month
    )

    # --------------------------------------------------------
    # SPESE GIÀ FATTE
    # --------------------------------------------------------

    spent = monthly_expenses_total(
        year,
        month
    )

    # --------------------------------------------------------
    # RICORRENTI ANCORA DA PAGARE
    # --------------------------------------------------------

    recurring_remaining = recurring_remaining_today()

    # --------------------------------------------------------
    # SPESE FUTURE
    # --------------------------------------------------------

    future = future_expenses_remaining(
        year,
        month
    )

    # --------------------------------------------------------
    # RISPARMIO PROGRAMMATO
    # --------------------------------------------------------

    saving_target = get_setting(
        "risparmio_mensile_target",
        0
    )

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
        - recurring_remaining
        - future
        - saving_target
        - safety_fund
    )

    last_day = calendar.monthrange(
        year,
        month
    )[1]

    days_remaining = (
        last_day - today.day + 1
    )

    daily_budget = (
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
            f"💚 Disponibilità residua del mese: **{euro(available)}**"
        )

    else:

        st.error(
            f"🔴 Il budget del mese è già oltre il limite di "
            f"**{euro(abs(available))}**"
        )

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "💰 Disponibile",
        euro(available)
    )

    c2.metric(
        "📅 Giorni rimanenti",
        days_remaining
    )

    c3.metric(
        "💸 Budget giornaliero",
        euro(max(0, daily_budget))
    )

    st.markdown("---")

    st.subheader(
        "🔎 Come viene calcolato"
    )

    calculation = pd.DataFrame({
        "Voce": [
            "Entrate del mese",
            "Spese già sostenute",
            "Ricorrenti ancora previste",
            "Spese future",
            "Risparmio programmato",
            "Fondo sicurezza",
            "Disponibilità residua"
        ],
        "Importo": [
            income,
            -spent,
            -recurring_remaining,
            -future,
            -saving_target,
            -safety_fund,
            available
        ]
    })

    st.dataframe(
        calculation.style.format({
            "Importo": lambda x: euro(x)
        }),
        use_container_width=True,
        hide_index=True
    )

    st.caption(
        "Il budget giornaliero è una stima: non considera eventuali "
        "entrate straordinarie non ancora registrate."
    )

    # --------------------------------------------------------
    # SIMULATORE
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "🧮 Simula una spesa"
    )

    simulated_expense = st.number_input(
        "Se oggi spendessi...",
        min_value=0.0,
        step=10.0
    )

    simulated_remaining = (
        available - simulated_expense
    )

    if simulated_remaining >= 0:

        st.success(
            f"Dopo questa spesa avresti ancora "
            f"**{euro(simulated_remaining)}** disponibili."
        )

    else:

        st.warning(
            f"Con questa spesa supereresti la disponibilità "
            f"di **{euro(abs(simulated_remaining))}**."
        )


# ============================================================
# ANALISI
# ============================================================

elif menu == "📊 Analisi":

    section_title(
        "📊 Analisi",
        "Analizza come vengono distribuite le spese"
    )

    selected = st.date_input(
        "Mese da analizzare",
        date.today().replace(day=1),
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

    savings = income - total_expenses

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Entrate",
        euro(income)
    )

    c2.metric(
        "Spese",
        euro(total_expenses)
    )

    c3.metric(
        "Risparmio",
        euro(savings)
    )

    if not expenses.empty:

        st.markdown("---")

        st.subheader(
            "🥧 Distribuzione delle spese"
        )

        c1, c2 = st.columns(2)

        with c1:

            fig, ax = plt.subplots(
                figsize=(6, 6)
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

        with c2:

            st.dataframe(
                expenses.style.format({
                    "importo": lambda x: euro(x)
                }),
                use_container_width=True,
                hide_index=True
            )

        st.markdown("---")

        st.subheader(
            "📊 Spese per categoria"
        )

        fig, ax = plt.subplots(
            figsize=(10, 5)
        )

        ax.bar(
            expenses["categoria"],
            expenses["importo"]
        )

        ax.set_ylabel("€")
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
            "Nessuna spesa disponibile per questo mese."
        )

    # --------------------------------------------------------
    # MEDIA STORICA
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "📈 Media mensile storica per categoria"
    )

    historical = query_df("""
        SELECT
            categoria,
            AVG(totale) AS media
        FROM (
            SELECT
                mese,
                categoria,
                SUM(importo) AS totale
            FROM spese_mensili
            GROUP BY mese, categoria
        )
        GROUP BY categoria
        ORDER BY media DESC
    """)

    if not historical.empty:

        st.dataframe(
            historical.style.format({
                "media": lambda x: euro(x)
            }),
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# IMPOSTAZIONI
# ============================================================

elif menu == "⚙️ Impostazioni":

    section_title(
        "⚙️ Impostazioni",
        "Personalizza il funzionamento dell'app"
    )

    # --------------------------------------------------------
    # RISPARMIO
    # --------------------------------------------------------

    st.subheader(
        "🎯 Obiettivo di risparmio"
    )

    saving_target = st.number_input(
        "Quanto vuoi accantonare ogni mese? (€)",
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
        "Fondo sicurezza da lasciare sempre disponibile (€)",
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

    st.markdown("---")

    # --------------------------------------------------------
    # CATEGORIE
    # --------------------------------------------------------

    st.subheader(
        "🏷️ Categorie"
    )

    categories = query_df("""
        SELECT
            nome AS Categoria,
            tipo AS Tipo,
            attiva AS Attiva
        FROM categorie
        ORDER BY tipo, nome
    """)

    st.dataframe(
        categories,
        use_container_width=True,
        hide_index=True
    )

    # --------------------------------------------------------
    # PERSONE
    # --------------------------------------------------------

    st.subheader(
        "👨‍👩‍👧 Persone"
    )

    people = query_df("""
        SELECT
            nome AS Persona,
            attiva AS Attiva
        FROM persone
        ORDER BY nome
    """)

    st.dataframe(
        people,
        use_container_width=True,
        hide_index=True
    )

    # --------------------------------------------------------
    # GESTIONE DATI
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "🗑️ Gestione dati"
    )

    st.warning(
        "Le operazioni seguenti modificano i dati presenti nel database."
    )

    delete_type = st.selectbox(
        "Cosa vuoi eliminare?",
        [
            "Nessuna operazione",
            "Tutte le spese singole",
            "Tutte le entrate",
            "Tutte le spese future",
            "Tutti i budget",
            "Tutti gli obiettivi"
        ]
    )

    confirm = st.checkbox(
        "Confermo di voler eliminare i dati selezionati"
    )

    if st.button(
        "🗑️ Esegui eliminazione"
    ):

        if not confirm:

            st.error(
                "Devi confermare l'operazione."
            )

        else:

            if delete_type == "Tutte le spese singole":

                execute("""
                    DELETE FROM movimenti
                    WHERE tipo='uscita'
                """)

            elif delete_type == "Tutte le entrate":

                execute("""
                    DELETE FROM movimenti
                    WHERE tipo='entrata'
                """)

            elif delete_type == "Tutte le spese future":

                execute(
                    "DELETE FROM spese_future"
                )

            elif delete_type == "Tutti i budget":

                execute(
                    "DELETE FROM budget_mensile"
                )

            elif delete_type == "Tutti gli obiettivi":

                execute(
                    "DELETE FROM obiettivi"
                )

            if delete_type != "Nessuna operazione":

                st.success(
                    "Operazione completata."
                )

                st.rerun()


# ============================================================
# IMPORT / EXPORT
# ============================================================

elif menu == "📥 Import / Export":

    section_title(
        "📥 Import / Export",
        "Backup completo dei dati dell'app"
    )

    # --------------------------------------------------------
    # STATO EXCEL
    # --------------------------------------------------------

    if EXCEL_FILE.exists():

        st.success(
            "🟢 File Excel trovato e storico importato."
        )

        st.caption(
            f"File: {EXCEL_FILE.name}"
        )

        historical = parse_excel_family()

        if not historical.empty:

            st.metric(
                "Righe storico Excel",
                len(historical)
            )

    else:

        st.warning(
            "Il file Spese casa -2.xlsx non è presente "
            "nella stessa cartella dell'app."
        )

    # --------------------------------------------------------
    # ESPORTAZIONE
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "⬇️ Esporta dati"
    )

    tables = {

        "Movimenti": query_df("""
            SELECT *
            FROM movimenti
            ORDER BY data
        """),

        "Spese mensili": query_df("""
            SELECT *
            FROM spese_mensili
            ORDER BY mese
        """),

        "Ricorrenti": query_df("""
            SELECT *
            FROM ricorrenti
        """),

        "Spese future": query_df("""
            SELECT *
            FROM spese_future
            ORDER BY data
        """),

        "Budget": query_df("""
            SELECT *
            FROM budget_mensile
            ORDER BY mese
        """),

        "Obiettivi": query_df("""
            SELECT *
            FROM obiettivi
        """),

        "Impostazioni": query_df("""
            SELECT *
            FROM impostazioni
        """),

        "Categorie": query_df("""
            SELECT *
            FROM categorie
        """),

        "Persone": query_df("""
            SELECT *
            FROM persone
        """)
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

        if EXCEL_FILE.exists():

            historical = parse_excel_family()

            if not historical.empty:

                historical.to_excel(
                    writer,
                    sheet_name="Storico Excel",
                    index=False
                )

    st.download_button(
        "⬇️ Scarica backup Excel",
        data=output.getvalue(),
        file_name="backup_finanze_famiglia.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )

    # --------------------------------------------------------
    # RIEPILOGO DATABASE
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "📊 Stato del database"
    )

    stats = {
        "Spese singole": query_df("""
            SELECT COUNT(*) AS n
            FROM movimenti
            WHERE tipo='uscita'
        """).iloc[0]["n"],

        "Entrate": query_df("""
            SELECT COUNT(*) AS n
            FROM movimenti
            WHERE tipo='entrata'
        """).iloc[0]["n"],

        "Mesi con spese": query_df("""
            SELECT COUNT(DISTINCT mese) AS n
            FROM spese_mensili
        """).iloc[0]["n"],

        "Ricorrenti": query_df("""
            SELECT COUNT(*) AS n
            FROM ricorrenti
        """).iloc[0]["n"],

        "Budget": query_df("""
            SELECT COUNT(*) AS n
            FROM budget_mensile
        """).iloc[0]["n"],

        "Obiettivi": query_df("""
            SELECT COUNT(*) AS n
            FROM obiettivi
        """).iloc[0]["n"]
    }

    stats_df = pd.DataFrame(
        list(stats.items()),
        columns=["Voce", "Numero"]
    )

    st.dataframe(
        stats_df,
        use_container_width=True,
        hide_index=True
    )
