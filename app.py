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
    page_title="Gestione Casa & Finanze",
    page_icon="🏡",
    layout="centered",
    initial_sidebar_state="expanded"
)

BASE_DIR = Path(__file__).resolve().parent

# Il file Excel deve stare nella stessa cartella di app.py
EXCEL_FILE = BASE_DIR / "Spese casa -2.xlsx"

# Database utilizzato SOLO per i nuovi dati inseriti nell'app
DB_FILE = BASE_DIR / "finanze_famiglia.db"


# ============================================================
# DATABASE
# ============================================================

def get_connection():
    return sqlite3.connect(DB_FILE)


def init_database():

    conn = get_connection()
    cursor = conn.cursor()

    # --------------------------------------------------------
    # ENTRATE
    # --------------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS entrate (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            descrizione TEXT NOT NULL,
            persona TEXT,
            categoria TEXT,
            importo REAL NOT NULL,
            data TEXT NOT NULL,
            ricorrente INTEGER DEFAULT 0
        )
    """)

    # --------------------------------------------------------
    # SPESE NUOVE
    # --------------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS spese (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            descrizione TEXT NOT NULL,
            persona TEXT,
            categoria TEXT,
            importo REAL NOT NULL,
            data TEXT NOT NULL,
            ricorrente INTEGER DEFAULT 0,
            note TEXT
        )
    """)

    # --------------------------------------------------------
    # SPESE RICORRENTI
    # --------------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS spese_ricorrenti (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            descrizione TEXT NOT NULL,
            categoria TEXT,
            importo REAL NOT NULL,
            giorno INTEGER DEFAULT 1,
            persona TEXT
        )
    """)

    # --------------------------------------------------------
    # BUDGET
    # --------------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS budget (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            categoria TEXT UNIQUE,
            importo REAL
        )
    """)

    # --------------------------------------------------------
    # OBIETTIVI
    # --------------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS obiettivi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT,
            obiettivo REAL,
            accumulato REAL DEFAULT 0,
            scadenza TEXT
        )
    """)

    # --------------------------------------------------------
    # SPESE FUTURE
    # --------------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS spese_future (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data TEXT,
            categoria TEXT,
            descrizione TEXT,
            importo REAL,
            note TEXT
        )
    """)

    # --------------------------------------------------------
    # COSTI NUOVA CASA
    # --------------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS costi_casa (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            voce TEXT,
            importo REAL
        )
    """)

    # --------------------------------------------------------
    # COPERTURE NUOVA CASA
    # --------------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS entrate_casa (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            voce TEXT,
            importo REAL
        )
    """)

    # --------------------------------------------------------
    # IMPOSTAZIONI
    # --------------------------------------------------------

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS impostazioni (
            chiave TEXT PRIMARY KEY,
            valore TEXT
        )
    """)

    conn.commit()
    conn.close()


init_database()


# ============================================================
# FUNZIONI DATABASE
# ============================================================

def query_df(query, params=()):

    conn = get_connection()

    try:
        df = pd.read_sql_query(
            query,
            conn,
            params=params
        )
        return df

    finally:
        conn.close()


def execute(query, params=()):

    conn = get_connection()

    try:
        cursor = conn.cursor()
        cursor.execute(query, params)
        conn.commit()
        return cursor.lastrowid

    finally:
        conn.close()


# ============================================================
# DATI STATICI
# ============================================================

CATEGORIE_NUOVE_SPESE = [
    "🍝 Alimentari",
    "🏠 Casa",
    "💡 Utenze",
    "🚗 Trasporti e auto",
    "👧 Figli",
    "🎓 Scuola e sport",
    "🛍️ Shopping",
    "🏥 Salute",
    "🍕 Ristoranti",
    "✈️ Vacanze e viaggi",
    "🎮 Tempo libero",
    "🪑 Arredi e nuova casa",
    "📦 Altro"
]

CATEGORIE_ENTRATE = [
    "Stipendio",
    "Bonus",
    "Rimborso",
    "Affitto",
    "Altre entrate"
]

PERSONE = [
    "Famiglia",
    "Io",
    "Partner"
]


# ============================================================
# FUNZIONI UTILI
# ============================================================

def euro(valore):

    try:
        valore = float(valore)
    except:
        valore = 0

    return (
        f"{valore:,.2f} €"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


def normalizza_importo(valore):

    if pd.isna(valore):
        return None

    if isinstance(valore, (int, float)):
        return float(valore)

    testo = str(valore).strip()

    if not testo:
        return None

    testo = (
        testo
        .replace("€", "")
        .replace(" ", "")
    )

    # Esempio:
    # 1.234,56 -> 1234.56
    if "," in testo:
        testo = testo.replace(".", "")
        testo = testo.replace(",", ".")

    try:
        return float(testo)

    except:
        return None


def mese_label(anno, mese):

    mesi = [
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
        "Dicembre"
    ]

    return f"{mesi[mese - 1]} {anno}"


# ============================================================
# LETTURA DELLO STORICO EXCEL
# ============================================================

@st.cache_data(ttl=30)
def leggi_storico_excel(percorso, modifica_file):

    colonne = [
        "anno",
        "mese",
        "data",
        "categoria",
        "importo",
        "fonte"
    ]

    if not Path(percorso).exists():

        return pd.DataFrame(
            columns=colonne
        )

    try:

        df = pd.read_excel(
            percorso,
            sheet_name="Costi famiglia",
            header=None
        )

    except Exception:

        return pd.DataFrame(
            columns=colonne
        )

    storico = []

    anno_corrente = None
    mese_corrente = None

    mesi = {
        "GENNAIO": 1,
        "FEBBRAIO": 2,
        "MARZO": 3,
        "APRILE": 4,
        "MAGGIO": 5,
        "GIUGNO": 6,
        "LUGLIO": 7,
        "AGOSTO": 8,
        "SETTEMBRE": 9,
        "OTTOBRE": 10,
        "NOVEMBRE": 11,
        "DICEMBRE": 12
    }

    # --------------------------------------------------------
    # Scorre tutte le righe del foglio
    # --------------------------------------------------------

    for indice, riga in df.iterrows():

        if len(riga) == 0:
            continue

        prima_cella = riga.iloc[0]

        if pd.isna(prima_cella):
            continue

        testo = str(
            prima_cella
        ).strip()

        testo_upper = testo.upper()

        # ----------------------------------------------------
        # RICONOSCE LA RIGA DEL MESE
        #
        # Esempio:
        # Spese OTTOBRE 2025
        # ----------------------------------------------------

        match = re.match(
            r"^SPESE\s+([A-ZÀ-Ù]+)\s+(\d{4})$",
            testo_upper
        )

        if match:

            nome_mese = match.group(1)
            anno_corrente = int(
                match.group(2)
            )

            if nome_mese in mesi:

                mese_corrente = mesi[
                    nome_mese
                ]

            continue

        # ----------------------------------------------------
        # Se non abbiamo ancora trovato un mese
        # ignoriamo la riga
        # ----------------------------------------------------

        if (
            anno_corrente is None
            or mese_corrente is None
        ):
            continue

        # ----------------------------------------------------
        # La prima colonna contiene la categoria
        # ----------------------------------------------------

        categoria = testo

        if categoria.lower() in [
            "nan",
            "media",
            "totale",
            ""
        ]:
            continue

        # ----------------------------------------------------
        # La seconda colonna contiene il costo
        # ----------------------------------------------------

        if len(riga) < 2:
            continue

        importo = normalizza_importo(
            riga.iloc[1]
        )

        if importo is None:
            continue

        if importo == 0:
            continue

        data_mese = date(
            anno_corrente,
            mese_corrente,
            1
        )

        storico.append(
            {
                "anno": anno_corrente,
                "mese": mese_corrente,
                "data": data_mese,
                "categoria": categoria,
                "importo": importo,
                "fonte": "Excel"
            }
        )

    risultato = pd.DataFrame(
        storico
    )

    if not risultato.empty:

        risultato["data"] = pd.to_datetime(
            risultato["data"]
        )

    return risultato


# ============================================================
# LETTURA FOGLIO MOBILI
# ============================================================

@st.cache_data(ttl=30)
def leggi_mobili_excel(percorso, modifica_file):

    if not Path(percorso).exists():
        return pd.DataFrame()

    try:

        df = pd.read_excel(
            percorso,
            sheet_name="Mobili",
            header=None
        )

    except Exception:

        return pd.DataFrame()

    mobili = []

    # Nel tuo Excel:
    #
    # Riga 0 = Acquisto mobili
    # Riga 1 = Articolo / Costo / RIFERIMENTO / NEGOZIO
    # Riga 2 in poi = prodotti

    for indice in range(2, len(df)):

        riga = df.iloc[indice]

        if len(riga) < 2:
            continue

        articolo = riga.iloc[0]

        if pd.isna(articolo):
            continue

        costo = normalizza_importo(
            riga.iloc[1]
        )

        if costo is None:
            costo = 0

        riferimento = ""

        negozio = ""

        if (
            len(riga) > 2
            and not pd.isna(riga.iloc[2])
        ):
            riferimento = str(
                riga.iloc[2]
            )

        if (
            len(riga) > 3
            and not pd.isna(riga.iloc[3])
        ):
            negozio = str(
                riga.iloc[3]
            )

        mobili.append(
            {
                "Articolo": str(articolo),
                "Costo": costo,
                "Riferimento": riferimento,
                "Negozio": negozio
            }
        )

    return pd.DataFrame(
        mobili
    )


# ============================================================
# CARICAMENTO EXCEL
# ============================================================

if EXCEL_FILE.exists():

    modifica_excel = EXCEL_FILE.stat().st_mtime

    storico_excel = leggi_storico_excel(
        str(EXCEL_FILE),
        modifica_excel
    )

    mobili_excel = leggi_mobili_excel(
        str(EXCEL_FILE),
        modifica_excel
    )

else:

    storico_excel = pd.DataFrame()

    mobili_excel = pd.DataFrame()


# ============================================================
# FUNZIONI SULLO STORICO
# ============================================================

def storico_mese(anno, mese):

    if storico_excel.empty:
        return pd.DataFrame()

    return storico_excel[
        (storico_excel["anno"] == anno)
        &
        (storico_excel["mese"] == mese)
    ].copy()


def totale_storico_mese(anno, mese):

    dati = storico_mese(
        anno,
        mese
    )

    if dati.empty:
        return 0

    return float(
        dati["importo"].sum()
    )


def totale_spese_app_mese(anno, mese):

    df = query_df(
        """
        SELECT
            COALESCE(SUM(importo), 0) AS totale
        FROM spese
        WHERE strftime('%Y', data) = ?
        AND strftime('%m', data) = ?
        """,
        (
            str(anno),
            f"{mese:02d}"
        )
    )

    return float(
        df.iloc[0]["totale"]
    )


def totale_spese_completo_mese(anno, mese):

    return (
        totale_storico_mese(
            anno,
            mese
        )
        +
        totale_spese_app_mese(
            anno,
            mese
        )
    )


def totale_entrate_mese(anno, mese):

    df = query_df(
        """
        SELECT
            COALESCE(SUM(importo), 0) AS totale
        FROM entrate
        WHERE strftime('%Y', data) = ?
        AND strftime('%m', data) = ?
        """,
        (
            str(anno),
            f"{mese:02d}"
        )
    )

    return float(
        df.iloc[0]["totale"]
    )


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.title(
    "🏡 Gestione Casa"
)

menu = st.sidebar.radio(
    "MENU",
    [
        "🏠 Home",
        "📜 Storico Excel",
        "💰 Entrate",
        "💸 Spese",
        "🔁 Spese ricorrenti",
        "📅 Budget",
        "🎯 Obiettivi",
        "🔮 Previsioni",
        "📊 Analisi",
        "🏡 Nuova Casa",
        "🪑 Mobili",
        "🖼️ Rendering & Planimetrie",
        "📥 Importa / Esporta",
        "⚙️ Impostazioni"
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

    st.sidebar.error(
        "🔴 Excel non trovato"
    )

    st.sidebar.caption(
        "Metti il file nella stessa cartella di app.py"
    )


# ============================================================
# HOME
# ============================================================

if menu == "🏠 Home":

    st.title(
        "🏡 Gestione Casa & Finanze"
    )

    st.write(
        "Dashboard finanziaria familiare"
    )

    st.markdown("---")

    # --------------------------------------------------------
    # MESE
    # --------------------------------------------------------

    mese_scelto = st.date_input(
        "📅 Seleziona mese",
        value=date.today().replace(day=1)
    )

    anno = mese_scelto.year
    mese = mese_scelto.month

    # --------------------------------------------------------
    # DATI
    # --------------------------------------------------------

    storico = totale_storico_mese(
        anno,
        mese
    )

    nuove_spese = totale_spese_app_mese(
        anno,
        mese
    )

    entrate = totale_entrate_mese(
        anno,
        mese
    )

    spese_totali = (
        storico +
        nuove_spese
    )

    risparmio = (
        entrate -
        spese_totali
    )

    percentuale_risparmio = (
        risparmio / entrate * 100
        if entrate > 0
        else 0
    )

    # --------------------------------------------------------
    # METRICHE PRINCIPALI
    # --------------------------------------------------------

    col1, col2 = st.columns(2)

    with col1:

        st.metric(
            "💰 Entrate",
            euro(entrate)
        )

    with col2:

        st.metric(
            "💸 Spese",
            euro(spese_totali)
        )

    col3, col4 = st.columns(2)

    with col3:

        st.metric(
            "💚 Risparmio",
            euro(risparmio)
        )

    with col4:

        st.metric(
            "📈 Risparmio %",
            f"{percentuale_risparmio:.1f}%"
        )

    # --------------------------------------------------------
    # ORIGINE DELLE SPESE
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        f"📌 Composizione spese - {mese_label(anno, mese)}"
    )

    col1, col2 = st.columns(2)

    with col1:

        st.metric(
            "📊 Storico Excel",
            euro(storico)
        )

    with col2:

        st.metric(
            "✏️ Inserite nell'app",
            euro(nuove_spese)
        )

    # --------------------------------------------------------
    # DETTAGLIO STORICO
    # --------------------------------------------------------

    dati_mese = storico_mese(
        anno,
        mese
    )

    if not dati_mese.empty:

        st.subheader(
            "📊 Dove sono finiti i soldi?"
        )

        fig, ax = plt.subplots(
            figsize=(8, 6)
        )

        ax.pie(
            dati_mese["importo"],
            labels=dati_mese["categoria"],
            autopct="%1.1f%%",
            startangle=90
        )

        ax.axis("equal")

        st.pyplot(
            fig,
            clear_figure=True
        )

    # --------------------------------------------------------
    # DISPONIBILITÀ
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "💳 Quanto possiamo spendere?"
    )

    ricorrenti_df = query_df(
        """
        SELECT
            COALESCE(SUM(importo),0) AS totale
        FROM spese_ricorrenti
        """
    )

    spese_ricorrenti = float(
        ricorrenti_df.iloc[0]["totale"]
    )

    future_df = query_df(
        """
        SELECT
            COALESCE(SUM(importo),0) AS totale
        FROM spese_future
        WHERE strftime('%Y',data)=?
        AND strftime('%m',data)=?
        """,
        (
            str(anno),
            f"{mese:02d}"
        )
    )

    spese_future = float(
        future_df.iloc[0]["totale"]
    )

    # Le spese ricorrenti e future sono una previsione
    # utile soprattutto per il mese corrente/futuro.
    if (
        anno == date.today().year
        and mese == date.today().month
    ):

        spese_da_affrontare = (
            spese_ricorrenti +
            spese_future
        )

    else:

        spese_da_affrontare = 0

    disponibilita = (
        entrate
        -
        spese_totali
        -
        spese_da_affrontare
    )

    if disponibilita >= 0:

        st.success(
            f"💚 Disponibilità stimata: "
            f"**{euro(disponibilita)}**"
        )

    else:

        st.error(
            f"⚠️ Disponibilità stimata: "
            f"**{euro(disponibilita)}**"
        )

    # --------------------------------------------------------
    # ULTIMI 12 MESI
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "📈 Andamento ultimi 12 mesi"
    )

    righe = []

    anno_temp = date.today().year
    mese_temp = date.today().month

    for _ in range(12):

        totale = totale_spese_completo_mese(
            anno_temp,
            mese_temp
        )

        righe.append(
            {
                "Mese": mese_label(
                    anno_temp,
                    mese_temp
                ),
                "Spese": totale
            }
        )

        mese_temp -= 1

        if mese_temp == 0:

            mese_temp = 12
            anno_temp -= 1

    df_grafico = pd.DataFrame(
        righe[::-1]
    )

    fig, ax = plt.subplots(
        figsize=(10, 5)
    )

    ax.plot(
        df_grafico["Mese"],
        df_grafico["Spese"],
        marker="o"
    )

    ax.set_title(
        "Spese mensili"
    )

    ax.set_ylabel(
        "Euro"
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
# STORICO EXCEL
# ============================================================

elif menu == "📜 Storico Excel":

    st.title(
        "📜 Storico spese Excel"
    )

    st.info(
        "Questa sezione legge direttamente il foglio "
        "'Costi famiglia' del file "
        "'Spese casa -2.xlsx'."
    )

    if storico_excel.empty:

        st.error(
            "Non sono riuscito a leggere lo storico Excel."
        )

        st.write(
            f"File cercato: `{EXCEL_FILE}`"
        )

    else:

        # ----------------------------------------------------
        # ANNO
        # ----------------------------------------------------

        anni = sorted(
            storico_excel["anno"]
            .unique(),
            reverse=True
        )

        anno = st.selectbox(
            "Anno",
            anni
        )

        # ----------------------------------------------------
        # MESE
        # ----------------------------------------------------

        mesi_disponibili = sorted(
            storico_excel[
                storico_excel["anno"] == anno
            ]["mese"]
            .unique()
        )

        mese = st.selectbox(
            "Mese",
            mesi_disponibili,
            format_func=lambda x:
                mese_label(anno, x)
        )

        dati = storico_mese(
            anno,
            mese
        )

        totale = (
            dati["importo"].sum()
            if not dati.empty
            else 0
        )

        st.metric(
            f"Totale {mese_label(anno, mese)}",
            euro(totale)
        )

        if not dati.empty:

            # ------------------------------------------------
            # GRAFICO TORTA
            # ------------------------------------------------

            st.subheader(
                "🍕 Distribuzione"
            )

            fig, ax = plt.subplots(
                figsize=(8, 7)
            )

            ax.pie(
                dati["importo"],
                labels=dati["categoria"],
                autopct="%1.1f%%",
                startangle=90
            )

            ax.axis("equal")

            st.pyplot(
                fig,
                clear_figure=True
            )

            # ------------------------------------------------
            # TABELLA
            # ------------------------------------------------

            tabella = dati[
                [
                    "categoria",
                    "importo"
                ]
            ].copy()

            tabella.columns = [
                "Categoria",
                "Importo"
            ]

            st.dataframe(
                tabella,
                use_container_width=True,
                hide_index=True
            )

            # ------------------------------------------------
            # GRAFICO BARRE
            # ------------------------------------------------

            st.subheader(
                "📊 Dettaglio categorie"
            )

            dati_barre = dati.sort_values(
                "importo",
                ascending=True
            )

            fig2, ax2 = plt.subplots(
                figsize=(10, 6)
            )

            ax2.barh(
                dati_barre["categoria"],
                dati_barre["importo"]
            )

            ax2.set_xlabel(
                "Euro"
            )

            st.pyplot(
                fig2,
                clear_figure=True
            )


# ============================================================
# ENTRATE
# ============================================================

elif menu == "💰 Entrate":

    st.title(
        "💰 Entrate"
    )

    st.subheader(
        "➕ Inserisci nuova entrata"
    )

    with st.form(
        "nuova_entrata"
    ):

        descrizione = st.text_input(
            "Descrizione",
            placeholder="Es. Stipendio"
        )

        col1, col2 = st.columns(2)

        with col1:

            persona = st.selectbox(
                "Persona",
                PERSONE
            )

        with col2:

            categoria = st.selectbox(
                "Categoria",
                CATEGORIE_ENTRATE
            )

        importo = st.number_input(
            "Importo (€)",
            min_value=0.0,
            step=50.0
        )

        data_entrata = st.date_input(
            "Data",
            value=date.today()
        )

        ricorrente = st.checkbox(
            "🔁 Entrata ricorrente"
        )

        salva = st.form_submit_button(
            "💾 Salva entrata"
        )

        if salva:

            if (
                descrizione.strip()
                and importo > 0
            ):

                execute(
                    """
                    INSERT INTO entrate
                    (
                        descrizione,
                        persona,
                        categoria,
                        importo,
                        data,
                        ricorrente
                    )
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        descrizione,
                        persona,
                        categoria,
                        importo,
                        data_entrata.isoformat(),
                        int(ricorrente)
                    )
                )

                st.success(
                    "Entrata salvata."
                )

                st.rerun()

            else:

                st.warning(
                    "Inserisci descrizione e importo."
                )

    st.markdown("---")

    st.subheader(
        "📋 Entrate inserite nell'app"
    )

    df = query_df(
        """
        SELECT
            id AS ID,
            descrizione AS Descrizione,
            persona AS Persona,
            categoria AS Categoria,
            importo AS Importo,
            data AS Data
        FROM entrate
        ORDER BY data DESC
        """
    )

    if df.empty:

        st.info(
            "Nessuna entrata inserita."
        )

    else:

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# SPESE
# ============================================================

elif menu == "💸 Spese":

    st.title(
        "💸 Spese"
    )

    st.info(
        "Le spese inserite qui vengono salvate nell'app. "
        "Lo storico precedente rimane nel file Excel."
    )

    with st.form(
        "nuova_spesa"
    ):

        descrizione = st.text_input(
            "Descrizione",
            placeholder="Es. Supermercato"
        )

        col1, col2 = st.columns(2)

        with col1:

            categoria = st.selectbox(
                "Categoria",
                CATEGORIE_NUOVE_SPESE
            )

        with col2:

            persona = st.selectbox(
                "Persona",
                PERSONE
            )

        importo = st.number_input(
            "Importo (€)",
            min_value=0.0,
            step=5.0
        )

        data_spesa = st.date_input(
            "Data",
            value=date.today()
        )

        note = st.text_input(
            "Note",
            placeholder="Facoltativo"
        )

        ricorrente = st.checkbox(
            "🔁 Spesa ricorrente"
        )

        salva = st.form_submit_button(
            "💾 Salva spesa"
        )

        if salva:

            if (
                descrizione.strip()
                and importo > 0
            ):

                execute(
                    """
                    INSERT INTO spese
                    (
                        descrizione,
                        persona,
                        categoria,
                        importo,
                        data,
                        ricorrente,
                        note
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        descrizione,
                        persona,
                        categoria,
                        importo,
                        data_spesa.isoformat(),
                        int(ricorrente),
                        note
                    )
                )

                st.success(
                    "Spesa salvata."
                )

                st.rerun()

            else:

                st.warning(
                    "Inserisci descrizione e importo."
                )

    st.markdown("---")

    st.subheader(
        "📋 Spese inserite nell'app"
    )

    df = query_df(
        """
        SELECT
            id AS ID,
            descrizione AS Descrizione,
            categoria AS Categoria,
            persona AS Persona,
            importo AS Importo,
            data AS Data,
            note AS Note
        FROM spese
        ORDER BY data DESC
        """
    )

    if df.empty:

        st.info(
            "Nessuna spesa inserita nell'app."
        )

    else:

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True
        )

        st.metric(
            "Totale spese app",
            euro(df["Importo"].sum())
        )

        st.markdown("---")

        id_delete = st.number_input(
            "ID della spesa da eliminare",
            min_value=0,
            step=1
        )

        if st.button(
            "🗑️ Elimina spesa"
        ):

            if id_delete > 0:

                execute(
                    "DELETE FROM spese WHERE id=?",
                    (id_delete,)
                )

                st.success(
                    "Spesa eliminata."
                )

                st.rerun()


# ============================================================
# SPESE RICORRENTI
# ============================================================

elif menu == "🔁 Spese ricorrenti":

    st.title(
        "🔁 Spese ricorrenti"
    )

    st.write(
        "Inserisci qui mutuo, assicurazioni, "
        "abbonamenti, scuola, ecc."
    )

    with st.form(
        "spesa_ricorrente"
    ):

        descrizione = st.text_input(
            "Descrizione",
            placeholder="Es. Mutuo"
        )

        categoria = st.selectbox(
            "Categoria",
            CATEGORIE_NUOVE_SPESE
        )

        importo = st.number_input(
            "Importo mensile (€)",
            min_value=0.0,
            step=10.0
        )

        giorno = st.number_input(
            "Giorno del mese",
            min_value=1,
            max_value=31,
            value=1
        )

        persona = st.selectbox(
            "Persona",
            PERSONE
        )

        salva = st.form_submit_button(
            "➕ Aggiungi"
        )

        if salva:

            if (
                descrizione.strip()
                and importo > 0
            ):

                execute(
                    """
                    INSERT INTO spese_ricorrenti
                    (
                        descrizione,
                        categoria,
                        importo,
                        giorno,
                        persona
                    )
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    (
                        descrizione,
                        categoria,
                        importo,
                        giorno,
                        persona
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
            persona AS Persona
        FROM spese_ricorrenti
        ORDER BY giorno
        """
    )

    if df.empty:

        st.info(
            "Nessuna spesa ricorrente."
        )

    else:

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True
        )

        st.metric(
            "💸 Totale spese fisse mensili",
            euro(df["Importo"].sum())
        )

        id_delete = st.number_input(
            "ID ricorrenza da eliminare",
            min_value=0,
            step=1
        )

        if st.button(
            "🗑️ Elimina ricorrenza"
        ):

            if id_delete > 0:

                execute(
                    """
                    DELETE FROM spese_ricorrenti
                    WHERE id=?
                    """,
                    (id_delete,)
                )

                st.rerun()


# ============================================================
# BUDGET
# ============================================================

elif menu == "📅 Budget":

    st.title(
        "📅 Budget familiare"
    )

    st.write(
        "Imposta il limite mensile per le nuove categorie "
        "di spesa inserite nell'app."
    )

    categoria = st.selectbox(
        "Categoria",
        CATEGORIE_NUOVE_SPESE
    )

    importo = st.number_input(
        "Budget mensile (€)",
        min_value=0.0,
        step=50.0
    )

    if st.button(
        "💾 Salva budget"
    ):

        execute(
            """
            INSERT INTO budget
            (
                categoria,
                importo
            )
            VALUES (?, ?)
            ON CONFLICT(categoria)
            DO UPDATE SET
                importo=excluded.importo
            """,
            (
                categoria,
                importo
            )
        )

        st.success(
            "Budget aggiornato."
        )

        st.rerun()

    st.markdown("---")

    mese_budget = st.date_input(
        "Controlla mese",
        value=date.today().replace(day=1)
    )

    anno = mese_budget.year
    mese = mese_budget.month

    budget_df = query_df(
        """
        SELECT
            categoria,
            importo
        FROM budget
        """
    )

    righe = []

    for _, row in budget_df.iterrows():

        categoria = row["categoria"]

        budget = float(
            row["importo"]
        )

        speso_df = query_df(
            """
            SELECT
                COALESCE(SUM(importo),0) AS totale
            FROM spese
            WHERE categoria=?
            AND strftime('%Y',data)=?
            AND strftime('%m',data)=?
            """,
            (
                categoria,
                str(anno),
                f"{mese:02d}"
            )
        )

        speso = float(
            speso_df.iloc[0]["totale"]
        )

        residuo = (
            budget -
            speso
        )

        righe.append(
            {
                "Categoria": categoria,
                "Budget": budget,
                "Speso": speso,
                "Residuo": residuo
            }
        )

    if righe:

        risultato = pd.DataFrame(
            righe
        )

        st.dataframe(
            risultato,
            use_container_width=True,
            hide_index=True
        )

        for _, row in risultato.iterrows():

            if row["Residuo"] < 0:

                st.error(
                    f"⚠️ {row['Categoria']} "
                    f"superato di "
                    f"{euro(abs(row['Residuo']))}"
                )

    else:

        st.info(
            "Non hai ancora impostato nessun budget."
        )


# ============================================================
# OBIETTIVI
# ============================================================

elif menu == "🎯 Obiettivi":

    st.title(
        "🎯 Obiettivi di risparmio"
    )

    with st.form(
        "nuovo_obiettivo"
    ):

        nome = st.text_input(
            "Nome obiettivo",
            placeholder="Es. Vacanza"
        )

        obiettivo = st.number_input(
            "Importo obiettivo (€)",
            min_value=0.0,
            step=500.0
        )

        accumulato = st.number_input(
            "Già accumulato (€)",
            min_value=0.0,
            step=100.0
        )

        scadenza = st.date_input(
            "Scadenza",
            value=date.today()
        )

        salva = st.form_submit_button(
            "🎯 Crea obiettivo"
        )

        if salva:

            if (
                nome.strip()
                and obiettivo > 0
            ):

                execute(
                    """
                    INSERT INTO obiettivi
                    (
                        nome,
                        obiettivo,
                        accumulato,
                        scadenza
                    )
                    VALUES (?, ?, ?, ?)
                    """,
                    (
                        nome,
                        obiettivo,
                        accumulato,
                        scadenza.isoformat()
                    )
                )

                st.success(
                    "Obiettivo creato."
                )

                st.rerun()

    st.markdown("---")

    df = query_df(
        "SELECT * FROM obiettivi"
    )

    if df.empty:

        st.info(
            "Nessun obiettivo."
        )

    else:

        for _, row in df.iterrows():

            obiettivo = float(
                row["obiettivo"]
            )

            accumulato = float(
                row["accumulato"]
            )

            percentuale = (
                accumulato / obiettivo
                if obiettivo > 0
                else 0
            )

            percentuale = min(
                percentuale,
                1
            )

            st.subheader(
                f"🎯 {row['nome']}"
            )

            st.progress(
                percentuale
            )

            st.write(
                f"{euro(accumulato)} / "
                f"{euro(obiettivo)}"
            )

            mancante = max(
                obiettivo - accumulato,
                0
            )

            st.caption(
                f"Mancano {euro(mancante)}"
            )

            if st.button(
                f"🗑️ Elimina {row['nome']}",
                key=f"elimina_obiettivo_{row['id']}"
            ):

                execute(
                    "DELETE FROM obiettivi WHERE id=?",
                    (row["id"],)
                )

                st.rerun()


# ============================================================
# PREVISIONI
# ============================================================

elif menu == "🔮 Previsioni":

    st.title(
        "🔮 Previsioni"
    )

    if storico_excel.empty:

        st.warning(
            "Nessun dato storico disponibile."
        )

    else:

        mensile = (
            storico_excel
            .groupby(
                ["anno", "mese"],
                as_index=False
            )["importo"]
            .sum()
        )

        mensile["Mese"] = mensile.apply(
            lambda r:
            mese_label(
                int(r["anno"]),
                int(r["mese"])
            ),
            axis=1
        )

        # ----------------------------------------------------
        # MEDIE
        # ----------------------------------------------------

        media_3 = (
            mensile.tail(3)["importo"].mean()
        )

        media_6 = (
            mensile.tail(6)["importo"].mean()
        )

        media_12 = (
            mensile.tail(12)["importo"].mean()
        )

        col1, col2, col3 = st.columns(3)

        with col1:

            st.metric(
                "Media 3 mesi",
                euro(media_3)
            )

        with col2:

            st.metric(
                "Media 6 mesi",
                euro(media_6)
            )

        with col3:

            st.metric(
                "Media 12 mesi",
                euro(media_12)
            )

        st.markdown("---")

        st.subheader(
            "🔮 Stima spese prossimo mese"
        )

        st.info(
            f"In base alla media degli ultimi "
            f"3 mesi: **{euro(media_3)}**"
        )

        # ----------------------------------------------------
        # GRAFICO
        # ----------------------------------------------------

        fig, ax = plt.subplots(
            figsize=(10, 5)
        )

        ax.plot(
            mensile["Mese"],
            mensile["importo"],
            marker="o"
        )

        ax.axhline(
            media_3,
            linestyle="--",
            label="Media 3 mesi"
        )

        ax.legend()

        ax.set_ylabel(
            "Euro"
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

        # ----------------------------------------------------
        # CATEGORIE
        # ----------------------------------------------------

        st.subheader(
            "📊 Categorie principali"
        )

        categorie = (
            storico_excel
            .groupby("categoria")["importo"]
            .sum()
            .sort_values(
                ascending=False
            )
        )

        st.dataframe(
            categorie
            .reset_index()
            .rename(
                columns={
                    "categoria": "Categoria",
                    "importo": "Totale"
                }
            ),
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# ANALISI
# ============================================================

elif menu == "📊 Analisi":

    st.title(
        "📊 Analisi finanziaria"
    )

    if storico_excel.empty:

        st.warning(
            "Nessun dato storico disponibile."
        )

    else:

        st.subheader(
            "📈 Andamento delle spese"
        )

        mensile = (
            storico_excel
            .groupby(
                ["anno", "mese"],
                as_index=False
            )["importo"]
            .sum()
        )

        mensile["Mese"] = mensile.apply(
            lambda r:
            mese_label(
                int(r["anno"]),
                int(r["mese"])
            ),
            axis=1
        )

        fig, ax = plt.subplots(
            figsize=(10, 5)
        )

        ax.plot(
            mensile["Mese"],
            mensile["importo"],
            marker="o"
        )

        ax.set_ylabel(
            "Euro"
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

        st.markdown("---")

        st.subheader(
            "🍕 Distribuzione per categoria"
        )

        categorie = (
            storico_excel
            .groupby("categoria")["importo"]
            .sum()
            .sort_values(
                ascending=False
            )
        )

        fig2, ax2 = plt.subplots(
            figsize=(8, 7)
        )

        ax2.pie(
            categorie.values,
            labels=categorie.index,
            autopct="%1.1f%%",
            startangle=90
        )

        ax2.axis("equal")

        st.pyplot(
            fig2,
            clear_figure=True
        )

        st.dataframe(
            categorie
            .reset_index()
            .rename(
                columns={
                    "categoria": "Categoria",
                    "importo": "Totale"
                }
            ),
            use_container_width=True,
            hide_index=True
        )

        st.markdown("---")

        st.subheader(
            "📅 Tabella completa"
        )

        tabella = (
            mensile[
                [
                    "Mese",
                    "importo"
                ]
            ]
            .rename(
                columns={
                    "importo": "Spese"
                }
            )
        )

        st.dataframe(
            tabella,
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# NUOVA CASA
# ============================================================

elif menu == "🏡 Nuova Casa":

    st.title(
        "🏡 Bilancio Nuova Casa"
    )

    tab1, tab2 = st.tabs(
        [
            "🔴 Costi",
            "🟢 Coperture"
        ]
    )

    # --------------------------------------------------------
    # COSTI
    # --------------------------------------------------------

    with tab1:

        with st.form(
            "nuovo_costo_casa"
        ):

            voce = st.text_input(
                "Voce di costo"
            )

            importo = st.number_input(
                "Importo (€)",
                min_value=0.0,
                step=500.0
            )

            salva = st.form_submit_button(
                "➕ Aggiungi costo"
            )

            if salva and voce.strip():

                execute(
                    """
                    INSERT INTO costi_casa
                    (voce, importo)
                    VALUES (?, ?)
                    """,
                    (
                        voce,
                        importo
                    )
                )

                st.rerun()

        df = query_df(
            """
            SELECT
                id AS ID,
                voce AS Voce,
                importo AS Importo
            FROM costi_casa
            """
        )

        if df.empty:

            st.info(
                "Nessun costo inserito."
            )

        else:

            st.dataframe(
                df,
                use_container_width=True,
                hide_index=True
            )

            st.metric(
                "🔴 Totale costi",
                euro(df["Importo"].sum())
            )

    # --------------------------------------------------------
    # COPERTURE
    # --------------------------------------------------------

    with tab2:

        with st.form(
            "nuova_copertura"
        ):

            voce = st.text_input(
                "Fonte di copertura"
            )

            importo = st.number_input(
                "Importo (€)",
                min_value=0.0,
                step=500.0
            )

            salva = st.form_submit_button(
                "➕ Aggiungi copertura"
            )

            if salva and voce.strip():

                execute(
                    """
                    INSERT INTO entrate_casa
                    (voce, importo)
                    VALUES (?, ?)
                    """,
                    (
                        voce,
                        importo
                    )
                )

                st.rerun()

        df = query_df(
            """
            SELECT
                id AS ID,
                voce AS Voce,
                importo AS Importo
            FROM entrate_casa
            """
        )

        if df.empty:

            st.info(
                "Nessuna copertura inserita."
            )

        else:

            st.dataframe(
                df,
                use_container_width=True,
                hide_index=True
            )

            st.metric(
                "🟢 Totale coperture",
                euro(df["Importo"].sum())
            )

    # --------------------------------------------------------
    # RIEPILOGO
    # --------------------------------------------------------

    st.markdown("---")

    costi = query_df(
        """
        SELECT
            COALESCE(SUM(importo),0) AS totale
        FROM costi_casa
        """
    )

    coperture = query_df(
        """
        SELECT
            COALESCE(SUM(importo),0) AS totale
        FROM entrate_casa
        """
    )

    totale_costi = float(
        costi.iloc[0]["totale"]
    )

    totale_coperture = float(
        coperture.iloc[0]["totale"]
    )

    residuo = (
        totale_coperture -
        totale_costi
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Costi",
            euro(totale_costi)
        )

    with col2:

        st.metric(
            "Coperture",
            euro(totale_coperture)
        )

    with col3:

        st.metric(
            "Residuo",
            euro(residuo)
        )


# ============================================================
# MOBILI
# ============================================================

elif menu == "🪑 Mobili":

    st.title(
        "🪑 Mobili & Arredi"
    )

    st.info(
        "I mobili vengono letti direttamente dal foglio "
        "'Mobili' del file Excel."
    )

    if mobili_excel.empty:

        st.warning(
            "Nessun mobile trovato nel foglio Excel."
        )

    else:

        budget_mobili = st.number_input(
            "💰 Budget mobili",
            min_value=0.0,
            value=15000.0,
            step=500.0
        )

        totale_mobili = (
            mobili_excel["Costo"]
            .fillna(0)
            .sum()
        )

        residuo = (
            budget_mobili -
            totale_mobili
        )

        percentuale = (
            totale_mobili / budget_mobili
            if budget_mobili > 0
            else 0
        )

        percentuale = min(
            percentuale,
            1
        )

        col1, col2, col3 = st.columns(3)

        with col1:

            st.metric(
                "Budget",
                euro(budget_mobili)
            )

        with col2:

            st.metric(
                "Totale",
                euro(totale_mobili)
            )

        with col3:

            st.metric(
                "Residuo",
                euro(residuo)
            )

        st.progress(
            percentuale
        )

        st.markdown("---")

        st.dataframe(
            mobili_excel,
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# RENDERING
# ============================================================

elif menu == "🖼️ Rendering & Planimetrie":

    st.title(
        "🖼️ Rendering & Planimetrie"
    )

    st.write(
        "Carica qui rendering e planimetrie "
        "dei vari ambienti."
    )

    stanza = st.selectbox(
        "Ambiente",
        [
            "Cucina",
            "Salotto",
            "Ingresso",
            "Armadio ingresso",
            "Camera matrimoniale",
            "Camera bimbe",
            "Bagno piano terra",
            "Bagno ammezzato",
            "Bagno piano primo",
            "Mansarda",
            "Ripostiglio",
            "Giardino",
            "Altro"
        ]
    )

    immagini = st.file_uploader(
        "📷 Carica immagini",
        type=[
            "png",
            "jpg",
            "jpeg",
            "webp"
        ],
        accept_multiple_files=True
    )

    if immagini:

        st.subheader(
            f"📐 {stanza}"
        )

        for immagine in immagini:

            st.image(
                immagine,
                caption=immagine.name,
                use_container_width=True
            )


# ============================================================
# IMPORTA / ESPORTA
# ============================================================

elif menu == "📥 Importa / Esporta":

    st.title(
        "📥 Importa / Esporta"
    )

    # --------------------------------------------------------
    # STATO EXCEL
    # --------------------------------------------------------

    st.subheader(
        "📊 Collegamento Excel"
    )

    if EXCEL_FILE.exists():

        st.success(
            "🟢 File Excel collegato correttamente."
        )

        st.write(
            f"File: `{EXCEL_FILE.name}`"
        )

        st.write(
            f"Mesi storici letti: "
            f"**{len(storico_excel[['anno','mese']].drop_duplicates())}**"
        )

        st.write(
            f"Totale storico: "
            f"**{euro(storico_excel['importo'].sum())}**"
        )

    else:

        st.error(
            "File Excel non trovato."
        )

    # --------------------------------------------------------
    # ESPORTAZIONE
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "⬇️ Esporta dati"
    )

    spese_df = query_df(
        "SELECT * FROM spese"
    )

    entrate_df = query_df(
        "SELECT * FROM entrate"
    )

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        storico_excel.to_excel(
            writer,
            sheet_name="Storico Excel",
            index=False
        )

        spese_df.to_excel(
            writer,
            sheet_name="Spese App",
            index=False
        )

        entrate_df.to_excel(
            writer,
            sheet_name="Entrate App",
            index=False
        )

    st.download_button(
        "⬇️ Scarica report Excel",
        data=output.getvalue(),
        file_name="Report_finanze_famiglia.xlsx",
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

    # --------------------------------------------------------
    # EXCEL
    # --------------------------------------------------------

    st.subheader(
        "📁 File Excel"
    )

    if EXCEL_FILE.exists():

        st.success(
            "🟢 Excel trovato."
        )

        st.write(
            f"Percorso: `{EXCEL_FILE}`"
        )

        st.write(
            "Foglio storico: `Costi famiglia`"
        )

        st.write(
            "Foglio mobili: `Mobili`"
        )

    else:

        st.error(
            "🔴 Excel non trovato."
        )

    # --------------------------------------------------------
    # CONTEGGI
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "📊 Dati dell'app"
    )

    n_entrate = query_df(
        """
        SELECT COUNT(*) AS totale
        FROM entrate
        """
    ).iloc[0]["totale"]

    n_spese = query_df(
        """
        SELECT COUNT(*) AS totale
        FROM spese
        """
    ).iloc[0]["totale"]

    n_ricorrenti = query_df(
        """
        SELECT COUNT(*) AS totale
        FROM spese_ricorrenti
        """
    ).iloc[0]["totale"]

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Entrate",
            int(n_entrate)
        )

    with col2:

        st.metric(
            "Spese",
            int(n_spese)
        )

    with col3:

        st.metric(
            "Ricorrenti",
            int(n_ricorrenti)
        )

    # --------------------------------------------------------
    # RESET
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "⚠️ Cancella dati app"
    )

    st.warning(
        "Questa operazione cancella solamente "
        "i dati inseriti nell'app. "
        "NON modifica il file Excel."
    )

    conferma = st.checkbox(
        "Confermo di voler cancellare i dati dell'app."
    )

    if st.button(
        "🗑️ Cancella database app"
    ):

        if conferma:

            conn = get_connection()
            cursor = conn.cursor()

            tabelle = [
                "entrate",
                "spese",
                "spese_ricorrenti",
                "budget",
                "obiettivi",
                "spese_future",
                "costi_casa",
                "entrate_casa"
            ]

            for tabella in tabelle:

                cursor.execute(
                    f"DELETE FROM {tabella}"
                )

            conn.commit()
            conn.close()

            st.success(
                "Dati dell'app cancellati."
            )

            st.rerun()

        else:

            st.error(
                "Devi prima confermare."
            )


# ============================================================
# FOOTER
# ============================================================

st.sidebar.markdown("---")

st.sidebar.caption(
    "Gestione Casa & Finanze"
)

st.sidebar.caption(
    datetime.now().strftime(
        "Aggiornato il %d/%m/%Y alle %H:%M"
    )
                )
