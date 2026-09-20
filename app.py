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

EXCEL_FILE = BASE_DIR / "Spese casa -2.xlsx"
DB_FILE = BASE_DIR / "finanze_famiglia.db"


# ============================================================
# STILE
# ============================================================

st.markdown(
    """
    <style>

    .main {
        background-color: #f8fafc;
    }

    h1 {
        color: #1e293b;
        font-weight: 800;
    }

    h2, h3 {
        color: #334155;
    }

    div.stButton > button {
        border-radius: 10px;
        font-weight: 700;
        min-height: 42px;
    }

    [data-testid="stMetric"] {
        background: linear-gradient(
            135deg,
            #eff6ff 0%,
            #dbeafe 100%
        );
        padding: 14px;
        border-radius: 14px;
        border-left: 5px solid #2563eb;
    }

    .green-card {
        background: linear-gradient(
            135deg,
            #ecfdf5 0%,
            #d1fae5 100%
        );
        padding: 20px;
        border-radius: 15px;
        border-left: 6px solid #10b981;
    }

    .warning-card {
        background: linear-gradient(
            135deg,
            #fffbeb 0%,
            #fef3c7 100%
        );
        padding: 20px;
        border-radius: 15px;
        border-left: 6px solid #f59e0b;
    }

    .danger-card {
        background: linear-gradient(
            135deg,
            #fef2f2 0%,
            #fee2e2 100%
        );
        padding: 20px;
        border-radius: 15px;
        border-left: 6px solid #ef4444;
    }

    </style>
    """,
    unsafe_allow_html=True
)


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

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS entrate (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            descrizione TEXT NOT NULL,
            persona TEXT,
            categoria TEXT,
            importo REAL NOT NULL,
            data TEXT NOT NULL,
            ricorrente INTEGER DEFAULT 0
        )
        """
    )

    # --------------------------------------------------------
    # SPESE
    # --------------------------------------------------------

    cursor.execute(
        """
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
        """
    )

    # --------------------------------------------------------
    # SPESE RICORRENTI
    # --------------------------------------------------------

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS spese_ricorrenti (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            descrizione TEXT NOT NULL,
            categoria TEXT,
            importo REAL NOT NULL,
            giorno INTEGER DEFAULT 1,
            persona TEXT
        )
        """
    )

    # --------------------------------------------------------
    # BUDGET
    # --------------------------------------------------------

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS budget (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            categoria TEXT UNIQUE,
            importo REAL
        )
        """
    )

    # --------------------------------------------------------
    # OBIETTIVI
    # --------------------------------------------------------

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS obiettivi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT,
            obiettivo REAL,
            accumulato REAL DEFAULT 0,
            scadenza TEXT
        )
        """
    )

    # --------------------------------------------------------
    # SPESE FUTURE
    # --------------------------------------------------------

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS spese_future (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data TEXT,
            categoria TEXT,
            descrizione TEXT,
            importo REAL,
            note TEXT
        )
        """
    )

    # --------------------------------------------------------
    # COSTI NUOVA CASA
    # --------------------------------------------------------

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS costi_casa (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            voce TEXT,
            importo REAL
        )
        """
    )

    # --------------------------------------------------------
    # COPERTURE NUOVA CASA
    # --------------------------------------------------------

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS entrate_casa (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            voce TEXT,
            importo REAL
        )
        """
    )

    # --------------------------------------------------------
    # IMPOSTAZIONI
    # --------------------------------------------------------

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS impostazioni (
            chiave TEXT PRIMARY KEY,
            valore TEXT
        )
        """
    )

    conn.commit()
    conn.close()


init_database()


# ============================================================
# FUNZIONI DATABASE
# ============================================================

def query_df(query, params=()):

    conn = get_connection()

    try:
        return pd.read_sql_query(
            query,
            conn,
            params=params
        )
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


def get_setting(nome, default=None):

    df = query_df(
        """
        SELECT valore
        FROM impostazioni
        WHERE chiave = ?
        """,
        (nome,)
    )

    if df.empty:
        return default

    return df.iloc[0]["valore"]


def set_setting(nome, valore):

    execute(
        """
        INSERT INTO impostazioni
        (chiave, valore)
        VALUES (?, ?)
        ON CONFLICT(chiave)
        DO UPDATE SET valore = excluded.valore
        """,
        (
            nome,
            str(valore)
        )
    )


# ============================================================
# DATI STATICI
# ============================================================

CATEGORIE_SPESE = [
    "Costo alimentare mensile",
    "Tempo libero e viaggi",
    "Utenze",
    "Scuola e sport",
    "Trasporti e auto",
    "Prelievi contanti",
    "Casa e assicurazioni",
    "Shopping",
    "Farmacia e cura della persona",
    "Arredi e Extra Nuova Casa",
    "Altro"
]

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

PERSONE = [
    "Famiglia",
    "Io",
    "Partner"
]

CATEGORIE_ENTRATE = [
    "Stipendio",
    "Bonus",
    "Rimborso",
    "Affitto",
    "Altre entrate"
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

    # formato italiano:
    # 1.234,56
    if "," in testo:

        testo = (
            testo
            .replace(".", "")
            .replace(",", ".")
        )

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
# LETTURA EXCEL - COSTI FAMIGLIA
# ============================================================

@st.cache_data(ttl=30)
def leggi_storico_excel(percorso, ultima_modifica):

    if not Path(percorso).exists():
        return pd.DataFrame(
            columns=[
                "anno",
                "mese",
                "data",
                "categoria",
                "importo",
                "fonte"
            ]
        )

    try:

        df = pd.read_excel(
            percorso,
            sheet_name="Costi famiglia",
            header=None
        )

    except Exception:
        return pd.DataFrame(
            columns=[
                "anno",
                "mese",
                "data",
                "categoria",
                "importo",
                "fonte"
            ]
        )

    storico = []

    anno_corrente = None
    mese_corrente = None

    pattern = re.compile(
        r"SPESE\s+([A-ZÀ-Ù]+)\s+(\d{4})",
        re.IGNORECASE
    )

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

    for indice, riga in df.iterrows():

        if len(riga) == 0:
            continue

        prima_colonna = riga.iloc[0]

        if pd.isna(prima_colonna):
            continue

        testo = str(
            prima_colonna
        ).strip()

        # ----------------------------------------------------
        # RIGA MESE
        # ----------------------------------------------------

        match = pattern.search(testo)

        if match:

            nome_mese = (
                match.group(1)
                .upper()
                .strip()
            )

            anno = int(
                match.group(2)
            )

            if nome_mese in mesi:

                mese_corrente = mesi[
                    nome_mese
                ]

                anno_corrente = anno

            continue

        # ----------------------------------------------------
        # CATEGORIA
        # ----------------------------------------------------

        if (
            mese_corrente is None
            or anno_corrente is None
        ):
            continue

        categoria = testo

        if categoria.lower() in [
            "nan",
            "media",
            "totale",
            ""
        ]:
            continue

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
# LETTURA EXCEL - MOBILI
# ============================================================

@st.cache_data(ttl=30)
def leggi_mobili_excel(percorso, ultima_modifica):

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

    # La prima riga è il titolo
    # La seconda contiene:
    # Articolo / Costo / Riferimento / Negozio

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

        if len(riga) > 2 and not pd.isna(riga.iloc[2]):
            riferimento = str(
                riga.iloc[2]
            )

        if len(riga) > 3 and not pd.isna(riga.iloc[3]):
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

    excel_mtime = EXCEL_FILE.stat().st_mtime

    storico_excel = leggi_storico_excel(
        str(EXCEL_FILE),
        excel_mtime
    )

    mobili_excel = leggi_mobili_excel(
        str(EXCEL_FILE),
        excel_mtime
    )

else:

    storico_excel = pd.DataFrame(
        columns=[
            "anno",
            "mese",
            "data",
            "categoria",
            "importo",
            "fonte"
        ]
    )

    mobili_excel = pd.DataFrame()


# ============================================================
# FUNZIONI STORICO
# ============================================================

def storico_totale():

    if storico_excel.empty:
        return 0

    return float(
        storico_excel["importo"].sum()
    )


def storico_mese(anno, mese):

    if storico_excel.empty:
        return pd.DataFrame()

    return storico_excel[
        (storico_excel["anno"] == anno)
        &
        (storico_excel["mese"] == mese)
    ].copy()


def totale_storico_mese(anno, mese):

    df = storico_mese(
        anno,
        mese
    )

    if df.empty:
        return 0

    return float(
        df["importo"].sum()
    )


def totale_nuove_spese_mese(anno, mese):

    df = query_df(
        """
        SELECT COALESCE(SUM(importo), 0) AS totale
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


def totale_entrate_mese(anno, mese):

    df = query_df(
        """
        SELECT COALESCE(SUM(importo), 0) AS totale
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


def totale_spese_mese_completo(anno, mese):

    storico = totale_storico_mese(
        anno,
        mese
    )

    nuove = totale_nuove_spese_mese(
        anno,
        mese
    )

    return storico + nuove


# ============================================================
# GENERA LISTA MESI
# ============================================================

def genera_mesi():

    risultati = []

    anno_inizio = 2025
    mese_inizio = 1

    anno_fine = date.today().year + 1
    mese_fine = 12

    anno = anno_inizio
    mese = mese_inizio

    while (
        anno < anno_fine
        or (
            anno == anno_fine
            and mese <= mese_fine
        )
    ):

        risultati.append(
            (
                f"{mese_label(anno, mese)}",
                f"{anno}-{mese:02d}"
            )
        )

        mese += 1

        if mese == 13:
            mese = 1
            anno += 1

    return risultati


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
        "Il tuo cruscotto finanziario familiare."
    )

    st.markdown("---")

    oggi = date.today()

    # --------------------------------------------------------
    # MESE
    # --------------------------------------------------------

    mese_scelto = st.date_input(
        "📅 Seleziona il mese",
        value=oggi.replace(day=1)
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

    nuove_spese = totale_nuove_spese_mese(
        anno,
        mese
    )

    entrate = totale_entrate_mese(
        anno,
        mese
    )

    totale_spese = (
        storico +
        nuove_spese
    )

    risparmio = (
        entrate -
        totale_spese
    )

    percentuale = (
        risparmio / entrate * 100
        if entrate > 0
        else 0
    )

    # --------------------------------------------------------
    # METRICHE
    # --------------------------------------------------------

    c1, c2 = st.columns(2)

    with c1:

        st.metric(
            "💰 Entrate",
            euro(entrate)
        )

    with c2:

        st.metric(
            "💸 Spese",
            euro(totale_spese)
        )

    c3, c4 = st.columns(2)

    with c3:

        st.metric(
            "💚 Risparmio",
            euro(risparmio)
        )

    with c4:

        st.metric(
            "📈 Risparmio %",
            f"{percentuale:.1f}%"
        )

    # --------------------------------------------------------
    # DETTAGLIO ORIGINE
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        f"📌 Spese {mese_label(anno, mese)}"
    )

    c1, c2 = st.columns(2)

    with c1:

        st.metric(
            "📊 Storico Excel",
            euro(storico)
        )

    with c2:

        st.metric(
            "✏️ Inserite nell'app",
            euro(nuove_spese)
        )

    # --------------------------------------------------------
    # QUANTO POSSIAMO SPENDERE
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "💳 Quanto possiamo spendere?"
    )

    ricorrenti_df = query_df(
        """
        SELECT COALESCE(SUM(importo),0) AS totale
        FROM spese_ricorrenti
        """
    )

    spese_ricorrenti = float(
        ricorrenti_df.iloc[0]["totale"]
    )

    future_df = query_df(
        """
        SELECT COALESCE(SUM(importo),0) AS totale
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

    # Per un mese già storico non aggiungiamo
    # le ricorrenti alla spesa storica.
    if (
        storico > 0
        and mese != oggi.month
        and anno != oggi.year
    ):
        spese_previste = 0

    else:

        spese_previste = (
            spese_ricorrenti +
            spese_future
        )

    disponibilita = (
        entrate -
        totale_spese -
        spese_previste
    )

    if disponibilita >= 0:

        st.markdown(
            f"""
            <div class="green-card">
                <h3>💚 Disponibilità stimata</h3>
                <h1>{euro(disponibilita)}</h1>
                <p>
                    Dopo le spese già registrate e
                    quelle previste.
                </p>
            </div>
            """,
            unsafe_allow_html=True
        )

    else:

        st.markdown(
            f"""
            <div class="danger-card">
                <h3>⚠️ Attenzione</h3>
                <h1>{euro(disponibilita)}</h1>
                <p>
                    Le spese previste superano
                    la disponibilità.
                </p>
            </div>
            """,
            unsafe_allow_html=True
        )

    # --------------------------------------------------------
    # ULTIMI MESI
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "📈 Andamento ultimi mesi"
    )

    righe = []

    oggi = date.today()

    for i in range(12):

        mese_temp = oggi.month - i
        anno_temp = oggi.year

        while mese_temp <= 0:

            mese_temp += 12
            anno_temp -= 1

        spese_temp = totale_spese_mese_completo(
            anno_temp,
            mese_temp
        )

        righe.append(
            {
                "Mese": mese_label(
                    anno_temp,
                    mese_temp
                ),
                "Spese": spese_temp
            }
        )

    df_andamento = pd.DataFrame(
        righe
    ).iloc[::-1]

    fig, ax = plt.subplots(
        figsize=(10, 4)
    )

    ax.plot(
        df_andamento["Mese"],
        df_andamento["Spese"],
        marker="o"
    )

    ax.set_ylabel(
        "Euro"
    )

    ax.set_title(
        "Spese mensili"
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
        "📜 Storico spese"
    )

    st.info(
        "Questa sezione legge direttamente "
        "il foglio 'Costi famiglia' del file Excel."
    )

    if storico_excel.empty:

        st.warning(
            "Nessun dato storico trovato."
        )

    else:

        # ----------------------------------------------------
        # FILTRI
        # ----------------------------------------------------

        anni = sorted(
            storico_excel["anno"]
            .dropna()
            .unique(),
            reverse=True
        )

        anno = st.selectbox(
            "Anno",
            anni
        )

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

        # ----------------------------------------------------
        # TOTALE
        # ----------------------------------------------------

        totale = (
            dati["importo"].sum()
            if not dati.empty
            else 0
        )

        st.metric(
            f"Totale {mese_label(anno, mese)}",
            euro(totale)
        )

        # ----------------------------------------------------
        # GRAFICO
        # ----------------------------------------------------

        if not dati.empty:

            fig, ax = plt.subplots(
                figsize=(9, 6)
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
            # BAR CHART
            # ------------------------------------------------

            fig2, ax2 = plt.subplots(
                figsize=(10, 5)
            )

            dati_ordinati = dati.sort_values(
                "importo",
                ascending=True
            )

            ax2.barh(
                dati_ordinati["categoria"],
                dati_ordinati["importo"]
            )

            ax2.set_xlabel(
                "Euro"
            )

            ax2.set_title(
                f"Spese {mese_label(anno, mese)}"
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
        "📋 Entrate registrate nell'app"
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
        "Le spese inserite qui vengono salvate nel database "
        "e sono separate dallo storico Excel."
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
            "Note"
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
            "Nessuna nuova spesa."
        )

    else:

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True
        )

        st.metric(
            "Totale",
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
        "Mutuo, assicurazioni, abbonamenti, scuola, "
        "utenze e tutte le spese che si ripetono."
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

    if not df.empty:

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True
        )

        st.metric(
            "💸 Spese fisse mensili",
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

    else:

        st.info(
            "Nessuna spesa ricorrente."
        )


# ============================================================
# BUDGET
# ============================================================

elif menu == "📅 Budget":

    st.title(
        "📅 Budget familiare"
    )

    st.write(
        "Il budget viene confrontato con le spese "
        "del mese selezionato."
    )

    mese_budget = st.date_input(
        "Mese",
        value=date.today().replace(day=1)
    )

    anno = mese_budget.year
    mese = mese_budget.month

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
            (categoria, importo)
            VALUES (?, ?)
            ON CONFLICT(categoria)
            DO UPDATE SET importo=excluded.importo
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

        # Spese nuove inserite nell'app
        df_nuove = query_df(
            """
            SELECT COALESCE(SUM(importo),0) AS totale
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
            df_nuove.iloc[0]["totale"]
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
            placeholder="Es. Vacanza, fondo emergenza..."
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
                accumulato /
                obiettivo
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
                f"**{euro(accumulato)}** "
                f"su **{euro(obiettivo)}**"
            )

            st.caption(
                f"Scadenza: {row['scadenza']}"
            )

            if st.button(
                f"🗑️ Elimina {row['nome']}",
                key=f"delete_goal_{row['id']}"
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
        "🔮 Previsioni finanziarie"
    )

    st.subheader(
        "📈 Previsione basata sullo storico Excel"
    )

    if storico_excel.empty:

        st.info(
            "Non sono disponibili dati storici."
        )

    else:

        # Totale per mese
        mensile = (
            storico_excel
            .groupby(
                ["anno", "mese"],
                as_index=False
            )["importo"]
            .sum()
        )

        mensile["label"] = (
            mensile.apply(
                lambda r:
                mese_label(
                    int(r["anno"]),
                    int(r["mese"])
                ),
                axis=1
            )
        )

        st.dataframe(
            mensile[
                [
                    "label",
                    "importo"
                ]
            ].rename(
                columns={
                    "label": "Mese",
                    "importo": "Spese"
                }
            ),
            use_container_width=True,
            hide_index=True
        )

        media_3 = (
            mensile
            .tail(3)["importo"]
            .mean()
        )

        media_6 = (
            mensile
            .tail(6)["importo"]
            .mean()
        )

        media_12 = (
            mensile
            .tail(12)["importo"]
            .mean()
        )

        c1, c2, c3 = st.columns(3)

        with c1:

            st.metric(
                "Media 3 mesi",
                euro(media_3)
            )

        with c2:

            st.metric(
                "Media 6 mesi",
                euro(media_6)
            )

        with c3:

            st.metric(
                "Media 12 mesi",
                euro(media_12)
            )

        st.markdown("---")

        st.subheader(
            "🔮 Stima prossimo mese"
        )

        st.info(
            f"Se il prossimo mese fosse in linea "
            f"con la media degli ultimi 3 mesi, "
            f"la spesa stimata sarebbe circa "
            f"**{euro(media_3)}**."
        )

        st.subheader(
            "📊 Categorie che incidono di più"
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

    st.subheader(
        "📈 Andamento storico"
    )

    if storico_excel.empty:

        st.info(
            "Nessun dato Excel disponibile."
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

        mensile["label"] = (
            mensile.apply(
                lambda r:
                mese_label(
                    int(r["anno"]),
                    int(r["mese"])
                ),
                axis=1
            )
        )

        fig, ax = plt.subplots(
            figsize=(10, 5)
        )

        ax.plot(
            mensile["label"],
            mensile["importo"],
            marker="o"
        )

        ax.set_title(
            "Andamento spese mensili"
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
        "🍝 Distribuzione per categoria"
    )

    if not storico_excel.empty:

        categoria_totali = (
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
            categoria_totali.values,
            labels=categoria_totali.index,
            autopct="%1.1f%%",
            startangle=90
        )

        ax2.axis("equal")

        st.pyplot(
            fig2,
            clear_figure=True
        )

        st.dataframe(
            categoria_totali
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
        "🆕 Spese inserite nell'app"
    )

    nuove = query_df(
        """
        SELECT
            strftime('%Y-%m',data) AS mese,
            SUM(importo) AS totale
        FROM spese
        GROUP BY mese
        ORDER BY mese
        """
    )

    if not nuove.empty:

        st.dataframe(
            nuove.rename(
                columns={
                    "mese": "Mese",
                    "totale": "Totale"
                }
            ),
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
                "Voce"
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

    st.markdown("---")

    costi = query_df(
        """
        SELECT COALESCE(SUM(importo),0) AS totale
        FROM costi_casa
        """
    )

    coperture = query_df(
        """
        SELECT COALESCE(SUM(importo),0) AS totale
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

    c1, c2, c3 = st.columns(3)

    with c1:

        st.metric(
            "Costi",
            euro(totale_costi)
        )

    with c2:

        st.metric(
            "Coperture",
            euro(totale_coperture)
        )

    with c3:

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

    if mobili_excel.empty:

        st.warning(
            "Il foglio 'Mobili' non contiene dati "
            "oppure il file Excel non è disponibile."
        )

    else:

        st.info(
            "Questi dati vengono letti direttamente "
            "dal foglio 'Mobili' del file Excel."
        )

        budget = st.number_input(
            "💰 Budget mobili (€)",
            min_value=0.0,
            value=15000.0,
            step=500.0
        )

        totale = (
            mobili_excel["Costo"]
            .fillna(0)
            .sum()
        )

        residuo = (
            budget -
            totale
        )

        c1, c2, c3 = st.columns(3)

        with c1:

            st.metric(
                "Budget",
                euro(budget)
            )

        with c2:

            st.metric(
                "Mobili",
                euro(totale)
            )

        with c3:

            st.metric(
                "Residuo",
                euro(residuo)
            )

        percentuale = (
            min(
                totale / budget,
                1
            )
            if budget > 0
            else 0
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

        st.markdown("---")

        st.subheader(
            "💰 Mobili più costosi"
        )

        top = mobili_excel.sort_values(
            "Costo",
            ascending=False
        ).head(10)

        st.dataframe(
            top,
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

    st.info(
        "Puoi caricare qui rendering, planimetrie "
        "e immagini di riferimento."
    )

    stanza = st.selectbox(
        "Ambiente",
        [
            "Cucina",
            "Ripostiglio",
            "Ingresso",
            "Armadio ingresso",
            "Bagno piano terra",
            "Salotto",
            "Bagno ammezzato",
            "Camera matrimoniale",
            "Camera bimbe",
            "Bagno piano primo",
            "Mansarda",
            "Altro"
        ]
    )

    file = st.file_uploader(
        "📷 Carica immagini",
        type=[
            "png",
            "jpg",
            "jpeg",
            "webp"
        ],
        accept_multiple_files=True
    )

    if file:

        st.subheader(
            f"📐 {stanza}"
        )

        for immagine in file:

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
    # EXCEL
    # --------------------------------------------------------

    st.subheader(
        "📊 File Excel collegato"
    )

    if EXCEL_FILE.exists():

        st.success(
            f"🟢 {EXCEL_FILE.name}"
        )

        st.write(
            f"Percorso: `{EXCEL_FILE}`"
        )

        st.write(
            f"Totale storico Excel: "
            f"**{euro(storico_totale())}**"
        )

    else:

        st.error(
            "File Excel non trovato."
        )

    # --------------------------------------------------------
    # ESPORTA SPESE
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "⬇️ Esporta dati inseriti nell'app"
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

        spese_df.to_excel(
            writer,
            sheet_name="Spese app",
            index=False
        )

        entrate_df.to_excel(
            writer,
            sheet_name="Entrate app",
            index=False
        )

        storico_excel.to_excel(
            writer,
            sheet_name="Storico Excel",
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

    st.subheader(
        "📁 Collegamento Excel"
    )

    if EXCEL_FILE.exists():

        st.success(
            "🟢 File Excel correttamente collegato."
        )

        st.write(
            f"**File:** {EXCEL_FILE.name}"
        )

        st.write(
            f"**Foglio storico:** Costi famiglia"
        )

        st.write(
            f"**Mesi letti:** "
            f"{storico_excel[['anno','mese']].drop_duplicates().shape[0]}"
        )

        st.write(
            f"**Totale storico:** "
            f"{euro(storico_totale())}"
        )

    else:

        st.error(
            "Il file 'Spese casa -2.xlsx' "
            "non è presente nella cartella dell'app."
        )

    st.markdown("---")

    st.subheader(
        "📊 Riepilogo database"
    )

    c1, c2 = st.columns(2)

    entrate_count = query_df(
        "SELECT COUNT(*) AS totale FROM entrate"
    ).iloc[0]["totale"]

    spese_count = query_df(
        "SELECT COUNT(*) AS totale FROM spese"
    ).iloc[0]["totale"]

    with c1:

        st.metric(
            "Entrate inserite",
            int(entrate_count)
        )

    with c2:

        st.metric(
            "Spese inserite",
            int(spese_count)
        )

    st.markdown("---")

    st.subheader(
        "🗑️ Database"
    )

    st.warning(
        "Questa operazione elimina solamente "
        "i dati inseriti nell'app. "
        "Il file Excel NON viene modificato."
    )

    conferma = st.checkbox(
        "Confermo di voler eliminare i dati dell'app"
    )

    if st.button(
        "⚠️ Cancella database app"
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
                "Database dell'app cancellato."
            )

            st.rerun()

        else:

            st.error(
                "Devi confermare prima dell'eliminazione."
            )


# ============================================================
# FINE
# ============================================================

st.sidebar.markdown("---")

st.sidebar.caption(
    f"Ultimo aggiornamento: "
    f"{datetime.now().strftime('%d/%m/%Y %H:%M')}"
        )
h2, h3 {
    color: #334155;
}

div.stButton > button {
    border-radius: 10px;
    font-weight: 700;
    border: none;
    min-height: 45px;
}

[data-testid="stMetric"] {
    background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%);
    padding: 15px;
    border-radius: 14px;
    border-left: 5px solid #2563eb;
}

.card {
    background: white;
    padding: 18px;
    border-radius: 15px;
    margin-bottom: 15px;
    border: 1px solid #e2e8f0;
}

.green-card {
    background: linear-gradient(135deg, #ecfdf5 0%, #d1fae5 100%);
    padding: 20px;
    border-radius: 15px;
    border-left: 6px solid #10b981;
}

.warning-card {
    background: linear-gradient(135deg, #fffbeb 0%, #fef3c7 100%);
    padding: 20px;
    border-radius: 15px;
    border-left: 6px solid #f59e0b;
}

.danger-card {
    background: linear-gradient(135deg, #fef2f2 0%, #fee2e2 100%);
    padding: 20px;
    border-radius: 15px;
    border-left: 6px solid #ef4444;
}

</style>
""", unsafe_allow_html=True)


# ============================================================
# DATABASE
# ============================================================

def get_connection():
    return sqlite3.connect(DB_FILE)


def init_database():

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS entrate (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            descrizione TEXT,
            persona TEXT,
            categoria TEXT,
            importo REAL,
            data TEXT,
            ricorrente INTEGER DEFAULT 0
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS spese (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            descrizione TEXT,
            persona TEXT,
            categoria TEXT,
            importo REAL,
            data TEXT,
            ricorrente INTEGER DEFAULT 0,
            note TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS spese_ricorrenti (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            descrizione TEXT,
            categoria TEXT,
            importo REAL,
            giorno INTEGER,
            persona TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS budget (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            categoria TEXT UNIQUE,
            importo REAL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS obiettivi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT,
            obiettivo REAL,
            accumulato REAL DEFAULT 0,
            scadenza TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS costi_casa (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            voce TEXT,
            importo REAL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS entrate_casa (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            voce TEXT,
            importo REAL
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS mobili (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            articolo TEXT,
            costo REAL,
            negozio TEXT,
            acquistato INTEGER DEFAULT 0
        )
    """)

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

    conn.commit()
    conn.close()


init_database()


# ============================================================
# FUNZIONI DATABASE
# ============================================================

def query_df(query, params=()):
    conn = get_connection()
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df


def execute(query, params=()):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(query, params)
    conn.commit()
    conn.close()


# ============================================================
# DATI
# ============================================================

CATEGORIE_SPESE = [
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

PERSONE = [
    "Famiglia",
    "Io",
    "Partner"
]

CATEGORIE_ENTRATE = [
    "Stipendio",
    "Bonus",
    "Rimborso",
    "Affitto",
    "Altre entrate"
]


# ============================================================
# FUNZIONI DI CALCOLO
# ============================================================

def totale_entrate_mese(anno, mese):

    df = query_df(
        """
        SELECT COALESCE(SUM(importo), 0) AS totale
        FROM entrate
        WHERE strftime('%Y', data) = ?
        AND strftime('%m', data) = ?
        """,
        (str(anno), f"{mese:02d}")
    )

    return float(df.iloc[0]["totale"])


def totale_spese_mese(anno, mese):

    df = query_df(
        """
        SELECT COALESCE(SUM(importo), 0) AS totale
        FROM spese
        WHERE strftime('%Y', data) = ?
        AND strftime('%m', data) = ?
        """,
        (str(anno), f"{mese:02d}")
    )

    return float(df.iloc[0]["totale"])


def euro(valore):
    return f"{valore:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.title("🏡 Gestione Casa")

menu = st.sidebar.radio(
    "MENU",
    [
        "🏠 Home",
        "💰 Entrate",
        "💸 Spese",
        "🔁 Spese ricorrenti",
        "📅 Budget",
        "🎯 Obiettivi",
        "📊 Analisi",
        "🔮 Previsioni",
        "🏡 Nuova Casa",
        "🪑 Mobili & Arredi",
        "🖼️ Rendering & Planimetrie",
        "📥 Importa / Esporta",
        "⚙️ Impostazioni"
    ]
)

st.sidebar.markdown("---")
st.sidebar.caption("Gestione Casa & Finanze")
st.sidebar.caption("Versione 2.0")


# ============================================================
# HOME
# ============================================================

if menu == "🏠 Home":

    oggi = date.today()

    entrate = totale_entrate_mese(oggi.year, oggi.month)
    spese = totale_spese_mese(oggi.year, oggi.month)
    risparmio = entrate - spese

    st.title("🏡 Gestione Casa & Finanze")
    st.write(
        f"Situazione finanziaria di **{oggi.strftime('%B %Y')}**"
    )

    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:
        st.metric(
            "💰 Entrate",
            euro(entrate)
        )

    with col2:
        st.metric(
            "💸 Spese",
            euro(spese)
        )

    col3, col4 = st.columns(2)

    with col3:
        st.metric(
            "💚 Risparmio",
            euro(risparmio),
            delta=euro(risparmio)
        )

    percentuale = (
        risparmio / entrate * 100
        if entrate > 0 else 0
    )

    with col4:
        st.metric(
            "📈 Risparmio %",
            f"{percentuale:.1f}%"
        )

    st.markdown("---")

    # DISPONIBILITA'

    st.subheader("💳 Quanto possiamo spendere?")

    budget_df = query_df(
        "SELECT * FROM budget"
    )

    budget_totale = (
        budget_df["importo"].sum()
        if not budget_df.empty else 0
    )

    disponibilita = entrate - spese

    if disponibilita >= 0:

        st.markdown(
            f"""
            <div class="green-card">
                <h3>💚 Disponibilità attuale</h3>
                <h1>{euro(disponibilita)}</h1>
                <p>Risparmio potenziale del mese.</p>
            </div>
            """,
            unsafe_allow_html=True
        )

    else:

        st.markdown(
            f"""
            <div class="danger-card">
                <h3>⚠️ Attenzione</h3>
                <h1>{euro(disponibilita)}</h1>
                <p>Le spese superano le entrate.</p>
            </div>
            """,
            unsafe_allow_html=True
        )

    # OBIETTIVI

    st.markdown("---")
    st.subheader("🎯 Obiettivi")

    obiettivi = query_df(
        "SELECT * FROM obiettivi"
    )

    if obiettivi.empty:

        st.info(
            "Non hai ancora creato obiettivi di risparmio."
        )

    else:

        for _, row in obiettivi.iterrows():

            percent = (
                row["accumulato"] /
                row["obiettivo"] * 100
                if row["obiettivo"] > 0 else 0
            )

            percent = min(percent, 100)

            st.write(
                f"**{row['nome']}** — "
                f"{euro(row['accumulato'])} / "
                f"{euro(row['obiettivo'])}"
            )

            st.progress(percent / 100)

            st.caption(
                f"{percent:.1f}% completato"
            )


# ============================================================
# ENTRATE
# ============================================================

elif menu == "💰 Entrate":

    st.title("💰 Entrate")

    st.subheader("➕ Inserisci nuova entrata")

    with st.form("nuova_entrata"):

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

            if descrizione.strip() and importo > 0:

                execute(
                    """
                    INSERT INTO entrate
                    (descrizione, persona, categoria,
                     importo, data, ricorrente)
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

                st.success("Entrata salvata!")
                st.rerun()

            else:

                st.warning(
                    "Inserisci descrizione e importo."
                )

    st.markdown("---")

    st.subheader("📋 Storico entrate")

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

        st.info("Nessuna entrata registrata.")

    else:

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True
        )

        id_delete = st.number_input(
            "ID entrata da eliminare",
            min_value=0,
            step=1
        )

        if st.button("🗑️ Elimina entrata"):

            if id_delete > 0:

                execute(
                    "DELETE FROM entrate WHERE id = ?",
                    (id_delete,)
                )

                st.success("Entrata eliminata.")
                st.rerun()


# ============================================================
# SPESE
# ============================================================

elif menu == "💸 Spese":

    st.title("💸 Spese")

    with st.form("nuova_spesa"):

        descrizione = st.text_input(
            "Descrizione",
            placeholder="Es. Spesa supermercato"
        )

        col1, col2 = st.columns(2)

        with col1:

            categoria = st.selectbox(
                "Categoria",
                CATEGORIE_SPESE
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

            if descrizione.strip() and importo > 0:

                execute(
                    """
                    INSERT INTO spese
                    (descrizione, persona, categoria,
                     importo, data, ricorrente, note)
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

                st.success("Spesa salvata!")
                st.rerun()

            else:

                st.warning(
                    "Inserisci descrizione e importo."
                )

    st.markdown("---")

    st.subheader("📋 Storico spese")

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

        st.info("Nessuna spesa registrata.")

    else:

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True
        )

        totale = df["Importo"].sum()

        st.metric(
            "Totale spese registrate",
            euro(totale)
        )

        id_delete = st.number_input(
            "ID spesa da eliminare",
            min_value=0,
            step=1
        )

        if st.button("🗑️ Elimina spesa"):

            if id_delete > 0:

                execute(
                    "DELETE FROM spese WHERE id = ?",
                    (id_delete,)
                )

                st.success("Spesa eliminata.")
                st.rerun()


# ============================================================
# SPESE RICORRENTI
# ============================================================

elif menu == "🔁 Spese ricorrenti":

    st.title("🔁 Spese ricorrenti")

    st.write(
        "Gestisci le spese che si ripetono ogni mese."
    )

    with st.form("spesa_ricorrente"):

        descrizione = st.text_input(
            "Descrizione",
            placeholder="Es. Mutuo"
        )

        categoria = st.selectbox(
            "Categoria",
            CATEGORIE_SPESE
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
            "➕ Aggiungi spesa ricorrente"
        )

        if salva:

            if descrizione.strip() and importo > 0:

                execute(
                    """
                    INSERT INTO spese_ricorrenti
                    (descrizione, categoria,
                     importo, giorno, persona)
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
                    "Spesa ricorrente aggiunta!"
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

    if not df.empty:

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True
        )

        st.metric(
            "💸 Totale spese fisse mensili",
            euro(df["Importo"].sum())
        )

    else:

        st.info(
            "Nessuna spesa ricorrente."
        )


# ============================================================
# BUDGET
# ============================================================

elif menu == "📅 Budget":

    st.title("📅 Budget familiare")

    st.write(
        "Imposta quanto vuoi spendere ogni mese per ogni categoria."
    )

    categoria = st.selectbox(
        "Categoria",
        CATEGORIE_SPESE
    )

    importo = st.number_input(
        "Budget mensile (€)",
        min_value=0.0,
        step=50.0
    )

    if st.button("💾 Salva budget"):

        execute(
            """
            INSERT INTO budget (categoria, importo)
            VALUES (?, ?)
            ON CONFLICT(categoria)
            DO UPDATE SET importo = excluded.importo
            """,
            (
                categoria,
                importo
            )
        )

        st.success("Budget aggiornato.")
        st.rerun()

    st.markdown("---")

    oggi = date.today()

    df_budget = query_df(
        """
        SELECT categoria, importo
        FROM budget
        """
    )

    if not df_budget.empty:

        righe = []

        for _, row in df_budget.iterrows():

            cat = row["categoria"]
            budget = float(row["importo"])

            df_spesa = query_df(
                """
                SELECT COALESCE(SUM(importo),0) AS totale
                FROM spese
                WHERE categoria = ?
                AND strftime('%Y',data)=?
                AND strftime('%m',data)=?
                """,
                (
                    cat,
                    str(oggi.year),
                    f"{oggi.month:02d}"
                )
            )

            speso = float(
                df_spesa.iloc[0]["totale"]
            )

            residuo = budget - speso

            righe.append({
                "Categoria": cat,
                "Budget": budget,
                "Speso": speso,
                "Residuo": residuo
            })

        risultato = pd.DataFrame(righe)

        st.dataframe(
            risultato,
            use_container_width=True,
            hide_index=True
        )

        for _, row in risultato.iterrows():

            if row["Residuo"] < 0:

                st.error(
                    f"⚠️ {row['Categoria']}: "
                    f"superato di {euro(abs(row['Residuo']))}"
                )


# ============================================================
# OBIETTIVI
# ============================================================

elif menu == "🎯 Obiettivi":

    st.title("🎯 Obiettivi di risparmio")

    with st.form("nuovo_obiettivo"):

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
            "Quanto hai già accumulato? (€)",
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

            if nome.strip() and obiettivo > 0:

                execute(
                    """
                    INSERT INTO obiettivi
                    (nome, obiettivo, accumulato, scadenza)
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
                    "Obiettivo creato!"
                )

                st.rerun()

    st.markdown("---")

    df = query_df(
        "SELECT * FROM obiettivi"
    )

    for _, row in df.iterrows():

        percent = (
            row["accumulato"] /
            row["obiettivo"]
            if row["obiettivo"] > 0 else 0
        )

        percent = min(percent, 1)

        st.subheader(
            f"🎯 {row['nome']}"
        )

        st.progress(percent)

        st.write(
            f"{euro(row['accumulato'])} / "
            f"{euro(row['obiettivo'])}"
        )

        mancante = max(
            row["obiettivo"] -
            row["accumulato"],
            0
        )

        st.caption(
            f"Mancano {euro(mancante)}"
        )

        if st.button(
            f"🗑️ Elimina {row['nome']}",
            key=f"delete_goal_{row['id']}"
        ):

            execute(
                "DELETE FROM obiettivi WHERE id=?",
                (row["id"],)
            )

            st.rerun()


# ============================================================
# ANALISI
# ============================================================

elif menu == "📊 Analisi":

    st.title("📊 Analisi finanziaria")

    anno = st.number_input(
        "Anno",
        min_value=2020,
        max_value=2100,
        value=date.today().year
    )

    df_spese = query_df(
        """
        SELECT
            strftime('%m', data) AS mese,
            SUM(importo) AS totale
        FROM spese
        WHERE strftime('%Y', data)=?
        GROUP BY mese
        ORDER BY mese
        """,
        (str(anno),)
    )

    if not df_spese.empty:

        df_spese["mese"] = df_spese["mese"].astype(int)

        mesi = pd.DataFrame({
            "mese": range(1, 13)
        })

        df_grafico = mesi.merge(
            df_spese,
            on="mese",
            how="left"
        )

        df_grafico["totale"] = (
            df_grafico["totale"]
            .fillna(0)
        )

        nomi_mesi = [
            "Gen", "Feb", "Mar", "Apr",
            "Mag", "Giu", "Lug", "Ago",
            "Set", "Ott", "Nov", "Dic"
        ]

        fig, ax = plt.subplots(
            figsize=(10, 5)
        )

        ax.plot(
            nomi_mesi,
            df_grafico["totale"],
            marker="o"
        )

        ax.set_title(
            "Andamento delle spese"
        )

        ax.set_ylabel(
            "Euro"
        )

        ax.grid(
            alpha=0.2
        )

        st.pyplot(fig)

    else:

        st.info(
            "Non ci sono ancora dati sufficienti."
        )

    st.markdown("---")

    st.subheader(
        "🍕 Distribuzione spese per categoria"
    )

    df_cat = query_df(
        """
        SELECT
            categoria AS Categoria,
            SUM(importo) AS Totale
        FROM spese
        WHERE strftime('%Y',data)=?
        GROUP BY categoria
        ORDER BY Totale DESC
        """,
        (str(anno),)
    )

    if not df_cat.empty:

        fig, ax = plt.subplots(
            figsize=(8, 7)
        )

        ax.pie(
            df_cat["Totale"],
            labels=df_cat["Categoria"],
            autopct="%1.1f%%",
            startangle=140
        )

        ax.axis("equal")

        st.pyplot(fig)

        st.dataframe(
            df_cat,
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# PREVISIONI
# ============================================================

elif menu == "🔮 Previsioni":

    st.title("🔮 Previsione finanziaria")

    oggi = date.today()

    entrate = totale_entrate_mese(
        oggi.year,
        oggi.month
    )

    spese = totale_spese_mese(
        oggi.year,
        oggi.month
    )

    ricorrenti = query_df(
        "SELECT COALESCE(SUM(importo),0) AS totale FROM spese_ricorrenti"
    )

    spese_fisse = (
        float(ricorrenti.iloc[0]["totale"])
        if not ricorrenti.empty else 0
    )

    media_spese_df = query_df(
        """
        SELECT AVG(totale) AS media
        FROM (
            SELECT
                strftime('%Y-%m',data),
                SUM(importo) AS totale
            FROM spese
            GROUP BY strftime('%Y-%m',data)
            ORDER BY strftime('%Y-%m',data) DESC
            LIMIT 6
        )
        """
    )

    media_spese = (
        float(media_spese_df.iloc[0]["media"])
        if pd.notna(media_spese_df.iloc[0]["media"])
        else 0
    )

    st.subheader(
        "📌 Stima fine mese"
    )

    st.metric(
        "Entrate attuali",
        euro(entrate)
    )

    st.metric(
        "Spese attuali",
        euro(spese)
    )

    st.metric(
        "Spese fisse mensili",
        euro(spese_fisse)
    )

    st.markdown("---")

    st.subheader(
        "📈 Media spese ultimi mesi"
    )

    st.metric(
        "Media mensile",
        euro(media_spese)
    )

    previsione = entrate - max(
        spese,
        media_spese
    )

    if previsione >= 0:

        st.markdown(
            f"""
            <div class="green-card">
                <h3>💚 Risparmio potenziale</h3>
                <h1>{euro(previsione)}</h1>
                <p>Stima basata sulle spese attuali.</p>
            </div>
            """,
            unsafe_allow_html=True
        )

    else:

        st.markdown(
            f"""
            <div class="danger-card">
                <h3>⚠️ Possibile disavanzo</h3>
                <h1>{euro(previsione)}</h1>
            </div>
            """,
            unsafe_allow_html=True
        )


# ============================================================
# NUOVA CASA
# ============================================================

elif menu == "🏡 Nuova Casa":

    st.title("🏡 Piano finanziario nuova casa")

    tab1, tab2 = st.tabs(
        [
            "🔴 Costi",
            "🟢 Coperture"
        ]
    )

    with tab1:

        with st.form("costo_casa"):

            voce = st.text_input(
                "Voce di spesa"
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
            "SELECT id AS ID, voce AS Voce, importo AS Importo FROM costi_casa"
        )

        if not df.empty:

            st.dataframe(
                df,
                use_container_width=True,
                hide_index=True
            )

            st.metric(
                "🔴 Totale costi",
                euro(df["Importo"].sum())
            )

    with tab2:

        with st.form("entrata_casa"):

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
            "SELECT id AS ID, voce AS Voce, importo AS Importo FROM entrate_casa"
        )

        if not df.empty:

            st.dataframe(
                df,
                use_container_width=True,
                hide_index=True
            )

            st.metric(
                "🟢 Totale coperture",
                euro(df["Importo"].sum())
            )

    st.markdown("---")

    costi = query_df(
        "SELECT COALESCE(SUM(importo),0) AS totale FROM costi_casa"
    )

    coperture = query_df(
        "SELECT COALESCE(SUM(importo),0) AS totale FROM entrate_casa"
    )

    totale_costi = float(
        costi.iloc[0]["totale"]
    )

    totale_coperture = float(
        coperture.iloc[0]["totale"]
    )

    differenza = (
        totale_coperture -
        totale_costi
    )

    st.subheader(
        "📊 Situazione progetto"
    )

    c1, c2, c3 = st.columns(3)

    with c1:
        st.metric(
            "Costi",
            euro(totale_costi)
        )

    with c2:
        st.metric(
            "Coperture",
            euro(totale_coperture)
        )

    with c3:
        st.metric(
            "Residuo",
            euro(differenza)
        )


# ============================================================
# MOBILI
# ============================================================

elif menu == "🪑 Mobili & Arredi":

    st.title("🪑 Mobili & Arredi")

    budget_arredi = st.number_input(
        "💰 Budget totale arredi (€)",
        min_value=0.0,
        value=15000.0,
        step=500.0
    )

    with st.form("nuovo_mobile"):

        articolo = st.text_input(
            "Articolo"
        )

        costo = st.number_input(
            "Costo (€)",
            min_value=0.0,
            step=50.0
        )

        negozio = st.text_input(
            "Negozio"
        )

        acquistato = st.checkbox(
            "Già acquistato"
        )

        salva = st.form_submit_button(
            "➕ Aggiungi mobile"
        )

        if salva and articolo.strip():

            execute(
                """
                INSERT INTO mobili
                (articolo, costo, negozio, acquistato)
                VALUES (?, ?, ?, ?)
                """,
                (
                    articolo,
                    costo,
                    negozio,
                    int(acquistato)
                )
            )

            st.rerun()

    df = query_df(
        """
        SELECT
            id AS ID,
            articolo AS Articolo,
            costo AS Costo,
            negozio AS Negozio,
            acquistato AS Acquistato
        FROM mobili
        """
    )

    if not df.empty:

        totale = df["Costo"].sum()

        acquistato_totale = df[
            df["Acquistato"] == 1
        ]["Costo"].sum()

        residuo = budget_arredi - totale

        c1, c2, c3 = st.columns(3)

        with c1:
            st.metric(
                "Budget",
                euro(budget_arredi)
            )

        with c2:
            st.metric(
                "Totale mobili",
                euro(totale)
            )

        with c3:
            st.metric(
                "Residuo",
                euro(residuo)
            )

        st.progress(
            min(totale / budget_arredi, 1)
            if budget_arredi > 0 else 0
        )

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True
        )

    else:

        st.info(
            "Nessun mobile inserito."
        )


# ============================================================
# RENDERING
# ============================================================

elif menu == "🖼️ Rendering & Planimetrie":

    st.title(
        "🖼️ Rendering & Planimetrie"
    )

    stanza = st.selectbox(
        "Seleziona ambiente",
        [
            "Cucina",
            "Salotto",
            "Ingresso",
            "Camera matrimoniale",
            "Camera bimbe",
            "Bagno piano terra",
            "Bagno ammezzato",
            "Bagno piano primo",
            "Mansarda",
            "Giardino",
            "Altro"
        ]
    )

    file = st.file_uploader(
        "📷 Carica rendering o planimetria",
        type=[
            "png",
            "jpg",
            "jpeg",
            "webp"
        ],
        accept_multiple_files=True
    )

    if file:

        st.subheader(
            f"📷 {stanza}"
        )

        for immagine in file:

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
        "📥 Importa / Esporta dati"
    )

    st.subheader(
        "📤 Esporta spese"
    )

    df_spese = query_df(
        "SELECT * FROM spese"
    )

    if not df_spese.empty:

        csv = df_spese.to_csv(
            index=False
        ).encode("utf-8")

        st.download_button(
            "⬇️ Scarica spese CSV",
            data=csv,
            file_name="spese_famiglia.csv",
            mime="text/csv"
        )

    st.markdown("---")

    st.subheader(
        "📥 Importa spese da CSV"
    )

    uploaded = st.file_uploader(
        "Carica CSV",
        type=["csv"]
    )

    if uploaded:

        try:

            df_import = pd.read_csv(
                uploaded
            )

            st.dataframe(
                df_import,
                use_container_width=True
            )

            st.info(
                "Controlla i dati prima di importarli."
            )

        except Exception as e:

            st.error(
                f"Errore: {e}"
            )


# ============================================================
# IMPOSTAZIONI
# ============================================================

elif menu == "⚙️ Impostazioni":

    st.title("⚙️ Impostazioni")

    st.subheader(
        "📊 Stato database"
    )

    st.write(
        f"Database: `{DB_FILE}`"
    )

    st.success(
        "Database SQLite attivo."
    )

    st.markdown("---")

    st.subheader(
        "🗑️ Attenzione"
    )

    st.warning(
        "Le operazioni seguenti cancellano definitivamente i dati."
    )

    if st.button(
        "⚠️ CANCELLA TUTTI I DATI"
    ):

        conn = get_connection()
        cursor = conn.cursor()

        tabelle = [
            "entrate",
            "spese",
            "spese_ricorrenti",
            "budget",
            "obiettivi",
            "costi_casa",
            "entrate_casa",
            "mobili",
            "spese_future"
        ]

        for tabella in tabelle:
            cursor.execute(
                f"DELETE FROM {tabella}"
            )

        conn.commit()
        conn.close()

        st.success(
            "Database svuotato."
        )

        st.rerun()
