import streamlit as st
import pandas as pd
import sqlite3
import matplotlib.pyplot as plt
from datetime import date, datetime
from pathlib import Path

# ============================================================
# CONFIGURAZIONE
# ============================================================

st.set_page_config(
    page_title="Gestione Casa & Finanze",
    page_icon="🏡",
    layout="centered",
    initial_sidebar_state="expanded"
)

DB_FILE = "finanze_famiglia.db"
EXCEL_FILE = "Spese casa -2.xlsx"

# ============================================================
# STILE
# ============================================================

st.markdown("""
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
