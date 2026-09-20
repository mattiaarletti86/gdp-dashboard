import streamlit as st
import pandas as pd
import sqlite3
import matplotlib.pyplot as plt

from pathlib import Path
from datetime import date, datetime
from io import BytesIO
import re
import hashlib


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

# ============================================================
# FILE PRINCIPALI
# ============================================================

EXCEL_FILE = BASE_DIR / "Spese casa -2.xlsx"

# Database per i dati inseriti direttamente nell'app
DB_FILE = BASE_DIR / "finanze_famiglia.db"

# Cartella dove vengono estratte le immagini dall'Excel
IMAGE_DIR = BASE_DIR / "renderings_excel"
IMAGE_DIR.mkdir(exist_ok=True)


# ============================================================
# DATI STATICI
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
    "Dicembre"
]

EXPENSE_CATEGORIES = [
    "Alimentari",
    "Tempo libero e viaggi",
    "Utenze",
    "Scuola e sport",
    "Trasporti e auto",
    "Prelievi contanti",
    "Casa e assicurazioni",
    "Shopping",
    "Farmacia e cura della persona",
    "Arredi e nuova casa",
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
    "Famiglia",
    "Mattia",
    "Virginia"
]

ROOM_SHEETS = [
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
    "Mansarda"
]

PLAN_SHEETS = [
    "Planimetrie",
    "Planimetrie con rendering",
    "Planimetrie con rendering 1°",
    "Planimetrie con rendering 2°"
]


# ============================================================
# FUNZIONI UTILI
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
    Gestisce anche formati italiani:
    1.234,56
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

    text = str(value).strip()

    if not text:
        return None

    text = (
        text
        .replace("€", "")
        .replace(" ", "")
    )

    if "," in text:
        text = text.replace(".", "")
        text = text.replace(",", ".")

    try:
        return float(text)
    except Exception:
        return None


def month_label(year, month):
    return f"{MONTHS[month - 1]} {year}"


def section_title(title, subtitle=None):

    st.title(title)

    if subtitle:
        st.caption(subtitle)


# ============================================================
# DATABASE
# ============================================================

def conn():
    return sqlite3.connect(DB_FILE)


def init_db():

    connection = conn()
    cursor = connection.cursor()

    cursor.executescript(
        """
        CREATE TABLE IF NOT EXISTS entrate (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            descrizione TEXT,
            persona TEXT,
            categoria TEXT,
            importo REAL,
            data TEXT,
            ricorrente INTEGER DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS spese (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            descrizione TEXT,
            persona TEXT,
            categoria TEXT,
            importo REAL,
            data TEXT,
            ricorrente INTEGER DEFAULT 0,
            note TEXT
        );

        CREATE TABLE IF NOT EXISTS ricorrenti (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            descrizione TEXT,
            categoria TEXT,
            importo REAL,
            giorno INTEGER DEFAULT 1,
            persona TEXT
        );

        CREATE TABLE IF NOT EXISTS spese_future (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data TEXT,
            categoria TEXT,
            descrizione TEXT,
            importo REAL,
            note TEXT
        );

        CREATE TABLE IF NOT EXISTS budget (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            categoria TEXT UNIQUE,
            importo REAL
        );

        CREATE TABLE IF NOT EXISTS obiettivi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT,
            obiettivo REAL,
            accumulato REAL DEFAULT 0,
            scadenza TEXT
        );

        CREATE TABLE IF NOT EXISTS costi_casa (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            voce TEXT UNIQUE,
            importo REAL
        );

        CREATE TABLE IF NOT EXISTS entrate_casa (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            voce TEXT UNIQUE,
            importo REAL
        );
        """
    )

    connection.commit()
    connection.close()


init_db()


def qdf(sql, params=()):

    connection = conn()

    try:
        return pd.read_sql_query(
            sql,
            connection,
            params=params
        )

    finally:
        connection.close()


def execute(sql, params=()):

    connection = conn()

    try:
        cursor = connection.cursor()

        cursor.execute(
            sql,
            params
        )

        connection.commit()

        return cursor.lastrowid

    finally:
        connection.close()


# ============================================================
# LETTURA EXCEL
# ============================================================

def read_excel():

    if not EXCEL_FILE.exists():
        return None, None

    import openpyxl

    # Workbook con formule
    wb_formula = openpyxl.load_workbook(
        EXCEL_FILE,
        data_only=False
    )

    # Workbook con valori calcolati
    wb_values = openpyxl.load_workbook(
        EXCEL_FILE,
        data_only=True
    )

    return wb_formula, wb_values


# ============================================================
# STORICO COSTI FAMIGLIA
# ============================================================

def parse_family(wb_values):

    if "Costi famiglia" not in wb_values.sheetnames:
        return pd.DataFrame()

    ws = wb_values["Costi famiglia"]

    rows = []

    year = None
    month = None

    months = {
        name.upper(): index
        for index, name in enumerate(
            MONTHS,
            1
        )
    }

    for row_number in range(
        1,
        ws.max_row + 1
    ):

        first_cell = ws.cell(
            row_number,
            1
        ).value

        if first_cell is None:
            continue

        text = str(
            first_cell
        ).strip()

        text_upper = text.upper()

        # ----------------------------------------------------
        # Esempio:
        #
        # SPESE OTTOBRE 2025
        # ----------------------------------------------------

        match = re.match(
            r"SPESE\s+([A-ZÀ-Ù]+)\s+(\d{4})",
            text_upper
        )

        if match:

            month_name = match.group(1)
            year = int(
                match.group(2)
            )

            if month_name in months:
                month = months[
                    month_name
                ]

            continue

        if year is None or month is None:
            continue

        # ----------------------------------------------------
        # CATEGORIA
        # ----------------------------------------------------

        category = text

        if category.upper() in [
            "MEDIA",
            "TOTALE",
            ""
        ]:
            continue

        # ----------------------------------------------------
        # IMPORTO
        # ----------------------------------------------------

        amount = num(
            ws.cell(
                row_number,
                2
            ).value
        )

        if amount is None:
            continue

        if amount == 0:
            continue

        rows.append(
            {
                "anno": year,
                "mese": month,
                "data": date(
                    year,
                    month,
                    1
                ),
                "categoria": category,
                "importo": amount
            }
        )

    return pd.DataFrame(
        rows
    )


# ============================================================
# PIANO ACQUISTO NUOVA CASA
# ============================================================

def parse_house(wb_values):

    if "Piano acquisto nuova casa" not in wb_values.sheetnames:
        return {
            "costs": [],
            "funds": [],
            "detrazioni": pd.DataFrame(),
            "scenarios": pd.DataFrame(),
            "fixed1": pd.DataFrame(),
            "fixed2": pd.DataFrame()
        }

    ws = wb_values[
        "Piano acquisto nuova casa"
    ]

    costs = []
    funds = []
    detrazioni = []
    scenarios = []
    fixed1 = []
    fixed2 = []

    # ========================================================
    # COSTI ACQUISTO CASA
    # ========================================================

    for row in range(3, 15):

        label = ws.cell(
            row,
            2
        ).value

        value = ws.cell(
            row,
            3
        ).value

        amount = num(value)

        if (
            label
            and amount is not None
        ):

            costs.append(
                (
                    str(label).strip(),
                    amount
                )
            )

    # ========================================================
    # COPERTURE / RICAVI
    # ========================================================

    for row in range(15, 18):

        label = ws.cell(
            row,
            2
        ).value

        value = ws.cell(
            row,
            4
        ).value

        amount = num(value)

        if (
            label
            and amount is not None
        ):

            funds.append(
                (
                    str(label).strip(),
                    amount
                )
            )

    # ========================================================
    # DETRAZIONI
    # ========================================================

    for row in range(21, 31):

        year = ws.cell(
            row,
            2
        ).value

        if not year:
            continue

        try:
            year = int(year)
        except Exception:
            continue

        detrazioni.append(
            {
                "Anno": year,

                "Ristrutturazione vecchio appartamento":
                    num(
                        ws.cell(
                            row,
                            3
                        ).value
                    ) or 0,

                "Bonus nuova casa":
                    num(
                        ws.cell(
                            row,
                            4
                        ).value
                    ) or 0,

                "Bonus mobili":
                    num(
                        ws.cell(
                            row,
                            5
                        ).value
                    ) or 0,

                "Totale detrazioni":
                    num(
                        ws.cell(
                            row,
                            6
                        ).value
                    ) or 0
            }
        )

    # ========================================================
    # SCENARI COSTO MENSILE
    # ========================================================

    scenario_rows = [
        34,
        35,
        36,
        37,
        39,
        40,
        41,
        43,
        44,
        45
    ]

    for row in scenario_rows:

        label = ws.cell(
            row,
            2
        ).value

        if not label:
            continue

        scenarios.append(
            {
                "Riga": row,

                "Voce":
                    str(label).strip(),

                "Base":
                    num(
                        ws.cell(
                            row,
                            3
                        ).value
                    ),

                "50%":
                    num(
                        ws.cell(
                            row,
                            4
                        ).value
                    ),

                "36%":
                    num(
                        ws.cell(
                            row,
                            5
                        ).value
                    )
            }
        )

    # ========================================================
    # SPESE FISSE
    # ========================================================

    for row in range(51, 61):

        label = ws.cell(
            row,
            2
        ).value

        if not label:
            continue

        value = num(
            ws.cell(
                row,
                3
            ).value
        )

        if value is not None:

            fixed1.append(
                {
                    "Voce":
                        str(label).strip(),

                    "Importo":
                        value,

                    "Tipo":
                        "Costo"
                }
            )

    # Valori colonna D
    for row in range(59, 61):

        label = ws.cell(
            row,
            2
        ).value

        value = num(
            ws.cell(
                row,
                4
            ).value
        )

        if (
            label
            and value is not None
        ):

            fixed1.append(
                {
                    "Voce":
                        str(label).strip(),

                    "Importo":
                        value,

                    "Tipo":
                        "Entrata"
                }
            )

    # ========================================================
    # SECONDO PERIODO
    # ========================================================

    for row in range(66, 77):

        label = ws.cell(
            row,
            2
        ).value

        if not label:
            continue

        value_c = num(
            ws.cell(
                row,
                3
            ).value
        )

        value_d = num(
            ws.cell(
                row,
                4
            ).value
        )

        if value_c is not None:

            fixed2.append(
                {
                    "Voce":
                        str(label).strip(),

                    "Importo":
                        value_c,

                    "Tipo":
                        "Costo"
                }
            )

        if value_d is not None:

            fixed2.append(
                {
                    "Voce":
                        str(label).strip(),

                    "Importo":
                        value_d,

                    "Tipo":
                        "Entrata"
                }
            )

    return {
        "costs": costs,
        "funds": funds,

        "detrazioni":
            pd.DataFrame(
                detrazioni
            ),

        "scenarios":
            pd.DataFrame(
                scenarios
            ),

        "fixed1":
            pd.DataFrame(
                fixed1
            ),

        "fixed2":
            pd.DataFrame(
                fixed2
            )
    }


# ============================================================
# MOBILI
# ============================================================

def parse_furniture(wb_values):

    if "Mobili" not in wb_values.sheetnames:
        return pd.DataFrame()

    ws = wb_values[
        "Mobili"
    ]

    rows = []

    for row in range(
        3,
        ws.max_row + 1
    ):

        article = ws.cell(
            row,
            1
        ).value

        cost = num(
            ws.cell(
                row,
                2
            ).value
        )

        if (
            article
            and cost is not None
        ):

            rows.append(
                {
                    "Articolo":
                        str(article).strip(),

                    "Costo":
                        cost,

                    "Riferimento":
                        ws.cell(
                            row,
                            3
                        ).value or "",

                    "Negozio":
                        ws.cell(
                            row,
                            4
                        ).value or ""
                }
            )

    return pd.DataFrame(
        rows
    )


# ============================================================
# ESTRAZIONE IMMAGINI EXCEL
# ============================================================

def extract_images(wb_formula):

    extracted = {}

    for sheet_name in (
        ROOM_SHEETS
        +
        PLAN_SHEETS
    ):

        if sheet_name not in wb_formula.sheetnames:
            continue

        ws = wb_formula[
            sheet_name
        ]

        images = []

        for index, image in enumerate(
            getattr(
                ws,
                "_images",
                []
            ),
            1
        ):

            try:

                data = image._data()

                extension = "png"

                if data[:3] == b"\xff\xd8\xff":
                    extension = "jpg"

                filename = (
                    f"{hashlib.md5(sheet_name.encode()).hexdigest()[:8]}"
                    f"_{index}.{extension}"
                )

                path = (
                    IMAGE_DIR
                    /
                    filename
                )

                path.write_bytes(
                    data
                )

                # Posizione dell'immagine nel foglio
                anchor = getattr(
                    image,
                    "anchor",
                    None
                )

                cell = ""

                if hasattr(
                    anchor,
                    "_from"
                ):

                    cell = (
                        f"{anchor._from.col + 1},"
                        f"{anchor._from.row + 1}"
                    )

                images.append(
                    {
                        "name":
                            filename,

                        "path":
                            str(path),

                        "cell":
                            cell
                    }
                )

            except Exception:
                continue

        if images:
            extracted[
                sheet_name
            ] = images

    return extracted


# ============================================================
# CARICA TUTTO L'EXCEL
# ============================================================

def load_all_excel():

    if not EXCEL_FILE.exists():
        return None

    wb_formula, wb_values = read_excel()

    return {
        "house":
            parse_house(
                wb_values
            ),

        "family":
            parse_family(
                wb_values
            ),

        "furniture":
            parse_furniture(
                wb_values
            ),

        "images":
            extract_images(
                wb_formula
            ),

        "sheets":
            wb_formula.sheetnames
    }


excel = load_all_excel()


# ============================================================
# IMPORT COSTI CASA NEL DATABASE
# ============================================================

def seed_house_from_excel():

    if not excel:
        return

    house = excel[
        "house"
    ]

    count_costs = qdf(
        """
        SELECT COUNT(*) AS n
        FROM costi_casa
        """
    ).iloc[0]["n"]

    if count_costs == 0:

        for label, value in house["costs"]:

            execute(
                """
                INSERT OR IGNORE INTO costi_casa
                (voce, importo)
                VALUES (?, ?)
                """,
                (
                    label,
                    value
                )
            )

    count_funds = qdf(
        """
        SELECT COUNT(*) AS n
        FROM entrate_casa
        """
    ).iloc[0]["n"]

    if count_funds == 0:

        for label, value in house["funds"]:

            execute(
                """
                INSERT OR IGNORE INTO entrate_casa
                (voce, importo)
                VALUES (?, ?)
                """,
                (
                    label,
                    value
                )
            )


seed_house_from_excel()


# ============================================================
# TOTALE MENSILE
# ============================================================

def month_total(
    year,
    month
):

    total = 0

    # Storico Excel
    if (
        excel
        and not excel["family"].empty
    ):

        family = excel[
            "family"
        ]

        total += family.loc[
            (
                family["anno"] == year
            )
            &
            (
                family["mese"] == month
            ),
            "importo"
        ].sum()

    # Nuove spese app
    app_total = qdf(
        """
        SELECT
            COALESCE(SUM(importo), 0) AS totale
        FROM spese
        WHERE strftime('%Y', data) = ?
        AND strftime('%m', data) = ?
        """,
        (
            str(year),
            f"{month:02d}"
        )
    )

    total += float(
        app_total.iloc[0]["totale"]
    )

    return float(total)


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.title(
    "🏡 Gestione Casa"
)

menu = st.sidebar.radio(
    "SEZIONI",
    [
        "🏠 Home",
        "📜 Storico famiglia",
        "🏡 Nuova Casa",
        "🪑 Mobili",
        "🖼️ Rendering",
        "💰 Entrate",
        "💸 Spese",
        "🔁 Ricorrenti",
        "📅 Budget",
        "🎯 Obiettivi",
        "📊 Analisi",
        "📥 Import/Export"
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

    section_title(
        "🏡 Dashboard Casa & Finanze",
        "Excel = storico e progetto casa · "
        "SQLite = nuovi dati inseriti nell'app"
    )

    mese = st.date_input(
        "📅 Mese",
        date.today().replace(day=1)
    )

    year = mese.year
    month = mese.month

    # --------------------------------------------------------
    # DATI MENSILI
    # --------------------------------------------------------

    spese = month_total(
        year,
        month
    )

    entrate = float(
        qdf(
            """
            SELECT
                COALESCE(SUM(importo),0) AS totale
            FROM entrate
            WHERE strftime('%Y',data)=?
            AND strftime('%m',data)=?
            """,
            (
                str(year),
                f"{month:02d}"
            )
        ).iloc[0]["totale"]
    )

    risparmio = (
        entrate -
        spese
    )

    percentuale = (
        risparmio / entrate * 100
        if entrate
        else 0
    )

    # --------------------------------------------------------
    # METRICHE
    # --------------------------------------------------------

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "💰 Entrate",
        euro(entrate)
    )

    c2.metric(
        "💸 Spese",
        euro(spese)
    )

    c3.metric(
        "💚 Risparmio",
        euro(risparmio)
    )

    c4.metric(
        "📈 Risparmio %",
        f"{percentuale:.1f}%"
    )

    # --------------------------------------------------------
    # NUOVA CASA
    # --------------------------------------------------------

    if excel:

        house = excel[
            "house"
        ]

        detrazioni = house[
            "detrazioni"
        ]

        total_costs = sum(
            value
            for _, value
            in house["costs"]
        )

        total_funds = sum(
            value
            for _, value
            in house["funds"]
        )

        total_detrazioni = (
            detrazioni[
                "Totale detrazioni"
            ].sum()
            if not detrazioni.empty
            else 0
        )

        st.markdown("---")

        st.subheader(
            "🏡 Nuova casa"
        )

        a, b, c, d = st.columns(4)

        a.metric(
            "Costi casa",
            euro(total_costs)
        )

        b.metric(
            "Coperture",
            euro(total_funds)
        )

        c.metric(
            "Residuo",
            euro(
                total_funds -
                total_costs
            )
        )

        d.metric(
            "Detrazioni",
            euro(total_detrazioni)
        )

        # ----------------------------------------------------
        # SCENARI
        # ----------------------------------------------------

        scenarios = house[
            "scenarios"
        ]

        if not scenarios.empty:

            rows = []

            for row_number in [
                37,
                41,
                45
            ]:

                row = scenarios[
                    scenarios["Riga"]
                    == row_number
                ]

                if not row.empty:

                    rows.append(
                        {
                            "Scenario":
                                row.iloc[0]["Voce"],

                            "50%":
                                row.iloc[0]["50%"],

                            "36%":
                                row.iloc[0]["36%"]
                        }
                    )

            if rows:

                scenario_df = pd.DataFrame(
                    rows
                )

                st.subheader(
                    "💳 Costo netto mensile"
                )

                st.dataframe(
                    scenario_df,
                    use_container_width=True,
                    hide_index=True
                )

        # ----------------------------------------------------
        # DETRAZIONI
        # ----------------------------------------------------

        if not detrazioni.empty:

            st.subheader(
                "🧾 Recuperi fiscali"
            )

            st.dataframe(
                detrazioni,
                use_container_width=True,
                hide_index=True
            )

    # --------------------------------------------------------
    # ANDAMENTO
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "📈 Andamento spese"
    )

    if (
        excel
        and not excel["family"].empty
    ):

        family = excel[
            "family"
        ]

        monthly = (
            family
            .groupby(
                [
                    "anno",
                    "mese"
                ],
                as_index=False
            )["importo"]
            .sum()
        )

        monthly["Mese"] = monthly.apply(
            lambda row:
            month_label(
                int(row["anno"]),
                int(row["mese"])
            ),
            axis=1
        )

        fig, ax = plt.subplots(
            figsize=(10, 4)
        )

        ax.plot(
            monthly["Mese"],
            monthly["importo"],
            marker="o"
        )

        ax.tick_params(
            axis="x",
            rotation=45
        )

        ax.set_ylabel(
            "€"
        )

        ax.grid(
            alpha=0.2
        )

        st.pyplot(
            fig,
            clear_figure=True
        )


# ============================================================
# STORICO FAMIGLIA
# ============================================================

elif menu == "📜 Storico famiglia":

    section_title(
        "📜 Storico spese",
        "Dati letti direttamente dal foglio "
        "'Costi famiglia'"
    )

    if (
        not excel
        or excel["family"].empty
    ):

        st.error(
            "Nessun dato storico trovato."
        )

    else:

        df = excel[
            "family"
        ]

        years = sorted(
            df["anno"].unique(),
            reverse=True
        )

        year = st.selectbox(
            "Anno",
            years
        )

        months = sorted(
            df.loc[
                df["anno"] == year,
                "mese"
            ].unique(),
            reverse=True
        )

        month = st.selectbox(
            "Mese",
            months,
            format_func=lambda x:
            month_label(
                year,
                x
            )
        )

        data = df[
            (
                df["anno"] == year
            )
            &
            (
                df["mese"] == month
            )
        ].copy()

        total = data[
            "importo"
        ].sum()

        st.metric(
            "Totale",
            euro(total)
        )

        col1, col2 = st.columns(2)

        with col1:

            fig, ax = plt.subplots(
                figsize=(6, 5)
            )

            ax.pie(
                data["importo"],
                labels=data["categoria"],
                autopct="%1.1f%%",
                startangle=90
            )

            ax.axis("equal")

            st.pyplot(
                fig,
                clear_figure=True
            )

        with col2:

            st.dataframe(
                data[
                    [
                        "categoria",
                        "importo"
                    ]
                ].rename(
                    columns={
                        "categoria":
                            "Categoria",

                        "importo":
                            "Importo"
                    }
                ),
                use_container_width=True,
                hide_index=True
            )


# ============================================================
# NUOVA CASA
# ============================================================

elif menu == "🏡 Nuova Casa":

    section_title(
        "🏡 Nuova Casa",
        "Dati importati dal foglio "
        "'Piano acquisto nuova casa'"
    )

    if not excel:

        st.error(
            "File Excel non trovato."
        )

    else:

        house = excel[
            "house"
        ]

        tabs = st.tabs(
            [
                "💰 Bilancio",
                "🧾 Detrazioni",
                "📊 Scenari",
                "📋 Spese fisse"
            ]
        )

        # ====================================================
        # BILANCIO
        # ====================================================

        with tabs[0]:

            costs = pd.DataFrame(
                house["costs"],
                columns=[
                    "Voce",
                    "Importo"
                ]
            )

            funds = pd.DataFrame(
                house["funds"],
                columns=[
                    "Voce",
                    "Importo"
                ]
            )

            total_costs = (
                costs["Importo"].sum()
                if not costs.empty
                else 0
            )

            total_funds = (
                funds["Importo"].sum()
                if not funds.empty
                else 0
            )

            residual = (
                total_funds -
                total_costs
            )

            c1, c2, c3 = st.columns(3)

            c1.metric(
                "Totale costi",
                euro(total_costs)
            )

            c2.metric(
                "Totale coperture",
                euro(total_funds)
            )

            c3.metric(
                "Residuo",
                euro(residual)
            )

            st.subheader(
                "🔴 Costi"
            )

            st.dataframe(
                costs,
                use_container_width=True,
                hide_index=True
            )

            st.subheader(
                "🟢 Coperture"
            )

            st.dataframe(
                funds,
                use_container_width=True,
                hide_index=True
            )

        # ====================================================
        # DETRAZIONI
        # ====================================================

        with tabs[1]:

            detrazioni = house[
                "detrazioni"
            ]

            if detrazioni.empty:

                st.warning(
                    "Nessuna tabella detrazioni trovata."
                )

            else:

                st.dataframe(
                    detrazioni,
                    use_container_width=True,
                    hide_index=True
                )

                total_detrazioni = (
                    detrazioni[
                        "Totale detrazioni"
                    ].sum()
                )

                st.metric(
                    "Totale recuperi fiscali",
                    euro(total_detrazioni)
                )

                # --------------------------------------------
                # GRAFICO
                # --------------------------------------------

                fig, ax = plt.subplots(
                    figsize=(10, 4)
                )

                ax.bar(
                    detrazioni[
                        "Anno"
                    ].astype(str),

                    detrazioni[
                        "Totale detrazioni"
                    ]
                )

                ax.set_ylabel(
                    "€"
                )

                ax.set_title(
                    "Detrazioni annuali"
                )

                st.pyplot(
                    fig,
                    clear_figure=True
                )

                # --------------------------------------------
                # MEDIA MENSILE
                # --------------------------------------------

                media_mensile = (
                    total_detrazioni
                    /
                    len(detrazioni)
                    /
                    12
                )

                st.info(
                    f"Media indicativa del recupero fiscale: "
                    f"**{euro(media_mensile)} al mese**."
                )

        # ====================================================
        # SCENARI
        # ====================================================

        with tabs[2]:

            scenarios = house[
                "scenarios"
            ]

            if scenarios.empty:

                st.warning(
                    "Nessun scenario trovato."
                )

            else:

                st.dataframe(
                    scenarios,
                    use_container_width=True,
                    hide_index=True
                )

                st.subheader(
                    "💳 Confronto scenari"
                )

                rows = []

                for row_number in [
                    37,
                    41,
                    45
                ]:

                    row = scenarios[
                        scenarios["Riga"]
                        == row_number
                    ]

                    if not row.empty:

                        rows.append(
                            {
                                "Scenario":
                                    row.iloc[0]["Voce"],

                                "Base":
                                    row.iloc[0]["Base"],

                                "50%":
                                    row.iloc[0]["50%"],

                                "36%":
                                    row.iloc[0]["36%"]
                            }
                        )

                if rows:

                    st.dataframe(
                        pd.DataFrame(rows),
                        use_container_width=True,
                        hide_index=True
                    )

        # ====================================================
        # SPESE FISSE
        # ====================================================

        with tabs[3]:

            fixed1 = house[
                "fixed1"
            ]

            fixed2 = house[
                "fixed2"
            ]

            st.subheader(
                "📅 Periodo 1"
            )

            st.dataframe(
                fixed1,
                use_container_width=True,
                hide_index=True
            )

            if not fixed1.empty:

                costs1 = fixed1.loc[
                    fixed1["Tipo"] == "Costo",
                    "Importo"
                ].sum()

                income1 = fixed1.loc[
                    fixed1["Tipo"] == "Entrata",
                    "Importo"
                ].sum()

                st.info(
                    f"Costi: {euro(costs1)} · "
                    f"Entrate: {euro(income1)} · "
                    f"Residuo: {euro(income1 - costs1)}"
                )

            st.subheader(
                "📅 Periodo 2"
            )

            st.dataframe(
                fixed2,
                use_container_width=True,
                hide_index=True
            )

            if not fixed2.empty:

                costs2 = fixed2.loc[
                    fixed2["Tipo"] == "Costo",
                    "Importo"
                ].sum()

                income2 = fixed2.loc[
                    fixed2["Tipo"] == "Entrata",
                    "Importo"
                ].sum()

                st.info(
                    f"Costi: {euro(costs2)} · "
                    f"Entrate: {euro(income2)} · "
                    f"Residuo: {euro(income2 - costs2)}"
                )


# ============================================================
# MOBILI
# ============================================================

elif menu == "🪑 Mobili":

    section_title(
        "🪑 Mobili",
        "Dati importati dal foglio 'Mobili'"
    )

    if (
        not excel
        or excel["furniture"].empty
    ):

        st.error(
            "Nessun mobile trovato."
        )

    else:

        furniture = excel[
            "furniture"
        ]

        budget = st.number_input(
            "💰 Budget mobili (€)",
            min_value=0.0,
            value=15000.0,
            step=500.0
        )

        total = furniture[
            "Costo"
        ].sum()

        residual = (
            budget -
            total
        )

        percentage = (
            total / budget
            if budget > 0
            else 0
        )

        percentage = min(
            percentage,
            1
        )

        c1, c2, c3 = st.columns(3)

        c1.metric(
            "Budget",
            euro(budget)
        )

        c2.metric(
            "Lista attuale",
            euro(total)
        )

        c3.metric(
            "Residuo",
            euro(residual)
        )

        st.progress(
            percentage
        )

        st.dataframe(
            furniture,
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# RENDERING
# ============================================================

elif menu == "🖼️ Rendering":

    section_title(
        "🖼️ Rendering & Planimetrie",
        "Le immagini incorporate nel file Excel "
        "vengono estratte automaticamente"
    )

    if not excel:

        st.error(
            "File Excel non trovato."
        )

    else:

        available = [
            sheet
            for sheet
            in ROOM_SHEETS + PLAN_SHEETS
            if sheet in excel["images"]
        ]

        if not available:

            st.warning(
                "Non sono state trovate immagini "
                "incorporate nei fogli previsti."
            )

        else:

            selected = st.selectbox(
                "📐 Seleziona ambiente",
                available
            )

            images = excel[
                "images"
            ][selected]

            columns = st.columns(2)

            for index, image in enumerate(
                images
            ):

                with columns[
                    index % 2
                ]:

                    st.image(
                        image["path"],
                        caption=(
                            f"{selected} - "
                            f"immagine {index + 1}"
                        ),
                        use_container_width=True
                    )

        # ----------------------------------------------------
        # PLANIMETRIE
        # ----------------------------------------------------

        st.markdown("---")

        st.subheader(
            "📐 Planimetrie"
        )

        for sheet in PLAN_SHEETS:

            if sheet not in excel["images"]:
                continue

            with st.expander(
                sheet
            ):

                images = excel[
                    "images"
                ][sheet]

                columns = st.columns(2)

                for index, image in enumerate(
                    images
                ):

                    with columns[
                        index % 2
                    ]:

                        st.image(
                            image["path"],
                            use_container_width=True
                        )


# ============================================================
# ENTRATE
# ============================================================

elif menu == "💰 Entrate":

    section_title(
        "💰 Entrate"
    )

    with st.form(
        "form_entrata"
    ):

        description = st.text_input(
            "Descrizione",
            placeholder="Es. Stipendio"
        )

        person = st.selectbox(
            "Persona",
            PEOPLE
        )

        category = st.selectbox(
            "Categoria",
            INCOME_CATEGORIES
        )

        amount = st.number_input(
            "Importo (€)",
            min_value=0.0,
            step=50.0
        )

        data = st.date_input(
            "Data",
            date.today()
        )

        recurring = st.checkbox(
            "Entrata ricorrente"
        )

        save = st.form_submit_button(
            "💾 Salva entrata"
        )

        if (
            save
            and description.strip()
            and amount > 0
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
                    description,
                    person,
                    category,
                    amount,
                    data.isoformat(),
                    int(recurring)
                )
            )

            st.success(
                "Entrata salvata."
            )

            st.rerun()

    st.markdown("---")

    df = qdf(
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

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# SPESE
# ============================================================

elif menu == "💸 Spese":

    section_title(
        "💸 Spese",
        "Le nuove spese vengono salvate nell'app; "
        "lo storico precedente rimane nel file Excel."
    )

    with st.form(
        "form_spesa"
    ):

        description = st.text_input(
            "Descrizione",
            placeholder="Es. Supermercato"
        )

        category = st.selectbox(
            "Categoria",
            EXPENSE_CATEGORIES
        )

        person = st.selectbox(
            "Persona",
            PEOPLE
        )

        amount = st.number_input(
            "Importo (€)",
            min_value=0.0,
            step=5.0
        )

        data = st.date_input(
            "Data",
            date.today()
        )

        note = st.text_input(
            "Note"
        )

        save = st.form_submit_button(
            "💾 Salva spesa"
        )

        if (
            save
            and description.strip()
            and amount > 0
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
                    note
                )
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    description,
                    person,
                    category,
                    amount,
                    data.isoformat(),
                    note
                )
            )

            st.success(
                "Spesa salvata."
            )

            st.rerun()

    st.markdown("---")

    df = qdf(
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

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True
    )

    if not df.empty:

        st.metric(
            "Totale spese app",
            euro(
                df["Importo"].sum()
            )
        )


# ============================================================
# RICORRENTI
# ============================================================

elif menu == "🔁 Ricorrenti":

    section_title(
        "🔁 Spese ricorrenti"
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
            EXPENSE_CATEGORIES
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
            PEOPLE
        )

        save = st.form_submit_button(
            "➕ Aggiungi"
        )

        if (
            save
            and description.strip()
            and amount > 0
        ):

            execute(
                """
                INSERT INTO ricorrenti
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

    df = qdf(
        """
        SELECT
            id AS ID,
            descrizione AS Descrizione,
            categoria AS Categoria,
            importo AS Importo,
            giorno AS Giorno,
            persona AS Persona
        FROM ricorrenti
        ORDER BY giorno
        """
    )

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True
    )

    total = (
        df["Importo"].sum()
        if not df.empty
        else 0
    )

    st.metric(
        "💸 Totale fisso mensile",
        euro(total)
    )


# ============================================================
# BUDGET
# ============================================================

elif menu == "📅 Budget":

    section_title(
        "📅 Budget mensile"
    )

    category = st.selectbox(
        "Categoria",
        EXPENSE_CATEGORIES
    )

    amount = st.number_input(
        "Budget (€)",
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
                category,
                amount
            )
        )

        st.success(
            "Budget salvato."
        )

        st.rerun()

    st.markdown("---")

    month = st.date_input(
        "Mese da controllare",
        date.today().replace(day=1)
    )

    year = month.year
    month_number = month.month

    budgets = qdf(
        """
        SELECT
            categoria,
            importo
        FROM budget
        """
    )

    rows = []

    for _, row in budgets.iterrows():

        spent = float(
            qdf(
                """
                SELECT
                    COALESCE(
                        SUM(importo),
                        0
                    ) AS totale
                FROM spese
                WHERE categoria=?
                AND strftime('%Y',data)=?
                AND strftime('%m',data)=?
                """,
                (
                    row["categoria"],
                    str(year),
                    f"{month_number:02d}"
                )
            ).iloc[0]["totale"]
        )

        rows.append(
            {
                "Categoria":
                    row["categoria"],

                "Budget":
                    row["importo"],

                "Speso":
                    spent,

                "Residuo":
                    row["importo"] - spent
            }
        )

    if rows:

        budget_result = pd.DataFrame(
            rows
        )

        st.dataframe(
            budget_result,
            use_container_width=True,
            hide_index=True
        )

    else:

        st.info(
            "Nessun budget impostato."
        )


# ============================================================
# OBIETTIVI
# ============================================================

elif menu == "🎯 Obiettivi":

    section_title(
        "🎯 Obiettivi di risparmio"
    )

    with st.form(
        "form_obiettivo"
    ):

        name = st.text_input(
            "Nome obiettivo",
            placeholder="Es. Vacanza"
        )

        target = st.number_input(
            "Importo obiettivo (€)",
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

        if (
            save
            and name.strip()
            and target > 0
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

    goals = qdf(
        """
        SELECT *
        FROM obiettivi
        ORDER BY scadenza
        """
    )

    if goals.empty:

        st.info(
            "Nessun obiettivo."
        )

    else:

        for _, row in goals.iterrows():

            target = float(
                row["obiettivo"]
            )

            saved = float(
                row["accumulato"]
            )

            percentage = (
                saved / target
                if target
                else 0
            )

            percentage = min(
                max(
                    percentage,
                    0
                ),
                1
            )

            st.subheader(
                f"🎯 {row['nome']}"
            )

            st.progress(
                percentage
            )

            st.write(
                f"{euro(saved)} / "
                f"{euro(target)}"
            )

            st.caption(
                f"Scadenza: {row['scadenza']}"
            )


# ============================================================
# ANALISI
# ============================================================

elif menu == "📊 Analisi":

    section_title(
        "📊 Analisi finanziaria"
    )

    if (
        not excel
        or excel["family"].empty
    ):

        st.warning(
            "Nessun dato storico disponibile."
        )

    else:

        family = excel[
            "family"
        ]

        # ----------------------------------------------------
        # ANDAMENTO MENSILE
        # ----------------------------------------------------

        st.subheader(
            "📈 Andamento mensile"
        )

        monthly = (
            family
            .groupby(
                [
                    "anno",
                    "mese"
                ],
                as_index=False
            )["importo"]
            .sum()
        )

        monthly["Mese"] = monthly.apply(
            lambda row:
            month_label(
                int(row["anno"]),
                int(row["mese"])
            ),
            axis=1
        )

        st.dataframe(
            monthly[
                [
                    "Mese",
                    "importo"
                ]
            ].rename(
                columns={
                    "importo":
                        "Spese"
                }
            ),
            use_container_width=True,
            hide_index=True
        )

        fig, ax = plt.subplots(
            figsize=(10, 4)
        )

        ax.plot(
            monthly["Mese"],
            monthly["importo"],
            marker="o"
        )

        ax.tick_params(
            axis="x",
            rotation=45
        )

        ax.grid(
            alpha=0.2
        )

        ax.set_ylabel(
            "€"
        )

        st.pyplot(
            fig,
            clear_figure=True
        )

        # ----------------------------------------------------
        # CATEGORIE
        # ----------------------------------------------------

        st.markdown("---")

        st.subheader(
            "🍕 Distribuzione per categoria"
        )

        categories = (
            family
            .groupby(
                "categoria"
            )["importo"]
            .sum()
            .sort_values(
                ascending=False
            )
        )

        fig2, ax2 = plt.subplots(
            figsize=(8, 7)
        )

        ax2.pie(
            categories.values,
            labels=categories.index,
            autopct="%1.1f%%",
            startangle=90
        )

        ax2.axis(
            "equal"
        )

        st.pyplot(
            fig2,
            clear_figure=True
        )

        st.dataframe(
            categories
            .reset_index()
            .rename(
                columns={
                    "categoria":
                        "Categoria",

                    "importo":
                        "Totale"
                }
            ),
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# IMPORT / EXPORT
# ============================================================

elif menu == "📥 Import/Export":

    section_title(
        "📥 Import / Export"
    )

    if EXCEL_FILE.exists():

        st.success(
            "🟢 File Excel collegato."
        )

        st.write(
            f"File: `{EXCEL_FILE.name}`"
        )

    else:

        st.error(
            "🔴 File Excel non trovato."
        )

    st.markdown("---")

    st.subheader(
        "⬇️ Backup dei dati dell'app"
    )

    tables = {
        "Spese app":
            qdf(
                "SELECT * FROM spese"
            ),

        "Entrate app":
            qdf(
                "SELECT * FROM entrate"
            ),

        "Ricorrenti":
            qdf(
                "SELECT * FROM ricorrenti"
            ),

        "Obiettivi":
            qdf(
                "SELECT * FROM obiettivi"
            )
    }

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        if excel:

            excel[
                "family"
            ].to_excel(
                writer,
                sheet_name="Storico Excel",
                index=False
            )

            pd.DataFrame(
                excel["house"]["costs"],
                columns=[
                    "Voce",
                    "Importo"
                ]
            ).to_excel(
                writer,
                sheet_name="Casa Costi",
                index=False
            )

            pd.DataFrame(
                excel["house"]["funds"],
                columns=[
                    "Voce",
                    "Importo"
                ]
            ).to_excel(
                writer,
                sheet_name="Casa Coperture",
                index=False
            )

            excel[
                "house"
            ]["detrazioni"].to_excel(
                writer,
                sheet_name="Detrazioni",
                index=False
            )

            excel[
                "furniture"
            ].to_excel(
                writer,
                sheet_name="Mobili",
                index=False
            )

        for name, dataframe in tables.items():

            dataframe.to_excel(
                writer,
                sheet_name=name[:31],
                index=False
            )

    st.download_button(
        "⬇️ Scarica backup/report",
        data=output.getvalue(),
        file_name="backup_finanze.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


# ============================================================
# FINE
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
