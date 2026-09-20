import streamlit as st
import pandas as pd
import sqlite3
import matplotlib.pyplot as plt
from pathlib import Path
from datetime import date, datetime
from io import BytesIO
import calendar
import re
import hashlib

# ============================================================
# CONFIGURAZIONE
# ============================================================

st.set_page_config(
    page_title="Gestione Casa & Finanze",
    page_icon="🏡",
    layout="centered",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent

EXCEL_FILE = BASE_DIR / "Spese casa -2.xlsx"
DB_FILE = BASE_DIR / "finanze_famiglia.db"

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
    "Mansarda",
    "mansarda"
]

PLAN_SHEETS = [
    "Planimetrie",
    "Planimetrie con rendering",
    "Planimetrie con rendering 1°",
    "Planimetrie con rendering 2°"
]


# ============================================================
# UTILITÀ
# ============================================================

def euro(x):

    try:
        x = float(x)
    except Exception:
        x = 0

    return (
        f"{x:,.2f} €"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


def num(x):

    if x is None:
        return None

    try:
        if pd.isna(x):
            return None
    except Exception:
        pass

    if isinstance(x, (int, float)):
        return float(x)

    s = str(x).strip()

    s = (
        s
        .replace("€", "")
        .replace(" ", "")
    )

    if not s:
        return None

    if "," in s:
        s = (
            s
            .replace(".", "")
            .replace(",", ".")
        )

    try:
        return float(s)

    except Exception:
        return None


def month_label(year, month):

    return (
        f"{MONTHS[month - 1]} {year}"
    )


def section_title(
    title,
    subtitle=None
):

    st.title(title)

    if subtitle:
        st.caption(subtitle)


# ============================================================
# DATABASE
# ============================================================

def conn():

    return sqlite3.connect(
        DB_FILE
    )


def qdf(
    sql,
    params=()
):

    connection = conn()

    try:

        return pd.read_sql_query(
            sql,
            connection,
            params=params
        )

    finally:

        connection.close()


def execute(
    sql,
    params=()
):

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


def init_db():

    connection = conn()

    connection.executescript(
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

        CREATE TABLE IF NOT EXISTS impostazioni (
            chiave TEXT PRIMARY KEY,
            valore REAL
        );
        """
    )

    connection.commit()
    connection.close()


init_db()


# ============================================================
# LETTURA EXCEL
# ============================================================

def read_excel():

    if not EXCEL_FILE.exists():
        return None, None

    import openpyxl

    wb_formula = openpyxl.load_workbook(
        EXCEL_FILE,
        data_only=False
    )

    wb_values = openpyxl.load_workbook(
        EXCEL_FILE,
        data_only=True
    )

    return (
        wb_formula,
        wb_values
    )


# ============================================================
# COSTI FAMIGLIA
# ============================================================

def parse_family(wb):

    if "Costi famiglia" not in wb.sheetnames:
        return pd.DataFrame()

    ws = wb[
        "Costi famiglia"
    ]

    rows = []

    year = None
    month = None

    months = {
        name.upper(): index
        for index, name
        in enumerate(
            MONTHS,
            1
        )
    }

    for row in range(
        1,
        ws.max_row + 1
    ):

        first = ws.cell(
            row,
            1
        ).value

        if first is None:
            continue

        text = str(
            first
        ).strip()

        upper = text.upper()

        match = re.match(
            r"SPESE\s+([A-ZÀ-Ù]+)\s+(\d{4})",
            upper
        )

        if (
            match
            and match.group(1)
            in months
        ):

            month = months[
                match.group(1)
            ]

            year = int(
                match.group(2)
            )

            continue

        if (
            year is None
            or month is None
        ):
            continue

        amount = num(
            ws.cell(
                row,
                2
            ).value
        )

        if (
            amount is None
            or amount == 0
            or upper in [
                "MEDIA",
                "TOTALE"
            ]
        ):
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
                "categoria": text,
                "importo": amount
            }
        )

    return pd.DataFrame(
        rows
    )


# ============================================================
# PIANO ACQUISTO NUOVA CASA
# ============================================================

def parse_house(wb):

    if (
        "Piano acquisto nuova casa"
        not in wb.sheetnames
    ):

        return {
            "costs": [],
            "funds": [],
            "detrazioni": pd.DataFrame(),
            "scenarios": pd.DataFrame(),
            "fixed1": pd.DataFrame(),
            "fixed2": pd.DataFrame()
        }

    ws = wb[
        "Piano acquisto nuova casa"
    ]

    costs = []
    funds = []
    detrazioni = []
    scenarios = []
    fixed1 = []
    fixed2 = []

    # --------------------------------------------------------
    # COSTI
    # --------------------------------------------------------

    for row in range(
        3,
        15
    ):

        label = ws.cell(
            row,
            2
        ).value

        value = num(
            ws.cell(
                row,
                3
            ).value
        )

        if (
            label
            and value is not None
        ):

            costs.append(
                (
                    str(label).strip(),
                    value
                )
            )

    # --------------------------------------------------------
    # COPERTURE
    # --------------------------------------------------------

    for row in range(
        15,
        18
    ):

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

            funds.append(
                (
                    str(label).strip(),
                    value
                )
            )

    # --------------------------------------------------------
    # DETRAZIONI
    # --------------------------------------------------------

    for row in range(
        21,
        31
    ):

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
                "Anno":
                    year,

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

    # --------------------------------------------------------
    # SCENARI
    # --------------------------------------------------------

    for row in [
        34, 35, 36, 37,
        39, 40, 41,
        43, 44, 45
    ]:

        label = ws.cell(
            row,
            2
        ).value

        if not label:
            continue

        scenarios.append(
            {
                "riga":
                    row,

                "voce":
                    str(label).strip(),

                "base":
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

    # --------------------------------------------------------
    # SPESE FISSE PERIODO 1
    # --------------------------------------------------------

    for row in range(
        51,
        61
    ):

        label = ws.cell(
            row,
            2
        ).value

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

        if (
            label
            and value_c is not None
        ):

            fixed1.append(
                {
                    "Voce":
                        str(label).strip(),

                    "Importo":
                        value_c,

                    "Tipo":
                        "Costo"
                }
            )

        if (
            label
            and value_d is not None
        ):

            fixed1.append(
                {
                    "Voce":
                        str(label).strip(),

                    "Importo":
                        value_d,

                    "Tipo":
                        "Entrata"
                }
            )

    # --------------------------------------------------------
    # SPESE FISSE PERIODO 2
    # --------------------------------------------------------

    for row in range(
        66,
        77
    ):

        label = ws.cell(
            row,
            2
        ).value

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

        if (
            label
            and value_c is not None
        ):

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

        if (
            label
            and value_d is not None
        ):

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
        "costs":
            costs,

        "funds":
            funds,

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

def parse_furniture(wb):

    if "Mobili" not in wb.sheetnames:
        return pd.DataFrame()

    ws = wb[
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
# RENDERING
# ============================================================

def column_width_pixels(width):

    if width is None:
        width = 8.43

    return int(
        width * 7 + 5
    )


def row_height_pixels(height):

    if height is None:
        height = 15

    return int(
        height * 96 / 72
    )


def anchor_xy(
    ws,
    image
):

    anchor = getattr(
        image,
        "anchor",
        None
    )

    if (
        anchor is None
        or not hasattr(
            anchor,
            "_from"
        )
    ):

        return 0, 0

    point = anchor._from

    x = 0

    for col in range(
        1,
        point.col + 1
    ):

        try:

            letter = (
                ws.cell(
                    1,
                    col
                ).column_letter
            )

            width = ws.column_dimensions[
                letter
            ].width

            x += column_width_pixels(
                width
            )

        except Exception:

            x += column_width_pixels(
                8.43
            )

    y = 0

    for row in range(
        1,
        point.row + 1
    ):

        y += row_height_pixels(
            ws.row_dimensions[
                row
            ].height
        )

    x += int(
        getattr(
            point,
            "colOff",
            0
        ) / 9525
    )

    y += int(
        getattr(
            point,
            "rowOff",
            0
        ) / 9525
    )

    return (
        max(0, x),
        max(0, y)
    )


def compose_sheet_images(
    wb,
    sheet_name
):

    from PIL import Image

    if sheet_name not in wb.sheetnames:
        return None, []

    ws = wb[
        sheet_name
    ]

    images = getattr(
        ws,
        "_images",
        []
    )

    if not images:
        return None, []

    items = []

    min_x = 10**9
    min_y = 10**9
    max_x = 0
    max_y = 0

    for index, image in enumerate(
        images,
        1
    ):

        try:

            data = image._data()

            pil = Image.open(
                BytesIO(data)
            ).convert(
                "RGBA"
            )

            x, y = anchor_xy(
                ws,
                image
            )

            width = int(
                getattr(
                    image,
                    "width",
                    pil.width
                )
                or pil.width
            )

            height = int(
                getattr(
                    image,
                    "height",
                    pil.height
                )
                or pil.height
            )

            pil = pil.resize(
                (
                    max(1, width),
                    max(1, height)
                )
            )

            items.append(
                (
                    index,
                    pil,
                    x,
                    y
                )
            )

            min_x = min(
                min_x,
                x
            )

            min_y = min(
                min_y,
                y
            )

            max_x = max(
                max_x,
                x + width
            )

            max_y = max(
                max_y,
                y + height
            )

        except Exception:

            continue

    if not items:
        return None, []

    canvas_width = (
        max_x - min_x
    )

    canvas_height = (
        max_y - min_y
    )

    if (
        canvas_width <= 0
        or canvas_height <= 0
        or canvas_width > 8000
        or canvas_height > 8000
    ):

        return None, items

    canvas = Image.new(
        "RGBA",
        (
            canvas_width,
            canvas_height
        ),
        (
            255,
            255,
            255,
            255
        )
    )

    for (
        index,
        image,
        x,
        y
    ) in items:

        canvas.alpha_composite(
            image,
            (
                x - min_x,
                y - min_y
            )
        )

    filename = (
        "composite_"
        +
        hashlib.md5(
            sheet_name.encode()
        ).hexdigest()[:10]
        +
        ".png"
    )

    output = (
        IMAGE_DIR
        /
        filename
    )

    canvas.save(
        output
    )

    return (
        str(output),
        items
    )


def extract_images(
    wb
):

    result = {}

    sheets = (
        ROOM_SHEETS
        +
        PLAN_SHEETS
    )

    for sheet in sheets:

        if sheet not in wb.sheetnames:
            continue

        composite, items = (
            compose_sheet_images(
                wb,
                sheet
            )
        )

        single = []

        ws = wb[
            sheet
        ]

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

                extension = (
                    "jpg"
                    if data[:3] == b"\xff\xd8\xff"
                    else "png"
                )

                filename = (
                    hashlib.md5(
                        (
                            sheet
                            +
                            str(index)
                        ).encode()
                    ).hexdigest()[:12]
                    +
                    "."
                    +
                    extension
                )

                path = (
                    IMAGE_DIR
                    /
                    filename
                )

                path.write_bytes(
                    data
                )

                x, y = anchor_xy(
                    ws,
                    image
                )

                single.append(
                    {
                        "path":
                            str(path),

                        "x":
                            x,

                        "y":
                            y,

                        "index":
                            index
                    }
                )

            except Exception:

                continue

        if (
            single
            or composite
        ):

            result[
                sheet
            ] = {
                "composite":
                    composite,

                "images":
                    single
            }

    return result


# ============================================================
# CARICA TUTTO EXCEL
# ============================================================

def load_excel():

    if not EXCEL_FILE.exists():
        return None

    wb_formula, wb_values = (
        read_excel()
    )

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


excel = load_excel()


# ============================================================
# LOGICA GARAGE
# ============================================================

def is_garage_label(
    label
):

    text = str(
        label
    ).lower()

    return (
        "garage" in text
        or
        "autorimessa" in text
    )


def house_costs_df(
    include_garage=False
):

    if not excel:

        return pd.DataFrame(
            columns=[
                "Voce",
                "Importo"
            ]
        )

    df = pd.DataFrame(
        excel["house"]["costs"],
        columns=[
            "Voce",
            "Importo"
        ]
    )

    if not include_garage:

        df = df[
            ~df["Voce"].apply(
                is_garage_label
            )
        ].copy()

    return df


def garage_amount():

    if not excel:
        return 0.0

    for label, value in (
        excel["house"]["costs"]
    ):

        if is_garage_label(
            label
        ):

            return float(
                value
            )

    return 0.0


def garage_monthly_amount():

    return (
        garage_amount()
        /
        60.0
    )


def house_initial_costs_total():

    df = house_costs_df(
        include_garage=False
    )

    if df.empty:
        return 0.0

    return float(
        df["Importo"].sum()
    )


def house_total_coverage():

    if not excel:
        return 0.0

    return float(
        sum(
            value
            for _, value
            in excel["house"]["funds"]
        )
    )


# ============================================================
# SEED DATI CASA
# ============================================================

def seed_house_from_excel():

    if not excel:
        return

    count_costs = qdf(
        """
        SELECT COUNT(*) AS n
        FROM costi_casa
        """
    ).iloc[0, 0]

    if count_costs == 0:

        for label, value in (
            excel["house"]["costs"]
        ):

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
    ).iloc[0, 0]

    if count_funds == 0:

        for label, value in (
            excel["house"]["funds"]
        ):

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
# FINANZE MENSILI
# ============================================================

def month_total_excel(
    year,
    month
):

    if (
        not excel
        or excel["family"].empty
    ):

        return 0.0

    df = excel[
        "family"
    ]

    return float(
        df.loc[
            (
                df["anno"] == year
            )
            &
            (
                df["mese"] == month
            ),
            "importo"
        ].sum()
    )


def month_total_app(
    year,
    month
):

    result = qdf(
        """
        SELECT
            COALESCE(
                SUM(importo),
                0
            ) AS totale

        FROM spese

        WHERE strftime(
            '%Y',
            data
        ) = ?

        AND strftime(
            '%m',
            data
        ) = ?
        """,
        (
            str(year),
            f"{month:02d}"
        )
    )

    return float(
        result.iloc[0, 0]
    )


def month_income_app(
    year,
    month
):

    result = qdf(
        """
        SELECT
            COALESCE(
                SUM(importo),
                0
            ) AS totale

        FROM entrate

        WHERE strftime(
            '%Y',
            data
        ) = ?

        AND strftime(
            '%m',
            data
        ) = ?
        """,
        (
            str(year),
            f"{month:02d}"
        )
    )

    return float(
        result.iloc[0, 0]
    )


# ============================================================
# MEDIA STORICA PER CATEGORIA
# ============================================================

def historical_average_by_category(
    months_window=12
):

    if (
        not excel
        or excel["family"].empty
    ):

        return pd.DataFrame(
            columns=[
                "Categoria",
                "Media mensile"
            ]
        )

    df = excel[
        "family"
    ].copy()

    periods = (
        df[
            [
                "anno",
                "mese"
            ]
        ]
        .drop_duplicates()
        .sort_values(
            [
                "anno",
                "mese"
            ]
        )
    )

    if len(periods) > months_window:

        periods = periods.tail(
            months_window
        )

    df = df.merge(
        periods,
        on=[
            "anno",
            "mese"
        ],
        how="inner"
    )

    number_of_months = max(
        1,
        len(periods)
    )

    result = (
        df
        .groupby(
            "categoria",
            as_index=False
        )["importo"]
        .sum()
    )

    result["Media mensile"] = (
        result["importo"]
        /
        number_of_months
    )

    result = (
        result
        .rename(
            columns={
                "categoria":
                    "Categoria"
            }
        )
        [
            [
                "Categoria",
                "Media mensile"
            ]
        ]
    )

    return result.sort_values(
        "Media mensile",
        ascending=False
    )


# ============================================================
# BUDGET GUIDA
# ============================================================

def savings_guideline_table(
    months_window,
    target_savings_pct,
    income_monthly
):

    history = (
        historical_average_by_category(
            months_window
        )
    )

    if history.empty:
        return history

    total_history = (
        history[
            "Media mensile"
        ].sum()
    )

    target_saving = (
        income_monthly
        *
        target_savings_pct
        /
        100
    )

    available = max(
        0,
        income_monthly
        -
        target_saving
    )

    if (
        income_monthly <= 0
        or total_history <= 0
    ):

        history[
            "Budget guida"
        ] = history[
            "Media mensile"
        ]

        history[
            "Riduzione"
        ] = 0.0

        return history

    scale = (
        available
        /
        total_history
    )

    history[
        "Budget guida"
    ] = (
        history[
            "Media mensile"
        ]
        *
        scale
    )

    history[
        "Riduzione"
    ] = (
        1 - scale
    ).clip(
        lower=0
    )

    return history


# ============================================================
# SPESE RICORRENTI
# ============================================================

def app_recurring_monthly_total():

    df = qdf(
        """
        SELECT
            COALESCE(
                SUM(importo),
                0
            ) AS totale
        FROM ricorrenti
        """
    )

    return float(
        df.iloc[0, 0]
    )


def app_future_month_total(
    year,
    month
):

    df = qdf(
        """
        SELECT
            COALESCE(
                SUM(importo),
                0
            ) AS totale

        FROM spese_future

        WHERE strftime(
            '%Y',
            data
        ) = ?

        AND strftime(
            '%m',
            data
        ) = ?
        """,
        (
            str(year),
            f"{month:02d}"
        )
    )

    return float(
        df.iloc[0, 0]
    )# ============================================================
# COSTO MENSILE NUOVA CASA
# ============================================================

def find_monthly_mortgage():

    if not excel:
        return 0.0

    house = excel[
        "house"
    ]

    mortgage = 0.0

    for fixed in [
        house["fixed1"],
        house["fixed2"]
    ]:

        if fixed.empty:
            continue

        for _, row in fixed.iterrows():

            voce = str(
                row["Voce"]
            ).lower()

            tipo = str(
                row["Tipo"]
            )

            if (
                "mutuo" in voce
                and tipo == "Costo"
            ):

                mortgage = max(
                    mortgage,
                    float(
                        row["Importo"]
                    )
                )

    return mortgage


def new_house_monthly_cost(
    period="Primi 5 anni"
):

    if not excel:

        return (
            0.0,
            0.0,
            0.0,
            0.0
        )

    house = excel[
        "house"
    ]

    mortgage = (
        find_monthly_mortgage()
    )

    # --------------------------------------------------------
    # GARAGE:
    #
    # NON entra nel costo iniziale.
    # Viene pagato in 60 mesi.
    # --------------------------------------------------------

    if period == "Primi 5 anni":

        garage = (
            garage_monthly_amount()
        )

    else:

        garage = 0.0

    # --------------------------------------------------------
    # DETRAZIONI
    # --------------------------------------------------------

    detrazioni = house[
        "detrazioni"
    ]

    if detrazioni.empty:

        credit = 0.0

    elif period == "Primi 5 anni":

        credit = (
            detrazioni
            .head(5)
            [
                "Totale detrazioni"
            ]
            .mean()
            /
            12
        )

    elif period == "6° anno":

        if len(detrazioni) > 5:

            credit = (
                detrazioni.iloc[5]
                [
                    "Totale detrazioni"
                ]
                /
                12
            )

        else:

            credit = 0.0

    else:

        if len(detrazioni) >= 4:

            credit = (
                detrazioni.tail(4)
                [
                    "Totale detrazioni"
                ]
                .mean()
                /
                12
            )

        else:

            credit = (
                detrazioni[
                    "Totale detrazioni"
                ]
                .mean()
                /
                12
            )

    gross = (
        mortgage
        +
        garage
    )

    net = max(
        0,
        gross - credit
    )

    return (
        net,
        mortgage,
        garage,
        credit
    )


# ============================================================
# QUANTO POSSO SPENDERE OGGI
# ============================================================

def daily_spending_plan(
    target_saving_pct,
    expected_income,
    selected_date,
    include_house=True,
    include_recurring=True,
    include_future=True,
    safety_buffer=0.0,
    house_period="Primi 5 anni"
):

    year = (
        selected_date.year
    )

    month = (
        selected_date.month
    )

    days_in_month = (
        calendar.monthrange(
            year,
            month
        )[1]
    )

    days_remaining = max(
        1,
        days_in_month
        -
        selected_date.day
        +
        1
    )

    # --------------------------------------------------------
    # SPESE GIÀ SOSTENUTE
    # --------------------------------------------------------

    actual_spent = (
        month_total_excel(
            year,
            month
        )
        +
        month_total_app(
            year,
            month
        )
    )

    # --------------------------------------------------------
    # RISPARMIO
    # --------------------------------------------------------

    savings_target = max(
        0,
        expected_income
        *
        target_saving_pct
        /
        100
    )

    # --------------------------------------------------------
    # RICORRENTI
    # --------------------------------------------------------

    recurring = (
        app_recurring_monthly_total()
        if include_recurring
        else 0
    )

    # --------------------------------------------------------
    # SPESE FUTURE
    # --------------------------------------------------------

    future = (
        app_future_month_total(
            year,
            month
        )
        if include_future
        else 0
    )

    # --------------------------------------------------------
    # NUOVA CASA
    # --------------------------------------------------------

    house_cost = (
        new_house_monthly_cost(
            house_period
        )[0]
        if include_house
        else 0
    )

    # --------------------------------------------------------
    # BUDGET RESIDUO
    # --------------------------------------------------------

    remaining = (
        expected_income
        -
        actual_spent
        -
        recurring
        -
        future
        -
        house_cost
        -
        savings_target
        -
        safety_buffer
    )

    daily = (
        max(
            0,
            remaining
        )
        /
        days_remaining
    )

    return {
        "income":
            expected_income,

        "actual_spent":
            actual_spent,

        "recurring":
            recurring,

        "future":
            future,

        "house":
            house_cost,

        "saving":
            savings_target,

        "buffer":
            safety_buffer,

        "remaining":
            remaining,

        "daily":
            daily,

        "days_remaining":
            days_remaining
    }


# ============================================================
# MENU
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
        "💳 Quanto posso spendere oggi?",
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
        "Storico e progetto casa dal tuo Excel + "
        "nuovi dati inseriti nell'app"
    )

    mese = st.date_input(
        "📅 Mese",
        date.today().replace(
            day=1
        )
    )

    year = mese.year
    month = mese.month

    # --------------------------------------------------------
    # SPESE
    # --------------------------------------------------------

    spese = (
        month_total_excel(
            year,
            month
        )
        +
        month_total_app(
            year,
            month
        )
    )

    # --------------------------------------------------------
    # ENTRATE
    # --------------------------------------------------------

    entrate = (
        month_income_app(
            year,
            month
        )
    )

    risparmio = (
        entrate
        -
        spese
    )

    percentuale = (
        risparmio
        /
        entrate
        *
        100
        if entrate
        else 0
    )

    # --------------------------------------------------------
    # METRICHE
    # --------------------------------------------------------

    c1, c2, c3, c4 = (
        st.columns(4)
    )

    c1.metric(
        "💰 Entrate app",
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

        st.markdown("---")

        st.subheader(
            "🏡 Situazione nuova casa"
        )

        initial = (
            house_initial_costs_total()
        )

        coverage = (
            house_total_coverage()
        )

        garage = (
            garage_amount()
        )

        garage_month = (
            garage_monthly_amount()
        )

        det = excel[
            "house"
        ][
            "detrazioni"
        ]

        total_det = (
            det[
                "Totale detrazioni"
            ].sum()
            if not det.empty
            else 0
        )

        a, b, c, d = (
            st.columns(4)
        )

        a.metric(
            "Costi iniziali",
            euro(initial),
            help=(
                "Esclude il garage: "
                "il garage viene pagato "
                "in 60 mesi."
            )
        )

        b.metric(
            "Coperture",
            euro(coverage)
        )

        c.metric(
            "Residuo iniziale",
            euro(
                coverage
                -
                initial
            )
        )

        d.metric(
            "Detrazioni totali",
            euro(total_det)
        )

        st.info(
            f"🚗 Garage: {euro(garage)} "
            f"distribuiti in 60 mesi = "
            f"**{euro(garage_month)}/mese**. "
            "La quota garage non viene "
            "sottratta dalla liquidità iniziale."
        )

    # --------------------------------------------------------
    # BUDGET RISPARMIO
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "🎯 Budget per risparmiare"
    )

    st.caption(
        "La guida parte dalle tue spese "
        "storiche reali e permette di "
        "scegliere un obiettivo di risparmio."
    )

    target = st.slider(
        "Obiettivo minimo di risparmio",
        0,
        30,
        15,
        1,
        format="%d%%"
    )

    if (
        not excel
        or excel["family"].empty
    ):

        st.info(
            "Non ci sono dati storici."
        )

    elif entrate <= 0:

        st.info(
            "Inserisci almeno un'entrata "
            "mensile nella sezione Entrate."
        )

    else:

        target_amount = (
            entrate
            *
            target
            /
            100
        )

        available = max(
            0,
            entrate
            -
            target_amount
        )

        c1, c2 = st.columns(2)

        c1.metric(
            "Risparmio obiettivo",
            euro(target_amount)
        )

        c2.metric(
            "Budget massimo spese",
            euro(available)
        )

        guide = (
            savings_guideline_table(
                12,
                target,
                entrate
            )
        )

        st.dataframe(
            guide,
            use_container_width=True,
            hide_index=True
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

        total = (
            data["importo"].sum()
        )

        st.metric(
            "Totale",
            euro(total)
        )

        c1, c2 = (
            st.columns(2)
        )

        with c1:

            fig, ax = plt.subplots(
                figsize=(6, 5)
            )

            ax.pie(
                data["importo"],
                labels=data["categoria"],
                autopct="%1.1f%%",
                startangle=90
            )

            ax.axis(
                "equal"
            )

            st.pyplot(
                fig,
                clear_figure=True
            )

        with c2:

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
        "Costi, coperture, detrazioni "
        "e scenari importati dal tuo Excel"
    )

    if not excel:

        st.error(
            "File Excel non trovato."
        )

    else:

        house = excel[
            "house"
        ]

        costs_all = pd.DataFrame(
            house["costs"],
            columns=[
                "Voce",
                "Importo"
            ]
        )

        costs_initial = (
            house_costs_df(
                False
            )
        )

        funds = pd.DataFrame(
            house["funds"],
            columns=[
                "Voce",
                "Importo"
            ]
        )

        garage = (
            garage_amount()
        )

        garage_month = (
            garage_monthly_amount()
        )

        tabs = st.tabs(
            [
                "💰 Bilancio",
                "🚗 Garage",
                "🧾 Detrazioni",
                "📊 Scenari",
                "📋 Spese fisse"
            ]
        )

        # ====================================================
        # BILANCIO
        # ====================================================

        with tabs[0]:

            initial = (
                costs_initial[
                    "Importo"
                ].sum()
                if not costs_initial.empty
                else 0
            )

            coverage = (
                funds[
                    "Importo"
                ].sum()
                if not funds.empty
                else 0
            )

            st.subheader(
                "Esborso iniziale"
            )

            a, b, c = (
                st.columns(3)
            )

            a.metric(
                "Costi iniziali",
                euro(initial)
            )

            b.metric(
                "Coperture",
                euro(coverage)
            )

            c.metric(
                "Liquidità iniziale residua",
                euro(
                    coverage
                    -
                    initial
                )
            )

            st.info(
                f"Il garage da {euro(garage)} "
                "è escluso da questo conteggio "
                "perché viene pagato in 60 mesi."
            )

            st.subheader(
                "Costi iniziali"
            )

            st.dataframe(
                costs_initial,
                use_container_width=True,
                hide_index=True
            )

            st.subheader(
                "Voce esclusa dal rogito"
            )

            garage_df = costs_all[
                costs_all["Voce"].apply(
                    is_garage_label
                )
            ]

            st.dataframe(
                garage_df,
                use_container_width=True,
                hide_index=True
            )

            st.subheader(
                "Coperture"
            )

            st.dataframe(
                funds,
                use_container_width=True,
                hide_index=True
            )

        # ====================================================
        # GARAGE
        # ====================================================

        with tabs[1]:

            st.subheader(
                "🚗 Pagamento garage"
            )

            a, b, c = (
                st.columns(3)
            )

            a.metric(
                "Totale garage",
                euro(garage)
            )

            b.metric(
                "Durata",
                "60 mesi"
            )

            c.metric(
                "Quota mensile",
                euro(garage_month)
            )

            st.progress(
                1 / 5
            )

            st.caption(
                "La quota mensile viene considerata "
                "nel costo della nuova casa nei "
                "primi 5 anni, non nell'esborso iniziale."
            )

            st.markdown(
                f"**Calcolo:** "
                f"{euro(garage)} ÷ 60 = "
                f"**{euro(garage_month)}/mese**"
            )

        # ====================================================
        # DETRAZIONI
        # ====================================================

        with tabs[2]:

            det = house[
                "detrazioni"
            ]

            if det.empty:

                st.warning(
                    "Nessuna tabella detrazioni trovata."
                )

            else:

                st.dataframe(
                    det,
                    use_container_width=True,
                    hide_index=True
                )

                st.metric(
                    "Totale recuperi fiscali",
                    euro(
                        det[
                            "Totale detrazioni"
                        ].sum()
                    )
                )

                fig, ax = plt.subplots(
                    figsize=(10, 4)
                )

                ax.bar(
                    det["Anno"].astype(
                        str
                    ),
                    det[
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

        # ====================================================
        # SCENARI
        # ====================================================

        with tabs[3]:

            scen = house[
                "scenarios"
            ]

            st.dataframe(
                scen,
                use_container_width=True,
                hide_index=True
            )

            st.subheader(
                "Scenario mensile corretto"
            )

            if not house[
                "detrazioni"
            ].empty:

                avg_det = (
                    house[
                        "detrazioni"
                    ]
                    .head(5)
                    [
                        "Totale detrazioni"
                    ]
                    .mean()
                    /
                    12
                )

            else:

                avg_det = 0

            loan = (
                find_monthly_mortgage()
            )

            if loan > 0:

                det6 = 0

                if len(
                    house["detrazioni"]
                ) > 5:

                    det6 = (
                        house[
                            "detrazioni"
                        ]
                        .iloc[5]
                        [
                            "Totale detrazioni"
                        ]
                        /
                        12
                    )

                corrected = pd.DataFrame(
                    [
                        {
                            "Periodo":
                                "Primi 5 anni",

                            "Mutuo":
                                loan,

                            "Garage":
                                garage_month,

                            "Detrazione media":
                                avg_det,

                            "Costo netto":
                                loan
                                +
                                garage_month
                                -
                                avg_det
                        },

                        {
                            "Periodo":
                                "6° anno",

                            "Mutuo":
                                loan,

                            "Garage":
                                0,

                            "Detrazione media":
                                det6,

                            "Costo netto":
                                loan
                                -
                                det6
                        }
                    ]
                )

                st.dataframe(
                    corrected,
                    use_container_width=True,
                    hide_index=True
                )

                st.info(
                    "Il garage viene aggiunto "
                    "al costo mensile soltanto "
                    "nei primi 60 mesi."
                )

        # ====================================================
        # SPESE FISSE
        # ====================================================

        with tabs[4]:

            st.subheader(
                "Periodo 1"
            )

            st.dataframe(
                house["fixed1"],
                use_container_width=True,
                hide_index=True
            )

            st.subheader(
                "Periodo 2"
            )

            st.dataframe(
                house["fixed2"],
                use_container_width=True,
                hide_index=True
            )
            # ============================================================
# MOBILI
# ============================================================

elif menu == "🪑 Mobili":

    section_title(
        "🪑 Mobili",
        "Lista importata dal foglio 'Mobili'"
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
            "Budget mobili (€)",
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

        c1, c2, c3 = (
            st.columns(3)
        )

        c1.metric(
            "Budget",
            euro(budget)
        )

        c2.metric(
            "Lista",
            euro(total)
        )

        c3.metric(
            "Residuo",
            euro(residual)
        )

        st.progress(
            min(
                percentage,
                1
            )
        )

        st.dataframe(
            furniture,
            use_container_width=True,
            hide_index=True
        )

        st.info(
            "Per risparmiare puoi impostare "
            "un budget massimo per stanza e "
            "ordinare gli acquisti per priorità."
        )


# ============================================================
# RENDERING
# ============================================================

elif menu == "🖼️ Rendering":

    section_title(
        "🖼️ Rendering & Planimetrie",
        "Le immagini vengono estratte dal file Excel "
        "mantenendo posizione e sovrapposizione."
    )

    if not excel:

        st.error(
            "Excel non trovato."
        )

    else:

        available = [
            sheet
            for sheet in (
                ROOM_SHEETS
                +
                PLAN_SHEETS
            )
            if sheet in excel["images"]
        ]

        if not available:

            st.warning(
                "Nessuna immagine trovata "
                "nei fogli previsti."
            )

        else:

            choice = st.selectbox(
                "Seleziona ambiente / planimetria",
                available
            )

            data = excel[
                "images"
            ][choice]

            # ------------------------------------------------
            # COMPOSIZIONE COMPLETA
            # ------------------------------------------------

            if data.get(
                "composite"
            ):

                st.subheader(
                    f"📐 Composizione completa — {choice}"
                )

                st.image(
                    data["composite"],
                    use_container_width=True
                )

                if choice in PLAN_SHEETS:

                    st.caption(
                        "La planimetria viene ricostruita "
                        "mantenendo la posizione relativa "
                        "delle diverse immagini incorporate "
                        "nel foglio Excel."
                    )

            else:

                st.warning(
                    "Non è stato possibile creare "
                    "la composizione automatica."
                )

            # ------------------------------------------------
            # IMMAGINI SINGOLE
            # ------------------------------------------------

            with st.expander(
                "🔎 Visualizza immagini singole"
            ):

                cols = st.columns(2)

                for i, image in enumerate(
                    data.get(
                        "images",
                        []
                    )
                ):

                    with cols[
                        i % 2
                    ]:

                        st.image(
                            image["path"],
                            caption=(
                                f"Immagine {i + 1} · "
                                f"posizione Excel "
                                f"{image['x']},{image['y']}"
                            ),
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
        "Le nuove spese vengono salvate "
        "nel database; lo storico resta nell'Excel."
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
            placeholder="Es. Assicurazione"
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
            "Giorno",
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

    st.metric(
        "💸 Totale fisso mensile",
        euro(
            df["Importo"].sum()
            if not df.empty
            else 0
        )
    )


# ============================================================
# BUDGET E LINEE GUIDA
# ============================================================

elif menu == "📅 Budget":

    section_title(
        "📅 Budget e linee guida per risparmiare",
        "Il budget viene costruito partendo "
        "dalle tue spese storiche reali."
    )

    # --------------------------------------------------------
    # OBIETTIVO
    # --------------------------------------------------------

    st.subheader(
        "1. Obiettivo di risparmio"
    )

    target = st.slider(
        "Percentuale di reddito da destinare al risparmio",
        0,
        30,
        15,
        1,
        format="%d%%"
    )

    income = st.number_input(
        "Entrate mensili di riferimento (€)",
        min_value=0.0,
        value=0.0,
        step=100.0
    )

    months_window = st.selectbox(
        "Periodo storico",
        [
            3,
            6,
            12
        ],
        index=2
    )

    target_amount = (
        income
        *
        target
        /
        100
    )

    variable_budget = max(
        0,
        income
        -
        target_amount
    )

    c1, c2 = (
        st.columns(2)
    )

    c1.metric(
        "Risparmio obiettivo",
        euro(target_amount)
    )

    c2.metric(
        "Spesa massima teorica",
        euro(variable_budget)
    )

    # --------------------------------------------------------
    # BUDGET GUIDA
    # --------------------------------------------------------

    st.subheader(
        "2. Budget guida per categoria"
    )

    guide = (
        savings_guideline_table(
            months_window,
            target,
            income
        )
    )

    if guide.empty:

        st.info(
            "Per generare la guida servono "
            "lo storico Excel e un'entrata "
            "mensile di riferimento."
        )

    else:

        st.dataframe(
            guide,
            use_container_width=True,
            hide_index=True
        )

        st.markdown(
            "**Regola utilizzata:** il budget "
            "guida riduce proporzionalmente "
            "la tua spesa storica media per "
            "lasciare spazio all'obiettivo "
            "di risparmio."
        )

        # ----------------------------------------------------
        # TRE LIVELLI
        # ----------------------------------------------------

        total_history = (
            guide[
                "Media mensile"
            ].sum()
        )

        st.subheader(
            "3. Tre livelli pratici"
        )

        levels = pd.DataFrame(
            [
                {
                    "Profilo":
                        "Prudente",

                    "Risparmio":
                        "10%",

                    "Spesa massima":
                        euro(
                            income * 0.90
                        ),

                    "Differenza vs media":
                        euro(
                            income * 0.90
                            -
                            total_history
                        )
                },

                {
                    "Profilo":
                        "Equilibrato",

                    "Risparmio":
                        "15%",

                    "Spesa massima":
                        euro(
                            income * 0.85
                        ),

                    "Differenza vs media":
                        euro(
                            income * 0.85
                            -
                            total_history
                        )
                },

                {
                    "Profilo":
                        "Risparmio forte",

                    "Risparmio":
                        "20%",

                    "Spesa massima":
                        euro(
                            income * 0.80
                        ),

                    "Differenza vs media":
                        euro(
                            income * 0.80
                            -
                            total_history
                        )
                }
            ]
        )

        st.dataframe(
            levels,
            use_container_width=True,
            hide_index=True
        )

    # --------------------------------------------------------
    # SALVA BUDGET
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "4. Imposta il budget di una categoria"
    )

    category = st.selectbox(
        "Categoria",
        EXPENSE_CATEGORIES
    )

    amount = st.number_input(
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
                importo =
                    excluded.importo
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

    # --------------------------------------------------------
    # CONTROLLO MESE
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "5. Verifica il mese"
    )

    selected_month = st.date_input(
        "Mese da controllare",
        date.today().replace(
            day=1
        )
    )

    year = (
        selected_month.year
    )

    month = (
        selected_month.month
    )

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

        spent_excel = 0

        if (
            excel
            and not excel["family"].empty
        ):

            spent_excel = (
                excel["family"]
                .loc[
                    (
                        excel["family"].anno
                        == year
                    )
                    &
                    (
                        excel["family"].mese
                        == month
                    )
                    &
                    (
                        excel["family"].categoria
                        == row["categoria"]
                    ),
                    "importo"
                ]
                .sum()
            )

        spent_app = float(
            qdf(
                """
                SELECT
                    COALESCE(
                        SUM(importo),
                        0
                    )
                FROM spese

                WHERE categoria=?

                AND strftime(
                    '%Y',
                    data
                )=?

                AND strftime(
                    '%m',
                    data
                )=?
                """,
                (
                    row["categoria"],
                    str(year),
                    f"{month:02d}"
                )
            ).iloc[0, 0]
        )

        spent = (
            float(spent_excel)
            +
            spent_app
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
                    row["importo"]
                    -
                    spent
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

        over = budget_result[
            budget_result[
                "Residuo"
            ] < 0
        ]

        if not over.empty:

            st.warning(
                "Alcune categorie "
                "sono oltre il budget."
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
        "goal"
    ):

        name = st.text_input(
            "Nome"
        )

        target = st.number_input(
            "Obiettivo (€)",
            min_value=0.0,
            step=500.0
        )

        saved = st.number_input(
            "Accumulato (€)",
            min_value=0.0,
            step=100.0
        )

        deadline = st.date_input(
            "Scadenza",
            date.today()
        )

        ok = st.form_submit_button(
            "🎯 Crea"
        )

        if (
            ok
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

    df = qdf(
        """
        SELECT *
        FROM obiettivi
        ORDER BY scadenza
        """
    )

    if df.empty:

        st.info(
            "Nessun obiettivo."
        )

    else:

        for _, row in df.iterrows():

            percentage = (
                row["accumulato"]
                /
                row["obiettivo"]
                if row["obiettivo"]
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
                f"{euro(row['accumulato'])} "
                f"/ "
                f"{euro(row['obiettivo'])}"
            )

            st.caption(
                f"Scadenza: "
                f"{row['scadenza']}"
            )


# ============================================================
# QUANTO POSSO SPENDERE OGGI?
# ============================================================

elif menu == "💳 Quanto posso spendere oggi?":

    section_title(
        "💳 Quanto posso spendere oggi?",
        "Calcola un limite giornaliero pratico "
        "per arrivare a fine mese mantenendo "
        "il risparmio desiderato."
    )

    # --------------------------------------------------------
    # DATA
    # --------------------------------------------------------

    selected_date = st.date_input(
        "📅 Data di riferimento",
        date.today()
    )

    year = (
        selected_date.year
    )

    month = (
        selected_date.month
    )

    # --------------------------------------------------------
    # ENTRATE
    # --------------------------------------------------------

    actual_income = (
        month_income_app(
            year,
            month
        )
    )

    default_income = (
        actual_income
        if actual_income > 0
        else 4500.0
    )

    st.subheader(
        "1. Regole del mese"
    )

    c1, c2, c3 = (
        st.columns(3)
    )

    with c1:

        expected_income = st.number_input(
            "Entrate previste (€)",
            min_value=0.0,
            value=float(
                default_income
            ),
            step=100.0,
            help=(
                "Inserisci le entrate "
                "familiari previste "
                "per il mese."
            )
        )

    with c2:

        target_saving_pct = st.slider(
            "Risparmio obiettivo",
            0,
            40,
            15,
            1,
            format="%d%%"
        )

    with c3:

        safety_buffer = st.number_input(
            "Cuscinetto sicurezza (€)",
            min_value=0.0,
            value=100.0,
            step=50.0
        )

    # --------------------------------------------------------
    # IMPEGNI
    # --------------------------------------------------------

    st.subheader(
        "2. Impegni da considerare"
    )

    c1, c2 = (
        st.columns(2)
    )

    with c1:

        include_house = st.checkbox(
            "🏡 Includi costo mensile nuova casa",
            True
        )

        include_recurring = st.checkbox(
            "🔁 Includi spese ricorrenti",
            True
        )

    with c2:

        include_future = st.checkbox(
            "📅 Includi spese future del mese",
            True
        )

        house_period = st.selectbox(
            "Periodo nuova casa",
            [
                "Primi 5 anni",
                "6° anno",
                "7°-10° anno"
            ],
            disabled=not include_house
        )

    # --------------------------------------------------------
    # CALCOLO
    # --------------------------------------------------------

    result = daily_spending_plan(
        target_saving_pct,
        expected_income,
        selected_date,
        include_house,
        include_recurring,
        include_future,
        safety_buffer,
        house_period
    )

    st.markdown("---")

    st.subheader(
        "3. Il risultato"
    )

    c1, c2 = (
        st.columns(2)
    )

    with c1:

        st.metric(
            "💳 Puoi spendere oggi",
            euro(
                result["daily"]
            ),
            help=(
                "Budget medio giornaliero "
                "disponibile da oggi alla "
                "fine del mese."
            )
        )

    with c2:

        st.metric(
            "📆 Budget residuo del mese",
            euro(
                max(
                    0,
                    result["remaining"]
                )
            )
        )

    if result["remaining"] < 0:

        st.error(
            f"⚠️ Il piano è in deficit di "
            f"{euro(abs(result['remaining']))}. "
            "Per rispettare il piano devi "
            "ridurre le spese, aumentare "
            "le entrate o modificare "
            "l'obiettivo di risparmio."
        )

    elif result["daily"] == 0:

        st.warning(
            "Il budget discrezionale "
            "giornaliero è pari a 0 €."
        )

    else:

        st.success(
            f"Con un obiettivo di risparmio "
            f"del {target_saving_pct}%, "
            f"il tuo limite medio è "
            f"**{euro(result['daily'])} al giorno**."
        )

    # --------------------------------------------------------
    # DETTAGLIO CALCOLO
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "4. Come viene calcolato"
    )

    breakdown = pd.DataFrame(
        [
            {
                "Voce":
                    "Entrate previste",

                "Importo":
                    result["income"]
            },

            {
                "Voce":
                    "Spese già sostenute",

                "Importo":
                    -result["actual_spent"]
            },

            {
                "Voce":
                    "Spese ricorrenti",

                "Importo":
                    -result["recurring"]
            },

            {
                "Voce":
                    "Spese future",

                "Importo":
                    -result["future"]
            },

            {
                "Voce":
                    "Nuova casa",

                "Importo":
                    -result["house"]
            },

            {
                "Voce":
                    f"Risparmio obiettivo "
                    f"({target_saving_pct}%)",

                "Importo":
                    -result["saving"]
            },

            {
                "Voce":
                    "Cuscinetto sicurezza",

                "Importo":
                    -result["buffer"]
            },

            {
                "Voce":
                    "Budget residuo",

                "Importo":
                    result["remaining"]
            }
        ]
    )

    st.dataframe(
        breakdown,
        use_container_width=True,
        hide_index=True
    )

    st.caption(
        f"Il calcolo considera "
        f"{result['days_remaining']} giorni "
        "compreso oggi."
    )

    # --------------------------------------------------------
    # CONFRONTO STORICO
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "5. Confronto con le tue abitudini"
    )

    history = (
        historical_average_by_category(
            12
        )
    )

    if history.empty:

        st.info(
            "Lo storico Excel non contiene "
            "abbastanza dati."
        )

    else:

        history_total = (
            history[
                "Media mensile"
            ].sum()
        )

        c1, c2 = (
            st.columns(2)
        )

        c1.metric(
            "Media storica mensile",
            euro(history_total)
        )

        c2.metric(
            "Budget giornaliero",
            euro(result["daily"])
        )

        if history_total > 0:

            history = history.copy()

            history["Peso"] = (
                history[
                    "Media mensile"
                ]
                /
                history_total
            )

            history[
                "Budget guida residuo"
            ] = (
                history["Peso"]
                *
                max(
                    0,
                    result["remaining"]
                )
            )

            st.dataframe(
                history[
                    [
                        "Categoria",
                        "Media mensile",
                        "Budget guida residuo"
                    ]
                ],
                use_container_width=True,
                hide_index=True
            )


# ============================================================
# ANALISI
# ============================================================

elif menu == "📊 Analisi":

    section_title(
        "📊 Analisi finanziaria"
    )

    if (
        excel
        and not excel["family"].empty
    ):

        df = excel[
            "family"
        ]

        monthly = (
            df
            .groupby(
                [
                    "anno",
                    "mese"
                ],
                as_index=False
            )["importo"]
            .sum()
        )

        monthly["Mese"] = (
            monthly.apply(
                lambda row:
                    month_label(
                        int(row["anno"]),
                        int(row["mese"])
                    ),
                axis=1
            )
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

        st.subheader(
            "Categorie con maggiore peso"
        )

        categories = (
            df
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

        st.dataframe(
            categories.rename(
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

    else:

        st.warning(
            "Nessun dato storico disponibile."
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
            "🟢 Il file Excel è la fonte ufficiale "
            "per storico, nuova casa, detrazioni, "
            "mobili e rendering."
        )

    else:

        st.error(
            "Metti 'Spese casa -2.xlsx' "
            "nella stessa cartella di app.py"
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

        "Budget":
            qdf(
                "SELECT * FROM budget"
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
                excel[
                    "house"
                ]["costs"],
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
                excel[
                    "house"
                ]["funds"],
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

        for name, dataframe in (
            tables.items()
        ):

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
