import streamlit as st
import pandas as pd
import numpy as np

import re
import ast
import io
import time
import hashlib
from datetime import datetime


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Proof-Carrying Data Analyst",
    page_icon="🤖",
    layout="wide"
)


# ============================================================
# CUSTOM UI
# ============================================================

st.markdown("""
<style>

.main {
    padding-top: 1rem;
}

.hero {
    padding: 30px;
    border-radius: 18px;
    text-align: center;
    border: 1px solid rgba(128,128,128,0.35);
    margin-bottom: 20px;
}

.hero h1 {
    margin-bottom: 5px;
}

.hero p {
    font-size: 17px;
}

.answer-box {
    padding: 25px;
    border-radius: 16px;
    text-align: center;
    border: 2px solid rgba(0,128,0,0.45);
    margin: 20px 0;
}

.answer-value {
    font-size: 38px;
    font-weight: bold;
}

.refusal-box {
    padding: 25px;
    border-radius: 16px;
    border: 2px solid rgba(200,0,0,0.45);
    margin: 20px 0;
}

.step-box {
    padding: 10px;
    border-radius: 10px;
    border: 1px solid rgba(128,128,128,0.25);
    text-align: center;
}

.small {
    font-size: 13px;
    opacity: 0.75;
}

</style>
""", unsafe_allow_html=True)


# ============================================================
# HEADER
# ============================================================

st.markdown("""
<div class="hero">

<h1>🤖 Proof-Carrying Data Analyst</h1>

<p>
HackNex 2026 • HNX26PSI08 • Agentic GenAI
</p>

<p class="small">
Answers questions over messy data using executable,
verified Python code and refuses unreliable answers.
</p>

</div>
""", unsafe_allow_html=True)


st.info(
    "🔐 Every numerical answer must have executable proof. "
    "If the available data is unreliable or ambiguous, "
    "the system returns CANNOT_DETERMINE."
)


# ============================================================
# GEMINI CLIENT
# ============================================================

def get_gemini_client():

    try:

        from google import genai

        api_key = st.secrets.get(
            "GEMINI_API_KEY"
        )

        if not api_key:
            return None

        return genai.Client(
            api_key=api_key
        )

    except Exception:

        return None


# ============================================================
# BASIC HELPERS
# ============================================================

def normalize_name(name):

    name = str(name)

    name = name.strip()

    name = re.sub(
        r"[^a-zA-Z0-9_]",
        "_",
        name
    )

    name = re.sub(
        r"_+",
        "_",
        name
    )

    return name.lower().strip("_")


def normalize_column(name):

    return normalize_name(
        name
    )


def normalize_table_alias(name):

    """
    Allows messy_sales, clean_sales, etc.
    to be recognized as sales-like tables.
    """

    name = normalize_name(
        name
    )

    prefixes = [
        "messy_",
        "clean_",
        "raw_",
        "data_",
        "final_"
    ]

    for prefix in prefixes:

        if name.startswith(prefix):

            name = name[
                len(prefix):
            ]

    return name


def safe_string(value):

    if pd.isna(value):
        return ""

    return str(value).strip()


def make_hash(value):

    text = str(value)

    return hashlib.sha256(
        text.encode("utf-8")
    ).hexdigest()[:12]


# ============================================================
# LOAD CSV
# ============================================================

def load_csv(uploaded_file):

    df = pd.read_csv(
        uploaded_file
    )

    name = normalize_name(
        uploaded_file.name.rsplit(
            ".",
            1
        )[0]
    )

    return {
        name: df
    }


# ============================================================
# LOAD EXCEL
# ============================================================

def load_excel(uploaded_file):

    result = {}

    excel_bytes = uploaded_file.getvalue()

    excel = pd.ExcelFile(
        io.BytesIO(excel_bytes)
    )

    base_name = normalize_name(
        uploaded_file.name.rsplit(
            ".",
            1
        )[0]
    )

    for sheet in excel.sheet_names:

        df = pd.read_excel(
            io.BytesIO(excel_bytes),
            sheet_name=sheet
        )

        sheet_name = normalize_name(
            sheet
        )

        if len(excel.sheet_names) == 1:

            table_name = base_name

        else:

            table_name = (
                base_name
                + "_"
                + sheet_name
            )

        result[
            table_name
        ] = df

    return result


# ============================================================
# LOAD TEXT DOCUMENT
# ============================================================

def load_text(uploaded_file):

    raw = uploaded_file.getvalue()

    text = raw.decode(
        "utf-8",
        errors="ignore"
    )

    return text


# ============================================================
# LOAD PDF
# ============================================================

def load_pdf(uploaded_file):

    try:

        from pypdf import PdfReader

        reader = PdfReader(
            io.BytesIO(
                uploaded_file.getvalue()
            )
        )

        pages = []

        for i, page in enumerate(
            reader.pages
        ):

            text = page.extract_text()

            if text:

                pages.append(
                    f"\n--- PAGE {i + 1} ---\n"
                    f"{text}"
                )

        return "\n".join(
            pages
        )

    except Exception as e:

        return (
            "PDF extraction failed: "
            + str(e)
        )


# ============================================================
# LOAD DOCX
# ============================================================

def load_docx(uploaded_file):

    try:

        from docx import Document

        document = Document(
            io.BytesIO(
                uploaded_file.getvalue()
            )
        )

        paragraphs = []

        for paragraph in document.paragraphs:

            if paragraph.text.strip():

                paragraphs.append(
                    paragraph.text
                )

        return "\n".join(
            paragraphs
        )

    except Exception as e:

        return (
            "DOCX extraction failed: "
            + str(e)
        )


# ============================================================
# LOAD ALL FILES
# ============================================================

def load_uploaded_files(
    uploaded_files
):

    tables = {}

    documents = {}

    errors = []

    for uploaded_file in uploaded_files:

        filename = uploaded_file.name

        lower = filename.lower()

        try:

            if lower.endswith(".csv"):

                loaded = load_csv(
                    uploaded_file
                )

                tables.update(
                    loaded
                )

            elif lower.endswith(".xlsx"):

                loaded = load_excel(
                    uploaded_file
                )

                tables.update(
                    loaded
                )

            elif lower.endswith(
                (".txt", ".md")
            ):

                name = normalize_name(
                    filename.rsplit(
                        ".",
                        1
                    )[0]
                )

                documents[
                    name
                ] = load_text(
                    uploaded_file
                )

            elif lower.endswith(".pdf"):

                name = normalize_name(
                    filename.rsplit(
                        ".",
                        1
                    )[0]
                )

                documents[
                    name
                ] = load_pdf(
                    uploaded_file
                )

            elif lower.endswith(".docx"):

                name = normalize_name(
                    filename.rsplit(
                        ".",
                        1
                    )[0]
                )

                documents[
                    name
                ] = load_docx(
                    uploaded_file
                )

            else:

                errors.append(
                    f"{filename}: "
                    "Unsupported file type."
                )

        except Exception as e:

            errors.append(
                f"{filename}: {str(e)}"
            )

    return (
        tables,
        documents,
        errors
    )


# ============================================================
# COLUMN TYPE DETECTION
# ============================================================

def detect_column_type(
    column,
    series
):

    name = column.lower()

    if (
        "currency" in name
        or "curr" in name
    ):

        return "currency"

    if (
        "date" in name
        or "time" in name
    ):

        return "date"

    if (
        "price" in name
        or "amount" in name
        or "revenue" in name
        or "cost" in name
        or "profit" in name
        or "salary" in name
    ):

        return "monetary"

    if (
        "quantity" in name
        or "qty" in name
        or "count" in name
        or "stock" in name
        or "age" in name
    ):

        return "numeric"

    if pd.api.types.is_numeric_dtype(
        series
    ):

        return "numeric"

    return "text"


# ============================================================
# DATE AMBIGUITY
# ============================================================

def detect_date_issues(
    column,
    series
):

    issues = []

    values = (
        series
        .dropna()
        .astype(str)
        .str.strip()
    )

    if len(values) == 0:

        return issues

    # Explicit slash date pattern
    slash_dates = values[
        values.str.match(
            r"^\d{1,2}/\d{1,2}/\d{2,4}$"
        )
    ]

    if len(slash_dates) > 0:

        issues.append(
            f"{column}: slash-formatted dates "
            "may be ambiguous (DD/MM vs MM/DD)."
        )

    # Mixed date styles
    styles = set()

    for value in values:

        if re.match(
            r"^\d{4}-\d{2}-\d{2}$",
            value
        ):

            styles.add(
                "ISO"
            )

        elif re.match(
            r"^\d{1,2}/\d{1,2}/\d{2,4}$",
            value
        ):

            styles.add(
                "SLASH"
            )

        elif re.match(
            r"^\d{1,2}-\d{1,2}-\d{2,4}$",
            value
        ):

            styles.add(
                "DASH"
            )

    if len(styles) > 1:

        issues.append(
            f"{column}: mixed date formats "
            f"detected ({', '.join(styles)})."
        )

    parsed = pd.to_datetime(
        values,
        errors="coerce"
    )

    invalid = int(
        parsed.isna().sum()
    )

    if invalid > 0:

        issues.append(
            f"{column}: {invalid} "
            "date values cannot be parsed."
        )

    return issues


# ============================================================
# UNIT DETECTION
# ============================================================

def detect_units(
    column,
    series
):

    units = set()

    name = column.lower()

    unit_patterns = {

        "usd": [
            "$",
            "usd"
        ],

        "eur": [
            "€",
            "eur"
        ],

        "inr": [
            "₹",
            "inr",
            "rs"
        ],

        "kg": [
            "kg",
            "kilogram"
        ],

        "g": [
            "gram"
        ],

        "m": [
            "meter"
        ],

        "cm": [
            "centimeter"
        ],

        "km": [
            "kilometer"
        ],

        "percent": [
            "%",
            "percent"
        ]
    }

    for unit, patterns in unit_patterns.items():

        for pattern in patterns:

            if (
                pattern in name
            ):

                units.add(
                    unit
                )

    values = (
        series
        .dropna()
        .astype(str)
        .head(100)
    )

    for value in values:

        value_lower = value.lower()

        for unit, patterns in unit_patterns.items():

            for pattern in patterns:

                if pattern in value_lower:

                    units.add(
                        unit
                    )

    return units


# ============================================================
# DATA QUALITY REPORT
# ============================================================

def profile_table(
    table_name,
    df
):

    duplicate_rows = int(
        df.duplicated().sum()
    )

    missing_total = int(
        df.isna().sum().sum()
    )

    columns = []

    currencies = set()

    units = set()

    date_issues = []

    for column in df.columns:

        series = df[column]

        column_type = detect_column_type(
            column,
            series
        )

        column_units = detect_units(
            column,
            series
        )

        units.update(
            column_units
        )

        if column_type == "currency":

            values = (
                series
                .dropna()
                .astype(str)
                .str.upper()
                .unique()
                .tolist()
            )

            currencies.update(
                values
            )

        if column_type == "date":

            date_issues.extend(
                detect_date_issues(
                    column,
                    series
                )
            )

        columns.append(
            {
                "name": column,
                "type": column_type,
                "missing": int(
                    series.isna().sum()
                ),
                "unique": int(
                    series.nunique(
                        dropna=True
                    )
                ),
                "units": list(
                    column_units
                )
            }
        )

    return {
        "table": table_name,
        "rows": len(df),
        "columns": len(df.columns),
        "duplicate_rows": duplicate_rows,
        "missing_values": missing_total,
        "currencies": sorted(
            currencies
        ),
        "units": sorted(
            units
        ),
        "date_issues": date_issues,
        "column_details": columns
    }


def build_quality_report(
    tables
):

    return [
        profile_table(
            name,
            df
        )
        for name, df in tables.items()
    ]


# ============================================================
# CONTRADICTION DETECTION
# ============================================================

def find_key_columns(
    df
):

    keys = []

    for column in df.columns:

        name = column.lower()

        if (
            name == "id"
            or name.endswith("_id")
            or name.endswith("id")
            or "code" in name
        ):

            keys.append(
                column
            )

    return keys


def compare_tables_for_contradictions(
    tables
):

    contradictions = []

    names = list(
        tables.keys()
    )

    for i in range(
        len(names)
    ):

        for j in range(
            i + 1,
            len(names)
        ):

            name1 = names[i]
            name2 = names[j]

            df1 = tables[name1]
            df2 = tables[name2]

            common_columns = [
                c
                for c in df1.columns
                if c in df2.columns
            ]

            if not common_columns:

                continue

            key_columns = [
                c
                for c in common_columns
                if (
                    c.lower() == "id"
                    or c.lower().endswith("_id")
                    or c.lower().endswith("id")
                    or "code" in c.lower()
                )
            ]

            if not key_columns:

                continue

            # Prefer strongest key
            key = key_columns[0]

            left = df1[
                [key] +
                [
                    c
                    for c in common_columns
                    if c != key
                ]
            ].copy()

            right = df2[
                [key] +
                [
                    c
                    for c in common_columns
                    if c != key
                ]
            ].copy()

            left = left.dropna(
                subset=[key]
            )

            right = right.dropna(
                subset=[key]
            )

            merged = left.merge(
                right,
                on=key,
                suffixes=(
                    "_left",
                    "_right"
                )
            )

            for column in common_columns:

                if column == key:

                    continue

                left_col = (
                    column
                    + "_left"
                )

                right_col = (
                    column
                    + "_right"
                )

                if (
                    left_col not in merged.columns
                    or
                    right_col not in merged.columns
                ):

                    continue

                for _, row in merged.iterrows():

                    a = row[left_col]
                    b = row[right_col]

                    if (
                        pd.isna(a)
                        or
                        pd.isna(b)
                    ):

                        continue

                    if str(a).strip() != str(b).strip():

                        contradictions.append(
                            {
                                "table1": name1,
                                "table2": name2,
                                "key": key,
                                "key_value": row[key],
                                "column": column,
                                "value1": a,
                                "value2": b
                            }
                        )

    return contradictions


# ============================================================
# REFERENTIAL INTEGRITY
# ============================================================

def find_broken_relationships(
    tables
):

    problems = []

    for child_name, child_df in tables.items():

        for column in child_df.columns:

            if not column.lower().endswith(
                "_id"
            ):

                continue

            if column not in child_df.columns:

                continue

            possible_parent = column[
                :-3
            ]

            if possible_parent.endswith("_"):

                possible_parent = possible_parent[
                    :-1
                ]

            for parent_name, parent_df in tables.items():

                if child_name == parent_name:

                    continue

                parent_columns = [
                    c
                    for c in parent_df.columns
                    if normalize_column(c)
                    == normalize_column(column)
                ]

                if not parent_columns:

                    continue

                parent_column = parent_columns[0]

                child_values = set(
                    child_df[column]
                    .dropna()
                    .astype(str)
                )

                parent_values = set(
                    parent_df[parent_column]
                    .dropna()
                    .astype(str)
                )

                missing_keys = (
                    child_values
                    - parent_values
                )

                if missing_keys:

                    problems.append(
                        {
                            "child_table":
                                child_name,
                            "child_column":
                                column,
                            "parent_table":
                                parent_name,
                            "missing_keys":
                                sorted(
                                    list(
                                        missing_keys
                                    )
                                )[:20]
                        }
                    )

                break

    return problems


# ============================================================
# GLOBAL QUALITY ANALYSIS
# ============================================================

def analyze_data_quality(
    tables
):

    report = build_quality_report(
        tables
    )

    contradictions = (
        compare_tables_for_contradictions(
            tables
        )
    )

    relationships = (
        find_broken_relationships(
            tables
        )
    )

    all_currencies = set()

    all_units = set()

    for item in report:

        all_currencies.update(
            item["currencies"]
        )

        all_units.update(
            item["units"]
        )

    return {
        "tables": report,
        "contradictions":
            contradictions,
        "broken_relationships":
            relationships,
        "all_currencies":
            sorted(
                all_currencies
            ),
        "all_units":
            sorted(
                all_units
            )
    }


# ============================================================
# QUESTION TYPE
# ============================================================

def classify_question(
    question
):

    q = question.lower()

    if any(
        x in q
        for x in [
            "predict",
            "forecast",
            "next year",
            "next month",
            "future",
            "will be"
        ]
    ):

        return "prediction"

    if any(
        x in q
        for x in [
            "total revenue",
            "revenue",
            "sales amount",
            "sales value"
        ]
    ):

        return "revenue"

    if any(
        x in q
        for x in [
            "total quantity",
            "total units",
            "units sold"
        ]
    ):

        return "quantity"

    if any(
        x in q
        for x in [
            "how many orders",
            "number of orders",
            "count of orders"
        ]
    ):

        return "order_count"

    if any(
        x in q
        for x in [
            "average",
            "mean"
        ]
    ):

        return "average"

    if any(
        x in q
        for x in [
            "maximum",
            "highest",
            "largest",
            "max"
        ]
    ):

        return "maximum"

    if any(
        x in q
        for x in [
            "minimum",
            "lowest",
            "smallest",
            "min"
        ]
    ):

        return "minimum"

    if any(
        x in q
        for x in [
            "percentage",
            "percent",
            "%"
        ]
    ):

        return "percentage"

    return "general"


# ============================================================
# NUMERIC QUESTION
# ============================================================

def is_numeric_question(
    question
):

    q = question.lower()

    numeric_words = [

        "total",
        "sum",
        "average",
        "mean",
        "count",
        "how many",
        "revenue",
        "sales",
        "quantity",
        "amount",
        "maximum",
        "minimum",
        "highest",
        "lowest",
        "percentage",
        "percent",
        "profit",
        "price",
        "value",
        "stock"
    ]

    return any(
        word in q
        for word in numeric_words
    )


# ============================================================
# TRAP DETECTION
# ============================================================

def detect_traps(
    question,
    quality
):

    traps = []

    q = question.lower()

    # --------------------------------------------------------
    # DUPLICATES
    # --------------------------------------------------------

    duplicate_tables = [
        item["table"]
        for item in quality["tables"]
        if item["duplicate_rows"] > 0
    ]

    if duplicate_tables:

        traps.append(
            "Duplicate rows detected in: "
            + ", ".join(
                duplicate_tables
            )
        )

    # --------------------------------------------------------
    # MISSING VALUES
    # --------------------------------------------------------

    missing_tables = [
        item["table"]
        for item in quality["tables"]
        if item["missing_values"] > 0
    ]

    if missing_tables:

        traps.append(
            "Missing values detected in: "
            + ", ".join(
                missing_tables
            )
        )

    # --------------------------------------------------------
    # CURRENCY
    # --------------------------------------------------------

    currencies = set(
        quality["all_currencies"]
    )

    if len(currencies) > 1:

        traps.append(
            "Multiple currencies detected: "
            + ", ".join(
                sorted(currencies)
            )
        )

    # --------------------------------------------------------
    # UNITS
    # --------------------------------------------------------

    units = set(
        quality["all_units"]
    )

    if len(units) > 1:

        # Not every unit combination is automatically
        # a contradiction. Report it for AI reasoning.

        traps.append(
            "Multiple units detected: "
            + ", ".join(
                sorted(units)
            )
        )

    # --------------------------------------------------------
    # DATES
    # --------------------------------------------------------

    date_issues = []

    for item in quality["tables"]:

        date_issues.extend(
            item["date_issues"]
        )

    if date_issues:

        traps.append(
            "Ambiguous or invalid date values detected."
        )

    # --------------------------------------------------------
    # CONTRADICTIONS
    # --------------------------------------------------------

    contradictions = (
        quality["contradictions"]
    )

    if contradictions:

        sample = contradictions[:3]

        details = []

        for item in sample:

            details.append(
                f"{item['table1']} vs "
                f"{item['table2']}: "
                f"{item['key']}={item['key_value']} "
                f"{item['column']} "
                f"({item['value1']} vs "
                f"{item['value2']})"
            )

        traps.append(
            "Contradictory tables detected. "
            + " | ".join(details)
        )

    # --------------------------------------------------------
    # BROKEN RELATIONSHIPS
    # --------------------------------------------------------

    if quality[
        "broken_relationships"
    ]:

        traps.append(
            "Some foreign-key-like values "
            "do not have matching records."
        )

    # --------------------------------------------------------
    # FUTURE/PREDICTION
    # --------------------------------------------------------

    future_words = [

        "predict",
        "forecast",
        "next year",
        "next month",
        "future",
        "will be"
    ]

    if any(
        word in q
        for word in future_words
    ):

        traps.append(
            "Question asks for a future prediction "
            "not directly supported by historical data."
        )

    # --------------------------------------------------------
    # TRICK / INVALID QUESTIONS
    # --------------------------------------------------------

    if any(
        phrase in q
        for phrase in [
            "prove that",
            "guarantee",
            "exactly what will",
            "can you guarantee"
        ]
    ):

        traps.append(
            "Question requests a guarantee or claim "
            "not directly supported by the data."
        )

    return traps


# ============================================================
# SHOULD REFUSE
# ============================================================

def decide_refusal(
    question,
    quality,
    tables
):

    q = question.lower()

    traps = detect_traps(
        question,
        quality
    )

    question_type = classify_question(
        question
    )

    # --------------------------------------------------------
    # Empty question
    # --------------------------------------------------------

    if not q.strip():

        return True, [
            "No question was provided."
        ]

    # --------------------------------------------------------
    # Prediction
    # --------------------------------------------------------

    if question_type == "prediction":

        return True, [
            "The current system is designed for "
            "data-supported analysis, not unsupported "
            "future prediction."
        ]

    # --------------------------------------------------------
    # No tables
    # --------------------------------------------------------

    if not tables:

        return True, [
            "No tabular data was uploaded."
        ]

    # --------------------------------------------------------
    # Numeric question + severe traps
    # --------------------------------------------------------

    if is_numeric_question(
        question
    ):

        if quality[
            "contradictions"
        ]:

            return True, traps

        if len(
            quality["all_currencies"]
        ) > 1:

            return True, traps

        # Missing values can make a numerical answer
        # unreliable when the question needs those values.

        if any(
            item["missing_values"] > 0
            for item in quality["tables"]
        ):

            return True, traps

        if any(
            item["duplicate_rows"] > 0
            for item in quality["tables"]
        ):

            return True, traps

        # Ambiguous dates matter when question refers to time.

        time_words = [
            "date",
            "month",
            "year",
            "between",
            "during",
            "before",
            "after",
            "january",
            "february",
            "march",
            "april",
            "may",
            "june",
            "july",
            "august",
            "september",
            "october",
            "november",
            "december"
        ]

        has_date_issue = any(
            item["date_issues"]
            for item in quality["tables"]
        )

        if (
            has_date_issue
            and any(
                word in q
                for word in time_words
            )
        ):

            return True, traps

    return False, traps


# ============================================================
# DATA DESCRIPTION
# ============================================================

def create_dataset_description(
    tables,
    documents
):

    sections = []

    for table_name, df in tables.items():

        sections.append(
            f"""
TABLE: {table_name}

Rows: {len(df)}

Columns:
{list(df.columns)}

Data types:
{df.dtypes.to_string()}

First rows:
{df.head(8).to_string(index=False)}
"""
        )

    for document_name, text in documents.items():

        limited_text = text[:12000]

        sections.append(
            f"""
DOCUMENT: {document_name}

CONTENT:
{limited_text}
"""
        )

    return "\n".join(
        sections
    )


# ============================================================
# CODE CLEANING
# ============================================================

def clean_code(
    code
):

    if not code:

        return ""

    code = code.strip()

    code = re.sub(
        r"```python",
        "",
        code,
        flags=re.IGNORECASE
    )

    code = re.sub(
        r"```",
        "",
        code
    )

    lines = []

    for line in code.splitlines():

        stripped = line.strip()

        if stripped.startswith(
            "import "
        ):

            continue

        if stripped.startswith(
            "from "
        ):

            continue

        lines.append(
            line
        )

    return "\n".join(
        lines
    ).strip()


# ============================================================
# AST SAFETY VALIDATION
# ============================================================

ALLOWED_AST_NODES = {

    ast.Module,
    ast.Assign,
    ast.AnnAssign,
    ast.Expr,
    ast.Name,
    ast.Load,
    ast.Store,

    ast.Constant,

    ast.BinOp,
    ast.UnaryOp,

    ast.Add,
    ast.Sub,
    ast.Mult,
    ast.Div,
    ast.FloorDiv,
    ast.Mod,
    ast.Pow,

    ast.USub,
    ast.UAdd,

    ast.Compare,
    ast.Eq,
    ast.NotEq,
    ast.Lt,
    ast.LtE,
    ast.Gt,
    ast.GtE,

    ast.BoolOp,
    ast.And,
    ast.Or,

    ast.IfExp,

    ast.Call,
    ast.Attribute,

    ast.Subscript,
    ast.Slice,

    ast.List,
    ast.Tuple,
    ast.Dict,
    ast.Set,

    ast.Index,

    ast.keyword
}


BLOCKED_NAMES = {

    "open",
    "eval",
    "exec",
    "compile",
    "__import__",
    "input",

    "globals",
    "locals",
    "vars",

    "getattr",
    "setattr",
    "delattr",

    "help",
    "dir",

    "breakpoint",

    "exit",
    "quit"
}


BLOCKED_ATTRIBUTES = {

    "system",
    "popen",
    "remove",
    "rmdir",
    "unlink",

    "subprocess",
    "socket",

    "__class__",
    "__dict__",
    "__globals__",
    "__builtins__",

    "to_pickle",
    "to_sql",
    "read_pickle"
}


def validate_code(
    code
):

    try:

        tree = ast.parse(
            code
        )

    except SyntaxError as e:

        return False, (
            "Syntax error: "
            + str(e)
        )

    for node in ast.walk(
        tree
    ):

        node_type = type(node)

        if node_type not in ALLOWED_AST_NODES:

            return False, (
                "Unsupported Python operation: "
                + node_type.__name__
            )

        if isinstance(
            node,
            ast.Name
        ):

            if node.id in BLOCKED_NAMES:

                return False, (
                    "Blocked name: "
                    + node.id
                )

        if isinstance(
            node,
            ast.Attribute
        ):

            if node.attr in BLOCKED_ATTRIBUTES:

                return False, (
                    "Blocked attribute: "
                    + node.attr
                )

    # Explicitly reject imports
    for node in ast.walk(tree):

        if isinstance(
            node,
            (
                ast.Import,
                ast.ImportFrom
            )
        ):

            return False, (
                "Imports are not allowed."
            )

    return True, None


# ============================================================
# SAFE EXECUTION
# ============================================================

def execute_code(
    code,
    tables
):

    safe, reason = validate_code(
        code
    )

    if not safe:

        return {
            "success": False,
            "result": None,
            "error": reason
        }

    safe_builtins = {

        "len": len,
        "sum": sum,
        "min": min,
        "max": max,

        "abs": abs,
        "round": round,

        "float": float,
        "int": int,
        "str": str,

        "list": list,
        "dict": dict,
        "set": set,

        "range": range,
        "enumerate": enumerate,

        "True": True,
        "False": False,
        "None": None
    }

    safe_globals = {

        "__builtins__":
            safe_builtins,

        "pd": pd,
        "np": np
    }

    for table_name, df in tables.items():

        safe_globals[
            table_name
        ] = df.copy()

    local_vars = {}

    try:

        exec(
            code,
            safe_globals,
            local_vars
        )

        if "result" not in local_vars:

            return {
                "success": False,
                "result": None,
                "error":
                    "Generated code did not create "
                    "the required 'result' variable."
            }

        result = local_vars[
            "result"
        ]

        if result is None:

            return {
                "success": False,
                "result": None,
                "error":
                    "Generated code returned None."
            }

        return {
            "success": True,
            "result": result,
            "error": None
        }

    except Exception as e:

        return {
            "success": False,
            "result": None,
            "error": str(e)
        }


# ============================================================
# EXECUTE TWICE FOR REPRODUCIBILITY
# ============================================================

def reproducibility_check(
    code,
    tables
):

    first = execute_code(
        code,
        tables
    )

    if not first["success"]:

        return {
            "passed": False,
            "first": first,
            "second": None,
            "message":
                "First execution failed."
        }

    second = execute_code(
        code,
        tables
    )

    if not second["success"]:

        return {
            "passed": False,
            "first": first,
            "second": second,
            "message":
                "Second execution failed."
        }

    first_value = first[
        "result"
    ]

    second_value = second[
        "result"
    ]

    try:

        equal = np.isclose(
            float(first_value),
            float(second_value)
        )

    except Exception:

        equal = (
            str(first_value)
            == str(second_value)
        )

    return {
        "passed": bool(equal),
        "first": first,
        "second": second,
        "message":
            "Code produced the same result "
            "on repeated execution."
            if equal
            else
            "Code produced different results."
    }


# ============================================================
# FIND TABLE BY ALIAS
# ============================================================

def find_table(
    tables,
    desired
):

    desired = normalize_table_alias(
        desired
    )

    for name, df in tables.items():

        if normalize_table_alias(
            name
        ) == desired:

            return df

    return None


# ============================================================
# INDEPENDENT VERIFIER
# ============================================================

def independent_verify(
    question,
    tables
):

    q = question.lower()

    # --------------------------------------------------------
    # TOTAL REVENUE
    # --------------------------------------------------------

    if (
        "total revenue" in q
        or "total sales" in q
    ):

        sales = find_table(
            tables,
            "sales"
        )

        products = find_table(
            tables,
            "products"
        )

        if sales is None:

            return {
                "verified": False,
                "value": None,
                "message":
                    "No sales table found."
            }

        if products is None:

            return {
                "verified": False,
                "value": None,
                "message":
                    "No products table found."
            }

        required_sales = [
            "product_id",
            "quantity"
        ]

        required_products = [
            "product_id",
            "price"
        ]

        missing_sales = [
            c
            for c in required_sales
            if c not in sales.columns
        ]

        missing_products = [
            c
            for c in required_products
            if c not in products.columns
        ]

        if missing_sales:

            return {
                "verified": False,
                "value": None,
                "message":
                    "Missing sales columns: "
                    + ", ".join(
                        missing_sales
                    )
            }

        if missing_products:

            return {
                "verified": False,
                "value": None,
                "message":
                    "Missing product columns: "
                    + ", ".join(
                        missing_products
                    )
            }

        merged = sales.merge(
            products[
                [
                    "product_id",
                    "price"
                ]
            ],
            on="product_id",
            how="left",
            validate="many_to_one"
        )

        if merged[
            "price"
        ].isna().any():

            return {
                "verified": False,
                "value": None,
                "message":
                    "Some products have no matching price."
            }

        if merged[
            "quantity"
        ].isna().any():

            return {
                "verified": False,
                "value": None,
                "message":
                    "Some sales records have missing quantity."
            }

        value = (
            merged["quantity"]
            * merged["price"]
        ).sum()

        return {
            "verified": True,
            "value": float(value),
            "message":
                "Independent calculation: "
                "sum(quantity × product price)."
        }

    # --------------------------------------------------------
    # TOTAL QUANTITY
    # --------------------------------------------------------

    if (
        "total quantity" in q
        or "total units" in q
        or "units sold" in q
    ):

        sales = find_table(
            tables,
            "sales"
        )

        if sales is None:

            return {
                "verified": False,
                "value": None,
                "message":
                    "No sales table found."
            }

        if "quantity" not in sales.columns:

            return {
                "verified": False,
                "value": None,
                "message":
                    "quantity column not found."
            }

        value = sales[
            "quantity"
        ].sum()

        return {
            "verified": True,
            "value": float(value),
            "message":
                "Independent calculation: "
                "sum(sales.quantity)."
        }

    # --------------------------------------------------------
    # ORDER COUNT
    # --------------------------------------------------------

    if (
        "number of orders" in q
        or "how many orders" in q
        or "count of orders" in q
    ):

        sales = find_table(
            tables,
            "sales"
        )

        if sales is None:

            return {
                "verified": False,
                "value": None,
                "message":
                    "No sales table found."
            }

        if "order_id" not in sales.columns:

            return {
                "verified": False,
                "value": None,
                "message":
                    "order_id column not found."
            }

        value = sales[
            "order_id"
        ].nunique()

        return {
            "verified": True,
            "value": int(value),
            "message":
                "Independent calculation: "
                "number of unique order IDs."
        }

    # --------------------------------------------------------
    # GENERAL QUESTION
    # --------------------------------------------------------

    return {
        "verified": None,
        "value": None,
        "message":
            "No specialized independent verifier "
            "is available for this question."
    }


# ============================================================
# GEMINI PROMPT
# ============================================================

def ask_gemini(
    question,
    dataset_description,
    quality
):

    client = get_gemini_client()

    if client is None:

        return (
            None,
            "Gemini API is not configured."
        )

    quality_summary = []

    for item in quality["tables"]:

        quality_summary.append(
            f"""
Table: {item['table']}
Rows: {item['rows']}
Columns: {item['columns']}
Duplicates: {item['duplicate_rows']}
Missing values: {item['missing_values']}
Currencies: {item['currencies']}
Units: {item['units']}
Date issues: {item['date_issues']}
"""
        )

    contradiction_text = ""

    if quality[
        "contradictions"
    ]:

        contradiction_text = (
            "\nCONTRADICTIONS:\n"
        )

        for item in quality[
            "contradictions"
        ][:10]:

            contradiction_text += (
                f"- {item['table1']} vs "
                f"{item['table2']}: "
                f"{item['key']}="
                f"{item['key_value']}, "
                f"{item['column']}: "
                f"{item['value1']} vs "
                f"{item['value2']}\n"
            )

    prompt = f"""
You are the reasoning engine of a Proof-Carrying Data Analyst
for HackNex 2026 problem HNX26PSI08.

USER QUESTION:
{question}

DATA:
{dataset_description}

DATA QUALITY:
{''.join(quality_summary)}

{contradiction_text}

CRITICAL RULES:

1. Use ONLY the supplied data.
2. Never invent missing values.
3. Never invent exchange rates.
4. Never silently choose between contradictory tables.
5. Never silently remove duplicates.
6. Never silently interpret an ambiguous date.
7. Never make unsupported predictions.
8. If the answer is unreliable, return CANNOT_DETERMINE.
9. Every numerical answer MUST be calculated by Python.
10. The Python code must be executable.
11. pandas is already available as pd.
12. numpy is already available as np.
13. Do NOT import anything.
14. The uploaded DataFrames are available by table name.
15. Store the final answer in a variable called result.
16. Do not hard-code the answer.
17. Do not use external websites or APIs.
18. Use joins when multiple tables are required.
19. Check whether joins can duplicate rows.
20. Respect currency and unit consistency.
21. If a question has no valid answer, refuse it.
22. If a question asks for future prediction and there is no
    forecasting model or sufficient basis, refuse it.

IMPORTANT:
If the data quality information shows a trap that affects
the requested answer, return CANNOT_DETERMINE.

OUTPUT FORMAT:

For an answer:

PYTHON_CODE:
<only executable Python code>

For an unreliable or unanswerable question:

CANNOT_DETERMINE:
<clear reason>

Do not provide any other format.
"""

    models = [
        "gemini-3.8-flash",
        "gemini-3.7-flash",
        "gemini-3.6-flash",
        "gemini-3.5-flash-lite"
    ]

    for model in models:

        for attempt in range(2):

            try:

                response = (
                    client.models.generate_content(
                        model=model,
                        contents=prompt
                    )
                )

                return (
                    response.text.strip(),
                    None
                )

            except Exception as e:

                error = str(e)

                if (
                    "503" in error
                    or "UNAVAILABLE"
                    in error
                    or "high demand"
                    in error.lower()
                ):

                    time.sleep(2)

                    continue

                return (
                    None,
                    error
                )

    return (
        None,
        "All Gemini models are temporarily unavailable."
    )


# ============================================================
# DISPLAY QUALITY
# ============================================================

def display_quality(
    quality
):

    st.subheader(
        "🔍 Data Quality & Trap Detection"
    )

    for item in quality[
        "tables"
    ]:

        st.markdown(
            f"### 📄 {item['table']}"
        )

        c1, c2, c3, c4 = st.columns(4)

        with c1:

            st.metric(
                "Rows",
                item["rows"]
            )

        with c2:

            st.metric(
                "Columns",
                item["columns"]
            )

        with c3:

            st.metric(
                "Duplicates",
                item["duplicate_rows"]
            )

        with c4:

            st.metric(
                "Missing",
                item["missing_values"]
            )

        if item[
            "currencies"
        ]:

            st.write(
                "💰 Currencies:",
                ", ".join(
                    item["currencies"]
                )
            )

        if item[
            "units"
        ]:

            st.write(
                "📏 Units:",
                ", ".join(
                    item["units"]
                )
            )

        if item[
            "date_issues"
        ]:

            for issue in item[
                "date_issues"
            ]:

                st.warning(
                    "📅 " + issue
                )

    if quality[
        "contradictions"
    ]:

        st.error(
            f"⚠️ {len(quality['contradictions'])} "
            "table contradiction(s) detected."
        )

        with st.expander(
            "View contradictions"
        ):

            for item in quality[
                "contradictions"
            ][:20]:

                st.write(
                    f"**{item['table1']}** vs "
                    f"**{item['table2']}** | "
                    f"{item['key']} = "
                    f"{item['key_value']} | "
                    f"{item['column']}: "
                    f"{item['value1']} ≠ "
                    f"{item['value2']}"
                )

    if quality[
        "broken_relationships"
    ]:

        st.warning(
            "🔗 Referential integrity problems detected."
        )

        with st.expander(
            "View relationship problems"
        ):

            for item in quality[
                "broken_relationships"
            ]:

                st.write(
                    item
                )


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.title(
    "📁 Data & Documents"
)

st.sidebar.write(
    "Upload messy tables and supporting documents."
)

uploaded_files = st.sidebar.file_uploader(
    "Choose files",
    type=[
        "csv",
        "xlsx",
        "txt",
        "md",
        "pdf",
        "docx"
    ],
    accept_multiple_files=True
)


# ============================================================
# LOAD DATA
# ============================================================

tables = {}
documents = {}
load_errors = []

if uploaded_files:

    (
        tables,
        documents,
        load_errors
    ) = load_uploaded_files(
        uploaded_files
    )

    for error in load_errors:

        st.error(
            error
        )


# ============================================================
# NO DATA
# ============================================================

if not tables and not documents:

    st.info(
        "👈 Upload CSV, Excel, PDF, DOCX, TXT or Markdown files."
    )

    st.markdown("""
### 🧪 Recommended HackNex demo

Upload:

**Normal**
- sales.csv
- products.csv
- customers.csv

Ask:

> What is the total revenue?

Then upload the messy files and ask the same question.

The system should detect:

- duplicate rows
- missing quantity
- currency mismatch
- ambiguous dates

and refuse the answer.
""")

    st.stop()


# ============================================================
# DATA SUMMARY
# ============================================================

st.success(
    f"✅ Loaded {len(tables)} table(s) "
    f"and {len(documents)} document(s)."
)


# ============================================================
# WORKFLOW
# ============================================================

st.subheader(
    "🔄 Agentic Workflow"
)

steps = [
    "Upload",
    "Understand",
    "Detect Traps",
    "Generate Code",
    "Execute",
    "Verify",
    "Answer / Refuse"
]

cols = st.columns(
    len(steps)
)

for i, step in enumerate(
    steps
):

    with cols[i]:

        st.markdown(
            f"""
            <div class="step-box">
            <b>{i + 1}</b><br>
            {step}
            </div>
            """,
            unsafe_allow_html=True
        )


# ============================================================
# TABLE PREVIEW
# ============================================================

if tables:

    st.subheader(
        "📊 Uploaded Tables"
    )

    for name, df in tables.items():

        with st.expander(
            f"📄 {name} "
            f"({len(df)} rows × "
            f"{len(df.columns)} columns)"
        ):

            st.dataframe(
                df,
                use_container_width=True
            )


# ============================================================
# DOCUMENT PREVIEW
# ============================================================

if documents:

    st.subheader(
        "📚 Uploaded Documents"
    )

    for name, text in documents.items():

        with st.expander(
            f"📄 {name}"
        ):

            st.text(
                text[:5000]
            )


# ============================================================
# QUALITY ANALYSIS
# ============================================================

quality = analyze_data_quality(
    tables
)

display_quality(
    quality
)


# ============================================================
# DATA DESCRIPTION
# ============================================================

dataset_description = (
    create_dataset_description(
        tables,
        documents
    )
)


# ============================================================
# QUESTION
# ============================================================

st.subheader(
    "💬 Ask Your Data"
)

question = st.text_input(
    "Enter your question",
    placeholder=
    "Example: What is the total revenue?"
)

analyze_button = st.button(
    "🚀 Analyze & Verify",
    type="primary",
    use_container_width=True
)


# ============================================================
# ANALYSIS
# ============================================================

if analyze_button:

    if not question.strip():

        st.warning(
            "Please enter a question."
        )

        st.stop()

    # --------------------------------------------------------
    # STEP 1: TRAP ANALYSIS
    # --------------------------------------------------------

    st.subheader(
        "🛡️ Trap Analysis"
    )

    refusal, reasons = decide_refusal(
        question,
        quality,
        tables
    )

    if reasons:

        for reason in reasons:

            st.warning(
                "⚠️ " + reason
            )

    # --------------------------------------------------------
    # REFUSE
    # --------------------------------------------------------

    if refusal:

        st.markdown(
            """
            <div class="refusal-box">

            <h2>❌ CANNOT_DETERMINE</h2>

            <p>
            The system refused to provide a numerical
            answer because the available evidence is
            not reliable enough.
            </p>

            </div>
            """,
            unsafe_allow_html=True
        )

        st.write(
            "### Why the system refused:"
        )

        for reason in reasons:

            st.write(
                "• " + reason
            )

        st.info(
            "This is intentional. For this problem, "
            "a justified refusal is safer than a confident "
            "wrong answer."
        )

        st.stop()

    # --------------------------------------------------------
    # GEMINI
    # --------------------------------------------------------

    with st.spinner(
        "🤖 Agent is reasoning about the data..."
    ):

        response, error = ask_gemini(
            question,
            dataset_description,
            quality
        )

    if error:

        st.error(
            "Gemini error:"
        )

        st.code(
            error
        )

        st.stop()

    if not response:

        st.error(
            "No AI response received."
        )

        st.stop()

    # --------------------------------------------------------
    # AI REFUSAL
    # --------------------------------------------------------

    if response.upper().startswith(
        "CANNOT_DETERMINE"
    ):

        reason = response

        st.markdown(
            """
            <div class="refusal-box">

            <h2>❌ CANNOT_DETERMINE</h2>

            </div>
            """,
            unsafe_allow_html=True
        )

        st.warning(
            reason
        )

        st.stop()

    # --------------------------------------------------------
    # EXTRACT CODE
    # --------------------------------------------------------

    if "PYTHON_CODE:" in response:

        code = response.split(
            "PYTHON_CODE:",
            1
        )[1].strip()

    else:

        code = response.strip()

    code = clean_code(
        code
    )

    if not code:

        st.error(
            "AI did not generate executable code."
        )

        st.stop()

    # --------------------------------------------------------
    # CODE SAFETY
    # --------------------------------------------------------

    st.subheader(
        "🔐 Code Safety Check"
    )

    safe, safety_reason = (
        validate_code(
            code
        )
    )

    if not safe:

        st.error(
            "❌ Generated code rejected."
        )

        st.warning(
            safety_reason
        )

        st.stop()

    st.success(
        "✅ Generated code passed AST safety validation."
    )

    # --------------------------------------------------------
    # SHOW CODE
    # --------------------------------------------------------

    st.subheader(
        "🐍 Generated Proof Code"
    )

    st.code(
        code,
        language="python"
    )

    # --------------------------------------------------------
    # EXECUTE
    # --------------------------------------------------------

    with st.spinner(
        "⚙️ Executing proof code..."
    ):

        execution = execute_code(
            code,
            tables
        )

    if not execution[
        "success"
    ]:

        st.error(
            "❌ Code execution failed."
        )

        st.code(
            execution["error"]
        )

        st.warning(
            "The system refuses to accept a number "
            "when its proof code does not execute."
        )

        st.stop()

    ai_result = execution[
        "result"
    ]

    st.success(
        "✅ Proof code executed successfully."
    )

    # --------------------------------------------------------
    # RESULT
    # --------------------------------------------------------

    st.subheader(
        "📊 Code Result"
    )

    st.write(
        ai_result
    )

    # --------------------------------------------------------
    # REPRODUCIBILITY
    # --------------------------------------------------------

    st.subheader(
        "🔁 Reproducibility Check"
    )

    reproducibility = (
        reproducibility_check(
            code,
            tables
        )
    )

    if reproducibility[
        "passed"
    ]:

        st.success(
            "✅ Reproducibility PASSED"
        )

        st.write(
            reproducibility["message"]
        )

    else:

        st.error(
            "❌ Reproducibility FAILED"
        )

        st.stop()

    # --------------------------------------------------------
    # INDEPENDENT VERIFICATION
    # --------------------------------------------------------

    st.subheader(
        "🔎 Independent Verification"
    )

    verification = (
        independent_verify(
            question,
            tables
        )
    )

    if verification[
        "verified"
    ] is True:

        st.write(
            "Independent calculation:",
            verification["value"]
        )

        try:

            ai_number = float(
                ai_result
            )

            verified_number = float(
                verification["value"]
            )

            difference = abs(
                ai_number
                - verified_number
            )

            if difference < 0.000001:

                st.success(
                    "✅ INDEPENDENT VERIFICATION PASSED"
                )

                st.caption(
                    verification["message"]
                )

            else:

                st.error(
                    "❌ INDEPENDENT VERIFICATION FAILED"
                )

                st.error(
                    f"AI result: {ai_number}"
                )

                st.error(
                    f"Independent result: "
                    f"{verified_number}"
                )

                st.stop()

        except Exception:

            if (
                str(ai_result)
                != str(
                    verification["value"]
                )
            ):

                st.error(
                    "❌ Verification mismatch."
                )

                st.stop()

    else:

        st.info(
            "ℹ️ No specialized independent "
            "verifier exists for this question type. "
            "The result was still checked by executing "
            "the generated proof code twice."
        )

    # --------------------------------------------------------
    # FINAL ANSWER
    # --------------------------------------------------------

    st.subheader(
        "🏆 Final Answer"
    )

    if verification[
        "verified"
    ] is True:

        final_value = (
            verification["value"]
        )

        verification_text = (
            "Independently verified"
        )

    else:

        final_value = ai_result

        verification_text = (
            "Verified by repeated execution "
            "of the generated proof code"
        )

    st.markdown(
        f"""
        <div class="answer-box">

        <h2>✅ Verified Answer</h2>

        <div class="answer-value">
        {final_value}
        </div>

        <p>
        {verification_text}
        </p>

        </div>
        """,
        unsafe_allow_html=True
    )

    # --------------------------------------------------------
    # PROOF
    # --------------------------------------------------------

    st.subheader(
        "📜 Reproducible Proof"
    )

    st.write(
        "Anyone can inspect and rerun the following "
        "Python code on the same data:"
    )

    st.code(
        code,
        language="python"
    )

    # --------------------------------------------------------
    # EVIDENCE
    # --------------------------------------------------------

    st.subheader(
        "🧾 Verification Evidence"
    )

    evidence = {

        "Question":
            question,

        "Question type":
            classify_question(
                question
            ),

        "Tables used":
            list(
                tables.keys()
            ),

        "Documents used":
            list(
                documents.keys()
            ),

        "Proof code hash":
            make_hash(
                code
            ),

        "Execution result":
            str(
                ai_result
            ),

        "Reproducibility":
            "PASSED"
            if reproducibility[
                "passed"
            ]
            else
            "FAILED",

        "Independent verification":
            (
                "PASSED"
                if verification[
                    "verified"
                ] is True
                else
                "N/A"
            )
    }

    st.json(
        evidence
    )

   