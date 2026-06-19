import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import re
import urllib.request
import io
import base64
from pathlib import Path

st.set_page_config(page_title="MFI Comparative Dashboard", layout="wide", page_icon="📊")


def apply_light_background_image():
    """Apply a lightened full-page background image if available in workspace."""
    candidate_paths = [
        Path(__file__).with_name("background.jpg"),
        Path(__file__).with_name("background.jpeg"),
        Path(__file__).with_name("fusion.jpg"),
    ]

    chosen = next((p for p in candidate_paths if p.exists()), None)
    if not chosen:
        return

    ext = chosen.suffix.lower().replace(".", "")
    if ext == "jpg":
        ext = "jpeg"

    encoded = base64.b64encode(chosen.read_bytes()).decode("utf-8")
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image:
                linear-gradient(rgba(255,255,255,0.80), rgba(255,255,255,0.80)),
                url("data:image/{ext};base64,{encoded}");
            background-size: cover;
            background-position: center;
            background-repeat: no-repeat;
            background-attachment: fixed;
        }}
        [data-testid="stHeader"] {{
            background: rgba(255, 255, 255, 0.55);
            backdrop-filter: blur(2px);
        }}
        [data-testid="stSidebar"] {{
            background: rgba(255, 255, 255, 0.88);
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


apply_light_background_image()

st.markdown("""
<style>
    .main-header {
        font-size: 2rem;
        font-weight: 700;
        color: #1a1a2e;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1rem;
        color: #666;
        margin-bottom: 1.5rem;
    }
    .metric-label {
        font-size: 0.85rem;
        color: #444;
    }
    .stMultiSelect > div {
        border-radius: 8px;
    }
    .upload-section {
        background: #f8f9fa;
        border-radius: 12px;
        padding: 1.5rem;
        margin-bottom: 1rem;
    }
    .cover-slide {
        background: linear-gradient(135deg, #f8fbff 0%, #eef4ff 100%);
        border: 1px solid #dbe6ff;
        border-radius: 16px;
        padding: 1.2rem 1.2rem 1.4rem;
        margin-bottom: 1rem;
        box-shadow: 0 6px 18px rgba(20, 35, 90, 0.08);
    }
    .cover-title {
        text-align: center;
        font-size: 2.2rem;
        font-weight: 800;
        color: #1a2a52;
        margin-bottom: 0.15rem;
    }
    .cover-subtitle {
        text-align: center;
        font-size: 0.95rem;
        color: #4b5a7a;
        margin-bottom: 1rem;
    }
    .cover-logo-caption {
        text-align: center;
        font-size: 0.84rem;
        font-weight: 650;
        color: #2f3f65;
        margin-top: 0.25rem;
        min-height: 2.1rem;
    }
    .logo-fallback {
        height: 72px;
        border: 1px dashed #c8d7ff;
        border-radius: 10px;
        background: #ffffff;
        display: flex;
        align-items: center;
        justify-content: center;
        color: #395a9a;
        font-weight: 800;
        font-size: 0.95rem;
    }
</style>
""", unsafe_allow_html=True)

QUARTERS = ["Q1 FY25", "Q2 FY25", "Q3 FY25", "Q4 FY25", "Q1 FY26", "Q2 FY26", "Q3 FY26", "Q4 FY26"]

# Fixed metrics-table row window from source sheets (0-indexed)
FIXED_TABLE_START_ROW = 50   # row 51 in Excel
FIXED_TABLE_END_ROW = 62     # row 63 in Excel
FIXED_TOTAL_ASSETS_ROW = 48  # row 49 in Excel

FIXED_TABLE_SECTIONS = [
    {"key": "rows_51_63", "title": "Table 1 — Rows 51 to 63", "start": 51, "end": 63},
    {"key": "rows_68_83", "title": "Table 2 — Rows 68 to 83", "start": 68, "end": 83},
    {"key": "rows_85_88", "title": "Table 3 — Rows 85 to 88", "start": 85, "end": 88},
    {"key": "rows_93_104", "title": "Table 4 — Rows 93 to 104", "start": 93, "end": 104},
    {"key": "rows_109_117", "title": "Table 5 — Rows 109 to 117", "start": 109, "end": 117},
    {"key": "rows_122_135", "title": "Table 6 — Rows 122 to 135", "start": 122, "end": 135},
    {"key": "rows_140_150", "title": "Table 7 — Rows 140 to 150", "start": 140, "end": 150},
]

FIXED_CHART_SECTIONS = [
    {
        "key": "aum_disbursement_85_87",
        "title": "💸 AUM & Disbursement (Rows 85 & 87)",
        "rows": [85, 87],
    },
    {
        "key": "total_assets_93_104",
        "title": "🏦 Total Assets Block (Rows 93–104)",
        "start": 93,
        "end": 104,
        "force_percentage": True,
    },
    {
        "key": "key_ratios_109_117",
        "title": "📐 Key Ratios Block (Rows 109–117)",
        "start": 109,
        "end": 117,
    },
    {
        "key": "avg_aum_122_135",
        "title": "📊 Avg AUM Block (Rows 122–135)",
        "start": 122,
        "end": 135,
        "force_percentage": True,
    },
    {
        "key": "rows_140_150",
        "title": "🧾 Rows 140–150",
        "start": 140,
        "end": 150,
        "force_percentage": True,
    },
]

EXCLUDED_CHART_ROWS = {93, 109, 114, 122, 140}
FORCED_PERCENTAGE_ROWS = {115, 116, 117}


def normalize_quarter_label(value: str):
    s = str(value or "").strip().upper().replace("-", " ")
    s = re.sub(r"\s+", " ", s)
    m = re.search(r"Q\s*([1-4])\s*FY\s*(\d{2})", s)
    if not m:
        return None
    return f"Q{m.group(1)} FY{m.group(2)}"


def quarter_sort_key(q: str):
    m = re.match(r"Q([1-4]) FY(\d{2})", str(q).strip())
    if not m:
        return (9999, 9)
    return (int(m.group(2)), int(m.group(1)))

# Row mapping (0-indexed) shared across all files
# Updated to match new Excel file layout (P&L "Absolute Fig" section shifted down by 3 rows)
ROW_MAP = {
    # P&L — from "Absolute Fig." block (rows 69–83 in Excel = 0-indexed 68–82)
    "Interest Income": 69,
    "Other Income": 70,
    "Total Income": 71,
    "Finance Costs (Interest Expense)": 72,
    "Gross Inocme (Net of Expense)": 73,
    "Operating Expenses": 75,
    "PPOP": 76,
    "Credit Cost (Provisions)": 77,
    "Profit Before Tax": 78,
    "Profit After Tax": 79,
    # Metrics — from KPI block (rows 51–63 in Excel = 0-indexed 50–62)
    "AUM (Crores)": 52,
    "GNPA %": 54,
    "NNPA %": 55,
    "CRAR": 57,
    "Branches": 56,
    "Active Clients": 58,
    "Debt to Equity Ratio": 60,
    "NIM": 62,
    "Cost to Income Ratio": 75,   # searched by label via METRIC_PERCENT_ALIASES
    "Total Assets": 48,
    "Basic EPS (₹)": 45,
    # Disbursement dashboard rows (Excel row 85 = 0-indexed 84; row 117 label-searched)
    "Disbursement": 84,
    "Disbursement %": 116,
}

METRIC_GROUPS = {
    "📈 Profitability": [
        "Interest Income", "Other Income", "Total Income",
        "Finance Costs (Interest Expense)", "Gross Inocme (Net of Expense)",
        "Operating Expenses", "PPOP", "Credit Cost (Provisions)",
        "Profit Before Tax", "Profit After Tax", "Basic EPS (₹)"
    ],
    "🏦 Balance Sheet": ["AUM (Crores)", "Debt to Equity Ratio"],
    "⚠️ Asset Quality": ["GNPA %", "NNPA %", "Credit Cost (Provisions)"],
    "📊 Key Ratios": ["NIM", "Cost to Income Ratio", "CRAR", "Opex to Average AUM"],
    "🌐 Operational": ["Branches", "Active Clients"],
    "💸 Disbursement": ["Disbursement", "Disbursement %"],
}

PERCENTAGE_METRICS = {"GNPA %", "NNPA %", "CRAR", "NIM", "Cost to Income Ratio", "Opex to Average AUM", "Disbursement %"}
RATIO_METRICS = {"Debt to Equity Ratio"}

# For these metrics, prefer percentage rows already present in the uploaded file
FILE_PERCENTAGE_METRICS = {
    "Interest Income",
    "Other Income",
    "Total Income",
    "Finance Costs (Interest Expense)",
    "Gross Inocme (Net of Expense)",
    "Operating Expenses",
    "PPOP",
    "Credit Cost (Provisions)",
    "Profit Before Tax",
    "Profit After Tax",
}

METRIC_PERCENT_ALIASES = {
    "Interest Income": ["interest income"],
    "Other Income": ["other income"],
    "Total Income": ["gross income", "total income"],
    "Finance Costs (Interest Expense)": ["interest expense", "finance cost"],
    "Gross Inocme (Net of Expense)": [
        "gross income net of expense",
        "gross inocme net of expense",
        "net income after interest expense",
    ],
    "Operating Expenses": ["operating expense", "operating expenses"],
    "PPOP": ["ppop"],
    "Credit Cost (Provisions)": ["credit cost", "credit cost provisions"],
    "Profit Before Tax": ["profit before tax"],
    "Profit After Tax": ["profit after tax", "profit after tax"],
    "GNPA %": ["gnpa"],
    "NNPA %": ["nnpa"],
    "Cost to Income Ratio": ["cost to income", "cost to income ratio"],
    "CRAR": ["crar", "capital adequacy"],
    "Opex to Average AUM": [
        "opex to average aum",
        "opex to avg aum",
        "operating expenses to average aum",
        "operating expense to average aum",
        "opex average aum",
    ],
}

PRODUCTIVITY_RATIO_METRICS = [
    "AUM per employee",
    "AUM per branch",
    "PAT per employee",
]

PRODUCTIVITY_RATIO_ALIASES = {
    "AUM per employee": ["aum per employee"],
    "AUM per branch": ["aum per branch"],
    "PAT per employee": ["pat per employee", "profit after tax per employee"],
}

FILE_SOURCE_ONLY_METRICS = set(METRIC_PERCENT_ALIASES.keys())

# Metrics that must be read directly from fixed rows (not calculated)
FIXED_ROW_METRICS = {
    "Disbursement": 84,     # row 85 in Excel (0-indexed 84) — updated for new file layout
    "Disbursement %": 116,  # row 117 in Excel — only present in Satin; others rely on label search
}

COLORS = [
    "#5182EF", "#E35151", "#44B56E", "#E19238",
    "#9662F1", "#E25292", "#39A7C1"
]

COMPANY_COLOR_MAP = {
    "creditaccess": "#3C8B72",    # Lighter dark green
    "credit access": "#3C8B72",   # Lighter dark green (variant)
    "muthoot": "#3C63A3",         # Lighter dark blue
    "annapurna": "#7C569C",       # Lighter dark violet
    "spandana": "#A5547B",        # Lighter dark pink
    "spandhan": "#A5547B",        # Lighter dark pink (variant)
    "fusion": "#C3753A",          # Lighter dark orange
    "satin": "#894152",           # Lighter dark maroon
}

COMPANY_LOGO_SOURCES = {
    "Fusion Finance": [
        "https://fusionfin.com/wp-content/uploads/2024/10/cropped-Fusion-Finance-Logo-Final-PNG-3.png",
        "https://logo.clearbit.com/fusionfin.com",
        "https://www.google.com/s2/favicons?sz=256&domain=fusionfin.com",
    ],
    "CreditAccess Grameen": [
        "https://logo.clearbit.com/creditaccessgrameen.in",
        "https://www.google.com/s2/favicons?sz=256&domain=creditaccessgrameen.in",
    ],
    "Muthoot Microfin Limited": [
        "https://logo.clearbit.com/muthootmicrofin.com",
        "https://www.google.com/s2/favicons?sz=256&domain=muthootmicrofin.com",
    ],
    "Annapurna Finance": [
        "https://logo.clearbit.com/annapurnafinance.in",
        "https://www.google.com/s2/favicons?sz=256&domain=annapurnafinance.in",
    ],
    "Satin Creditcare": [
        "https://logo.clearbit.com/satincreditcare.com",
        "https://www.google.com/s2/favicons?sz=256&domain=satincreditcare.com",
    ],
    "Spandhan Sphoorthy": [
        "https://logo.clearbit.com/spandanasphoorty.com",
        "https://www.google.com/s2/favicons?sz=256&domain=spandanasphoorty.com",
    ],
}


@st.cache_data(show_spinner=False)
def fetch_logo_bytes(url: str):
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=6) as resp:
            return resp.read()
    except Exception:
        return None


def get_logo_bytes(company_name: str):
    for url in COMPANY_LOGO_SOURCES.get(company_name, []):
        content = fetch_logo_bytes(url)
        if content:
            return content
    return None


def render_cover_slide():
    st.markdown('<div class="cover-slide">', unsafe_allow_html=True)
    st.markdown('<div class="cover-title">Peer Analysis Report</div>', unsafe_allow_html=True)
    st.markdown('<div class="cover-subtitle">Comparative microfinance performance dashboard</div>', unsafe_allow_html=True)

    ordered_names = [
        "Fusion Finance",
        "CreditAccess Grameen",
        "Muthoot Microfin Limited",
        "Annapurna Finance",
        "Satin Creditcare",
        "Spandhan Sphoorthy",
    ]
    logo_cols = st.columns(6)
    for i, name in enumerate(ordered_names):
        with logo_cols[i]:
            logo_bytes = get_logo_bytes(name)
            if logo_bytes:
                st.image(logo_bytes, width=95)
            else:
                short = short_company_label(name)
                st.markdown(f'<div class="logo-fallback">{short}</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="cover-logo-caption">{name}</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

def get_company_color(company_name: str, fallback_idx: int = 0):
    nm = str(company_name or "").strip().lower()
    for key, color in COMPANY_COLOR_MAP.items():
        if key in nm:
            return color
    return COLORS[fallback_idx % len(COLORS)]


def short_company_label(company_name: str):
    nm = str(company_name or "").strip()
    low = nm.lower()
    if "fusion" in low:
        return "Fusion"
    if "creditaccess" in low:
        return "CreditAccess"
    if "muthoot" in low:
        return "Muthoot"
    if "annapurna" in low:
        return "Annapurna"
    if "satin" in low:
        return "Satin"
    if "spandhan" in low or "spandana" in low:
        return "Spandhan"
    return nm[:14]


PREFERRED_COMPANY_ORDER = [
    "creditaccess",
    "muthoot",
    "satin",
    "annapurna",
    "fusion",
    "spandhan",
]


def company_order_key(company_name: str):
    low = str(company_name or "").strip().lower()
    for idx, token in enumerate(PREFERRED_COMPANY_ORDER):
        if token in low or (token == "spandhan" and "spandana" in low):
            return (idx, low)
    return (len(PREFERRED_COMPANY_ORDER), low)


def sort_companies_by_preference(companies):
    return sorted(companies, key=company_order_key)


def checkbox_selector(options, key_prefix: str, columns: int = 2, default_checked: bool = True):
    """Render a checkbox list and return selected options in original order."""
    if not options:
        return []

    col_count = min(max(1, columns), len(options))
    cols = st.columns(col_count)
    selected = []

    for idx, opt in enumerate(options):
        safe_key = re.sub(r"[^a-zA-Z0-9]+", "_", str(opt)).strip("_").lower()
        widget_key = f"{key_prefix}_{safe_key}"
        if widget_key not in st.session_state:
            st.session_state[widget_key] = default_checked

        with cols[idx % col_count]:
            checked = st.checkbox(str(opt), key=widget_key)

        if checked:
            selected.append(opt)

    return selected


def parse_value(val):
    """Convert cell value to float, handling '-', 'L' suffix, etc."""
    if pd.isna(val):
        return np.nan
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(",", "")
    if s in ("-", "", "nan"):
        return np.nan
    # Handle "L" suffix (lakh) - just strip it, values are already in lakh
    if s.endswith("L"):
        s = s[:-1]
    try:
        return float(s)
    except ValueError:
        return np.nan


def norm_label(s):
    txt = str(s or "").strip().lower()
    txt = txt.replace("%", " ")
    txt = re.sub(r"[^a-z0-9]+", " ", txt)
    return re.sub(r"\s+", " ", txt).strip()


def to_percent_display(val):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return np.nan
    v = float(val)
    return v * 100 if abs(v) <= 1 else v


def find_metric_row_idx(df, aliases):
    max_rows = min(len(df), 220)
    matched_rows = []
    for r in range(max_rows):
        row_vals = [norm_label(df.iloc[r, c]) for c in range(min(8, len(df.columns)))]
        for alias in aliases:
            a = norm_label(alias)
            if any(a and a in cell for cell in row_vals):
                matched_rows.append(r)
                break
    if not matched_rows:
        return None
    # Prefer later occurrence because percentage sections are often placed below absolute sections.
    return matched_rows[-1]


def find_metric_row_idx_from_rows(norm_rows, aliases, prefer_first=False):
    matched_rows = []
    alias_keys = [norm_label(a) for a in aliases]
    for r, row_vals in enumerate(norm_rows):
        if any(a and any(a in cell for cell in row_vals) for a in alias_keys):
            matched_rows.append(r)
    if not matched_rows:
        return None
    return matched_rows[0] if prefer_first else matched_rows[-1]


@st.cache_data(show_spinner=False)
def load_company_data_from_bytes(file_bytes: bytes, file_name: str):
    """Cached parser: returns (company_name, dict of quarter->metric->value)."""
    if not file_bytes:
        return None, None

    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    if not xls.sheet_names:
        return None, None

    first_sheet = xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=first_sheet, header=None)
    all_sheets = {
        sheet_name: pd.read_excel(xls, sheet_name=sheet_name, header=None)
        for sheet_name in xls.sheet_names
    }

    # Detect company name from row 0, and quarter row
    # Company name is in the cell that has actual text
    # Pick the LAST non-empty cell in row 0 as the company name.
    # This handles the Muthoot file where a legacy label appears in col 0
    # and the real company name is further to the right (col 4).
    company_name = None
    for col in range(min(6, len(df.columns))):
        val = df.iloc[0, col]
        if pd.notna(val) and str(val).strip():
            company_name = str(val).strip()
            # do NOT break — keep scanning so we end up with the last non-nan

    # Detect quarter columns dynamically from all sheets (supports future quarters)
    quarter_cols = {}
    for sheet_df in all_sheets.values():
        max_rows = min(120, len(sheet_df))
        for row_idx in range(max_rows):
            row = sheet_df.iloc[row_idx]
            for col_idx, cell in enumerate(row):
                q_lbl = normalize_quarter_label(cell)
                if q_lbl:
                    quarter_cols[q_lbl] = col_idx

    if not quarter_cols:
        return None, None

    max_rows = min(len(df), 220)
    max_cols = min(8, len(df.columns))
    norm_rows = [
        [norm_label(df.iloc[r, c]) for c in range(max_cols)]
        for r in range(max_rows)
    ]
    # For absolute P&L metrics (FILE_SOURCE_ONLY_METRICS), prefer the FIRST label match
    # (the absolute figures section) over the last match (which may be a % / ratio section
    # with near-zero values that would display as ₹0.0 Cr).
    metric_percent_rows = {
        metric: find_metric_row_idx_from_rows(
            norm_rows, aliases, prefer_first=(metric in FILE_SOURCE_ONLY_METRICS)
        )
        for metric, aliases in METRIC_PERCENT_ALIASES.items()
    }
    productivity_ratio_rows = {
        metric: find_metric_row_idx_from_rows(norm_rows, aliases)
        for metric, aliases in PRODUCTIVITY_RATIO_ALIASES.items()
    }

    data = {}
    absolute_opex_by_quarter = {}
    for quarter, col_idx in quarter_cols.items():
        data[quarter] = {}
        for metric, row_idx in ROW_MAP.items():
            if metric in FIXED_ROW_METRICS:
                val = np.nan
                for sheet_df in all_sheets.values():
                    if row_idx < len(sheet_df) and col_idx < len(sheet_df.columns):
                        candidate = parse_value(sheet_df.iloc[row_idx, col_idx])
                        if not (isinstance(candidate, float) and np.isnan(candidate)):
                            val = candidate
                            break
                data[quarter][metric] = val
                continue

            if metric in FILE_SOURCE_ONLY_METRICS:
                data[quarter][metric] = np.nan
            elif row_idx < len(df):
                data[quarter][metric] = parse_value(df.iloc[row_idx, col_idx])
            else:
                data[quarter][metric] = np.nan

        for metric in PRODUCTIVITY_RATIO_METRICS:
            data[quarter][metric] = np.nan

        opex_row_idx = ROW_MAP.get("Operating Expenses")
        if opex_row_idx is not None and opex_row_idx < len(df):
            absolute_opex_by_quarter[quarter] = parse_value(df.iloc[opex_row_idx, col_idx])
        else:
            absolute_opex_by_quarter[quarter] = np.nan

        # Fill with percentage values directly from file (preferred source for these metrics)
        for metric, aliases in METRIC_PERCENT_ALIASES.items():
            row_idx = metric_percent_rows.get(metric)
            if row_idx is not None and row_idx < len(df):
                pct_val = parse_value(df.iloc[row_idx, col_idx])
                if not (isinstance(pct_val, float) and np.isnan(pct_val)):
                    data[quarter][metric] = pct_val

        for metric, row_idx in productivity_ratio_rows.items():
            if row_idx is not None and row_idx < len(df):
                ratio_val = parse_value(df.iloc[row_idx, col_idx])
                if not (isinstance(ratio_val, float) and np.isnan(ratio_val)):
                    data[quarter][metric] = ratio_val

    sorted_quarters = sorted(data.keys(), key=quarter_sort_key)
    for idx, quarter in enumerate(sorted_quarters):
        metric_name = "Opex to Average AUM"
        current_val = data[quarter].get(metric_name, np.nan)
        if pd.notna(current_val):
            continue

        current_aum = data[quarter].get("AUM (Crores)", np.nan)
        current_opex = absolute_opex_by_quarter.get(quarter, np.nan)
        if pd.isna(current_aum) or pd.isna(current_opex) or float(current_aum) == 0:
            data[quarter][metric_name] = np.nan
            continue

        prev_aum = np.nan
        if idx > 0:
            prev_q = sorted_quarters[idx - 1]
            prev_aum = data[prev_q].get("AUM (Crores)", np.nan)

        avg_aum = (float(current_aum) + float(prev_aum)) / 2 if pd.notna(prev_aum) else float(current_aum)
        data[quarter][metric_name] = (float(current_opex) / avg_aum) if avg_aum != 0 else np.nan

    return company_name, data


def load_company_data(uploaded_file):
    """Load a company file and return (company_name, dict of quarter->metric->value)."""
    file_bytes = uploaded_file.getvalue()
    file_name = getattr(uploaded_file, "name", "uploaded_file.xlsx")
    return load_company_data_from_bytes(file_bytes, file_name)


def format_value(val, metric):
    if np.isnan(val) if isinstance(val, float) else False:
        return "N/A"
    if metric in PERCENTAGE_METRICS or metric in FILE_PERCENTAGE_METRICS:
        return f"{to_percent_display(val):.2f}%"
    elif metric in RATIO_METRICS:
        return f"{val:.2f}x"
    else:
        return f"₹{val:,.1f} Cr"


def format_k_short(val):
    if val is None or pd.isna(val):
        return ""
    v = float(val)
    if abs(v) >= 1000:
        return f"<b>{v / 1000:.1f}K</b>"
    return f"<b>{v:,.1f}</b>"


def clean_timeline_label(label):
    if pd.isna(label):
        return None
    txt = str(label).strip()
    if not txt or txt.lower() == "nan":
        return None
    return re.sub(r"\s+", " ", txt)


def format_metrics_table_value(raw_val, timeline_label):
    if pd.isna(raw_val):
        return ""

    s = str(raw_val).strip()
    if not s or s.lower() == "nan":
        return ""

    # Preserve already formatted strings from source (%, L, etc.)
    if any(tok in s for tok in ["%", "L", "l", "₹", "x", "X"]):
        return s

    try:
        v = float(str(s).replace(",", ""))
    except Exception:
        return s

    row_key = norm_label(timeline_label)
    if row_key in {"gnpa", "nnpa", "crar", "nim"}:
        return f"{to_percent_display(v):.2f}%"

    if abs(v - round(v)) < 1e-9:
        return str(int(round(v)))

    return f"{v:.2f}"


def extract_fixed_row_table(df: pd.DataFrame, ordered_quarters, quarter_cols, start_row: int, end_row: int):
    records = []
    for excel_row in range(start_row, end_row + 1):
        row_idx = excel_row - 1  # convert 1-indexed Excel row to 0-index
        if row_idx >= len(df):
            continue

        timeline = None
        for label_col in range(min(3, len(df.columns))):
            timeline = clean_timeline_label(df.iloc[row_idx, label_col])
            if timeline:
                break

        if not timeline:
            timeline = f"Row {excel_row}"

        rec = {"Timeline": timeline}
        for q in ordered_quarters:
            col_idx = quarter_cols[q]
            raw_val = df.iloc[row_idx, col_idx] if col_idx < len(df.columns) else np.nan
            rec[q] = format_metrics_table_value(raw_val, timeline)
        records.append(rec)

    return pd.DataFrame(records)


@st.cache_data(show_spinner=False)
def load_all_fixed_tables_from_bytes(file_bytes: bytes, file_name: str):
    """Extract all requested fixed row-range tables from source sheet."""
    if not file_bytes:
        return None, {}

    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    if not xls.sheet_names:
        return None, {}

    first_sheet = xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=first_sheet, header=None)

    # Pick the LAST non-empty cell in row 0 as the company name.
    # This handles the Muthoot file where a legacy label appears in col 0
    # and the real company name is further to the right (col 4).
    company_name = None
    for col in range(min(6, len(df.columns))):
        val = df.iloc[0, col]
        if pd.notna(val) and str(val).strip():
            company_name = str(val).strip()
            # do NOT break — keep scanning so we end up with the last non-nan

    quarter_cols = {}
    max_rows = min(120, len(df))
    for row_idx in range(max_rows):
        row = df.iloc[row_idx]
        for col_idx, cell in enumerate(row):
            q_lbl = normalize_quarter_label(cell)
            if q_lbl:
                quarter_cols[q_lbl] = col_idx

    if not quarter_cols:
        return company_name, {}

    ordered_quarters = sorted(quarter_cols.keys(), key=quarter_sort_key)
    tables = {}
    for section in FIXED_TABLE_SECTIONS:
        tables[section["key"]] = extract_fixed_row_table(
            df=df,
            ordered_quarters=ordered_quarters,
            quarter_cols=quarter_cols,
            start_row=section["start"],
            end_row=section["end"],
        )

    return company_name, tables


def load_all_fixed_tables(uploaded_file):
    file_bytes = uploaded_file.getvalue()
    file_name = getattr(uploaded_file, "name", "uploaded_file.xlsx")
    return load_all_fixed_tables_from_bytes(file_bytes, file_name)


def extract_fixed_row_numeric_table(
    df: pd.DataFrame,
    ordered_quarters,
    quarter_cols,
    start_row: int = None,
    end_row: int = None,
    rows=None,
):
    row_numbers = []
    if rows is not None:
        row_numbers = list(rows)
    elif start_row is not None and end_row is not None:
        row_numbers = list(range(start_row, end_row + 1))

    records = []
    for excel_row in row_numbers:
        row_idx = excel_row - 1
        if row_idx >= len(df):
            continue

        timeline = None
        for label_col in range(min(3, len(df.columns))):
            timeline = clean_timeline_label(df.iloc[row_idx, label_col])
            if timeline:
                break

        if not timeline:
            timeline = f"Row {excel_row}"

        rec = {"Excel Row": excel_row, "Timeline": timeline}
        for q in ordered_quarters:
            col_idx = quarter_cols[q]
            raw_val = df.iloc[row_idx, col_idx] if col_idx < len(df.columns) else np.nan
            rec[q] = parse_value(raw_val)
        records.append(rec)

    return pd.DataFrame(records)


@st.cache_data(show_spinner=False)
def load_fixed_chart_sections_from_bytes(file_bytes: bytes, file_name: str):
    if not file_bytes:
        return None, {}

    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    if not xls.sheet_names:
        return None, {}

    first_sheet = xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=first_sheet, header=None)

    # Pick the LAST non-empty cell in row 0 as the company name.
    # This handles the Muthoot file where a legacy label appears in col 0
    # and the real company name is further to the right (col 4).
    company_name = None
    for col in range(min(6, len(df.columns))):
        val = df.iloc[0, col]
        if pd.notna(val) and str(val).strip():
            company_name = str(val).strip()
            # do NOT break — keep scanning so we end up with the last non-nan

    quarter_cols = {}
    max_rows = min(120, len(df))
    for row_idx in range(max_rows):
        row = df.iloc[row_idx]
        for col_idx, cell in enumerate(row):
            q_lbl = normalize_quarter_label(cell)
            if q_lbl:
                quarter_cols[q_lbl] = col_idx

    if not quarter_cols:
        return company_name, {}

    ordered_quarters = sorted(quarter_cols.keys(), key=quarter_sort_key)
    sections = {}
    for section in FIXED_CHART_SECTIONS:
        if "rows" in section:
            sections[section["key"]] = extract_fixed_row_numeric_table(
                df=df,
                ordered_quarters=ordered_quarters,
                quarter_cols=quarter_cols,
                rows=section["rows"],
            )
        else:
            sections[section["key"]] = extract_fixed_row_numeric_table(
                df=df,
                ordered_quarters=ordered_quarters,
                quarter_cols=quarter_cols,
                start_row=section["start"],
                end_row=section["end"],
            )

    return company_name, sections


def load_fixed_chart_sections(uploaded_file):
    file_bytes = uploaded_file.getvalue()
    file_name = getattr(uploaded_file, "name", "uploaded_file.xlsx")
    return load_fixed_chart_sections_from_bytes(file_bytes, file_name)


@st.cache_data(show_spinner=False)
def load_metrics_table_from_bytes(file_bytes: bytes, file_name: str):
    """Extract fixed metrics table rows (51-63) + Total Assets row (49)."""
    if not file_bytes:
        return None, pd.DataFrame()

    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    if not xls.sheet_names:
        return None, pd.DataFrame()

    first_sheet = xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=first_sheet, header=None)

    # Pick the LAST non-empty cell in row 0 as the company name.
    # This handles the Muthoot file where a legacy label appears in col 0
    # and the real company name is further to the right (col 4).
    company_name = None
    for col in range(min(6, len(df.columns))):
        val = df.iloc[0, col]
        if pd.notna(val) and str(val).strip():
            company_name = str(val).strip()
            # do NOT break — keep scanning so we end up with the last non-nan

    quarter_cols = {}
    max_rows = min(120, len(df))
    for row_idx in range(max_rows):
        row = df.iloc[row_idx]
        for col_idx, cell in enumerate(row):
            q_lbl = normalize_quarter_label(cell)
            if q_lbl:
                quarter_cols[q_lbl] = col_idx

    if not quarter_cols:
        return company_name, pd.DataFrame()

    ordered_quarters = sorted(quarter_cols.keys(), key=quarter_sort_key)
    row_indices = list(range(FIXED_TABLE_START_ROW, FIXED_TABLE_END_ROW + 1)) + [FIXED_TOTAL_ASSETS_ROW]

    records = []
    for row_idx in row_indices:
        if row_idx >= len(df):
            continue

        timeline = None
        for label_col in range(min(3, len(df.columns))):
            timeline = clean_timeline_label(df.iloc[row_idx, label_col])
            if timeline:
                break

        if not timeline:
            continue

        rec = {"Timeline": timeline}
        for q in ordered_quarters:
            col_idx = quarter_cols[q]
            raw_val = df.iloc[row_idx, col_idx] if col_idx < len(df.columns) else np.nan
            rec[q] = format_metrics_table_value(raw_val, timeline)
        records.append(rec)

    table_df = pd.DataFrame(records)
    if not table_df.empty:
        total_mask = table_df["Timeline"].astype(str).str.lower().str.contains(r"total\s+asset")
        if total_mask.any():
            table_df = pd.concat([table_df.loc[~total_mask], table_df.loc[total_mask]], ignore_index=True)

    return company_name, table_df


def load_metrics_table(uploaded_file):
    file_bytes = uploaded_file.getvalue()
    file_name = getattr(uploaded_file, "name", "uploaded_file.xlsx")
    return load_metrics_table_from_bytes(file_bytes, file_name)


def make_bar_chart(companies_data, selected_companies, selected_quarters, metric):
    fig = go.Figure()

    is_productivity_ratio = metric in PRODUCTIVITY_RATIO_METRICS
    non_percentage_metrics = {"AUM (Crores)", "Disbursement"}
    is_pct = (not is_productivity_ratio) and (metric not in non_percentage_metrics)
    is_ratio = metric in RATIO_METRICS
    use_compact_k_labels = metric in {"AUM (Crores)", "Branches"}

    show_bar_labels = True
    all_numeric_values = []

    for i, company in enumerate(selected_companies):
        if company not in companies_data:
            continue
        values = []
        quarters_available = []
        for q in selected_quarters:
            val = companies_data[company].get(q, {}).get(metric, np.nan)
            if val is None or pd.isna(val):
                display_val = None
            else:
                display_val = to_percent_display(val) if is_pct else float(val)
            values.append(display_val)
            quarters_available.append(q)

        if not any(v is not None for v in values):
            continue

        color = get_company_color(company, i)

        bar_text = None
        bar_text_positions = None
        if show_bar_labels:
            bar_text = []
            bar_text_positions = []
            for v in values:
                if v is None:
                    bar_text.append("")
                    bar_text_positions.append("outside")
                elif is_ratio and not is_pct:
                    bar_text.append(f"<b>{v:.2f}x</b>")
                    bar_text_positions.append("outside" if v >= 0 else "inside")
                elif is_productivity_ratio:
                    bar_text.append(f"<b>{v:.2f}</b>")
                    bar_text_positions.append("outside" if v >= 0 else "inside")
                elif use_compact_k_labels:
                    bar_text.append(f"<b>{v:.2f}</b>")
                    bar_text_positions.append("outside" if v >= 0 else "inside")
                else:
                    bar_text.append(f"<b>{v:.2f}</b>")
                    bar_text_positions.append("outside" if v >= 0 else "inside")

        all_numeric_values.extend([v for v in values if v is not None and not pd.isna(v)])

        fig.add_trace(go.Bar(
            name=company,
            x=[[short_company_label(company)] * len(quarters_available), quarters_available],
            y=values,
            marker=dict(
                color=color,
                line=dict(color="rgba(255,255,255,0.98)", width=2.0),
            ),
            opacity=0.96,
            text=bar_text,
            texttemplate="%{text}",
            textposition=bar_text_positions if show_bar_labels else "none",
            textfont=dict(size=10 if use_compact_k_labels else 12, family="Arial Black", color="#111111"),
            cliponaxis=False,
            customdata=quarters_available,
            hovertemplate=(
                f"<b>{company}</b><br>"
                "Quarter: %{customdata}<br>"
                f"{'Value: %{y:.2f}%<extra></extra>' if is_pct else 'Value: %{y:.2f}<extra></extra>'}"
            ),
        ))

    y_suffix = "%" if is_pct else ""
    if is_pct:
        y_title = metric if "%" in str(metric) else f"{metric} (%)"
    elif is_ratio:
        y_title = f"{metric} (x)"
    else:
        y_title = metric

    y_range = None
    y_dtick = None
    y_nticks = None
    if all_numeric_values:
        y_min = min(all_numeric_values)
        y_max = max(all_numeric_values)
        if y_min >= 0:
            # Add explicit bins for cleaner scaling and keep bars visually shorter
            # when values are tightly clustered near the top.
            if metric in {"NNPA %", "GNPA %"}:
                # Custom y-axis scaling for NNPA % and GNPA % with 0.25 gap
                y_dtick = 0.25
            elif is_pct:
                if y_max <= 25:
                    y_dtick = 5
                elif y_max <= 60:
                    y_dtick = 10
                else:
                    y_dtick = 20
            else:
                target_bins = 8
                raw_step = max(y_max / target_bins, 1e-9)
                base = 10 ** np.floor(np.log10(raw_step))
                norm = raw_step / base
                if norm <= 1:
                    nice_mult = 1
                elif norm <= 2:
                    nice_mult = 2
                elif norm <= 5:
                    nice_mult = 5
                else:
                    nice_mult = 10
                y_dtick = nice_mult * base

            y_upper = np.ceil(y_max / y_dtick) * y_dtick
            y_upper += y_dtick  # base headroom

            # Additional headroom for better visibility of bars.
            # Reduced headroom factors to make bars taller and more visible.
            headroom_factor = 1.2 if is_pct else 1.15
            y_upper = max(y_upper, float(y_max) * headroom_factor)
            y_upper = np.ceil(y_upper / y_dtick) * y_dtick

            y_range = [0, y_upper]
            y_nticks = int((y_upper / y_dtick) + 1)
        else:
            span = max(abs(y_max - y_min), 1.0)
            upper_pad = span * 0.22
            lower_pad = span * 0.12
            y_range = [y_min - lower_pad, y_max + upper_pad]

    fig.update_layout(
        barmode="group",
        barcornerradius=8,
        bargap=0.18,
        bargroupgap=0.08,
        title=dict(text=y_title, x=0.5, xanchor="center", y=0.98, yanchor="top", font=dict(size=16, family="Arial Black", color="#111111")),
        plot_bgcolor="#fcfcfd",
        paper_bgcolor="white",
        font=dict(family="Inter, Arial, sans-serif", size=12, color="#1f2937"),
        xaxis=dict(
            tickfont=dict(size=13, family="Arial Black", color="#111111"),
            tickangle=-25,
            showgrid=False,
            automargin=True,
            ticklabeloverflow="hide past div",
            linecolor="#d1d5db",
            linewidth=1,
        ),
        yaxis=dict(
            ticksuffix=y_suffix,
            showticklabels=False,
            showgrid=False,
            linecolor="#d1d5db",
            linewidth=1,
            zeroline=False,
            automargin=True,
            range=y_range,
            dtick=y_dtick,
            nticks=y_nticks,
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.03,
            xanchor="left",
            x=0,
            font=dict(size=11),
        ),
        margin=dict(l=50, r=20, t=50, b=110),
        height=390,
        uniformtext=dict(minsize=9, mode="show"),
    )

    return fig


PLOTLY_DOWNLOAD_CONFIG = {
    "toImageButtonOptions": {
        "format": "png",
        "filename": "mfi_chart",
        "scale": 2,
    }
}


def render_plotly_chart(fig, use_container_width=True, key=None):
    st.plotly_chart(
        fig,
        use_container_width=use_container_width,
        key=key,
        config=PLOTLY_DOWNLOAD_CONFIG,
    )


def render_metric_group(group_name, metrics, all_companies_data, selected_companies, selected_quarters, charts_per_row):
    st.markdown(f"### {group_name}")

    if charts_per_row == 2:
        for j in range(0, len(metrics), 2):
            row_cols = st.columns(2)
            for k, col in enumerate(row_cols):
                if j + k < len(metrics):
                    metric = metrics[j + k]
                    with col:
                        st.markdown(f"**{metric}**")
                        fig = make_bar_chart(all_companies_data, selected_companies, selected_quarters, metric)
                        render_plotly_chart(fig, use_container_width=True, key=f"{group_name}_{metric}")
    else:
        for metric in metrics:
            st.markdown(f"**{metric}**")
            fig = make_bar_chart(all_companies_data, selected_companies, selected_quarters, metric)
            render_plotly_chart(fig, use_container_width=True, key=f"{group_name}_{metric}")

    st.divider()


def make_source_row_chart(
    all_company_fixed_chart_sections,
    selected_companies,
    selected_quarters,
    section_key,
    row_number,
    row_label,
    force_percentage=False,
):
    fig = go.Figure()
    is_pct = section_key != "aum_disbursement_85_87"
    all_numeric_values = []

    for i, company in enumerate(selected_companies):
        section_df = all_company_fixed_chart_sections.get(company, {}).get(section_key)
        if section_df is None or section_df.empty:
            continue

        row_match = section_df[section_df["Excel Row"] == row_number]
        if row_match.empty:
            continue

        row_data = row_match.iloc[0]
        values = []
        quarters_available = []
        for q in selected_quarters:
            if q not in section_df.columns:
                continue
            val = row_data.get(q, np.nan)
            if pd.isna(val):
                display_val = None
            else:
                numeric_val = float(val)
                display_val = to_percent_display(numeric_val) if is_pct else numeric_val
            values.append(display_val)
            quarters_available.append(q)

        if not any(v is not None for v in values):
            continue

        all_numeric_values.extend([v for v in values if v is not None and not pd.isna(v)])

        bar_text = ["" if v is None else f"<b>{v:.0f}</b>" if not is_pct else f"<b>{v:.2f}</b>" for v in values]
        bar_pos = ["outside" if (v is not None and v >= 0) else "inside" for v in values]

        fig.add_trace(go.Bar(
            name=company,
            x=[[short_company_label(company)] * len(quarters_available), quarters_available],
            y=values,
            marker=dict(
                color=get_company_color(company, i),
                line=dict(color="rgba(255,255,255,0.98)", width=2.0),
            ),
            opacity=0.96,
            text=bar_text,
            texttemplate="%{text}",
            textposition=bar_pos,
            textfont=dict(size=11, family="Arial Black", color="#111111"),
            cliponaxis=False,
            customdata=quarters_available,
            hovertemplate=(
                f"<b>{company}</b><br>"
                "Quarter: %{customdata}<br>"
                f"{'Value: %{y:.2f}%<extra></extra>' if is_pct else 'Value: %{y:.2f}<extra></extra>'}"
            ),
        ))

    y_range = None
    if all_numeric_values:
        y_min = min(all_numeric_values)
        y_max = max(all_numeric_values)
        if y_min >= 0:
            y_upper = max(float(y_max) * 1.35, float(y_max) + 1)
            y_range = [0, y_upper]

    fig.update_layout(
        barmode="group",
        barcornerradius=8,
        bargap=0.18,
        bargroupgap=0.08,
        title=dict(text=f"{row_label}{' (%)' if is_pct else ''}", x=0.5, xanchor="center", y=0.98, yanchor="top", font=dict(size=14, family="Arial Black", color="#111111")),
        plot_bgcolor="#fcfcfd",
        paper_bgcolor="white",
        font=dict(family="Inter, Arial, sans-serif", size=12, color="#1f2937"),
        xaxis=dict(
            tickfont=dict(size=12, family="Arial Black", color="#111111"),
            tickangle=-25,
            showgrid=False,
            automargin=True,
            linecolor="#d1d5db",
            linewidth=1,
        ),
        yaxis=dict(
            ticksuffix="%" if is_pct else "",
            showticklabels=False,
            showgrid=False,
            linecolor="#d1d5db",
            linewidth=1,
            zeroline=False,
            automargin=True,
            range=y_range,
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.03,
            xanchor="left",
            x=0,
            font=dict(size=11),
        ),
        margin=dict(l=50, r=20, t=40, b=100),
        height=360,
    )
    return fig


def render_fixed_section_charts(
    section,
    all_company_fixed_chart_sections,
    selected_companies,
    selected_quarters,
    charts_per_row,
):
    st.markdown(f"### {section['title']}")

    row_meta = []
    for company in selected_companies:
        section_df = all_company_fixed_chart_sections.get(company, {}).get(section["key"])
        if section_df is None or section_df.empty:
            continue
        for _, r in section_df[["Excel Row", "Timeline"]].drop_duplicates().iterrows():
            row_meta.append((int(r["Excel Row"]), str(r["Timeline"])))

    ordered_rows = sorted(
        {
            item[0]: item[1]
            for item in row_meta
            if int(item[0]) not in EXCLUDED_CHART_ROWS
        }.items(),
        key=lambda x: x[0],
    )
    if not ordered_rows:
        st.warning("No chartable rows were found for this section in selected files.")
        return

    if charts_per_row == 2:
        for j in range(0, len(ordered_rows), 2):
            row_cols = st.columns(2)
            for k, col in enumerate(row_cols):
                if j + k < len(ordered_rows):
                    row_no, row_label = ordered_rows[j + k]
                    with col:
                        st.markdown(f"**Row {row_no}: {row_label}**")
                        fig = make_source_row_chart(
                            all_company_fixed_chart_sections=all_company_fixed_chart_sections,
                            selected_companies=selected_companies,
                            selected_quarters=selected_quarters,
                            section_key=section["key"],
                            row_number=row_no,
                            row_label=row_label,
                            force_percentage=section.get("force_percentage", False),
                        )
                        render_plotly_chart(fig, use_container_width=True, key=f"section_{section['key']}_row_{row_no}")
    else:
        for row_no, row_label in ordered_rows:
            st.markdown(f"**Row {row_no}: {row_label}**")
            fig = make_source_row_chart(
                all_company_fixed_chart_sections=all_company_fixed_chart_sections,
                selected_companies=selected_companies,
                selected_quarters=selected_quarters,
                section_key=section["key"],
                row_number=row_no,
                row_label=row_label,
                force_percentage=section.get("force_percentage", False),
            )
            render_plotly_chart(fig, use_container_width=True, key=f"section_{section['key']}_row_{row_no}")

    st.divider()


# ──────────────────────────────────────────────
# SIDEBAR
# ──────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📁 Upload Company Files")
    st.markdown("Upload 2–6 MFI quarterly result files (`.xlsx`)")

    uploaded_files = st.file_uploader(
        "Upload files",
        type=["xlsx"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    st.divider()

    if uploaded_files:
        all_companies_data = {}
        all_company_metric_tables = {}
        all_company_fixed_tables = {}
        all_company_fixed_chart_sections = {}
        all_company_names = []
        available_quarters = set()


        for uf in uploaded_files:
            company_name, data = load_company_data(uf)
            table_company_name, metrics_table_df = load_metrics_table(uf)
            fixed_tables_company_name, fixed_tables = load_all_fixed_tables(uf)
            chart_sections_company_name, chart_sections = load_fixed_chart_sections(uf)

            # Debug output: show extracted company name for each file
            st.write(f"**Extracted company name for file `{getattr(uf, 'name', 'uploaded_file.xlsx')}`:** `{company_name}`")

            resolved_name = company_name or table_company_name or fixed_tables_company_name or chart_sections_company_name

            if company_name and data:
                all_companies_data[company_name] = data
                all_company_names.append(company_name)
                available_quarters.update(data.keys())

            if resolved_name and metrics_table_df is not None and not metrics_table_df.empty:
                all_company_metric_tables[resolved_name] = metrics_table_df

            if resolved_name and fixed_tables:
                all_company_fixed_tables[resolved_name] = fixed_tables

            if resolved_name and chart_sections:
                all_company_fixed_chart_sections[resolved_name] = chart_sections

        all_company_names = sort_companies_by_preference(set(all_company_names))

        available_quarters = sorted(available_quarters, key=quarter_sort_key)

        if all_company_names:
            st.markdown("### 🏢 Tick Companies")
            selected_companies = checkbox_selector(
                options=all_company_names,
                key_prefix="company_select",
                columns=2,
                default_checked=True,
            )
            selected_companies = sort_companies_by_preference(selected_companies)

            st.markdown("### 📅 Tick Quarters")
            selected_quarters = checkbox_selector(
                options=available_quarters,
                key_prefix="quarter_select",
                columns=2,
                default_checked=True,
            )

            st.divider()
            st.markdown("### 📐 Chart Layout")
            charts_per_row = st.radio(
                "Charts per row",
                options=[1, 2],
                index=1,
                horizontal=True,
                label_visibility="visible",
            )


# ──────────────────────────────────────────────
# MAIN CONTENT
# ──────────────────────────────────────────────
render_cover_slide()

st.markdown('<div class="main-header">📊 MFI Comparative Dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Microfinance Institution — Quarterly Performance Comparison</div>', unsafe_allow_html=True)
st.caption("Using direct source rows for disbursement charts: row 85 = Disbursement, row 87 = AUM growth (no calculation).")

if not uploaded_files:
    st.info("👈 **Upload your company Excel files** from the sidebar to get started. Each file should follow the standard MFI quarterly results format.")
    st.markdown("""
    **How to use:**
    1. Upload 2–6 company `.xlsx` files using the sidebar
    2. Select which companies to compare
    3. Select one or more quarters
    4. Explore charts across all metric categories
    """)
    st.stop()

if not all_company_names:
    st.error("Could not parse any company data from the uploaded files. Please check the file format.")
    st.stop()

if len(selected_companies) < 2:
    st.warning("Please select at least 2 companies to compare.")
    st.stop()

if not selected_quarters:
    st.warning("Please select at least one quarter.")
    st.stop()

# Summary KPI row
st.markdown("#### Quick Summary — Latest Quarter")
latest_q = selected_quarters[-1]
cols = st.columns(len(selected_companies))
for i, company in enumerate(selected_companies):
    with cols[i]:
        pat = companies_data_val = all_companies_data[company].get(latest_q, {}).get("Profit After Tax", np.nan)
        aum = all_companies_data[company].get(latest_q, {}).get("AUM (Crores)", np.nan)
        gnpa = all_companies_data[company].get(latest_q, {}).get("GNPA %", np.nan)
        color = get_company_color(company, i)
        st.markdown(f"""
        <div style="border-left: 4px solid {color}; padding: 10px 14px; border-radius: 6px; background: #fafafa; margin-bottom: 4px;">
            <div style="font-weight:700; font-size:1rem; color:{color}">{company}</div>
            <div style="font-size:0.8rem; color:#555; margin-top:4px;">
                PAT: <b>{'₹{:,.1f} Cr'.format(pat) if not np.isnan(pat) else 'N/A'}</b><br>
                AUM: <b>{'₹{:,.0f} Cr'.format(aum) if not np.isnan(aum) else 'N/A'}</b><br>
                GNPA: <b>{'{:.2f}%'.format(to_percent_display(gnpa)) if not np.isnan(gnpa) else 'N/A'}</b>
            </div>
        </div>
        """, unsafe_allow_html=True)

st.divider()

overview_tab, balance_sheet_tab, asset_quality_tab, key_ratios_tab, all_tables_tab, aum_disb_tab, total_assets_rows_tab, key_ratio_rows_tab, avg_aum_rows_tab, rows_140_150_tab = st.tabs([
    "📊 Overview",
    "🏦 Balance Sheet",
    "⚠️ Asset Quality Ratios",
    "📈 Key Ratios",
    "📚 All Source Tables",
    "💸 AUM & Disbursement",
    "🏦 Total Assets Rows 93–104",
    "📐 Key Ratios Rows 109–117",
    "📊 Avg AUM Rows 122–135",
    "🧾 Rows 140–150",
])

with overview_tab:
    render_metric_group(
        "📈 Profitability",
        METRIC_GROUPS["📈 Profitability"],
        all_companies_data,
        selected_companies,
        selected_quarters,
        charts_per_row,
    )
    render_metric_group(
        "🌐 Operational",
        METRIC_GROUPS["🌐 Operational"],
        all_companies_data,
        selected_companies,
        selected_quarters,
        charts_per_row,
    )

with balance_sheet_tab:
    render_metric_group(
        "🏦 Balance Sheet",
        METRIC_GROUPS["🏦 Balance Sheet"],
        all_companies_data,
        selected_companies,
        selected_quarters,
        charts_per_row,
    )

with asset_quality_tab:
    render_metric_group(
        "⚠️ Asset Quality",
        METRIC_GROUPS["⚠️ Asset Quality"],
        all_companies_data,
        selected_companies,
        selected_quarters,
        charts_per_row,
    )

with key_ratios_tab:
    render_metric_group(
        "📊 Key Ratios",
        METRIC_GROUPS["📊 Key Ratios"],
        all_companies_data,
        selected_companies,
        selected_quarters,
        charts_per_row,
    )

with all_tables_tab:
    st.markdown("### 📚 All Source Tables")
    st.caption(
        "Includes fixed row-range tables from source sheets: "
        "51–63, 68–83, 85–88, 93–104, 109–117, 122–135, and 140–150."
    )

    for company in selected_companies:
        st.markdown(f"#### {company}")
        company_tables = all_company_fixed_tables.get(company, {})
        if not company_tables:
            st.warning(f"Could not extract fixed row-range tables for {company}.")
            st.divider()
            continue

        for section in FIXED_TABLE_SECTIONS:
            st.markdown(f"**{section['title']}**")
            table_df = company_tables.get(section["key"])

            if table_df is None or table_df.empty:
                st.info(f"No rows found for {section['start']}–{section['end']}.")
                continue

            quarter_cols = [q for q in selected_quarters if q in table_df.columns]
            display_cols = ["Timeline"] + quarter_cols if quarter_cols else table_df.columns.tolist()
            st.dataframe(table_df[display_cols], use_container_width=True, hide_index=True)

        st.divider()

with aum_disb_tab:
    render_fixed_section_charts(
        section=next(s for s in FIXED_CHART_SECTIONS if s["key"] == "aum_disbursement_85_87"),
        all_company_fixed_chart_sections=all_company_fixed_chart_sections,
        selected_companies=selected_companies,
        selected_quarters=selected_quarters,
        charts_per_row=charts_per_row,
    )

with total_assets_rows_tab:
    render_fixed_section_charts(
        section=next(s for s in FIXED_CHART_SECTIONS if s["key"] == "total_assets_93_104"),
        all_company_fixed_chart_sections=all_company_fixed_chart_sections,
        selected_companies=selected_companies,
        selected_quarters=selected_quarters,
        charts_per_row=charts_per_row,
    )

with key_ratio_rows_tab:
    render_fixed_section_charts(
        section=next(s for s in FIXED_CHART_SECTIONS if s["key"] == "key_ratios_109_117"),
        all_company_fixed_chart_sections=all_company_fixed_chart_sections,
        selected_companies=selected_companies,
        selected_quarters=selected_quarters,
        charts_per_row=charts_per_row,
    )

with avg_aum_rows_tab:
    render_fixed_section_charts(
        section=next(s for s in FIXED_CHART_SECTIONS if s["key"] == "avg_aum_122_135"),
        all_company_fixed_chart_sections=all_company_fixed_chart_sections,
        selected_companies=selected_companies,
        selected_quarters=selected_quarters,
        charts_per_row=charts_per_row,
    )

with rows_140_150_tab:
    render_fixed_section_charts(
        section=next(s for s in FIXED_CHART_SECTIONS if s["key"] == "rows_140_150"),
        all_company_fixed_chart_sections=all_company_fixed_chart_sections,
        selected_companies=selected_companies,
        selected_quarters=selected_quarters,
        charts_per_row=charts_per_row,
    )

# Raw data table
with st.expander("🔍 View Raw Data Table"):
    table_metrics = list(dict.fromkeys(list(ROW_MAP.keys()) + ["Opex to Average AUM"] + PRODUCTIVITY_RATIO_METRICS))
    rows = []
    for company in selected_companies:
        for q in selected_quarters:
            row = {"Company": company, "Quarter": q}
            for metric in table_metrics:
                val = all_companies_data[company].get(q, {}).get(metric, np.nan)
                row[metric] = val
            rows.append(row)
    raw_df = pd.DataFrame(rows)
    for metric in PRODUCTIVITY_RATIO_METRICS:
        if metric in raw_df.columns:
            raw_df[metric] = raw_df[metric].apply(
                lambda x: round(float(x), 2) if pd.notna(x) else x
            )
    st.dataframe(raw_df, use_container_width=True)