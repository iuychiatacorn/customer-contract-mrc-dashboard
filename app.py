import os
import streamlit as st
import pandas as pd
import plotly.express as px

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="Customer Tracking Dashboard",
    page_icon="📊",
    layout="wide"
)

# =========================================================
# STYLING
# =========================================================
st.markdown(
    """
    <style>
    .stApp {
        background: linear-gradient(180deg, #07111f 0%, #0b1728 100%);
        color: #e8eef8;
    }
    .block-container {
        padding-top: 3rem;
        padding-bottom: 2rem;
        max-width: 1500px;
    }
    .dashboard-title {
        font-size: 2rem;
        font-weight: 700;
        color: #f4f7fb;
        margin-bottom: 0.15rem;
    }
    .dashboard-subtitle {
        color: #aebcd0;
        margin-bottom: 1.2rem;
    }
    .metric-card {
        background: #12233b;
        border: 1px solid #213753;
        border-radius: 18px;
        padding: 18px 20px;
        box-shadow: 0 8px 24px rgba(0,0,0,0.22);
    }
    .metric-label {
        color: #9fb3c8;
        font-size: 0.95rem;
        margin-bottom: 0.35rem;
    }
    .metric-value {
        color: #ffffff;
        font-size: 2rem;
        font-weight: 700;
        line-height: 1.1;
    }
    div[data-testid="stDataFrame"] {
        border-radius: 14px;
        overflow: hidden;
    }
    div[data-baseweb="select"] > div,
    div[data-testid="stTextInput"] input {
        background-color: #0f1d31 !important;
        color: #e8eef8 !important;
        border-radius: 10px !important;
    }
    .login-wrap {
        text-align: center;
        margin-top: 22vh;
    }
    .login-title {
        font-size: 2rem;
        font-weight: 700;
        color: white;
        margin-bottom: 0.3rem;
    }
    .login-subtitle {
        color: #9fb3c8;
        margin-bottom: 1rem;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# =========================================================
# OPTIONAL PASSWORD PROTECTION
# =========================================================
def check_password() -> bool:
    app_password = st.secrets.get("APP_PASSWORD", "")
    if not app_password:
        return True
    if "password_ok" not in st.session_state:
        st.session_state.password_ok = False
    if st.session_state.password_ok:
        return True
    c1, c2, c3 = st.columns([1, 1.2, 1])
    with c2:
        st.markdown('<div class="login-wrap">', unsafe_allow_html=True)
        st.markdown('<div class="login-title">🔐 Customer Tracking Dashboard</div>', unsafe_allow_html=True)
        st.markdown('<div class="login-subtitle">Authorized Access Only</div>', unsafe_allow_html=True)
        pwd = st.text_input(
            "Enter Password",
            type="password",
            label_visibility="collapsed",
            placeholder="Enter password"
        )
        if pwd:
            if pwd == app_password:
                st.session_state.password_ok = True
                st.rerun()
            else:
                st.error("Incorrect password")
        st.markdown("</div>", unsafe_allow_html=True)
    return False


if not check_password():
    st.stop()

# =========================================================
# CONFIG — column name candidates (order = priority)
# =========================================================
FILE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Customer Contract and MRC Tracking.xlsx")

CODE_CANDIDATES      = ["Customer Code", "CustomerCode", "Cust Code", "Code", "Customer ID", "CustomerID"]
NAME_CANDIDATES      = ["Customer Name", "Customer", "Name", "Account Name", "Client Name", "Company"]
STATUS_CANDIDATES    = ["Status", "Customer Status", "Contract Status"]
AM_CANDIDATES        = ["Account Manager", "AM", "Owner", "Sales Rep"]
TIER_CANDIDATES      = ["Customer Category", "Category", "Tier", "Service Tier"]
MRR_CANDIDATES       = ["MRR", "MRC", "Monthly Recurring Revenue"]
IT_MRC_CANDIDATES    = [
    "Current IT Services MRC",
    "Current IT Services MR",
    "Current IT-Services MRC",
    "IT Services MRC",
    "IT Services MR",
    "IT-Services MRC",
    "IT MRC",
    "IT MR",
    "Current MRC",
    "MRC",
]
EXP_CANDIDATES       = ["Contract Expiration", "Contract Expiry", "Expiration", "Renewal Date", "Contract End"]
NEXT_REVIEW_CANDIDATES = ["Next Business Review", "Next Review", "Next QBR"]

# =========================================================
# HELPERS
# =========================================================
def canonical(value: str) -> str:
    """Lowercase alphanumeric only — used for fuzzy column/sheet matching."""
    return "".join(ch.lower() for ch in str(value).strip() if ch.isalnum())


def find_col(df: pd.DataFrame, candidates: list) -> str | None:
    """Return the first matching column name from candidates list, or None."""
    canonical_map = {canonical(c): c for c in df.columns}
    for candidate in candidates:
        key = canonical(candidate)
        if key in canonical_map:
            return canonical_map[key]
    return None


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df.dropna(how="all")
    df.columns = [str(c).replace("\n", " ").replace("\r", " ").strip() for c in df.columns]
    return df


def safe_str(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.strip()


def to_numeric(s: pd.Series) -> pd.Series:
    return pd.to_numeric(
        s.astype(str)
         .str.replace("$", "", regex=False)
         .str.replace(",", "", regex=False)
         .str.replace(" ", "", regex=False)
         .str.replace("\u00A0", "", regex=False)
         .str.strip(),
        errors="coerce"
    )


def to_dt(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce")


def fmt_currency(v) -> str:
    try:
        if pd.isna(v):
            return "$0.00"
        return f"${float(v):,.2f}"
    except Exception:
        return str(v)


def fmt_value(v) -> str:
    if pd.isna(v):
        return ""
    if isinstance(v, str) and "month" in v.lower():
        return "Month-to-Month"
    try:
        return pd.to_datetime(v, errors="raise").strftime("%b %d, %Y")
    except Exception:
        return str(v)


def format_contract_cell(val):
    if pd.isna(val) or val == "":
        return ""
    if isinstance(val, str) and "month" in str(val).lower():
        return "Month-to-Month"
    try:
        return pd.to_datetime(val).strftime("%b %d, %Y")
    except Exception:
        return val


def format_currency_cell(val):
    if pd.isna(val) or val == "":
        return ""
    try:
        return "${:,.2f}".format(float(str(val).replace("$", "").replace(",", "")))
    except Exception:
        return val


def card(label: str, value: str):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def section_open(title: str, subtitle: str = ""):
    st.markdown(f"### {title}")
    if subtitle:
        st.caption(subtitle)


def section_close():
    pass


# =========================================================
# WORKBOOK LOADING
# =========================================================
def detect_header_row(path: str, sheet: str, max_scan: int = 10) -> int:
    """
    Scan the first max_scan rows and return the index of the row that looks
    most like a header — i.e. has the most non-null string values.
    Handles sheets where rows 1-N are legend/title/blank rows before the
    real column headers (e.g. your MRC sheet with headers on row 5).
    """
    raw = pd.read_excel(path, sheet_name=sheet, header=None, nrows=max_scan)
    best_row = 0
    best_score = -1
    for i, row in raw.iterrows():
        score = int(row.dropna().apply(lambda v: isinstance(v, str)).sum())
        if score > best_score:
            best_score = score
            best_row = i
    return int(best_row)


@st.cache_data
def load_workbook(path: str) -> dict[str, pd.DataFrame]:
    """
    Load every sheet, auto-detecting the true header row for each sheet.
    This handles workbooks where rows 1-N are legend/title rows before the
    real column headers (like the MRC Contracted Rate sheet with headers on row 5).
    """
    xls = pd.ExcelFile(path)
    result = {}
    for sheet in xls.sheet_names:
        header_row = detect_header_row(path, sheet)
        df = pd.read_excel(path, sheet_name=sheet, header=header_row)
        result[sheet] = normalize_df(df)
    return result


# =========================================================
# MRC SHEET LOOKUP  (fuzzy name match — fixes cross-sheet bug)
# =========================================================
def get_mrc_sheet(sheets: dict[str, pd.DataFrame]) -> tuple[str | None, pd.DataFrame]:
    """
    Find the MRC contracted-rate sheet using fuzzy name matching.
    Matches any sheet whose name contains 'mrc' OR 'contracted'.
    """
    # Exact preferred name first
    for sheet_name, df in sheets.items():
        if sheet_name.strip().lower() == "mrc contracted rate":
            return sheet_name, df

    # Fuzzy fallback
    for sheet_name, df in sheets.items():
        norm = sheet_name.strip().lower()
        if "mrc" in norm or "contracted" in norm:
            return sheet_name, df

    return None, pd.DataFrame()


# =========================================================
# IT SERVICES MRC HELPERS  (use find_col consistently)
# =========================================================
def get_customer_mrc_record(
    sheets: dict[str, pd.DataFrame],
    customer_code: str = "",
    customer_name: str = ""
) -> pd.DataFrame:
    _, target_df = get_mrc_sheet(sheets)
    if target_df.empty:
        return pd.DataFrame()

    mrc_code_col = find_col(target_df, CODE_CANDIDATES)
    mrc_name_col = find_col(target_df, NAME_CANDIDATES)
    match = pd.DataFrame()

    if mrc_code_col and customer_code:
        match = target_df[safe_str(target_df[mrc_code_col]) == str(customer_code).strip()]

    if match.empty and mrc_name_col and customer_name:
        match = target_df[
            safe_str(target_df[mrc_name_col]).str.lower() == str(customer_name).strip().lower()
        ]

    return match


def find_it_mrc_col(df: pd.DataFrame) -> str | None:
    """
    First tries the explicit candidates list.
    Falls back to any column whose canonical name contains both 'it' and 'mrc'
    OR both 'it' and 'services', to survive any naming variation in the workbook.
    """
    col = find_col(df, IT_MRC_CANDIDATES)
    if col:
        return col
    for c in df.columns:
        norm = canonical(c)
        if "it" in norm and ("mrc" in norm or "services" in norm):
            return c
    return None


def get_it_services_value_for_customer(
    sheets: dict[str, pd.DataFrame],
    customer_code: str = "",
    customer_name: str = ""
) -> float:
    mrc_match = get_customer_mrc_record(
        sheets=sheets,
        customer_code=customer_code,
        customer_name=customer_name
    )
    if mrc_match.empty:
        return 0.0

    mrc_it_col = find_it_mrc_col(mrc_match)
    if not mrc_it_col:
        return 0.0

    return float(to_numeric(pd.Series([mrc_match.iloc[0][mrc_it_col]])).fillna(0).iloc[0])


def get_total_it_services_mrc_for_filtered(
    sheets: dict[str, pd.DataFrame],
    filtered_df: pd.DataFrame,
    code_col: str | None,
    name_col: str | None
) -> float:
    _, mrc_df = get_mrc_sheet(sheets)
    if mrc_df.empty:
        return 0.0

    mrc_code_col = find_col(mrc_df, CODE_CANDIDATES)
    mrc_name_col = find_col(mrc_df, NAME_CANDIDATES)

    # Use smart finder that falls back to keyword scan if candidates list misses
    mrc_it_col = find_it_mrc_col(mrc_df)
    if not mrc_it_col:
        return 0.0

    matched_mrc = pd.DataFrame()

    # Match by customer code first
    if code_col and mrc_code_col and code_col in filtered_df.columns:
        visible_codes = set(safe_str(filtered_df[code_col]))
        matched_mrc = mrc_df[safe_str(mrc_df[mrc_code_col]).isin(visible_codes)].copy()

    # Fallback: match by customer name
    if matched_mrc.empty and name_col and mrc_name_col and name_col in filtered_df.columns:
        visible_names = set(safe_str(filtered_df[name_col]).str.lower())
        matched_mrc = mrc_df[safe_str(mrc_df[mrc_name_col]).str.lower().isin(visible_names)].copy()

    if matched_mrc.empty:
        return 0.0

    return float(to_numeric(matched_mrc[mrc_it_col]).fillna(0).sum())


def add_it_services_to_display_df(
    display_df: pd.DataFrame,
    sheets: dict[str, pd.DataFrame],
    code_col: str | None,
    name_col: str | None
) -> pd.DataFrame:
    result_df = display_df.copy()
    values = []
    for _, row in result_df.iterrows():
        code_val = row.get(code_col, "") if code_col else ""
        name_val = row.get(name_col, "") if name_col else ""
        it_val = get_it_services_value_for_customer(
            sheets=sheets,
            customer_code=str(code_val),
            customer_name=str(name_val)
        )
        values.append(it_val)
    result_df["Current IT Services MRC"] = values
    return result_df


# =========================================================
# CROSS-SHEET RELATED ROWS
# =========================================================
def get_related_rows(
    sheets: dict[str, pd.DataFrame],
    customer_code: str,
    customer_name: str
) -> dict[str, pd.DataFrame]:
    related = {}
    for sheet_name, df in sheets.items():
        code_col_local = find_col(df, CODE_CANDIDATES)
        name_col_local = find_col(df, NAME_CANDIDATES)
        matches = pd.DataFrame()

        if code_col_local and customer_code:
            matches = df[safe_str(df[code_col_local]) == str(customer_code).strip()]

        if matches.empty and name_col_local and customer_name:
            matches = df[
                safe_str(df[name_col_local]).str.lower() == str(customer_name).strip().lower()
            ]

        if not matches.empty:
            related[sheet_name] = matches

    return related


# =========================================================
# FILTER UI
# =========================================================
def filter_customer_df(df: pd.DataFrame, key_prefix: str = "main") -> pd.DataFrame:
    code_col_local   = find_col(df, CODE_CANDIDATES)
    name_col_local   = find_col(df, NAME_CANDIDATES)
    status_col_local = find_col(df, STATUS_CANDIDATES)
    am_col_local     = find_col(df, AM_CANDIDATES)
    tier_col_local   = find_col(df, TIER_CANDIDATES)

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        search = st.text_input(
            "Search",
            placeholder="Code, customer, AM...",
            key=f"{key_prefix}_search"
        )

    with c3:
        am_sel = []
        if am_col_local:
            opts = sorted([x for x in safe_str(df[am_col_local]).unique() if x])
            am_sel = st.multiselect("Account Manager", opts, key=f"{key_prefix}_am")

    with c4:
        tier_sel = []
        if tier_col_local:
            opts = sorted([x for x in safe_str(df[tier_col_local]).unique() if x])
            tier_sel = st.multiselect("Tier / Category", opts, key=f"{key_prefix}_tier")

    filtered_local = df.copy()

    if am_col_local and am_sel:
        filtered_local = filtered_local[safe_str(filtered_local[am_col_local]).isin(am_sel)]

    if tier_col_local and tier_sel:
        filtered_local = filtered_local[safe_str(filtered_local[tier_col_local]).isin(tier_sel)]

    if search:
        search_cols = [c for c in [code_col_local, name_col_local, am_col_local, status_col_local, tier_col_local] if c]
        if not search_cols:
            search_cols = filtered_local.columns.tolist()
        mask = pd.Series(False, index=filtered_local.index)
        for col in search_cols:
            mask = mask | safe_str(filtered_local[col]).str.contains(search, case=False, na=False)
        filtered_local = filtered_local[mask]

    return filtered_local


# =========================================================
# LOAD DATA
# =========================================================
if not os.path.exists(FILE_PATH):
    st.error(f"❌ Excel file not found at: `{FILE_PATH}`\n\nMake sure **Customer Contract and MRC Tracking.xlsx** is in the same folder as this script.")
    st.stop()

sheets = load_workbook(FILE_PATH)

# Debug sidebar
with st.sidebar:
    st.markdown("### 🔧 Debug Info")
    st.write("**Loaded sheets:**", list(sheets.keys()))
    mrc_sheet_name, mrc_debug_df = get_mrc_sheet(sheets)
    st.write("**MRC sheet detected:**", mrc_sheet_name or "❌ NOT FOUND")

    if not mrc_debug_df.empty:
        st.write("**MRC sheet columns:**")
        for c in mrc_debug_df.columns:
            st.code(repr(c))

        detected_it_col = find_it_mrc_col(mrc_debug_df)
        st.write("**IT MRC column matched:**", detected_it_col or "❌ NOT FOUND")

        mrc_code_col_debug = find_col(mrc_debug_df, CODE_CANDIDATES)
        mrc_name_col_debug = find_col(mrc_debug_df, NAME_CANDIDATES)
        st.write("**MRC code col:**", mrc_code_col_debug or "❌ NOT FOUND")
        st.write("**MRC name col:**", mrc_name_col_debug or "❌ NOT FOUND")

        st.markdown("**Sample MRC codes (first 5):**")
        if mrc_code_col_debug:
            for v in mrc_debug_df[mrc_code_col_debug].dropna().astype(str).str.strip().head(5):
                st.code(repr(v))
        else:
            st.warning("No code column found in MRC sheet")

        st.markdown("**Sample customer codes (first 5):**")
        if code_col:
            for v in customer_df[code_col].dropna().astype(str).str.strip().head(5):
                st.code(repr(v))

        if detected_it_col:
            st.markdown("**Sample IT MRC values (first 5):**")
            for v in mrc_debug_df[detected_it_col].dropna().head(5):
                st.code(repr(v))

    if st.button("Clear cache"):
        st.cache_data.clear()
        st.rerun()


# Resolve customer status sheet (case-insensitive)
customer_sheet_name = None
for name in sheets:
    if name.strip().lower() == "customer status":
        customer_sheet_name = name
        break
if not customer_sheet_name:
    customer_sheet_name = list(sheets.keys())[0]

customer_df = sheets[customer_sheet_name]

code_col        = find_col(customer_df, CODE_CANDIDATES)
name_col        = find_col(customer_df, NAME_CANDIDATES)
status_col      = find_col(customer_df, STATUS_CANDIDATES)
am_col          = find_col(customer_df, AM_CANDIDATES)
tier_col        = find_col(customer_df, TIER_CANDIDATES)
mrr_col         = find_col(customer_df, MRR_CANDIDATES)
exp_col         = find_col(customer_df, EXP_CANDIDATES)
next_review_col = find_col(customer_df, NEXT_REVIEW_CANDIDATES)

# =========================================================
# HEADER
# =========================================================
st.markdown('<div class="dashboard-title">Customer Tracking Dashboard</div>', unsafe_allow_html=True)
st.markdown(
    f'<div class="dashboard-subtitle">Workbook source: <strong>{customer_sheet_name}</strong> &nbsp;|&nbsp; '
    f'MRC sheet: <strong>{get_mrc_sheet(sheets)[0] or "not found"}</strong></div>',
    unsafe_allow_html=True
)

# =========================================================
# TABS
# =========================================================
tabs = st.tabs(["Dashboard", "Customer Discovery"])

# =========================================================
# DASHBOARD TAB
# =========================================================
with tabs[0]:
    filtered = filter_customer_df(customer_df, key_prefix="dashboard")

    total_customers = filtered[code_col].nunique() if code_col else len(filtered)
    total_mrr       = to_numeric(filtered[mrr_col]).fillna(0).sum() if mrr_col else 0
    total_it_mrc    = get_total_it_services_mrc_for_filtered(
        sheets=sheets,
        filtered_df=filtered,
        code_col=code_col,
        name_col=name_col
    )

    expiring_90 = 0
    if exp_col:
        exp_dates   = to_dt(filtered[exp_col])
        today       = pd.Timestamp.today().normalize()
        expiring_90 = ((exp_dates >= today) & (exp_dates <= today + pd.Timedelta(days=90))).sum()

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        card("Total Customers", f"{int(total_customers):,}")
    with k2:
        card("Total MRR", fmt_currency(total_mrr))
    with k3:
        card("IT Services MRC", fmt_currency(total_it_mrc))
    with k4:
        card("Expiring in 90 Days", f"{int(expiring_90):,}")

    c1, c2 = st.columns([1, 1.4])

    with c1:
        section_open("Customers by Tier", "Current filtered view")
        if tier_col and not filtered.empty:
            tier_counts = (
                safe_str(filtered[tier_col])
                .replace("", pd.NA)
                .dropna()
                .value_counts()
                .reset_index()
            )
            tier_counts.columns = ["Tier", "Count"]
            fig = px.pie(tier_counts, names="Tier", values="Count", hole=0.6)
            fig.update_layout(
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                font_color="#e8eef8",
                margin=dict(l=0, r=0, t=10, b=0),
                showlegend=True
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No tier/category data found.")
        section_close()

    with c2:
        section_open("Renewal / Review Forecast", "Upcoming items")
        forecast_df = filtered.copy()
        cols_for_forecast = [c for c in [code_col, name_col, am_col, exp_col, next_review_col] if c and c in forecast_df.columns]

        if cols_for_forecast:
            forecast_df = forecast_df[cols_for_forecast].copy()
            if exp_col and exp_col in forecast_df.columns:
                forecast_df["_sort_exp"] = to_dt(forecast_df[exp_col])
                forecast_df = forecast_df.sort_values("_sort_exp", ascending=True, na_position="last").drop(columns=["_sort_exp"])
                forecast_df[exp_col] = forecast_df[exp_col].apply(format_contract_cell)
            st.dataframe(forecast_df.head(15), use_container_width=True, hide_index=True)
        else:
            st.info("No renewal/review forecast columns found.")
        section_close()

    b1, b2 = st.columns([1.3, 1])

    with b1:
        section_open("Top Customers by MRR", "Highest monthly recurring revenue")
        if mrr_col and name_col:
            chart_df = filtered.copy()
            chart_df["_mrr_num"] = to_numeric(chart_df[mrr_col]).fillna(0)
            group_col = name_col if name_col else code_col
            top_mrr = (
                chart_df.groupby(group_col, dropna=False)["_mrr_num"]
                .sum()
                .sort_values(ascending=False)
                .head(10)
                .reset_index()
            )
            top_mrr.columns = ["Customer", "MRR"]
            fig = px.bar(top_mrr, x="Customer", y="MRR")
            fig.update_layout(
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                font_color="#e8eef8",
                margin=dict(l=0, r=0, t=10, b=0),
                xaxis_title="",
                yaxis_title=""
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No MRR/customer columns found.")
        section_close()

    with b2:
        section_open("At-Risk / Upcoming Renewals", "Soonest contract expirations")
        risk_cols = [c for c in [code_col, name_col, status_col, tier_col, am_col, mrr_col, exp_col] if c]
        if risk_cols:
            risk_df = filtered[risk_cols].copy()
            if exp_col and exp_col in risk_df.columns:
                risk_df["_exp_dt"] = to_dt(risk_df[exp_col])
                risk_df = risk_df.sort_values("_exp_dt", ascending=True, na_position="last").drop(columns=["_exp_dt"])
                risk_df[exp_col] = risk_df[exp_col].apply(format_contract_cell)
            if mrr_col and mrr_col in risk_df.columns:
                risk_df[mrr_col] = risk_df[mrr_col].apply(format_currency_cell)
            st.dataframe(risk_df.head(15), use_container_width=True, hide_index=True)
        else:
            st.info("No risk/renewal fields found.")
        section_close()

    section_open("Customer Table", "Filtered master customer view")
    preferred_cols = [c for c in [code_col, name_col, tier_col, status_col, am_col, exp_col, mrr_col] if c]
    display_df = filtered[preferred_cols].copy() if preferred_cols else filtered.copy()
    display_df = add_it_services_to_display_df(
        display_df=display_df,
        sheets=sheets,
        code_col=code_col,
        name_col=name_col
    )
    if exp_col and exp_col in display_df.columns:
        display_df[exp_col] = display_df[exp_col].apply(format_contract_cell)
    if mrr_col and mrr_col in display_df.columns:
        display_df[mrr_col] = display_df[mrr_col].apply(format_currency_cell)
    if "Current IT Services MRC" in display_df.columns:
        display_df["Current IT Services MRC"] = display_df["Current IT Services MRC"].apply(format_currency_cell)
    st.dataframe(display_df, use_container_width=True, hide_index=True)
    section_close()

# =========================================================
# CUSTOMER DISCOVERY TAB
# =========================================================
with tabs[1]:
    st.subheader("Customer Discovery")
    st.caption("Select a customer to see detail and related sheet records")

    selected_code = ""
    selected_name = ""

    d1, d2 = st.columns(2)

    with d1:
        if code_col:
            code_options = sorted(customer_df[code_col].dropna().astype(str).unique().tolist())
            selected_code = st.selectbox(
                "Select customer code",
                [""] + code_options,
                key="drilldown_code"
            )

    with d2:
        if name_col:
            name_options = sorted(customer_df[name_col].dropna().astype(str).unique().tolist())
            selected_name = st.selectbox(
                "Or select customer name",
                [""] + name_options,
                key="drilldown_name"
            )
            # Resolve code from name selection
            if selected_name and code_col and not selected_code:
                match = customer_df[safe_str(customer_df[name_col]) == selected_name]
                if not match.empty:
                    selected_code = str(match.iloc[0][code_col]).strip()

    if selected_code and code_col:
        main_row = customer_df[safe_str(customer_df[code_col]) == selected_code]

        if not main_row.empty:
            record = main_row.iloc[0]
            selected_customer_code = str(record.get(code_col, "")).strip() if code_col else ""
            selected_customer_name = str(record.get(name_col, "")).strip() if name_col else ""

            it_services_value = get_it_services_value_for_customer(
                sheets=sheets,
                customer_code=selected_customer_code,
                customer_name=selected_customer_name
            )

            r1, r2, r3 = st.columns(3, gap="medium")
            with r1:
                card("Customer Code", fmt_value(record.get(code_col, "")))
            with r2:
                card("Customer Name", fmt_value(record.get(name_col, "")))
            with r3:
                card("Tier / Category", fmt_value(record.get(tier_col, "")))

            st.markdown("<div style='height:16px;'></div>", unsafe_allow_html=True)

            r5, r6, r7, r8 = st.columns(4, gap="medium")
            with r5:
                card("Account Manager", fmt_value(record.get(am_col, "")))
            with r6:
                card("Contract Expiration", fmt_value(record.get(exp_col, "")))
            with r7:
                value = to_numeric(pd.Series([record.get(mrr_col, None)])).fillna(0).iloc[0] if mrr_col else 0
                card("MRR", fmt_currency(value))
            with r8:
                card("IT Services MRC", fmt_currency(it_services_value))

            st.markdown("#### Full Customer Status Record")
            main_row_display = main_row.copy()
            if exp_col and exp_col in main_row_display.columns:
                main_row_display[exp_col] = main_row_display[exp_col].apply(format_contract_cell)
            if mrr_col and mrr_col in main_row_display.columns:
                main_row_display[mrr_col] = main_row_display[mrr_col].apply(format_currency_cell)
            st.dataframe(main_row_display, use_container_width=True, hide_index=True)

            # ---- Cross-sheet related records ----
            related = get_related_rows(
                sheets=sheets,
                customer_code=selected_customer_code,
                customer_name=selected_customer_name
            )
            st.markdown("#### Related Records Across Sheets")

            if not related:
                st.info("No related records found in other sheets for this customer.")
            else:
                for sheet_name, rel_df in related.items():
                    rel_display = rel_df.copy()
                    rel_exp_col    = find_col(rel_display, EXP_CANDIDATES)
                    rel_mrr_col    = find_col(rel_display, MRR_CANDIDATES)
                    rel_it_mrc_col = find_col(rel_display, IT_MRC_CANDIDATES)

                    if rel_exp_col and rel_exp_col in rel_display.columns:
                        rel_display[rel_exp_col] = rel_display[rel_exp_col].apply(format_contract_cell)
                    if rel_mrr_col and rel_mrr_col in rel_display.columns:
                        rel_display[rel_mrr_col] = rel_display[rel_mrr_col].apply(format_currency_cell)
                    if rel_it_mrc_col and rel_it_mrc_col in rel_display.columns:
                        rel_display[rel_it_mrc_col] = rel_display[rel_it_mrc_col].apply(format_currency_cell)

                    if sheet_name == customer_sheet_name:
                        st.markdown(f"### {sheet_name}")
                        st.dataframe(rel_display, use_container_width=True, hide_index=True)
                    else:
                        with st.expander(f"{sheet_name} ({len(rel_display)} row(s))"):
                            st.dataframe(rel_display, use_container_width=True, hide_index=True)
        else:
            st.warning(f"No record found for code: `{selected_code}`")
