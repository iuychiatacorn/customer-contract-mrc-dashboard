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
EXP_CANDIDATES         = ["Contract Expiration", "Contract Expiry", "Expiration", "Renewal Date", "Contract End"]
NEXT_REVIEW_CANDIDATES = ["Next Business Review", "Next Review", "Next QBR"]
CHECKIN_CANDIDATES     = ["Pre/Check-in meetings?", "Pre/Check-in meetings", "Pre Check-in meetings", "Check-in meetings", "Pre/Checkin"]
SMARTSHEET_CANDIDATES  = ["Smartsheet", "Smart Sheet", "SmartSheet Link", "Smartsheet Link"]
BOOLEAN_CANDIDATES     = CHECKIN_CANDIDATES + ["Signed off by C/U", "Signed off by C U", "Signed off", "Sign off"]
QBR_GEN_CANDIDATES     = ["QBR vCIO Generated", "QBR Generated", "QBR vCIO", "QBR Date"]

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
        score = int(sum(isinstance(v, str) for v in row.dropna()))
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
            raw_opts = [x for x in safe_str(df[tier_col_local]).unique() if x]
            opts = sorted(raw_opts, key=lambda t: (int(x) if (x := ''.join(filter(str.isdigit, t))) else 999, t))
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
            # Sort Tier 1 → Tier 2 → Tier 3
            tier_counts["_sort"] = tier_counts["Tier"].str.extract(r"(\d+)").astype(float)
            tier_counts = tier_counts.sort_values("_sort").drop(columns=["_sort"])
            fig = px.pie(tier_counts, names="Tier", values="Count", hole=0.6,
                         category_orders={"Tier": tier_counts["Tier"].tolist()})
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

    # ── Search bar ──────────────────────────────────────────
    st.markdown("""
    <style>
    .profile-hero {
        background: linear-gradient(135deg, #0d1f38 0%, #112240 60%, #0a1628 100%);
        border: 1px solid #1e3a5f;
        border-radius: 24px;
        padding: 32px 36px;
        margin-bottom: 24px;
        position: relative;
        overflow: hidden;
    }
    .profile-hero::before {
        content: '';
        position: absolute;
        top: -60px; right: -60px;
        width: 220px; height: 220px;
        background: radial-gradient(circle, rgba(56,139,253,0.12) 0%, transparent 70%);
        border-radius: 50%;
    }
    .profile-avatar {
        width: 72px; height: 72px;
        border-radius: 18px;
        background: linear-gradient(135deg, #1c4f8a, #2d7dd2);
        display: flex; align-items: center; justify-content: center;
        font-size: 2rem; font-weight: 800;
        color: #fff;
        margin-bottom: 16px;
        box-shadow: 0 8px 24px rgba(45,125,210,0.35);
        letter-spacing: -1px;
    }
    .profile-name {
        font-size: 1.75rem;
        font-weight: 800;
        color: #f0f6ff;
        margin: 0 0 4px 0;
        letter-spacing: -0.5px;
    }
    .profile-code-badge {
        display: inline-block;
        background: rgba(56,139,253,0.15);
        border: 1px solid rgba(56,139,253,0.3);
        color: #58a6ff;
        font-size: 0.78rem;
        font-weight: 600;
        letter-spacing: 1.5px;
        text-transform: uppercase;
        padding: 3px 10px;
        border-radius: 20px;
        margin-right: 8px;
    }
    .profile-tier-badge {
        display: inline-block;
        font-size: 0.78rem;
        font-weight: 600;
        letter-spacing: 1px;
        text-transform: uppercase;
        padding: 3px 10px;
        border-radius: 20px;
    }
    .tier-1 { background: rgba(255,215,0,0.12); border: 1px solid rgba(255,215,0,0.35); color: #ffd700; }
    .tier-2 { background: rgba(192,192,192,0.12); border: 1px solid rgba(192,192,192,0.35); color: #c0c0c0; }
    .tier-3 { background: rgba(205,127,50,0.12); border: 1px solid rgba(205,127,50,0.35); color: #cd7f32; }
    .tier-other { background: rgba(100,160,255,0.12); border: 1px solid rgba(100,160,255,0.3); color: #64a0ff; }

    .stat-block {
        background: rgba(255,255,255,0.04);
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 16px;
        padding: 20px 22px;
        height: 100%;
    }
    .stat-block-label {
        font-size: 0.75rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 1.2px;
        color: #6b8aad;
        margin-bottom: 8px;
    }
    .stat-block-value {
        font-size: 1.5rem;
        font-weight: 800;
        color: #e8f0fe;
        line-height: 1.1;
    }
    .stat-block-sub {
        font-size: 0.8rem;
        color: #4a6fa5;
        margin-top: 4px;
    }
    .stat-accent-green { color: #3fb950; }
    .stat-accent-yellow { color: #e3b341; }
    .stat-accent-blue { color: #58a6ff; }

    .info-row {
        display: flex;
        align-items: center;
        padding: 14px 0;
        border-bottom: 1px solid rgba(255,255,255,0.05);
    }
    .info-row:last-child { border-bottom: none; }
    .info-row-icon {
        font-size: 1.1rem;
        width: 32px;
        flex-shrink: 0;
    }
    .info-row-label {
        font-size: 0.78rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: #4a6fa5;
        width: 160px;
        flex-shrink: 0;
    }
    .info-row-value {
        font-size: 0.95rem;
        color: #c9d8ec;
        font-weight: 500;
    }
    .section-panel {
        background: #0d1f38;
        border: 1px solid #1e3a5f;
        border-radius: 18px;
        padding: 24px 28px;
        margin-bottom: 20px;
    }
    .section-panel-title {
        font-size: 0.7rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 2px;
        color: #3d6494;
        margin-bottom: 18px;
        padding-bottom: 12px;
        border-bottom: 1px solid #1a3457;
    }
    .expiry-urgent { color: #f85149; font-weight: 700; }
    .expiry-soon   { color: #e3b341; font-weight: 600; }
    .expiry-ok     { color: #3fb950; }
    .no-selection-state {
        text-align: center;
        padding: 80px 20px;
        color: #2d4a6e;
    }
    .no-selection-state .big-icon { font-size: 4rem; margin-bottom: 16px; }
    .no-selection-state h3 { color: #3d6494; font-size: 1.2rem; font-weight: 600; margin: 0; }
    </style>
    """, unsafe_allow_html=True)

    # ── Selectors ───────────────────────────────────────────
    sel_c1, sel_c2 = st.columns(2)
    selected_code = ""
    selected_name = ""

    with sel_c1:
        if code_col:
            code_options = sorted(customer_df[code_col].dropna().astype(str).unique().tolist())
            selected_code = st.selectbox("Search by Customer Code", [""] + code_options, key="drilldown_code")

    with sel_c2:
        if name_col:
            name_options = sorted(customer_df[name_col].dropna().astype(str).unique().tolist())
            selected_name = st.selectbox("Search by Customer Name", [""] + name_options, key="drilldown_name")
            if selected_name and code_col and not selected_code:
                match = customer_df[safe_str(customer_df[name_col]) == selected_name]
                if not match.empty:
                    selected_code = str(match.iloc[0][code_col]).strip()

    # ── Profile ─────────────────────────────────────────────
    if not selected_code or not code_col:
        st.markdown("""
        <div class="no-selection-state">
            <div class="big-icon">🔍</div>
            <h3>Select a customer above to view their profile</h3>
        </div>
        """, unsafe_allow_html=True)
    else:
        main_row = customer_df[safe_str(customer_df[code_col]) == selected_code]

        if main_row.empty:
            st.warning(f"No record found for code: `{selected_code}`")
        else:
            record = main_row.iloc[0]
            cust_code   = str(record.get(code_col, "")).strip() if code_col else ""
            cust_name   = str(record.get(name_col, "")).strip() if name_col else ""
            cust_tier   = str(record.get(tier_col, "")).strip() if tier_col else ""
            cust_am     = str(record.get(am_col, "")).strip() if am_col else ""
            cust_status = str(record.get(status_col, "")).strip() if status_col else ""
            cust_mrr    = to_numeric(pd.Series([record.get(mrr_col, None)])).fillna(0).iloc[0] if mrr_col else 0
            cust_exp    = record.get(exp_col, None) if exp_col else None
            it_mrc      = get_it_services_value_for_customer(sheets=sheets, customer_code=cust_code, customer_name=cust_name)

            # Avatar initials
            initials = "".join(w[0].upper() for w in cust_name.split()[:2]) if cust_name else "??"

            # Tier badge class
            tier_num = ''.join(filter(str.isdigit, cust_tier))
            tier_class = f"tier-{tier_num}" if tier_num in ["1","2","3"] else "tier-other"
            tier_label = cust_tier if cust_tier else "—"

            # Expiry color
            exp_display = format_contract_cell(cust_exp)
            exp_class = ""
            if cust_exp and exp_display != "Month-to-Month":
                try:
                    days_left = (pd.to_datetime(cust_exp) - pd.Timestamp.today()).days
                    if days_left < 0:
                        exp_class = "expiry-urgent"
                        exp_display = f"{exp_display} (Expired)"
                    elif days_left <= 90:
                        exp_class = "expiry-soon"
                        exp_display = f"{exp_display} ({days_left}d left)"
                    else:
                        exp_class = "expiry-ok"
                except Exception:
                    pass

            # ── Hero card ──────────────────────────────────
            st.markdown(f"""
            <div class="profile-hero">
                <div class="profile-avatar">{initials}</div>
                <div class="profile-name">{cust_name or cust_code}</div>
                <div style="margin-top:10px;">
                    <span class="profile-code-badge">{cust_code}</span>
                    <span class="profile-tier-badge {tier_class}">{tier_label}</span>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # ── KPI row ────────────────────────────────────
            k1, k2, k3, k4 = st.columns(4)
            with k1:
                st.markdown(f"""
                <div class="stat-block">
                    <div class="stat-block-label">💰 Monthly Revenue</div>
                    <div class="stat-block-value stat-accent-green">{fmt_currency(cust_mrr)}</div>
                    <div class="stat-block-sub">MRR</div>
                </div>""", unsafe_allow_html=True)
            with k2:
                st.markdown(f"""
                <div class="stat-block">
                    <div class="stat-block-label">🖥 IT Services MRC</div>
                    <div class="stat-block-value stat-accent-blue">{fmt_currency(it_mrc)}</div>
                    <div class="stat-block-sub">Contracted rate</div>
                </div>""", unsafe_allow_html=True)
            with k3:
                st.markdown(f"""
                <div class="stat-block">
                    <div class="stat-block-label">📋 Contract Expiry</div>
                    <div class="stat-block-value {exp_class}" style="font-size:1.1rem; padding-top:4px;">{exp_display or "—"}</div>
                    <div class="stat-block-sub">Renewal date</div>
                </div>""", unsafe_allow_html=True)
            with k4:
                st.markdown(f"""
                <div class="stat-block">
                    <div class="stat-block-label">👤 Account Manager</div>
                    <div class="stat-block-value" style="font-size:1.05rem; padding-top:4px;">{cust_am or "—"}</div>
                    <div class="stat-block-sub">Owner</div>
                </div>""", unsafe_allow_html=True)

            st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

            # ── Two-column detail + related ────────────────
            left_col, right_col = st.columns([1, 1.6], gap="large")

            with left_col:
                # Identify special columns
                checkin_col    = find_col(customer_df, CHECKIN_CANDIDATES)
                smartsheet_col = find_col(customer_df, SMARTSHEET_CANDIDATES)
                signoff_col    = find_col(customer_df, ["Signed off by C/U", "Signed off by C U", "Signed off", "Sign off"])
                qbr_gen_col    = find_col(customer_df, QBR_GEN_CANDIDATES)
                special_cols   = {checkin_col, smartsheet_col, signoff_col, qbr_gen_col} - {None}

                EXCLUDE_COLS = {c for c in customer_df.columns if any(
                    kw in c.lower() for kw in ["seat", "seats", "license", "qty", "quantity"]
                )}
                shown_cols = {code_col, name_col, tier_col, am_col, mrr_col, exp_col, status_col} | special_cols | EXCLUDE_COLS

                # Helper: render TRUE/FALSE as styled badge
                def bool_badge(val) -> str:
                    s = str(val).strip().lower()
                    if s in ("true", "yes", "1", "1.0"):
                        return '<span style="background:rgba(63,185,80,0.15);border:1px solid rgba(63,185,80,0.4);color:#3fb950;font-size:0.8rem;font-weight:700;padding:3px 12px;border-radius:20px;">✓ Yes</span>'
                    elif s in ("false", "no", "0", "0.0"):
                        return '<span style="background:rgba(248,81,73,0.12);border:1px solid rgba(248,81,73,0.35);color:#f85149;font-size:0.8rem;font-weight:700;padding:3px 12px;border-radius:20px;">✗ No</span>'
                    return f'<span style="color:#6b8aad;">{val}</span>'

                # Helper: render a URL as a styled link button
                def link_badge(url) -> str:
                    s = str(url).strip()
                    if s.lower().startswith("http"):
                        return f'<a href="{s}" target="_blank" style="display:inline-block;background:rgba(56,139,253,0.12);border:1px solid rgba(56,139,253,0.35);color:#58a6ff;font-size:0.8rem;font-weight:600;padding:4px 14px;border-radius:20px;text-decoration:none;">🔗 Open Smartsheet</a>'
                    # Not a URL yet — show placeholder
                    return f'<span style="color:#4a6fa5;font-style:italic;font-size:0.85rem;">No link stored</span>'

                # Build regular info rows
                icon_map = {
                    "next": "📅", "review": "📅", "qbr": "📅",
                    "note": "📝", "address": "📍", "phone": "📞",
                    "email": "✉️", "website": "🌐", "industry": "🏭",
                    "proposed": "💡", "alignment": "🎯", "strategy": "🗺",
                    "budget": "📊", "mrc": "💲", "mrr": "💲",
                }
                extra_rows = []
                for col in customer_df.columns:
                    if col in shown_cols:
                        continue
                    val = record.get(col, None)
                    if pd.isna(val) or str(val).strip() in ("", "nan", "NaT"):
                        continue
                    try:
                        # Skip plain numbers — they'd parse as Unix epoch (Jan 1 1970)
                        if not isinstance(val, (int, float)):
                            parsed_dt = pd.to_datetime(val, errors="raise")
                            val = parsed_dt.strftime("%b %d, %Y")
                        else:
                            raise ValueError("numeric, skip date parse")
                    except Exception:
                        try:
                            fval = float(str(val).replace("$","").replace(",",""))
                            val = f"${fval:,.2f}"
                        except Exception:
                            val = str(val).strip()
                    icon = "•"
                    for kw, ic in icon_map.items():
                        if kw in col.lower():
                            icon = ic
                            break
                    extra_rows.append((icon, col, val))

                st.markdown('<div class="section-panel">', unsafe_allow_html=True)
                st.markdown('<div class="section-panel-title">Customer Details</div>', unsafe_allow_html=True)

                # Status row
                if cust_status:
                    st.markdown(f"""
                    <div class="info-row">
                        <div class="info-row-icon">🔵</div>
                        <div class="info-row-label">Status</div>
                        <div class="info-row-value">{cust_status}</div>
                    </div>""", unsafe_allow_html=True)

                # Pre/Check-in meetings row
                if checkin_col:
                    checkin_val = record.get(checkin_col, None)
                    if not (pd.isna(checkin_val) if not isinstance(checkin_val, str) else False):
                        st.markdown(f"""
                        <div class="info-row">
                            <div class="info-row-icon">📋</div>
                            <div class="info-row-label">Pre/Check-in Meeting</div>
                            <div class="info-row-value">{bool_badge(checkin_val)}</div>
                        </div>""", unsafe_allow_html=True)

                # Signed off by C/U row
                if signoff_col:
                    signoff_val = record.get(signoff_col, None)
                    if not (pd.isna(signoff_val) if not isinstance(signoff_val, str) else False):
                        st.markdown(f"""
                        <div class="info-row">
                            <div class="info-row-icon">✅</div>
                            <div class="info-row-label">Signed off by C/U</div>
                            <div class="info-row-value">{bool_badge(signoff_val)}</div>
                        </div>""", unsafe_allow_html=True)

                # QBR vCIO Generated row
                if qbr_gen_col:
                    qbr_val = record.get(qbr_gen_col, None)
                    qbr_is_blank = pd.isna(qbr_val) if not isinstance(qbr_val, str) else str(qbr_val).strip() in ("", "nan", "NaT")
                    if qbr_is_blank:
                        qbr_display = '<span style="color:#4a6fa5;font-style:italic;font-size:0.85rem;">No QBR Generated</span>'
                    else:
                        try:
                            qbr_display = f'<span style="background:rgba(63,185,80,0.15);border:1px solid rgba(63,185,80,0.4);color:#3fb950;font-size:0.8rem;font-weight:700;padding:3px 12px;border-radius:20px;">📅 {pd.to_datetime(qbr_val).strftime("%b %d, %Y")}</span>'
                        except Exception:
                            qbr_display = f'<span style="color:#c9d8ec;">{str(qbr_val).strip()}</span>'
                    st.markdown(f"""
                    <div class="info-row">
                        <div class="info-row-icon">📑</div>
                        <div class="info-row-label">QBR vCIO Generated</div>
                        <div class="info-row-value">{qbr_display}</div>
                    </div>""", unsafe_allow_html=True)

                # Smartsheet row
                if smartsheet_col:
                    ss_val = record.get(smartsheet_col, None)
                    if not (pd.isna(ss_val) if not isinstance(ss_val, str) else False):
                        st.markdown(f"""
                        <div class="info-row">
                            <div class="info-row-icon">📊</div>
                            <div class="info-row-label">Smartsheet</div>
                            <div class="info-row-value">{link_badge(ss_val)}</div>
                        </div>""", unsafe_allow_html=True)

                # Remaining regular rows
                for icon, label, val in extra_rows:
                    st.markdown(f"""
                    <div class="info-row">
                        <div class="info-row-icon">{icon}</div>
                        <div class="info-row-label">{label}</div>
                        <div class="info-row-value">{val}</div>
                    </div>""", unsafe_allow_html=True)

                st.markdown('</div>', unsafe_allow_html=True)

            with right_col:
                # Related records across sheets
                related = get_related_rows(sheets=sheets, customer_code=cust_code, customer_name=cust_name)

                st.markdown('<div class="section-panel-title" style="font-size:0.7rem;font-weight:700;text-transform:uppercase;letter-spacing:2px;color:#3d6494;margin-bottom:16px;">Records Across Sheets</div>', unsafe_allow_html=True)

                if not related:
                    st.info("No related records found in other sheets.")
                else:
                    for sheet_name, rel_df in related.items():
                        if sheet_name == customer_sheet_name:
                            continue   # already shown in hero/details
                        rel_display = rel_df.copy()
                        rel_exp_col    = find_col(rel_display, EXP_CANDIDATES)
                        rel_mrr_col    = find_col(rel_display, MRR_CANDIDATES)
                        rel_it_mrc_col = find_col(rel_display, IT_MRC_CANDIDATES)
                        if rel_exp_col:
                            rel_display[rel_exp_col] = rel_display[rel_exp_col].apply(format_contract_cell)
                        if rel_mrr_col:
                            rel_display[rel_mrr_col] = rel_display[rel_mrr_col].apply(format_currency_cell)
                        if rel_it_mrc_col:
                            rel_display[rel_it_mrc_col] = rel_display[rel_it_mrc_col].apply(format_currency_cell)
                        with st.expander(f"📄 {sheet_name}  ({len(rel_display)} row(s))", expanded=True):
                            st.dataframe(rel_display, use_container_width=True, hide_index=True)
