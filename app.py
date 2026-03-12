import streamlit as st
import pandas as pd
from datetime import datetime

# -------------------------------------------------
# PAGE CONFIG
# -------------------------------------------------
st.set_page_config(
    page_title="Customer Contract Dashboard",
    page_icon="📊",
    layout="wide"
)

# -------------------------------------------------
# PASSWORD PROTECTION
# -------------------------------------------------
def check_password():
    st.markdown(
        """
        <style>
        .center-box {
            text-align: center;
            margin-top: 20vh;
        }

        div[data-testid="stTextInput"] {
            max-width: 350px;
            margin: 0 auto;
        }

        .customer-hero {
            padding: 1.5rem 1rem 0.5rem 1rem;
            border: 1px solid #d9d9d9;
            border-radius: 12px;
            background: #fafafa;
            margin-bottom: 1rem;
        }

        .customer-code {
            font-size: 3.2rem;
            font-weight: 700;
            line-height: 1;
            margin-bottom: 0.25rem;
        }

        .customer-name {
            font-size: 2rem;
            font-weight: 600;
            margin-bottom: 1rem;
        }

        .info-card {
            border: 1px solid #d9d9d9;
            border-radius: 12px;
            padding: 1rem;
            background: white;
            min-height: 260px;
        }

        .label {
            font-weight: 700;
        }

        .detail-row {
            margin-bottom: 0.55rem;
            font-size: 1.02rem;
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    if "password_correct" not in st.session_state:
        st.session_state.password_correct = False

    if st.session_state.password_correct:
        return True

    col1, col2, col3 = st.columns([1, 2, 1])

    with col2:
        st.markdown('<div class="center-box">', unsafe_allow_html=True)
        st.markdown("## 🔐 Enter Password")
        st.caption("Authorized Access Only")

        password = st.text_input("", type="password", placeholder="Enter password")

        if password:
            if password == st.secrets["APP_PASSWORD"]:
                st.session_state.password_correct = True
                st.rerun()
            else:
                st.error("Incorrect password")

        st.markdown("</div>", unsafe_allow_html=True)

    return False


if not check_password():
    st.stop()

# -------------------------------------------------
# FILE PATH
# -------------------------------------------------
FILE_PATH = "Customer Contract and MRC Tracking (1).xlsx"

# -------------------------------------------------
# COLUMN CANDIDATES
# -------------------------------------------------
POSSIBLE_CODE_COLUMNS = [
    "Customer Code", "CustomerCode", "Cust Code", "Code", "Customer ID", "CustomerID", "Customer"
]

POSSIBLE_NAME_COLUMNS = [
    "Customer Name", "Customer", "Name", "Client Name", "Account Name", "Company Name"
]

POSSIBLE_STATUS_COLUMNS = [
    "Status", "Customer Status", "Contract Status"
]

POSSIBLE_AM_COLUMNS = [
    "AM", "Account Manager", "Owner", "Sales Rep"
]

POSSIBLE_CATEGORY_COLUMNS = [
    "Customer Category", "Category", "Tier", "Service Tier"
]

POSSIBLE_CONTRACT_EXP_COLUMNS = [
    "Contract Expiration", "Contract Expiry", "Expiration", "Renewal Date", "Contract End"
]

POSSIBLE_MRR_COLUMNS = [
    "MRR", "Monthly Recurring Revenue", "MRC"
]

POSSIBLE_CURRENT_IT_MRC_COLUMNS = [
    "Current IT-Services MRC", "Current IT Services MRC", "IT Services MRC", "Current MRC"
]

POSSIBLE_QBR_GENERATED_COLUMNS = [
    "QBR Generated", "QBR Date"
]

POSSIBLE_SIGNED_OFF_COLUMNS = [
    "Signed off by C/U", "Signed Off", "Signed Off By Customer"
]

POSSIBLE_LAST_BUSINESS_REVIEW_COLUMNS = [
    "Last Business Review", "Last Review", "Last QBR"
]

POSSIBLE_NEXT_BUSINESS_REVIEW_COLUMNS = [
    "Next Business Review", "Next Review", "Next QBR"
]

POSSIBLE_PRECHECK_COLUMNS = [
    "Pre/Checking Meeting", "Pre-Checking Meeting", "Pre Meeting", "Checking Meeting"
]

POSSIBLE_FISCAL_YEAR_COLUMNS = [
    "Fiscal Year", "FY"
]

# -------------------------------------------------
# HELPERS
# -------------------------------------------------
def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df.dropna(how="all")
    df.columns = [str(col).strip() for col in df.columns]
    return df


def find_first_matching_column(df: pd.DataFrame, candidates: list[str]):
    for col in candidates:
        if col in df.columns:
            return col

    lower_map = {str(col).strip().lower(): col for col in df.columns}
    for col in candidates:
        if col.lower() in lower_map:
            return lower_map[col.lower()]

    return None


def safe_str_series(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip()


def format_value(val):
    if pd.isna(val):
        return ""

    if isinstance(val, (pd.Timestamp, datetime)):
        return pd.to_datetime(val).strftime("%m/%d/%Y")

    return str(val)


def format_currency(val):
    if pd.isna(val) or val == "":
        return ""
    try:
        return "${:,.2f}".format(float(val))
    except Exception:
        return str(val)


def get_value(record, key):
    return record.get(key, "") if key else ""


@st.cache_data
def load_workbook(path: str):
    xls = pd.ExcelFile(path)
    sheets = {}
    for sheet in xls.sheet_names:
        sheets[sheet] = normalize_df(pd.read_excel(path, sheet_name=sheet))
    return sheets


def get_customer_related_rows(sheets, customer_code: str):
    related = {}
    for sheet_name, df in sheets.items():
        code_col = find_first_matching_column(df, POSSIBLE_CODE_COLUMNS)
        if code_col:
            matches = df[safe_str_series(df[code_col]) == str(customer_code).strip()]
            if not matches.empty:
                related[sheet_name] = matches
    return related


def filter_dataframe_ui(df: pd.DataFrame) -> pd.DataFrame:
    filtered_df = df.copy()

    st.markdown("### Filters")
    c1, c2, c3, c4 = st.columns(4)

    search_term = ""
    selected_statuses = []
    selected_ams = []
    selected_categories = []

    status_col = find_first_matching_column(df, POSSIBLE_STATUS_COLUMNS)
    am_col = find_first_matching_column(df, POSSIBLE_AM_COLUMNS)
    category_col = find_first_matching_column(df, POSSIBLE_CATEGORY_COLUMNS)

    with c1:
        search_term = st.text_input("Search", placeholder="Customer name, code, manager...")

    with c2:
        if status_col:
            statuses = sorted([x for x in safe_str_series(df[status_col]).unique() if x])
            selected_statuses = st.multiselect("Status", statuses)

    with c3:
        if am_col:
            ams = sorted([x for x in safe_str_series(df[am_col]).unique() if x])
            selected_ams = st.multiselect("Account Manager", ams)

    with c4:
        if category_col:
            categories = sorted([x for x in safe_str_series(df[category_col]).unique() if x])
            selected_categories = st.multiselect("Category / Tier", categories)

    if status_col and selected_statuses:
        filtered_df = filtered_df[safe_str_series(filtered_df[status_col]).isin(selected_statuses)]

    if am_col and selected_ams:
        filtered_df = filtered_df[safe_str_series(filtered_df[am_col]).isin(selected_ams)]

    if category_col and selected_categories:
        filtered_df = filtered_df[safe_str_series(filtered_df[category_col]).isin(selected_categories)]

    if search_term:
        mask = pd.Series(False, index=filtered_df.index)
        for col in filtered_df.columns:
            mask = mask | safe_str_series(filtered_df[col]).str.contains(search_term, case=False, na=False)
        filtered_df = filtered_df[mask]

    return filtered_df


def render_customer_profile(customer_df: pd.DataFrame, sheets: dict, customer_code: str):
    if customer_df.empty:
        st.warning("No customer record found.")
        return

    record = customer_df.iloc[0]

    code_col = find_first_matching_column(customer_df, POSSIBLE_CODE_COLUMNS)
    name_col = find_first_matching_column(customer_df, POSSIBLE_NAME_COLUMNS)
    status_col = find_first_matching_column(customer_df, POSSIBLE_STATUS_COLUMNS)
    am_col = find_first_matching_column(customer_df, POSSIBLE_AM_COLUMNS)
    category_col = find_first_matching_column(customer_df, POSSIBLE_CATEGORY_COLUMNS)
    contract_exp_col = find_first_matching_column(customer_df, POSSIBLE_CONTRACT_EXP_COLUMNS)
    mrr_col = find_first_matching_column(customer_df, POSSIBLE_MRR_COLUMNS)
    current_it_mrc_col = find_first_matching_column(customer_df, POSSIBLE_CURRENT_IT_MRC_COLUMNS)
    qbr_generated_col = find_first_matching_column(customer_df, POSSIBLE_QBR_GENERATED_COLUMNS)
    signed_off_col = find_first_matching_column(customer_df, POSSIBLE_SIGNED_OFF_COLUMNS)
    last_review_col = find_first_matching_column(customer_df, POSSIBLE_LAST_BUSINESS_REVIEW_COLUMNS)
    next_review_col = find_first_matching_column(customer_df, POSSIBLE_NEXT_BUSINESS_REVIEW_COLUMNS)
    precheck_col = find_first_matching_column(customer_df, POSSIBLE_PRECHECK_COLUMNS)
    fiscal_year_col = find_first_matching_column(customer_df, POSSIBLE_FISCAL_YEAR_COLUMNS)

    customer_code_val = format_value(get_value(record, code_col)) or customer_code
    customer_name_val = format_value(get_value(record, name_col))

    st.markdown('<div class="customer-hero">', unsafe_allow_html=True)
    st.markdown(f'<div class="customer-code">{customer_code_val}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="customer-name">{customer_name_val}</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    left, right = st.columns([1, 1])

    with left:
        st.markdown('<div class="info-card">', unsafe_allow_html=True)
        st.markdown(
            f"""
            <div class="detail-row"><span class="label">Customer Category:</span> {format_value(get_value(record, category_col))}</div>
            <div class="detail-row"><span class="label">Account Manager:</span> {format_value(get_value(record, am_col))}</div>
            <div class="detail-row"><span class="label">Customer Status:</span> {format_value(get_value(record, status_col))}</div>
            <div class="detail-row"><span class="label">Contract Expiration:</span> {format_value(get_value(record, contract_exp_col))}</div>
            <div class="detail-row"><span class="label">MRR:</span> {format_currency(get_value(record, mrr_col))}</div>
            <div class="detail-row"><span class="label">Current IT-Services MRC:</span> {format_currency(get_value(record, current_it_mrc_col))}</div>
            """,
            unsafe_allow_html=True
        )
        st.markdown('</div>', unsafe_allow_html=True)

    with right:
        st.markdown('<div class="info-card">', unsafe_allow_html=True)
        st.markdown(
            f"""
            <div class="detail-row"><span class="label">Fiscal Year:</span> {format_value(get_value(record, fiscal_year_col))}</div>
            <div class="detail-row"><span class="label">QBR Generated:</span> {format_value(get_value(record, qbr_generated_col))}</div>
            <div class="detail-row"><span class="label">Signed off by C/U:</span> {format_value(get_value(record, signed_off_col))}</div>
            <div class="detail-row"><span class="label">Last Business Review:</span> {format_value(get_value(record, last_review_col))}</div>
            <div class="detail-row"><span class="label">Next Business Review:</span> {format_value(get_value(record, next_review_col))}</div>
            <div class="detail-row"><span class="label">Pre/Checking Meeting:</span> {format_value(get_value(record, precheck_col))}</div>
            """,
            unsafe_allow_html=True
        )
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("### Full Customer Record")
    st.dataframe(customer_df, use_container_width=True, hide_index=True)

    related = get_customer_related_rows(sheets, customer_code)

    st.markdown("### Related Sheet Data")
    for sheet_name, related_df in related.items():
        with st.expander(f"{sheet_name} ({len(related_df)} record(s))", expanded=(sheet_name == "Customer Status")):
            st.dataframe(related_df, use_container_width=True, hide_index=True)


# -------------------------------------------------
# LOAD DATA
# -------------------------------------------------
sheets = load_workbook(FILE_PATH)

sheet_names = list(sheets.keys())
ordered_tabs = ["Customer Status"] if "Customer Status" in sheets else []
ordered_tabs += [name for name in sheet_names if name != "Customer Status"]

# -------------------------------------------------
# HEADER
# -------------------------------------------------
st.title("📊 Customer Contract Dashboard")
st.caption("Customer contract and MRC tracking")

# -------------------------------------------------
# TABS
# -------------------------------------------------
tabs = st.tabs(ordered_tabs)

for i, tab_name in enumerate(ordered_tabs):
    with tabs[i]:
        df = sheets[tab_name]

        if tab_name == "Customer Status":
            st.subheader("Customer Status")

            filtered_df = filter_dataframe_ui(df)

            code_col = find_first_matching_column(filtered_df, POSSIBLE_CODE_COLUMNS)
            name_col = find_first_matching_column(filtered_df, POSSIBLE_NAME_COLUMNS)

            k1, k2, k3 = st.columns(3)
            k1.metric("Visible Records", len(filtered_df))
            k2.metric("Total Records", len(df))
            k3.metric("Unique Customers", filtered_df[code_col].nunique() if code_col else len(filtered_df))

            st.markdown("### Customer List")

            display_cols = filtered_df.columns.tolist()
            event = st.dataframe(
                filtered_df,
                use_container_width=True,
                hide_index=True,
                on_select="rerun",
                selection_mode="single-row"
            )

            selected_customer_code = None

            try:
                selected_rows = event.selection.rows
            except Exception:
                selected_rows = []

            if selected_rows and code_col:
                selected_idx = selected_rows[0]
                selected_customer_code = str(filtered_df.iloc[selected_idx][code_col]).strip()

            st.markdown("### Open Customer Profile")
            selector_col1, selector_col2 = st.columns([2, 3])

            with selector_col1:
                if code_col:
                    code_options = sorted(filtered_df[code_col].dropna().astype(str).unique().tolist())
                    manual_code = st.selectbox("Select customer code", [""] + code_options)
                    if manual_code:
                        selected_customer_code = manual_code

            with selector_col2:
                if name_col:
                    name_options = sorted(filtered_df[name_col].dropna().astype(str).unique().tolist())
                    manual_name = st.selectbox("Or select customer name", [""] + name_options)
                    if manual_name and code_col:
                        matched = filtered_df[safe_str_series(filtered_df[name_col]) == manual_name]
                        if not matched.empty:
                            selected_customer_code = str(matched.iloc[0][code_col]).strip()

            if selected_customer_code:
                st.divider()
                customer_main = df[
                    safe_str_series(df[find_first_matching_column(df, POSSIBLE_CODE_COLUMNS)]) == str(selected_customer_code).strip()
                ]
                render_customer_profile(customer_main, sheets, selected_customer_code)

            st.download_button(
                label="Download filtered Customer Status CSV",
                data=filtered_df.to_csv(index=False).encode("utf-8"),
                file_name="customer_status_filtered.csv",
                mime="text/csv",
                key="download_customer_status_filtered"
            )

        else:
            st.subheader(tab_name)

            search_term = st.text_input(
                f"Search in {tab_name}",
                placeholder="Type to filter this sheet...",
                key=f"search_{tab_name}"
            )

            filtered_df = df.copy()
            if search_term:
                mask = pd.Series(False, index=filtered_df.index)
                for col in filtered_df.columns:
                    mask = mask | safe_str_series(filtered_df[col]).str.contains(search_term, case=False, na=False)
                filtered_df = filtered_df[mask]

            st.dataframe(filtered_df, use_container_width=True, hide_index=True)

            code_col = find_first_matching_column(filtered_df, POSSIBLE_CODE_COLUMNS)
            if code_col:
                options = sorted(filtered_df[code_col].dropna().astype(str).unique().tolist())
                selected_code = st.selectbox(
                    f"View customer details from {tab_name}",
                    [""] + options,
                    key=f"select_{tab_name}"
                )

                if selected_code:
                    related_rows = filtered_df[safe_str_series(filtered_df[code_col]) == selected_code]
                    st.markdown(f"### Records for Customer Code: `{selected_code}`")
                    st.dataframe(related_rows, use_container_width=True, hide_index=True)

            st.download_button(
                label=f"Download {tab_name} CSV",
                data=filtered_df.to_csv(index=False).encode("utf-8"),
                file_name=f"{tab_name}.csv",
                mime="text/csv",
                key=f"download_{tab_name}"
            )
